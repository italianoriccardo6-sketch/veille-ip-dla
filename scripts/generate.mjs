import fs from "node:fs/promises";
import { AlignmentType, Document, ExternalHyperlink, Footer, HeadingLevel, Packer, PageNumber, Paragraph, TextRun } from "docx";

const sources = JSON.parse(await fs.readFile(new URL("../data/sources.json", import.meta.url), "utf8"));
if (!process.env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY manquante");
const domains = sources.map((source) => source.domain);
const sourceCount = sources.length;
const discoveryModel = process.env.OPENAI_DISCOVERY_MODEL || "gpt-5.6-luna";
const editorialModel = process.env.OPENAI_EDITORIAL_MODEL || "gpt-5.6-terra";
const responsesEndpoint = process.env.OPENAI_RESPONSES_ENDPOINT || "https://api.openai.com/v1/responses";
const usageTotals = { input_tokens: 0, output_tokens: 0, total_tokens: 0, requests: 0, web_search_calls: 0 };
const usageByModel = {};

const callOpenAI = async (body, label) => {
  const maxAttempts = 5;
  let response;
  let raw;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    response = await fetch(responsesEndpoint, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });
    raw = await response.json();
    if (response.ok) {
      if (raw.status === "incomplete") {
        throw new Error(`${label}: réponse incomplète (${raw.incomplete_details?.reason || "raison inconnue"})`);
      }
      usageTotals.requests += 1;
      usageTotals.input_tokens += raw.usage?.input_tokens || 0;
      usageTotals.output_tokens += raw.usage?.output_tokens || 0;
      usageTotals.total_tokens += raw.usage?.total_tokens || 0;
      usageTotals.web_search_calls += raw.output?.filter((entry) => entry.type === "web_search_call").length || 0;
      usageByModel[body.model] ||= { input_tokens: 0, output_tokens: 0 };
      usageByModel[body.model].input_tokens += raw.usage?.input_tokens || 0;
      usageByModel[body.model].output_tokens += raw.usage?.output_tokens || 0;
      return raw;
    }
    const errorCode = raw?.error?.code;
    if (response.status !== 429 || errorCode !== "rate_limit_exceeded" || attempt === maxAttempts) {
      throw new Error(`${label}: ${JSON.stringify(raw)}`);
    }
    const message = raw?.error?.message || "";
    const suggestedSeconds = Number(message.match(/try again in ([0-9.]+)s/i)?.[1] || 0);
    const delayMs = Math.max(Math.ceil(suggestedSeconds * 1000) + 2000, 10000 * (2 ** (attempt - 1)));
    console.warn(`${label}: limite OpenAI, nouvel essai ${attempt + 1}/${maxAttempts} dans ${Math.ceil(delayMs / 1000)} s.`);
    await new Promise((resolve) => setTimeout(resolve, delayMs));
  }
  throw new Error(`${label}: échec inattendu`);
};

const extractOutputText = (raw, label) => {
  const text = raw.output?.flatMap((entry) => entry.content || [])
    .find((entry) => entry.type === "output_text")?.text;
  if (!text) throw new Error(`${label}: sortie structurée absente`);
  return text;
};

// Une recherche distincte et obligatoire est exécutée pour chaque source. Cette
// étape empêche le modèle éditorial de se concentrer uniquement sur les domaines
// qui remontent le plus facilement dans une recherche globale.
const discoverySchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    source_name: { type: "string" },
    domain: { type: "string" },
    searched: { type: "boolean" },
    search_note: { type: "string" },
    candidates: {
      type: "array",
      maxItems: 3,
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          title: { type: "string" },
          url: { type: "string" },
          publication_date: { type: "string" },
          content_type: { type: "string" },
          relevance: { type: "string" }
        },
        required: ["title", "url", "publication_date", "content_type", "relevance"]
      }
    }
  },
  required: ["source_name", "domain", "searched", "search_note", "candidates"]
};

const sourceCoverage = [];
for (const source of sources) {
  console.log(`Analyse obligatoire: ${source.name} (${source.domain})`);
  const discoveryRaw = await callOpenAI({
    model: discoveryModel,
    reasoning: { effort: "low" },
    tools: [{
      type: "web_search",
      filters: { allowed_domains: [source.domain] },
      external_web_access: true
    }],
    tool_choice: "required",
    input: [
      `Analyse obligatoirement la source ${source.name} (${source.domain}).`,
      ...(source.reference_url ? [`URL de référence prioritaire: ${source.reference_url}.`] : []),
      `Thèmes attendus: ${source.themes.join(", ")}.`,
      "Recherche d'abord toutes les publications pertinentes des 7 derniers jours, puis élargis aux 30 derniers jours.",
      "Repère jusqu'à trois décisions, textes, rapports ou actualités substantiels en propriété intellectuelle.",
      "Chaque résultat doit avoir une date vérifiable et une URL directe. N'invente rien.",
      "Même si aucun résultat pertinent n'est trouvé, confirme que la source a été analysée, laisse candidates vide et explique brièvement pourquoi dans search_note."
    ].join("\n"),
    max_output_tokens: 1200,
    text: { format: { type: "json_schema", name: "analyse_source_ip", strict: true, schema: discoverySchema } }
  }, `Analyse ${source.name}`);
  const discovery = JSON.parse(extractOutputText(discoveryRaw, `Analyse ${source.name}`));
  discovery.source_name = source.name;
  discovery.domain = source.domain;
  discovery.searched = true;
  sourceCoverage.push(discovery);
}

if (sourceCoverage.length !== sourceCount || sourceCoverage.some((entry) => !entry.searched)) {
  throw new Error("Veille refusée: toutes les sources obligatoires n'ont pas été analysées.");
}

const selectionProperties = {
  type: { type: "string", enum: ["JURISPRUDENCE", "ACTUALITE"] },
  category: { type: "string", minLength: 3 },
  title: { type: "string", minLength: 10 },
  court_reference: { type: "string", minLength: 3 },
  source: { type: "string", minLength: 3 },
  source_url: { type: "string", minLength: 10 },
  publication_date: { type: "string", minLength: 8 }
};
const selectionSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    week: { type: "string", minLength: 10 },
    editorial_note: { type: "string", minLength: 20 },
    selected_items: {
      type: "array",
      minItems: 5,
      maxItems: 6,
      items: {
        type: "object",
        additionalProperties: false,
        properties: selectionProperties,
        required: Object.keys(selectionProperties)
      }
    }
  },
  required: ["week", "editorial_note", "selected_items"]
};

const selectionRaw = await callOpenAI({
  model: discoveryModel,
  reasoning: { effort: "medium" },
  input: [
    "Tu es le secrétaire de rédaction d'une veille française de propriété intellectuelle.",
    `Sélectionne 5 ou 6 sujets parmi les résultats issus des ${sourceCount} sources effectivement contrôlées ci-dessous.`,
    "Privilégie les sources primaires, la date récente, la substance juridique et un équilibre entre marques, brevets, dessins et modèles, droit d'auteur, IA et numérique.",
    "Une newsletter secondaire ne sert qu'à détecter un sujet; préfère l'URL primaire lorsqu'elle figure dans les résultats.",
    "N'invente ni référence ni URL. Écarte les doublons et les sujets insuffisamment vérifiables.",
    JSON.stringify(sourceCoverage)
  ].join("\n"),
  max_output_tokens: 3000,
  text: { format: { type: "json_schema", name: "selection_veille_ip", strict: true, schema: selectionSchema } }
}, "Sélection éditoriale");
const selection = JSON.parse(extractOutputText(selectionRaw, "Sélection éditoriale"));

const requiredText = { type: "string", minLength: 20 };
const itemProperties = {
  ...selectionProperties,
  summary: requiredText,
  introduction: requiredText,
  facts_and_procedure: requiredText,
  parties_arguments: requiredText,
  legal_question: requiredText,
  reasoning: requiredText,
  outcome: requiredText,
  practical_relevance: requiredText
};
const itemSchema = {
  type: "object",
  additionalProperties: false,
  properties: itemProperties,
  required: Object.keys(itemProperties)
};

const items = [];
for (const [index, selected] of selection.selected_items.entries()) {
  console.log(`Rédaction ${index + 1}/${selection.selected_items.length}: ${selected.title}`);
  const itemRaw = await callOpenAI({
    model: editorialModel,
    reasoning: { effort: "medium" },
    tools: [{
      type: "web_search",
      filters: { allowed_domains: domains },
      external_web_access: true
    }],
    tool_choice: "required",
    input: [
      "Tu rédiges une fiche pour la veille Propriété intellectuelle d'un grand cabinet d'avocats international en France.",
      `Sujet sélectionné: ${JSON.stringify(selected)}`,
      "Ouvre et analyse le document primaire. L'URL finale doit mener directement à la décision, au texte ou au document institutionnel utilisé.",
      "Rédige en français juridique, sobre, impersonnel, précis et approfondi, exclusivement à partir d'informations vérifiables.",
      "Pour une jurisprudence, rédige 650 à 900 mots au total: litige, faits, procédure, arguments, question de droit explicitement formulée, raisonnement détaillé, solution et portée pratique.",
      "Pour une actualité, rédige 250 à 450 mots: contexte, contenu précis, conséquences pratiques et prochaines étapes.",
      "Le résumé destiné à la dashboard doit faire 35 à 55 mots. Les autres champs doivent être des paragraphes continus, sans listes.",
      "N'invente jamais une référence, une citation, un argument ou une étape procédurale. Si un élément manque, indique qu'il n'est pas précisé.",
      "N'utilise des guillemets que pour une citation réellement présente dans la source."
    ].join("\n"),
    max_output_tokens: 6000,
    text: { format: { type: "json_schema", name: "fiche_veille_ip", strict: true, schema: itemSchema } }
  }, `Rédaction fiche ${index + 1}`);
  items.push(JSON.parse(extractOutputText(itemRaw, `Rédaction fiche ${index + 1}`)));
}

const report = {
  week: selection.week,
  editorial_note: selection.editorial_note,
  source_coverage: sourceCoverage.map((entry) => ({
    source_name: entry.source_name,
    domain: entry.domain,
    searched: true,
    candidate_count: entry.candidates.length,
    search_note: entry.search_note
  })),
  items
};
const mandatoryFields = ["category", "title", "source", "source_url", "publication_date", "summary", "introduction", "reasoning", "outcome", "practical_relevance"];
const incompleteItems = report.items.filter((item) =>
  mandatoryFields.some((field) => typeof item[field] !== "string" || item[field].trim().length < 3)
);
if (incompleteItems.length || report.items.length < 5) {
  throw new Error(`Veille refusée: ${incompleteItems.length} fiche(s) incomplète(s), ${report.items.length} sujet(s) au total.`);
}
report.generated_at = new Date().toISOString();
report.status = "generated";
const slug = new Date().toISOString().slice(0, 10);
report.report_url = `/public/reports/veille-${slug}.docx`;
report.docx_url = report.report_url;
await fs.mkdir("public/reports", { recursive: true });

const jurisprudences = report.items.filter((item) => item.type === "JURISPRUDENCE");
const actualites = report.items.filter((item) => item.type === "ACTUALITE");
const textParagraphs = (value) => String(value || "")
  .split(/\n\s*\n/)
  .map((paragraph) => paragraph.trim())
  .filter(Boolean)
  .map((text) => new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 180, line: 300 },
    children: [new TextRun({ text })]
  }));

const sourceParagraph = (item) => new Paragraph({
  spacing: { after: 240 },
  children: [
    new ExternalHyperlink({
      link: item.source_url,
      children: [new TextRun({
        text: `${item.court_reference || item.source}, ${item.publication_date}`,
        color: "507D82",
        underline: {}
      })]
    })
  ]
});

const practicalParagraph = (item) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { before: 80, after: 260 },
  border: { left: { color: "507D82", size: 10, space: 8, style: "single" } },
  indent: { left: 180 },
  children: [
    new TextRun({ text: "Portée pratique. ", bold: true }),
    new TextRun({ text: item.practical_relevance })
  ]
});

const renderItem = (item) => [
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    keepNext: true,
    spacing: { before: 180, after: 100 },
    children: [new TextRun({
      text: item.type === "JURISPRUDENCE" ? `${item.category} – ${item.title}` : item.title,
      bold: true,
      underline: {}
    })]
  }),
  sourceParagraph(item),
  ...textParagraphs(item.introduction),
  ...textParagraphs(item.facts_and_procedure),
  ...textParagraphs(item.parties_arguments),
  ...textParagraphs(item.legal_question),
  ...textParagraphs(item.reasoning),
  ...textParagraphs(item.outcome),
  practicalParagraph(item)
];

const children = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 },
    children: [new TextRun({ text: "Veille Propriété intellectuelle", bold: true, size: 32, font: "Arial" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 420 },
    children: [new TextRun({ text: report.week, size: 22, font: "Arial" })]
  })
];

if (jurisprudences.length) {
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 180 },
    border: { bottom: { color: "222222", size: 6, space: 6, style: "single" } },
    children: [new TextRun({ text: "JURISPRUDENCES", bold: true })]
  }));
  jurisprudences.forEach((item) => children.push(...renderItem(item)));
}

if (actualites.length) {
  children.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 180 },
    border: { bottom: { color: "222222", size: 6, space: 6, style: "single" } },
    children: [new TextRun({ text: "ACTUALITÉS", bold: true })]
  }));
  actualites.forEach((item) => children.push(...renderItem(item)));
}

children.push(
  new Paragraph({
    spacing: { before: 260, after: 140 },
    children: [new TextRun({ text: report.editorial_note, color: "666666", size: 18, italics: true })]
  })
);

const document = new Document({
  creator: "Veille IP DLA",
  title: `Veille Propriété intellectuelle – ${report.week}`,
  description: "Veille hebdomadaire de propriété intellectuelle",
  styles: {
    default: {
      document: {
        run: { font: "Arial", size: 22, color: "171717" },
        paragraph: { spacing: { after: 160, line: 300 } }
      }
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Arial", size: 24, bold: true, color: "171717" }
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Arial", size: 22, bold: true, color: "171717" }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1020, right: 1020, bottom: 900, left: 1020 }
      }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({ text: "Veille IP · " }),
            new TextRun({ children: [PageNumber.CURRENT] })
          ]
        })]
      })
    },
    children
  }]
});

const buffer = await Packer.toBuffer(document);
await fs.writeFile(`public/reports/veille-${slug}.docx`, buffer);
await fs.writeFile("public/latest.json", JSON.stringify(report, null, 2) + "\n");
const tokenRates = {
  "gpt-5.6-luna": { input: 0.20, output: 1.20 },
  "gpt-5.6-terra": { input: 2.00, output: 12.00 },
  "gpt-5.6-sol": { input: 4.00, output: 20.00 }
};
const estimatedTokenCost = Object.entries(usageByModel).reduce((total, [model, usage]) => {
  const rate = tokenRates[model];
  if (!rate) return total;
  return total + (usage.input_tokens * rate.input + usage.output_tokens * rate.output) / 1_000_000;
}, 0);
const estimatedSearchCost = usageTotals.web_search_calls * 0.01;
console.log(`Usage: ${usageTotals.requests} requêtes, ${usageTotals.input_tokens} tokens d'entrée, ${usageTotals.output_tokens} tokens de sortie, ${usageTotals.web_search_calls} recherches web.`);
console.log(`Coût API estimé: $${(estimatedTokenCost + estimatedSearchCost).toFixed(3)} (hors éventuelle tarification long contexte/régionale).`);
console.log(`Generated Word report with ${jurisprudences.length} jurisprudences and ${actualites.length} actualités for ${report.week}`);
