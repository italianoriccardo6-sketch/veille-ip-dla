import fs from "node:fs/promises";
import { AlignmentType, Document, ExternalHyperlink, Footer, HeadingLevel, Packer, PageNumber, Paragraph, TextRun } from "docx";

const sources = JSON.parse(await fs.readFile(new URL("../data/sources.json", import.meta.url), "utf8"));
if (!process.env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY manquante");
const domains = sources.map((source) => source.domain);

const requiredText = { type: "string", minLength: 20 };
const itemProperties = {
  type: { type: "string", enum: ["JURISPRUDENCE", "ACTUALITE"] },
  category: { type: "string", minLength: 3 },
  title: { type: "string", minLength: 10 },
  court_reference: { type: "string", minLength: 3 },
  source: { type: "string", minLength: 3 },
  source_url: { type: "string", minLength: 10 },
  publication_date: { type: "string", minLength: 8 },
  summary: requiredText,
  introduction: requiredText,
  facts_and_procedure: requiredText,
  parties_arguments: requiredText,
  legal_question: requiredText,
  reasoning: requiredText,
  outcome: requiredText,
  practical_relevance: requiredText
};
const schema = {
  type: "object",
  additionalProperties: false,
  properties: {
    week: { type: "string" },
    editorial_note: { type: "string" },
    items: {
      type: "array",
      minItems: 5,
      maxItems: 6,
      items: {
        type: "object",
        additionalProperties: false,
        properties: itemProperties,
        required: Object.keys(itemProperties)
      }
    }
  },
  required: ["week", "editorial_note", "items"]
};

const input = [
  "Tu es le rédacteur d'une veille hebdomadaire destinée à l'équipe Propriété intellectuelle d'un grand cabinet d'avocats international en France.",
  "",
  "MISSION",
  "Recherche en priorité les développements des 7 derniers jours exclusivement dans les domaines autorisés et rédige une veille complète en français, directement exploitable par des avocats. Si la semaine est pauvre en publications, élargis progressivement la recherche aux 30 derniers jours afin de produire une sélection substantielle; indique alors clairement la date de chaque sujet.",
  "",
  "SÉLECTION",
  "- Sélectionne exactement 5 ou 6 sujets substantiels: idéalement 4 jurisprudences et 1 ou 2 actualités. En période de faible activité juridictionnelle, accepte 3 jurisprudences et 2 ou 3 actualités.",
  "- Recherche un équilibre entre marques, brevets, dessins et modèles, droit d'auteur et un sujet connexe pertinent (IA, numérique, médias ou concurrence déloyale).",
  "- Privilégie les décisions, textes, communiqués et dossiers législatifs provenant de sources primaires.",
  "- Ne retiens un sujet que si sa date, sa référence et son URL directe sont vérifiables.",
  "- N'invente jamais un numéro de décision, une juridiction, une citation, un fait, une position de partie ou une étape procédurale.",
  "- Si une information manque dans la source, indique sobrement qu'elle n'est pas précisée au lieu de la compléter.",
  "",
  "STYLE À REPRODUIRE",
  "- Reproduis le style d'une veille juridique française professionnelle: sobre, impersonnel, fluide, précis et approfondi.",
  "- Pour une jurisprudence, rédige environ 650 à 900 mots au total.",
  "- Commence par: « Un litige opposait… » ou une formulation équivalente présentant immédiatement le différend.",
  "- Développe successivement, sous forme de paragraphes continus: les parties et les faits; la procédure; les prétentions et arguments; la question de droit; le raisonnement détaillé de la juridiction; la solution et le dispositif.",
  "- Formule explicitement: « La question de droit posée à la cour était de savoir si… ».",
  "- Termine par un paragraphe de synthèse commençant par « La cour… », « Le tribunal… » ou « En conséquence… ».",
  "- Les champs du JSON servent à structurer le document, mais leur contenu doit être rédigé en paragraphes complets, sans listes à puces.",
  "- Pour une actualité, rédige 250 à 450 mots: contexte, contenu, modifications principales, conséquences pratiques et prochaines étapes.",
  "- Le résumé destiné à la dashboard doit faire 35 à 55 mots. Le document complet ne doit pas être une simple amplification du résumé.",
  "- N'utilise des guillemets que pour des citations réellement présentes dans la source.",
  "",
  "CONTENU DES CHAMPS",
  "- introduction: présentation du litige ou de l'actualité en une phrase dense.",
  "- facts_and_procedure: faits détaillés et historique procédural.",
  "- parties_arguments: prétentions et moyens essentiels des parties; pour une actualité, contexte et positions institutionnelles.",
  "- legal_question: question de droit explicitement formulée; pour une actualité, enjeu juridique central.",
  "- reasoning: motivation détaillée, distinctions et principes appliqués; pour une actualité, contenu précis des mesures.",
  "- outcome: dispositif et conséquences concrètes; pour une actualité, état de la procédure et calendrier.",
  "- practical_relevance: portée professionnelle en un paragraphe bref, sans conseil juridique personnalisé.",
  "",
  "CONTRÔLE QUALITÉ",
  "- Chaque URL doit mener directement au document primaire utilisé, pas à une page d'accueil.",
  "- Privilégie les 7 derniers jours. À défaut de matière suffisante, utilise des publications des 30 derniers jours, sans jamais masquer leur date réelle.",
  "- Ne produis jamais de fiche vide, de chaîne vide ou de sujet fictif pour atteindre le nombre demandé.",
  "- Évite les doublons et les sujets trop faibles.",
  "- L'editorial_note doit signaler honnêtement toute limite de couverture ou d'accès.",
  "",
  "Domaines autorisés: " + domains.join(", ")
].join("\n");

const requestBody = {
  model: process.env.OPENAI_MODEL || "gpt-5.6-sol",
  reasoning: { effort: "high" },
  tools: [{
    type: "web_search",
    filters: { allowed_domains: domains },
    external_web_access: true
  }],
  tool_choice: "required",
  input,
  max_output_tokens: 15000,
  text: { format: { type: "json_schema", name: "veille_ip_editoriale", strict: true, schema } }
};

let response;
let raw;
const maxAttempts = 5;

for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
  response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(requestBody)
  });

  raw = await response.json();
  if (response.ok) break;

  if (response.status !== 429 || attempt === maxAttempts) {
    throw new Error(JSON.stringify(raw));
  }

  const message = raw?.error?.message || "";
  const suggestedSeconds = Number(message.match(/try again in ([0-9.]+)s/i)?.[1] || 0);
  const delayMs = Math.max(Math.ceil(suggestedSeconds * 1000) + 2000, 10000 * (2 ** (attempt - 1)));
  console.warn(`Rate limit OpenAI: nouvel essai ${attempt + 1}/${maxAttempts} dans ${Math.ceil(delayMs / 1000)} s.`);
  await new Promise((resolve) => setTimeout(resolve, delayMs));
}

if (!response?.ok) throw new Error(JSON.stringify(raw));

const output = raw.output?.flatMap((entry) => entry.content || []).find((entry) => entry.type === "output_text")?.text;
if (!output) throw new Error("Sortie structurée absente");

const report = JSON.parse(output);
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
console.log(`Generated Word report with ${jurisprudences.length} jurisprudences and ${actualites.length} actualités for ${report.week}`);
