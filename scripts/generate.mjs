import fs from "node:fs/promises";
import { AlignmentType, BorderStyle, Document, ExternalHyperlink, Footer, HeadingLevel, Packer, PageNumber, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from "docx";

const sources = JSON.parse(await fs.readFile(new URL("../data/sources.json", import.meta.url), "utf8"));
if (!process.env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY manquante");
const domains = sources.map((source) => source.domain);
const sourceCount = sources.length;
const discoveryModel = process.env.OPENAI_DISCOVERY_MODEL || "gpt-5.6-luna";
const editorialModel = process.env.OPENAI_EDITORIAL_MODEL || "gpt-5.6-terra";
const responsesEndpoint = process.env.OPENAI_RESPONSES_ENDPOINT || "https://api.openai.com/v1/responses";
const usageTotals = { input_tokens: 0, output_tokens: 0, total_tokens: 0, requests: 0, web_search_calls: 0 };
const usageByModel = {};
const runStartedAt = new Date();
const weekStart = new Date(runStartedAt);
const daysSinceMonday = (runStartedAt.getUTCDay() + 6) % 7;
weekStart.setUTCDate(runStartedAt.getUTCDate() - daysSinceMonday);
weekStart.setUTCHours(0, 0, 0, 0);
const weekStartIso = weekStart.toISOString().slice(0, 10);
const weekEndIso = runStartedAt.toISOString().slice(0, 10);
const frenchMonths = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
const frenchDate = (value) => {
  const date = new Date(`${value}T00:00:00.000Z`);
  return `${date.getUTCDate()} ${frenchMonths[date.getUTCMonth()]} ${date.getUTCFullYear()}`;
};
const weekLabel = weekStart.getUTCFullYear() === runStartedAt.getUTCFullYear()
  ? `Du ${weekStart.getUTCDate()} ${frenchMonths[weekStart.getUTCMonth()]} au ${runStartedAt.getUTCDate()} ${frenchMonths[runStartedAt.getUTCMonth()]} ${runStartedAt.getUTCFullYear()}`
  : `Du ${frenchDate(weekStartIso)} au ${frenchDate(weekEndIso)}`;
const currentWeekInstruction = `Ne retiens que les contenus publiés entre le ${weekStartIso} et le ${weekEndIso}, dates incluses. N'utilise aucun contenu antérieur, même pour compléter la sélection.`;
const isCurrentWeekPublication = (value) => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) return false;
  const timestamp = Date.parse(`${value}T00:00:00.000Z`);
  return Number.isFinite(timestamp) && timestamp >= weekStart.getTime() && timestamp <= runStartedAt.getTime();
};

const callOpenAI = async (body, label) => {
  const maxAttempts = 5;
  let response;
  let raw;
  let incompleteRetries = 0;

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
        if (raw.incomplete_details?.reason === "max_output_tokens" && incompleteRetries < 1) {
          incompleteRetries += 1;
          body.max_output_tokens = Math.min((body.max_output_tokens || 4000) * 2, 12000);
          console.warn(`${label}: sortie tronquée, nouvel essai limité à cette étape avec ${body.max_output_tokens} tokens.`);
          continue;
        }
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

const decodeXml = (value = "") => value
  .replace(/&lt;/g, "<")
  .replace(/&gt;/g, ">")
  .replace(/&quot;/g, '"')
  .replace(/&#39;|&apos;/g, "'")
  .replace(/&amp;/g, "&")
  .replace(/<[^>]+>/g, " ")
  .replace(/\s+/g, " ")
  .trim();

const extractTag = (xml, tag) => decodeXml(xml.match(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, "i"))?.[1] || "");

const originalGoogleAlertUrl = (rawHref = "") => {
  const decoded = decodeXml(rawHref);
  try {
    const url = new URL(decoded);
    return url.hostname.includes("google.") && url.pathname === "/url"
      ? url.searchParams.get("url") || decoded
      : decoded;
  } catch {
    return decoded;
  }
};

const ipSignals = [
  /propri[eé]t[eé] intellectuelle/i,
  /droits? d['’]auteur/i,
  /droits? voisins?/i,
  /copyright/i,
  /contrefa[cç]on/i,
  /\bbrevets?\b/i,
  /\bmarques?\b/i,
  /dessins? et mod[eè]les?/i,
  /risque de confusion/i,
  /secret des affaires/i,
  /\b(?:EUIPO|INPI|OEB|EPO|OMPI|WIPO|CSPLA)\b/i,
  /originalit[eé].{0,30}(?:œuvre|oeuvre)/i
];

const googleAlertFeedUrls = [...new Set((process.env.GOOGLE_ALERT_RSS_URLS || "")
  .split(/[\s,;]+/)
  .map((url) => url.trim())
  .filter(Boolean))];

const loadGoogleAlertCandidates = async () => {
  if (!googleAlertFeedUrls.length) {
    console.log("Google Alerts: aucun flux configuré.");
    return [];
  }

  const cutoff = weekStart.getTime();
  const settled = await Promise.all(googleAlertFeedUrls.map(async (feedUrl) => {
    try {
      const response = await fetch(feedUrl, { signal: AbortSignal.timeout(15000) });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const xml = await response.text();
      const alertTitle = extractTag(xml, "title");
      return [...xml.matchAll(/<entry>([\s\S]*?)<\/entry>/gi)].map((match) => {
        const entry = match[1];
        const rawHref = entry.match(/<link[^>]+href="([^"]+)"/i)?.[1] || "";
        const title = extractTag(entry, "title");
        const summary = extractTag(entry, "content") || extractTag(entry, "summary");
        const published = extractTag(entry, "published") || extractTag(entry, "updated");
        const searchable = `${title} ${summary}`;
        const titleScore = ipSignals.filter((signal) => signal.test(title)).length;
        const totalScore = ipSignals.filter((signal) => signal.test(searchable)).length;
        return {
          alert_title: alertTitle,
          title,
          url: originalGoogleAlertUrl(rawHref),
          publication_date: published,
          summary: summary.slice(0, 500),
          relevance_score: titleScore * 2 + totalScore
        };
      });
    } catch (error) {
      console.warn(`Google Alerts: flux ignoré (${error.message})`);
      return [];
    }
  }));

  const deduplicated = new Map();
  for (const candidate of settled.flat()) {
    const timestamp = Date.parse(candidate.publication_date);
    if (!candidate.url || !candidate.title || !Number.isFinite(timestamp) || timestamp < cutoff || timestamp > runStartedAt.getTime() || candidate.relevance_score < 2) continue;
    let key;
    try {
      const url = new URL(candidate.url);
      ["utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term"].forEach((param) => url.searchParams.delete(param));
      key = url.toString();
    } catch {
      key = candidate.title.toLowerCase();
    }
    const existing = deduplicated.get(key);
    if (!existing || candidate.relevance_score > existing.relevance_score) deduplicated.set(key, candidate);
  }

  const candidates = [...deduplicated.values()]
    .sort((a, b) => b.relevance_score - a.relevance_score || Date.parse(b.publication_date) - Date.parse(a.publication_date))
    .slice(0, 30);
  console.log(`Google Alerts: ${googleAlertFeedUrls.length} flux uniques, ${candidates.length} candidats IP après filtrage.`);
  return candidates;
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
          publication_date: { type: "string", pattern: "^\\d{4}-\\d{2}-\\d{2}$" },
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
      currentWeekInstruction,
      "La fraîcheur est impérative. Si la source n'a rien publié pendant cette semaine, laisse candidates vide au lieu de rechercher une publication plus ancienne.",
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

const googleAlertCandidates = (await loadGoogleAlertCandidates()).map((candidate, index) => ({
  ...candidate,
  alert_id: `juliette-${index + 1}`
}));

const selectionProperties = {
  type: { type: "string", enum: ["JURISPRUDENCE", "ACTUALITE"] },
  category: { type: "string", minLength: 3 },
  title: { type: "string", minLength: 10 },
  court_reference: { type: "string", minLength: 3 },
  source: { type: "string", minLength: 3 },
  source_url: { type: "string", minLength: 10 },
  publication_date: { type: "string", pattern: "^\\d{4}-\\d{2}-\\d{2}$" },
  discovered_via_juliette_alert: { type: "boolean" },
  juliette_alert_id: { type: "string" }
};

const frenchEditorialRules = [
  "Rédige systématiquement dans un français soutenu, précis, élégant et juridiquement rigoureux, conforme aux standards rédactionnels d'un cabinet d'avocats.",
  "Respecte irréprochablement la grammaire, l'orthographe, la syntaxe et la ponctuation françaises.",
  "N'utilise jamais les caractères « – » ou « — ». Reformule la phrase ou emploie une ponctuation française appropriée."
].join(" ");

const normalizeFrenchTypography = (value) => String(value || "")
  .replace(/\s+[–—]\s+/g, ", ")
  .replace(/[–—]/g, ",")
  .replace(/\s+,/g, ",")
  .replace(/,{2,}/g, ",")
  .trim();
const selectionSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    week: { type: "string", minLength: 10 },
    editorial_note: { type: "string", minLength: 20 },
    selected_items: {
      type: "array",
      minItems: 6,
      maxItems: 10,
      items: {
        type: "object",
        additionalProperties: false,
        properties: selectionProperties,
        required: Object.keys(selectionProperties)
      }
    },
    alert_decisions: {
      type: "array",
      minItems: googleAlertCandidates.length,
      maxItems: googleAlertCandidates.length,
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          alert_id: { type: "string" },
          selected: { type: "boolean" },
          reason: { type: "string", minLength: 3 }
        },
        required: ["alert_id", "selected", "reason"]
      }
    }
  },
  required: ["week", "editorial_note", "selected_items", "alert_decisions"]
};

const selectionRaw = await callOpenAI({
  model: discoveryModel,
  reasoning: { effort: "medium" },
  input: [
    "Tu es le secrétaire de rédaction d'une veille française de propriété intellectuelle.",
    `Présélectionne entre 6 et 10 sujets parmi les résultats issus des ${sourceCount} sources effectivement contrôlées ci-dessous. Une phase distincte vérifiera ensuite l'accès au document primaire et retiendra les six meilleurs.`,
    "La cible éditoriale est de quatre jurisprudences et deux actualités substantielles. Présélectionne suffisamment de candidats de chaque type pour atteindre cette composition après vérification.",
    "Privilégie les sources primaires, la date récente, la substance juridique et un équilibre réel entre marques, brevets, dessins et modèles, droit d'auteur, IA et numérique. N'annonce jamais un équilibre qui ne ressort pas des sujets effectivement sélectionnés.",
    "Ne retiens pas plus de deux sujets provenant de la même institution. Privilégie la diversité institutionnelle et thématique plutôt que plusieurs décisions proches rendues le même jour.",
    "Une newsletter secondaire ne sert qu'à détecter un sujet; préfère l'URL primaire lorsqu'elle figure dans les résultats.",
    "Les résultats Google Alerts ci-dessous constituent des pistes de veille fournies par Juliette. Pour chaque sujet qui provient effectivement de ces résultats, indique discovered_via_juliette_alert=true et reporte son alert_id exact dans juliette_alert_id. Sinon, utilise false et une chaîne vide.",
    "Pour chacun des résultats Google Alerts, complète alert_decisions avec son identifiant exact, l'indication de sa sélection et une justification éditoriale concise.",
    currentWeekInstruction,
    "La veille doit présenter exclusivement les nouveautés de la semaine en cours. Ne complète jamais la sélection avec un sujet plus ancien.",
    "N'invente ni référence ni URL. Écarte les doublons et les sujets insuffisamment vérifiables.",
    frenchEditorialRules,
    `SOURCES OBLIGATOIRES: ${JSON.stringify(sourceCoverage)}`,
    `GOOGLE ALERTS FILTRÉS: ${JSON.stringify(googleAlertCandidates)}`
  ].join("\n"),
  max_output_tokens: 3000,
  text: { format: { type: "json_schema", name: "selection_veille_ip", strict: true, schema: selectionSchema } }
}, "Sélection éditoriale");
const selection = JSON.parse(extractOutputText(selectionRaw, "Sélection éditoriale"));
const staleSelections = selection.selected_items.filter((item) => !isCurrentWeekPublication(item.publication_date));
if (staleSelections.length) {
  throw new Error(`Veille refusée avant rédaction: ${staleSelections.length} sujet(s) hors de la semaine du ${weekStartIso} au ${weekEndIso}.`);
}

const alertIds = new Set(googleAlertCandidates.map((candidate) => candidate.alert_id));
const invalidAlertOrigins = selection.selected_items.filter((item) =>
  item.discovered_via_juliette_alert
    ? !alertIds.has(item.juliette_alert_id)
    : item.juliette_alert_id !== ""
);
if (invalidAlertOrigins.length) {
  throw new Error(`Veille refusée avant rédaction: ${invalidAlertOrigins.length} attribution(s) Google Alerts incohérente(s).`);
}
const alertDecisionIds = selection.alert_decisions.map((decision) => decision.alert_id);
const invalidAlertDecisions = selection.alert_decisions.filter((decision) =>
  !alertIds.has(decision.alert_id)
  || (decision.selected && !selection.selected_items.some((item) => item.juliette_alert_id === decision.alert_id))
  || (!decision.selected && selection.selected_items.some((item) => item.juliette_alert_id === decision.alert_id))
);
if (new Set(alertDecisionIds).size !== alertIds.size || invalidAlertDecisions.length) {
  throw new Error("Veille refusée avant rédaction: la traçabilité entre les alertes de Juliette et les sujets sélectionnés est incomplète ou contradictoire.");
}

const selectionByInstitution = new Map();
for (const selected of selection.selected_items) {
  let institution = selected.source.toLowerCase();
  try { institution = new URL(selected.source_url).hostname.replace(/^www\./, ""); } catch {}
  selectionByInstitution.set(institution, (selectionByInstitution.get(institution) || 0) + 1);
}
if ([...selectionByInstitution.values()].some((count) => count > 2)) {
  throw new Error("Veille refusée avant rédaction: plus de deux sujets proviennent de la même institution.");
}

const resolverSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    access_level: { type: "string", enum: ["COMPLET", "PARTIEL", "MINIMAL"] },
    full_text_confirmed: { type: "boolean" },
    primary_source_name: { type: "string", minLength: 3 },
    primary_source_url: { type: "string", minLength: 10 },
    publication_date: { type: "string", pattern: "^\\d{4}-\\d{2}-\\d{2}$" },
    document_type: { type: "string", minLength: 3 },
    retrieval_note: { type: "string", minLength: 20 },
    verified_facts: {
      type: "array",
      minItems: 1,
      maxItems: 12,
      items: { type: "string", minLength: 10 }
    }
  },
  required: ["access_level", "full_text_confirmed", "primary_source_name", "primary_source_url", "publication_date", "document_type", "retrieval_note", "verified_facts"]
};

const resolvedSelections = [];
for (const [index, selected] of selection.selected_items.entries()) {
  console.log(`Résolution de la source ${index + 1}/${selection.selected_items.length}: ${selected.title}`);
  const resolutionRaw = await callOpenAI({
    model: discoveryModel,
    reasoning: { effort: "medium" },
    tools: [{ type: "web_search", external_web_access: true }],
    tool_choice: "required",
    input: [
      "Tu vérifies l'accessibilité du document primaire avant toute rédaction juridique.",
      `Sujet: ${JSON.stringify(selected)}`,
      currentWeekInstruction,
      "Recherche successivement l'URL directe, les pièces jointes PDF, les versions linguistiques alternatives, le numéro d'affaire, le numéro de décision et l'ECLI.",
      "Une page de sommaire ou une interface dynamique ne suffit pas si elle contient un lien vers un PDF ou vers une version intégrale. Recherche et ouvre ce document avant de qualifier l'accès.",
      "Utilise COMPLET uniquement si le texte intégral et son dispositif ou ses dispositions ont été effectivement consultés. Utilise PARTIEL si plusieurs éléments substantiels sont vérifiables mais qu'une partie du document manque. Utilise MINIMAL si seuls le titre, la date ou des métadonnées sont accessibles.",
      "primary_source_url doit mener directement au document utilisé, de préférence au PDF ou au texte intégral officiel, jamais à une page d'accueil lorsque le document direct existe.",
      "Énumère uniquement des faits effectivement vérifiés. N'extrapole rien.",
      frenchEditorialRules
    ].join("\n"),
    max_output_tokens: 1800,
    text: { format: { type: "json_schema", name: "resolution_source_primaire", strict: true, schema: resolverSchema } }
  }, `Résolution source ${index + 1}`);
  const resolution = JSON.parse(extractOutputText(resolutionRaw, `Résolution source ${index + 1}`));
  if (resolution.access_level === "COMPLET" && !resolution.full_text_confirmed) {
    resolution.access_level = "PARTIEL";
    resolution.retrieval_note = `${resolution.retrieval_note} Le texte intégral n’a pas été confirmé.`;
  }
  if (!isCurrentWeekPublication(resolution.publication_date)) {
    console.warn(`Sujet écarté, date hors période après résolution: ${selected.title}`);
    continue;
  }
  resolvedSelections.push({ ...selected, resolution });
  const resolvedSubstantive = resolvedSelections.filter(({ resolution: current }) =>
    current.access_level !== "MINIMAL" && current.verified_facts.length >= 4
  );
  if (
    resolvedSubstantive.filter((item) => item.type === "JURISPRUDENCE").length >= 4
    && resolvedSubstantive.filter((item) => item.type === "ACTUALITE").length >= 2
  ) {
    console.log("Composition vérifiée atteinte: arrêt des résolutions supplémentaires afin de limiter le coût API.");
    break;
  }
}

const verifiedSelections = resolvedSelections.filter(({ resolution }) =>
  resolution.access_level !== "MINIMAL" && resolution.verified_facts.length >= 4
);
const verifiedJurisprudences = verifiedSelections.filter((item) => item.type === "JURISPRUDENCE");
const verifiedActualites = verifiedSelections.filter((item) => item.type === "ACTUALITE");
const substantiveSelections = [
  ...verifiedJurisprudences.slice(0, 4),
  ...verifiedActualites.slice(0, 2)
];
const minimalSelections = resolvedSelections.filter(({ resolution }) =>
  resolution.access_level === "MINIMAL" || resolution.verified_facts.length < 4
).slice(0, 2);

if (verifiedJurisprudences.length < 4 || verifiedActualites.length < 2) {
  throw new Error(`Veille refusée avant rédaction approfondie: ${verifiedJurisprudences.length} jurisprudence(s) et ${verifiedActualites.length} actualité(s) suffisamment substantielles. La cible obligatoire est de quatre jurisprudences et deux actualités de la semaine courante.`);
}

const requiredText = { type: "string", minLength: 20 };
const commonItemProperties = {
  ...selectionProperties,
  source_access: { type: "string", enum: ["COMPLET", "PARTIEL"] },
  access_warning: { type: "string" },
  summary: requiredText,
  retrieval_note: requiredText
};
const jurisprudenceProperties = {
  ...commonItemProperties,
  introduction: requiredText,
  facts_and_procedure: requiredText,
  parties_arguments: requiredText,
  legal_question: requiredText,
  reasoning: requiredText,
  outcome: requiredText,
  practical_relevance: requiredText
};
const actualiteProperties = {
  ...commonItemProperties,
  context: requiredText,
  legal_basis_and_scope: requiredText,
  main_provisions: requiredText,
  implementation_timeline: requiredText,
  practical_relevance: requiredText
};
const makeItemSchema = (properties) => ({
  type: "object",
  additionalProperties: false,
  properties,
  required: Object.keys(properties)
});

const finalizeEditorialItem = (draft, selected, resolution, properties) => {
  const item = { ...draft };
  item.type = selected.type;
  item.category = selected.category;
  item.title = selected.title;
  item.court_reference = selected.court_reference;
  item.source = resolution.primary_source_name;
  item.source_url = resolution.primary_source_url;
  item.publication_date = resolution.publication_date;
  item.source_access = resolution.access_level;
  item.retrieval_note = resolution.retrieval_note;
  item.discovered_via_juliette_alert = selected.discovered_via_juliette_alert;
  item.juliette_alert_id = selected.juliette_alert_id;
  for (const field of Object.keys(properties)) {
    if (field !== "source_url" && typeof item[field] === "string") {
      item[field] = normalizeFrenchTypography(item[field]);
    }
  }
  item.access_warning = item.source_access === "PARTIEL"
    ? "Attention : les informations présentées dans cette section doivent être vérifiées, l’accès à la source étant incomplet, limité ou restreint."
    : "";
  return item;
};

const items = [];
for (const [index, resolved] of substantiveSelections.entries()) {
  const { resolution, ...selected } = resolved;
  const isJurisprudence = selected.type === "JURISPRUDENCE";
  const properties = isJurisprudence ? jurisprudenceProperties : actualiteProperties;
  const editorialDomains = [...new Set([
    ...domains,
    (() => {
      try { return new URL(resolution.primary_source_url).hostname; } catch { return ""; }
    })()
  ].filter(Boolean))];
  const accessInstruction = resolution.access_level === "COMPLET"
    ? (isJurisprudence
        ? "Rédige une fiche approfondie de 600 à 850 mots."
        : "Rédige une actualité structurée de 180 à 350 mots.")
    : (isJurisprudence
        ? "La source est partielle mais substantielle. Rédige une analyse resserrée de 450 à 650 mots, limitée aux éléments vérifiés."
        : "La source est partielle mais substantielle. Rédige une actualité resserrée de 180 à 300 mots, limitée aux éléments vérifiés.");
  console.log(`Rédaction ${index + 1}/${substantiveSelections.length}: ${selected.title}`);
  const itemRaw = await callOpenAI({
    model: editorialModel,
    reasoning: { effort: "medium" },
    tools: [{
      type: "web_search",
      filters: { allowed_domains: editorialDomains },
      external_web_access: true
    }],
    tool_choice: "required",
    input: [
      "Tu rédiges une fiche pour la veille Propriété intellectuelle d'un grand cabinet d'avocats international en France.",
      `Sujet sélectionné: ${JSON.stringify(selected)}`,
      `Résolution préalable de la source: ${JSON.stringify(resolution)}`,
      "Fonde la rédaction sur le document primaire résolu et sur les faits vérifiés. L'URL finale doit être exactement primary_source_url et la source exactement primary_source_name.",
      currentWeekInstruction,
      "Rédige en français juridique, sobre, impersonnel, précis et approfondi, exclusivement à partir d'informations vérifiables.",
      frenchEditorialRules,
      accessInstruction,
      isJurisprudence
        ? "Structure la fiche autour du litige, des faits, de la procédure, des arguments, de la question de droit, du raisonnement, de la solution et de la portée pratique."
        : "Structure l'actualité autour du contexte, de son fondement et champ d'application, de ses principales dispositions, de son calendrier de mise en œuvre et de sa portée pratique. N'utilise jamais les rubriques contentieuses lorsqu'elles sont sans objet.",
      "Le résumé destiné à la dashboard doit faire 35 à 55 mots. Les autres champs doivent être des paragraphes continus, sans listes.",
      "Rédige obligatoirement le titre en français, même lorsque le document primaire utilise un titre anglais. Conserve uniquement les noms propres, sigles et marques qui ne doivent pas être traduits.",
      "N'invente jamais une référence, une citation, un argument ou une étape procédurale. Si un élément manque, indique qu'il n'est pas précisé.",
      "N'utilise des guillemets que pour une citation réellement présente dans la source.",
      `Reproduis source_access avec la valeur ${resolution.access_level}.`,
      "Si source_access vaut COMPLET, laisse access_warning vide. S'il vaut PARTIEL, inscris exactement: Attention : les informations présentées dans cette section doivent être vérifiées, l’accès à la source étant incomplet, limité ou restreint.",
      "La portée pratique doit découler directement du contenu vérifié. N'infère jamais un secteur, une solution ou une conséquence à partir du seul titre ou du nom des parties."
    ].join("\n"),
    max_output_tokens: 6000,
    text: { format: { type: "json_schema", name: isJurisprudence ? "fiche_jurisprudence" : "fiche_actualite", strict: true, schema: makeItemSchema(properties) } }
  }, `Rédaction fiche ${index + 1}`);
  const draft = JSON.parse(extractOutputText(itemRaw, `Rédaction fiche ${index + 1}`));
  items.push(finalizeEditorialItem(draft, selected, resolution, properties));
}

const limitWords = (value, maximum) => String(value || "").split(/\s+/).slice(0, maximum).join(" ");
const briefs = minimalSelections.map(({ resolution, ...selected }) => ({
  type: selected.type,
  category: normalizeFrenchTypography(selected.category),
  title: normalizeFrenchTypography(selected.title),
  source: normalizeFrenchTypography(resolution.primary_source_name),
  source_url: resolution.primary_source_url,
  publication_date: resolution.publication_date,
  source_access: "RESTREINT",
  access_warning: "Attention : les informations présentées dans cette section doivent être vérifiées, l’accès à la source étant incomplet, limité ou restreint.",
  discovered_via_juliette_alert: selected.discovered_via_juliette_alert,
  juliette_alert_id: selected.juliette_alert_id,
  summary: normalizeFrenchTypography(limitWords(resolution.verified_facts.join(" "), 180)),
  retrieval_note: normalizeFrenchTypography(resolution.retrieval_note)
}));

const contentWordCount = (item) => {
  const fields = item.type === "JURISPRUDENCE"
    ? ["introduction", "facts_and_procedure", "parties_arguments", "legal_question", "reasoning", "outcome", "practical_relevance"]
    : ["context", "legal_basis_and_scope", "main_provisions", "implementation_timeline", "practical_relevance"];
  return fields.flatMap((field) => String(item[field] || "").split(/\s+/)).filter(Boolean).length;
};
const limitationCount = (item) => (JSON.stringify(item).match(/ne peut|n’est pas possible|ne permettent|ne précise|n’a pas pu/gi) || []).length;
const titleLooksUntranslated = (item) => ((item.title.match(/\b(?:the|for|and|with|from|available|new|now)\b/gi) || []).length >= 2);
const qualityReasons = (item) => {
  const reasons = [];
  const words = contentWordCount(item);
  const limitations = limitationCount(item);
  if (!isCurrentWeekPublication(item.publication_date)) reasons.push(`date hors période: ${item.publication_date}`);
  if (titleLooksUntranslated(item)) reasons.push("titre non traduit intégralement en français");
  if (item.source_access === "COMPLET" && limitations > 2) reasons.push(`${limitations} réserves répétitives pour une source complète, maximum 2`);
  if (item.source_access === "PARTIEL" && limitations > 6) reasons.push(`${limitations} réserves répétitives pour une source partielle, maximum 6`);
  if (item.type === "ACTUALITE" && (words < 180 || words > 350)) reasons.push(`${words} mots, attendu entre 180 et 350`);
  if (item.type === "JURISPRUDENCE") {
    const [minimum, maximum] = item.source_access === "COMPLET" ? [600, 850] : [450, 650];
    if (words < minimum || words > maximum) reasons.push(`${words} mots, attendu entre ${minimum} et ${maximum}`);
  }
  return reasons;
};

let qualityFailures = items
  .map((item, index) => ({ index, title: item.title, reasons: qualityReasons(item) }))
  .filter((failure) => failure.reasons.length);

for (const failure of qualityFailures) {
  const { resolution, ...selected } = substantiveSelections[failure.index];
  const isJurisprudence = selected.type === "JURISPRUDENCE";
  const properties = isJurisprudence ? jurisprudenceProperties : actualiteProperties;
  console.warn(`Contrôle éditorial: « ${failure.title} » refusée: ${failure.reasons.join("; ")}. Correction ciblée unique.`);
  const correctedRaw = await callOpenAI({
    model: editorialModel,
    reasoning: { effort: "medium" },
    input: [
      "Tu corriges une fiche juridique déjà rédigée. Ne recommence aucune recherche et n'ajoute aucun fait nouveau.",
      `Sujet vérifié: ${JSON.stringify(selected)}`,
      `Résolution de la source: ${JSON.stringify(resolution)}`,
      `Fiche à corriger: ${JSON.stringify(items[failure.index])}`,
      `Motifs précis du refus: ${failure.reasons.join("; ")}.`,
      "Corrige uniquement les défauts signalés tout en conservant chaque information vérifiée, la structure prescrite et le niveau de précision juridique.",
      isJurisprudence
        ? "La fiche doit conserver les faits, la procédure, les arguments, la question de droit, le raisonnement, la solution et la portée pratique."
        : "L'actualité doit conserver le contexte, le fondement et le champ d'application, les principales dispositions, le calendrier de mise en œuvre et la portée pratique.",
      "Une réserve factuellement nécessaire peut être conservée, mais ne répète jamais la même limite dans plusieurs rubriques.",
      "Le titre doit être intégralement rédigé en français, à l’exception des noms propres, sigles et marques.",
      frenchEditorialRules
    ].join("\n"),
    max_output_tokens: 6000,
    text: { format: { type: "json_schema", name: isJurisprudence ? "fiche_jurisprudence" : "fiche_actualite", strict: true, schema: makeItemSchema(properties) } }
  }, `Correction ciblée: ${failure.title}`);
  const correctedDraft = JSON.parse(extractOutputText(correctedRaw, `Correction ciblée: ${failure.title}`));
  items[failure.index] = finalizeEditorialItem(correctedDraft, selected, resolution, properties);
}

qualityFailures = items
  .map((item, index) => ({ index, title: item.title, reasons: qualityReasons(item) }))
  .filter((failure) => failure.reasons.length);
if (qualityFailures.length) {
  const details = qualityFailures.map((failure) => `« ${failure.title} »: ${failure.reasons.join("; ")}`).join(" | ");
  throw new Error(`Veille refusée après correction ciblée: ${details}. Aucun document incomplet n’a été publié.`);
}

const jurisprudenceCount = items.filter((item) => item.type === "JURISPRUDENCE").length;
const actualiteCount = items.filter((item) => item.type === "ACTUALITE").length;
const totalEditorialWords = items.reduce((total, item) => total + contentWordCount(item), 0);
const editorialCategories = new Set(items.map((item) => item.category.trim().toLowerCase()));
if (items.length !== 6 || jurisprudenceCount !== 4 || actualiteCount !== 2) {
  throw new Error(`Veille refusée: composition finale de ${jurisprudenceCount} jurisprudence(s) et ${actualiteCount} actualité(s), au lieu de quatre jurisprudences et deux actualités.`);
}
if (totalEditorialWords < 2800 || totalEditorialWords > 4100) {
  throw new Error(`Veille refusée: longueur éditoriale totale de ${totalEditorialWords} mots, attendue entre 2 800 et 4 100 mots.`);
}
if (editorialCategories.size < 4) {
  throw new Error(`Veille refusée: seulement ${editorialCategories.size} catégories éditoriales distinctes, quatre au minimum sont requises.`);
}
const selectedAlertCount = items.filter((item) => item.discovered_via_juliette_alert).length;
const alertEditorialSentence = selectedAlertCount === 0
  ? "Aucun des sujets retenus ne provient des alertes de Juliette."
  : selectedAlertCount === 1
    ? "Un sujet a été initialement signalé par les alertes de Juliette, puis vérifié sur sa source primaire."
    : `${selectedAlertCount} sujets ont été initialement signalés par les alertes de Juliette, puis vérifiés sur leur source primaire.`;
const editorialNote = [
  `Cette édition réunit quatre jurisprudences et deux actualités publiées au cours de la semaine ${weekLabel.toLowerCase()}.`,
  `Les six sujets ont été retenus après vérification de leur source primaire et représentent ${new Set(items.map((item) => item.category)).size} catégories éditoriales distinctes.`,
  alertEditorialSentence
].join(" ");

const report = {
  week: weekLabel,
  editorial_note: normalizeFrenchTypography(editorialNote),
  source_coverage: sourceCoverage.map((entry) => ({
    source_name: entry.source_name,
    domain: entry.domain,
    searched: true,
    candidate_count: entry.candidates.length,
    search_note: entry.search_note
  })),
  scouting_coverage: {
    google_alert_feeds: googleAlertFeedUrls.length,
    google_alert_candidates: googleAlertCandidates.length
  },
  alert_trace: googleAlertCandidates.map((candidate) => {
    const decision = selection.alert_decisions.find((entry) => entry.alert_id === candidate.alert_id);
    return {
      alert_id: candidate.alert_id,
      title: candidate.title,
      url: candidate.url,
      publication_date: candidate.publication_date,
      selected: decision?.selected || false,
      reason: normalizeFrenchTypography(decision?.reason || "Non retenu lors de la sélection éditoriale.")
    };
  }),
  items,
  briefs: []
};
const mandatoryFields = ["category", "title", "source", "source_url", "publication_date", "source_access", "summary", "practical_relevance"];
const incompleteItems = report.items.filter((item) =>
  mandatoryFields.some((field) => typeof item[field] !== "string" || item[field].trim().length < 3)
  || (item.source_access === "PARTIEL" && item.access_warning.length < 20)
  || !isCurrentWeekPublication(item.publication_date)
);
if (incompleteItems.length || report.items.length !== 6) {
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
    keepLines: true,
    spacing: { after: 130, line: 276 },
    children: [new TextRun({ text })]
  }));

const labeledParagraph = (label, value) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  keepLines: true,
  spacing: { after: 130, line: 276 },
  children: [
    new TextRun({ text: `${label}. `, bold: true, color: "1F4E79" }),
    new TextRun({ text: String(value || "") })
  ]
});

const sourceParagraph = (item) => new Paragraph({
  keepNext: true,
  spacing: { after: 240 },
  children: [
    new ExternalHyperlink({
      link: item.source_url,
      children: [new TextRun({
        text: `${item.court_reference || item.source}, ${frenchDate(item.publication_date)}`,
        color: "507D82",
        underline: {}
      })]
    })
  ]
});

const accessWarningParagraph = (item) => item.source_access !== "COMPLET"
  ? new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      keepNext: true,
      spacing: { after: 220 },
      children: [new TextRun({
        text: item.access_warning,
        color: "C55A11",
        bold: true,
        italics: true
      })]
    })
  : null;

const julietteAlertParagraph = (item) => item.discovered_via_juliette_alert
  ? new Paragraph({
      keepNext: true,
      spacing: { after: 180 },
      children: [new TextRun({
        text: "Signalé par l’alerte de Juliette.",
        color: "2E75B6",
        bold: true,
        italics: true
      })]
    })
  : null;

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

const renderItem = (item) => {
  const warning = accessWarningParagraph(item);
  const alertOrigin = julietteAlertParagraph(item);
  const body = item.type === "JURISPRUDENCE"
    ? [
        ...textParagraphs(item.introduction),
        ...textParagraphs(item.facts_and_procedure),
        ...textParagraphs(item.parties_arguments),
        ...textParagraphs(item.legal_question),
        ...textParagraphs(item.reasoning),
        ...textParagraphs(item.outcome)
      ]
    : [
        labeledParagraph("Contexte", item.context),
        labeledParagraph("Fondement et champ d’application", item.legal_basis_and_scope),
        labeledParagraph("Principales dispositions", item.main_provisions),
        labeledParagraph("Mise en œuvre", item.implementation_timeline)
      ];
  const heading = new Paragraph({
    heading: HeadingLevel.HEADING_2,
    keepNext: true,
    spacing: { before: 180, after: 100 },
    children: [new TextRun({
      text: item.type === "JURISPRUDENCE" ? `${item.category} : ${item.title}` : item.title,
      bold: true,
      color: "1F4E79"
    })]
  });
  const content = [
    sourceParagraph(item),
    ...(alertOrigin ? [alertOrigin] : []),
    ...(warning ? [warning] : []),
    ...body,
    practicalParagraph(item)
  ];
  if (item.type !== "JURISPRUDENCE") return [heading, ...content];
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const boxBorder = { style: BorderStyle.SINGLE, size: 6, color: "7F7F7F" };
  return [
    heading,
    new Table({
      width: { size: 9866, type: WidthType.DXA },
      columnWidths: [9866],
      borders: {
        top: noBorder,
        bottom: noBorder,
        left: noBorder,
        right: noBorder,
        insideHorizontal: noBorder,
        insideVertical: noBorder
      },
      rows: content.map((paragraph, index) => new TableRow({
        children: [new TableCell({
          width: { size: 9866, type: WidthType.DXA },
          margins: { top: index === 0 ? 120 : 20, bottom: index === content.length - 1 ? 120 : 20, left: 180, right: 180 },
          borders: {
            top: index === 0 ? boxBorder : noBorder,
            bottom: index === content.length - 1 ? boxBorder : noBorder,
            left: boxBorder,
            right: boxBorder,
            insideHorizontal: noBorder,
            insideVertical: noBorder
          },
          children: [paragraph]
        })]
      }))
    })
  ];
};

const renderBrief = (item) => {
  const alertOrigin = julietteAlertParagraph(item);
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      keepNext: true,
      spacing: { before: 160, after: 90 },
      children: [new TextRun({ text: `${item.category} : ${item.title}`, bold: true, color: "1F4E79" })]
    }),
    sourceParagraph(item),
    ...(alertOrigin ? [alertOrigin] : []),
    accessWarningParagraph(item),
    ...textParagraphs(item.summary)
  ];
};

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

const createFooter = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    children: [
      new TextRun({ text: "Veille IP · " }),
      new TextRun({ children: [PageNumber.CURRENT] })
    ]
  })]
});

const document = new Document({
  creator: "Veille IP DLA",
  title: `Veille Propriété intellectuelle : ${report.week}`,
  description: "Veille hebdomadaire de propriété intellectuelle",
  styles: {
    default: {
      document: {
        run: { font: "Arial", size: 22, color: "171717" },
        paragraph: { spacing: { after: 130, line: 276 } }
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
        margin: { top: 1020, right: 1020, bottom: 1020, left: 1020 }
      }
    },
    footers: {
      default: createFooter()
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
