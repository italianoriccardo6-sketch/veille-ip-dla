import fs from "node:fs/promises";

const sources = JSON.parse(await fs.readFile(new URL("../data/sources.json", import.meta.url), "utf8"));
if (!process.env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY manquante");
const domains = sources.map((source) => source.domain);

const itemProperties = {
  type: { type: "string", enum: ["JURISPRUDENCE", "ACTUALITE"] },
  category: { type: "string" },
  title: { type: "string" },
  court_reference: { type: "string" },
  source: { type: "string" },
  source_url: { type: "string" },
  publication_date: { type: "string" },
  summary: { type: "string" },
  introduction: { type: "string" },
  facts_and_procedure: { type: "string" },
  parties_arguments: { type: "string" },
  legal_question: { type: "string" },
  reasoning: { type: "string" },
  outcome: { type: "string" },
  practical_relevance: { type: "string" }
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
  "Recherche les développements des 7 derniers jours exclusivement dans les domaines autorisés et rédige une veille complète en français, directement exploitable par des avocats.",
  "",
  "SÉLECTION",
  "- Sélectionne exactement 5 ou 6 sujets: idéalement 4 jurisprudences et 1 ou 2 actualités.",
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
  "- La date doit appartenir à la période examinée ou correspondre à une décision nouvellement publiée/commentée pendant cette période.",
  "- Évite les doublons et les sujets trop faibles.",
  "- L'editorial_note doit signaler honnêtement toute limite de couverture ou d'accès.",
  "",
  "Domaines autorisés: " + domains.join(", ")
].join("\n");

const response = await fetch("https://api.openai.com/v1/responses", {
  method: "POST",
  headers: {
    Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
    "Content-Type": "application/json"
  },
  body: JSON.stringify({
    model: process.env.OPENAI_MODEL || "gpt-5.6-sol",
    reasoning: { effort: "high" },
    tools: [{
      type: "web_search",
      filters: { allowed_domains: domains },
      external_web_access: true,
      return_token_budget: "unlimited"
    }],
    tool_choice: "required",
    input,
    max_output_tokens: 18000,
    text: { format: { type: "json_schema", name: "veille_ip_editoriale", strict: true, schema } }
  })
});

const raw = await response.json();
if (!response.ok) throw new Error(JSON.stringify(raw));
const output = raw.output?.flatMap((entry) => entry.content || []).find((entry) => entry.type === "output_text")?.text;
if (!output) throw new Error("Sortie structurée absente");

const report = JSON.parse(output);
report.generated_at = new Date().toISOString();
report.status = "generated";
const slug = new Date().toISOString().slice(0, 10);
report.report_url = `/public/reports/veille-${slug}.html`;
await fs.mkdir("public/reports", { recursive: true });

const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
}[character]));
const paragraphs = (value) => String(value || "").split(/\n\s*\n/).filter(Boolean).map((paragraph) => `<p>${esc(paragraph)}</p>`).join("");

const jurisprudences = report.items.filter((item) => item.type === "JURISPRUDENCE");
const actualites = report.items.filter((item) => item.type === "ACTUALITE");

const renderJurisprudence = (item) => `
  <article class="item jurisprudence">
    <h2>${esc(item.category)} – ${esc(item.title)}</h2>
    <p class="reference"><a href="${esc(item.source_url)}">${esc(item.court_reference || item.source)}, ${esc(item.publication_date)}</a></p>
    ${paragraphs(item.introduction)}
    ${paragraphs(item.facts_and_procedure)}
    ${paragraphs(item.parties_arguments)}
    ${paragraphs(item.legal_question)}
    ${paragraphs(item.reasoning)}
    ${paragraphs(item.outcome)}
    <p class="portee"><strong>Portée pratique.</strong> ${esc(item.practical_relevance)}</p>
  </article>`;

const renderActualite = (item) => `
  <article class="item actualite">
    <h2>${esc(item.title)}</h2>
    ${paragraphs(item.introduction)}
    ${paragraphs(item.facts_and_procedure)}
    ${paragraphs(item.parties_arguments)}
    ${paragraphs(item.legal_question)}
    ${paragraphs(item.reasoning)}
    ${paragraphs(item.outcome)}
    <p class="portee"><strong>Portée pratique.</strong> ${esc(item.practical_relevance)}</p>
    <p class="reference"><a href="${esc(item.source_url)}">Voir la source</a></p>
  </article>`;

const html = `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width">
<title>Veille Propriété intellectuelle – ${esc(report.week)}</title>
<style>
  @page { size: A4; margin: 18mm 18mm 16mm; }
  * { box-sizing: border-box; }
  body { max-width: 820px; margin: 42px auto; padding: 0 28px; color: #171717; font: 13.5px/1.58 Arial, Helvetica, sans-serif; }
  header { text-align: center; margin-bottom: 34px; }
  h1 { margin: 0; font: 700 22px Georgia, "Times New Roman", serif; }
  .week { margin-top: 7px; font-size: 13px; }
  .section { margin: 34px 0 15px; font-size: 14px; letter-spacing: .02em; border-bottom: 1px solid #222; padding-bottom: 5px; }
  .item { margin: 0 0 38px; break-inside: auto; }
  .item h2 { margin: 0 0 10px; font-size: 14px; line-height: 1.35; text-decoration: underline; text-underline-offset: 3px; }
  .reference { margin: 0 0 17px; }
  .reference a { color: #507d82; text-decoration: underline; }
  p { margin: 0 0 13px; text-align: justify; }
  .portee { border-left: 2px solid #507d82; padding-left: 12px; }
  .editorial { margin-top: 30px; color: #666; font-size: 11px; }
  footer { text-align: right; color: #777; font-size: 10px; margin-top: 25px; }
  @media print { body { margin: 0; padding: 0; max-width: none; } a { color: #333 !important; } }
</style>
</head>
<body>
<header><h1>Veille Propriété intellectuelle</h1><p class="week">${esc(report.week)}</p></header>
${jurisprudences.length ? `<h2 class="section">JURISPRUDENCES</h2>${jurisprudences.map(renderJurisprudence).join("")}` : ""}
${actualites.length ? `<h2 class="section">ACTUALITÉS</h2>${actualites.map(renderActualite).join("")}` : ""}
<p class="editorial">${esc(report.editorial_note)}</p>
<footer>Veille générée le ${esc(new Date(report.generated_at).toLocaleDateString("fr-FR"))} · Sources primaires accessibles par les liens ci-dessus</footer>
</body>
</html>`;

await fs.writeFile(`public/reports/veille-${slug}.html`, html);
await fs.writeFile("public/latest.json", JSON.stringify(report, null, 2) + "\n");
console.log(`Generated ${jurisprudences.length} jurisprudences and ${actualites.length} actualités for ${report.week}`);