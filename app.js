let current;

const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
}[character]));

fetch("/public/latest.json")
  .then((response) => response.json())
  .then((data) => {
    current = data;
    document.querySelector("#week").textContent = "SEMAINE DU " + data.week.toUpperCase();
    const displayedItems = [...data.items, ...(data.briefs || [])];
    document.querySelector("#itemCount").textContent = displayedItems.length;
    document.querySelector("#items").innerHTML = displayedItems.map((item, index) => `
      <article class="item">
        <span class="number">${String(index + 1).padStart(2, "0")}</span>
        <div>
          <span class="tag">${esc(item.category)}</span>
          <h3>${esc(item.title)}</h3>
          ${item.discovered_via_juliette_alert ? '<small class="alert-origin">Signalé par l’alerte de Juliette.</small>' : ''}
          <p>${esc(item.source)} · ${esc(item.summary)}</p>
        </div>
        <a href="${esc(item.source_url)}" target="_blank" rel="noopener">↗</a>
      </article>
    `).join("");
  })
  .catch(() => {
    document.querySelector("#items").innerHTML = '<p class="loading">La veille n’est pas encore disponible.</p>';
  });

document.querySelector("#download").addEventListener("click", () => {
  if (!current) return;
  if (current.report_url) {
    const link = document.createElement("a");
    link.href = current.report_url;
    link.download = "";
    link.click();
    return;
  }
  window.print();
});
