// =============================================================================
// PharmaTrack v2 — stockout-risk.js
// "📉 Stockout Risk" — NATIONAL-LEVEL ONLY (no per-plant breakdown).
//
// RULE (per product decision)
// ----------------------------
//   A material is AT RISK OF STOCKOUT when:
//       National MOS < 4 months
//   where National MOS is EXACTLY mos.js's existing computeNationalMOS():
//       National MOS = (SOH at every plant, including HO01)
//                     ÷ (AMC at every BRANCH plant, excluding HO01)
//   This is the same number already shown elsewhere (MOS by Plant, National
//   Table) — this page just applies one fixed lens to it: "is the WHOLE
//   network running low," not "is any single branch running low." Per-plant
//   shortfalls that redistribution can fix are intentionally out of scope
//   here (that's a hub→branch logistics problem, not a national stockout).
//
// SCOPE DECISIONS (confirmed)
// ---------------------------
//   - Threshold is a fixed constant (STOCKOUT_MOS_THRESHOLD = 4), not a UI input.
//   - Respects the global sidebar Person Filter, same as every other MOS-derived page.
//   - Materials with National MOS = null (no branch committed at all) or
//     = Infinity (stock but zero branch demand) are EXCLUDED from this page
//     entirely — neither case is "at risk of running out," and null has no
//     basis to judge risk against in the first place.
//   - A minimum National SOH Value (ETB) filter is available to cut noise
//     from very-low-value items cluttering the at-risk list.
//
// Requires: script.js (fmtQty, fmtETB, escHtml, buildTable, wireTableExport,
//           downloadCSV, downloadExcel, kpiCard, PLOTLY_LAYOUT, PLOTLY_CONFIG,
//           waitForPlotly, PAGE_RENDERERS, renderPage, currentPage, personFilter,
//           getReconciledBase)
//           mos.js (HUB_PLANT, mosMerged, mosPlants, buildMosSohMap,
//           computeNationalMOS, getMosFilteredRows, fmtMosVal)
// Must be loaded AFTER both script.js and mos.js.
// =============================================================================

const STOCKOUT_MOS_THRESHOLD = 4; // months — fixed per product decision, network-wide only

// ── NATIONAL UNIT VALUE LOOKUP (material → ETB per unit) ──────────────────────
// Same approximation expiry-risk.js uses: unit value = total Value of
// Unrestricted Stock ÷ total Unrestricted Stock qty, summed across every
// plant nationally (not just one plant) since this page has no per-plant
// view. Used only to turn National SOH into an ETB figure for the KPI/table/
// min-value filter — it does not affect the MOS or at-risk determination.
function buildNationalUnitValueMap() {
  const map = new Map(); // code → { valSum, qtySum }
  const base = (typeof getReconciledBase === "function") ? getReconciledBase() : (typeof rawDf !== "undefined" ? rawDf : []);
  if (!base.length) return map;

  for (const row of base) {
    const mat = String(row._mappedMaterial || row["Material"] || "").trim();
    if (!mat) continue;
    const qty = Number(row["Unrestricted Stock"] || 0);
    const val = Number(row["Value of Unrestricted Stock"] || 0);
    if (qty <= 0) continue;

    if (!map.has(mat)) map.set(mat, { valSum: 0, qtySum: 0 });
    const entry = map.get(mat);
    entry.valSum += val;
    entry.qtySum += qty;
  }
  return map;
}

function nationalUnitValueFor(map, code) {
  const entry = map.get(code);
  if (!entry || !entry.qtySum) return 0;
  return entry.valSum / entry.qtySum;
}

// ── BUILD THE NATIONAL STOCKOUT-RISK SNAPSHOT ─────────────────────────────────
// Returns an array of { code, desc, type, person, totalSoh, totalAmc, mos,
//                        unitVal, value, atRisk }
// Only includes materials where National MOS is a real, finite number — null
// (nothing committed anywhere) and Infinity (demand-free) are dropped, per
// the confirmed scope decision above.
function buildStockoutSnapshot(typeFilter, searchQ) {
  if (typeof mosMerged === "undefined" || !mosMerged.length) return [];

  const sohMap   = buildMosSohMap();          // from mos.js
  const valueMap = buildNationalUnitValueMap();

  // getMosFilteredRows() already applies the global personFilter before
  // type/search, exactly like MOS by Plant / Expiry Risk.
  const rows = (typeof getMosFilteredRows === "function")
    ? getMosFilteredRows(typeFilter || "", searchQ || "")
    : mosMerged.filter(r => (!typeFilter || r.type === typeFilter));

  const out = [];
  for (const r of rows) {
    const nat = computeNationalMOS(r, sohMap); // from mos.js
    if (nat.mos === null || nat.mos === Infinity) continue; // no basis / no real demand

    const unitVal = nationalUnitValueFor(valueMap, r.code);
    const value   = nat.totalSoh * unitVal;

    out.push({
      code: r.code, desc: r.desc, type: r.type, person: r.person || "",
      isMerged: r.isMerged, origCodes: r.origCodes,
      totalSoh: nat.totalSoh, totalAmc: nat.totalAmc, mos: nat.mos,
      unitVal, value,
      atRisk: nat.mos < STOCKOUT_MOS_THRESHOLD,
    });
  }
  return out;
}

// ── FORMATTING HELPERS ────────────────────────────────────────────────────────
function stkoKpiCard(label, value, sub, color) {
  return `<div class="kpi-card"><div class="kpi-label">${escHtml(label)}</div><div class="kpi-value" style="color:var(--${color||'blue'})">${value}</div>${sub ? `<div class="kpi-sub">${sub}</div>` : ""}</div>`;
}
function stkoKpiRow(cards) {
  const el = document.getElementById("stko-kpis");
  if (el) el.innerHTML = cards.join("");
}
function stkoMosCellStyle(mos) {
  return mos < STOCKOUT_MOS_THRESHOLD ? "color:var(--red);font-weight:700" : "color:var(--text)";
}

// ── MAIN RENDER ────────────────────────────────────────────────────────────────
async function renderStockoutRisk() {
  await waitForPlotly();

  const hasInventory = typeof rawDf !== "undefined" && rawDf.length > 0;
  const hasAmc        = typeof mosMerged !== "undefined" && mosMerged.length > 0;

  if (!hasInventory || !hasAmc) {
    document.getElementById("stko-no-data").style.display = "block";
    document.getElementById("stko-content").style.display = "none";
    return;
  }
  document.getElementById("stko-no-data").style.display  = "none";
  document.getElementById("stko-content").style.display  = "block";

  const searchEl   = document.getElementById("stko-search");
  const typeEl     = document.getElementById("stko-type");
  const minValEl   = document.getElementById("stko-min-value");
  const atRiskOnly = document.getElementById("stko-at-risk-only");

  const searchQ  = searchEl ? searchEl.value.trim() : "";
  const typeVal  = typeEl   ? typeEl.value.trim()   : "";
  const minValue = minValEl ? (Number(minValEl.value) || 0) : 0;
  const riskOnly = atRiskOnly ? atRiskOnly.checked : true;

  let snapshot = buildStockoutSnapshot(typeVal, searchQ);

  // Min-value filter — cuts noise from near-zero-ETB items
  if (minValue > 0) snapshot = snapshot.filter(r => r.value >= minValue);

  const screenedCount = snapshot.length;
  const atRiskRows    = snapshot.filter(r => r.atRisk);
  const atRiskValue   = atRiskRows.reduce((s, r) => s + r.value, 0);
  const avgAtRiskMos  = atRiskRows.length
    ? atRiskRows.reduce((s, r) => s + r.mos, 0) / atRiskRows.length
    : null;

  // ── KPIs ──────────────────────────────────────────────────────────────────
  stkoKpiRow([
    stkoKpiCard("Materials Screened", screenedCount.toLocaleString(), typeVal || "All types · National MOS only", "blue"),
    stkoKpiCard(`At Risk of Stockout (<${STOCKOUT_MOS_THRESHOLD}mo)`, atRiskRows.length.toLocaleString(), `of ${screenedCount.toLocaleString()} screened nationally`, "red"),
    stkoKpiCard("At-Risk Value", fmtETB(atRiskValue), "Total National SOH value of at-risk materials", "amber"),
    stkoKpiCard("Avg. MOS (At-Risk Only)", avgAtRiskMos !== null ? `${avgAtRiskMos.toFixed(1)} mo` : "—", `Threshold: ${STOCKOUT_MOS_THRESHOLD} mo`, "purple"),
  ]);

  // ── CHART: distribution of National MOS for screened items ────────────────
  const chartRows = riskOnly ? atRiskRows : snapshot;
  const sortedForChart = [...chartRows].sort((a, b) => a.mos - b.mos).slice(0, 40);
  Plotly.newPlot("chart-stko", [{
    type: "bar",
    orientation: "h",
    y: sortedForChart.map(r => `${r.code}`).reverse(),
    x: sortedForChart.map(r => r.mos).reverse(),
    marker: { color: sortedForChart.map(r => r.mos < STOCKOUT_MOS_THRESHOLD ? "#f85149" : "#3a8fd4").reverse() },
    text: sortedForChart.map(r => `${r.mos.toFixed(1)}mo`).reverse(),
    textposition: "outside",
    textfont: { size: 9 },
    hovertemplate: "<b>%{y}</b><br>National MOS: %{x:.1f} months<extra></extra>",
  }], {
    ...PLOTLY_LAYOUT,
    height: Math.max(320, sortedForChart.length * 18),
    margin: { l: 90, r: 40, t: 20, b: 50 },
    xaxis: { title: `National MOS (months) — lowest first, capped at 40 items`, tickfont: { size: 10 } },
    yaxis: { tickfont: { size: 9 } },
    shapes: [{
      type: "line", x0: STOCKOUT_MOS_THRESHOLD, x1: STOCKOUT_MOS_THRESHOLD, y0: -0.5, y1: sortedForChart.length - 0.5,
      line: { color: "#f85149", width: 1.5, dash: "dot" },
    }],
    paper_bgcolor: "rgba(0,0,0,0)", plot_bgcolor: "rgba(0,0,0,0)",
  }, PLOTLY_CONFIG);

  // ── TABLE ──────────────────────────────────────────────────────────────────
  const tableRows = (riskOnly ? atRiskRows : snapshot).sort((a, b) => a.mos - b.mos); // most urgent first

  const cols = [
    { key: "code", label: "Material Code",
      fmt: (v, r) => r.isMerged
        ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-mapped-badge" title="Merged from: ${escHtml(r.origCodes)}">MERGED</span>`
        : `<span class="col-mat-code">${escHtml(v)}</span>`,
      raw: true, cellClass: "col-mat-code-wrap" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "type", label: "Type" },
    { key: "person", label: "Person Assigned", fmt: v => v || "—" },
    { key: "totalSoh", label: "National SOH", fmt: fmtQty },
    { key: "totalAmc", label: "National AMC (branches)", fmt: fmtQty },
    { key: "mos", label: "National MOS",
      fmt: v => `<span style="${stkoMosCellStyle(v)}">${fmtMosVal(v)}</span>`, raw: true },
    { key: "value", label: "National SOH Value", fmt: fmtETB },
  ];

  document.getElementById("stko-table").innerHTML = tableRows.length
    ? buildTable(tableRows, cols, (row) => row.atRisk ? "row-critical" : "", "", { id: "stko-export", title: "" })
    : '<div class="alert-info" style="margin:0.5rem 0">✓ No materials match the current filters at national stockout risk.</div>';

  // ── EXPORT ────────────────────────────────────────────────────────────────
  const exportCols = [
    { key: "code", label: "Material Code" }, { key: "desc", label: "Description" }, { key: "type", label: "Type" },
    { key: "person", label: "Person Assigned" },
    { key: "totalSoh", label: "National SOH", fmt: v => Number(v || 0).toFixed(2) },
    { key: "totalAmc", label: "National AMC (branches only)", fmt: v => Number(v || 0).toFixed(2) },
    { key: "mos", label: "National MOS (months)", fmt: v => Number(v).toFixed(2) },
    { key: "unitVal", label: "Unit Value (ETB)", fmt: v => Number(v || 0).toFixed(2) },
    { key: "value", label: "National SOH Value (ETB)", fmt: v => Number(v || 0).toFixed(2) },
    { key: "atRisk", label: `At Risk (<${STOCKOUT_MOS_THRESHOLD}mo)?`, fmt: v => v ? "Yes" : "No" },
  ];
  if (tableRows.length) wireTableExport("stko-export", tableRows, exportCols, "national_stockout_risk");

  const dlRow = document.getElementById("stko-dl-row");
  if (dlRow) {
    dlRow.innerHTML = '<button class="dl-btn">⬇ CSV</button><button class="dl-btn">⬇ Excel</button>';
    dlRow.querySelectorAll(".dl-btn")[0].onclick = () => downloadCSV(tableRows,   exportCols, "national_stockout_risk.csv");
    dlRow.querySelectorAll(".dl-btn")[1].onclick = () => downloadExcel(tableRows, exportCols, "national_stockout_risk.xlsx");
  }
}

// ── WIRE INTO PAGE_RENDERERS AND EVENT LISTENERS ──────────────────────────────
(function wireStockoutRiskModule() {
  function extend() {
    if (typeof PAGE_RENDERERS !== "undefined") {
      PAGE_RENDERERS["stockout-risk"] = renderStockoutRisk;
    }

    const _origRenderPage = window.renderPage;
    window.renderPage = function (id) {
      if (id === "stockout-risk") {
        currentPage = id;
        document.getElementById("landingView").style.display = "none";
        document.querySelectorAll(".page").forEach(el => { el.style.display = "none"; });
        const pg = document.getElementById("page-stockout-risk");
        if (pg) pg.style.display = "block";
        document.querySelectorAll(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.page === id));
        try { renderStockoutRisk(); } catch (e) { console.error(e); }
        return;
      }
      _origRenderPage(id);
    };

    const filterMap = {
      "stko-apply": renderStockoutRisk,
      "stko-clear": () => {
        const s = document.getElementById("stko-search");        if (s) s.value = "";
        const t = document.getElementById("stko-type");          if (t) t.value = "";
        const v = document.getElementById("stko-min-value");     if (v) v.value = "";
        const c = document.getElementById("stko-at-risk-only");  if (c) c.checked = true;
        renderStockoutRisk();
      },
    };

    document.body.addEventListener("click", (e) => {
      const btn = e.target.closest("button[id]");
      if (!btn) return;
      const fn = filterMap[btn.id];
      if (fn) { e.stopPropagation(); fn(); }
    }, true);

    const atRiskToggle = document.getElementById("stko-at-risk-only");
    if (atRiskToggle) atRiskToggle.addEventListener("change", () => { if (mosMerged.length) renderStockoutRisk(); });

    // Enter-to-apply in the search box, same UX as other search filters
    const searchInput = document.getElementById("stko-search");
    if (searchInput) {
      searchInput.addEventListener("keydown", (e) => { if (e.key === "Enter") renderStockoutRisk(); });
    }

    // Re-render if currently on this page and either source file changes
    const fileInput  = document.getElementById("fileInput");
    const mosAmcInput = document.getElementById("mosAmcFileInput");
    [fileInput, mosAmcInput].forEach(inp => {
      if (!inp) return;
      inp.addEventListener("change", () => {
        setTimeout(() => { if (currentPage === "stockout-risk") renderStockoutRisk(); }, 350);
      });
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", extend);
  } else {
    extend();
  }
})();
