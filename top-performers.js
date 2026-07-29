// =============================================================================
// PharmaTrack v2 — top-performers.js
// "🏆 Top Risk Exposure" — ranks plants and materials by risk exposure, split
// into two independent sections, each with the same 4 reports:
//
//   📉 EXPIRY RISK EXPOSURE — stock that will outlive its shelf life before
//      it can be consumed (MOS > months of shelf life remaining). Same
//      definition as expiry-risk.js's "at risk" flag.
//   📦 OVERSTOCK RISK EXPOSURE — stock sitting well above normal consumption
//      pace, REGARDLESS of expiry timing (MOS > OVERSTOCK_MOS_THRESHOLD_MO).
//
// Each section has:
//   1. Top 10 Plants by total exposure value
//   2. Top 10 Items at a chosen plant (defaults to all plants combined)
//   3. Top 10 Items nationally (summed across all plants)
//   4. A breakdown-by-plant table, side-by-side windows, rows sorted
//      ascending by total exposure value (lowest-risk plants first):
//        Expiry section    → < 3mo / 3–6mo / 6–12mo shelf life remaining
//        Overstock section → 6–9mo / 9–12mo / 12mo+ of stock on hand
//
// Requires: script.js (rawDf, fmtQty, fmtETB, escHtml, buildTable,
//           wireTableExport, setKpis, PLOTLY_LAYOUT, PLOTLY_CONFIG,
//           waitForPlotly, PAGE_RENDERERS, renderPage)
//           mos.js (mosMerged, mosPlants, HUB_PLANT, buildMosSohMap,
//           computeRowMOS, getMosFilteredRows)
//           expiry-risk.js (buildRiskSnapshot, buildExpiryMap, unitValueFor)
// Must be loaded AFTER script.js, mos.js, AND expiry-risk.js.
// =============================================================================

// Assumption: 6 months of stock on hand is treated as "overstocked" — a
// common pharma inventory heuristic. Adjust here if the business defines
// overstock differently.
const OVERSTOCK_MOS_THRESHOLD_MO = 6;

// ── Build the overstock snapshot: same shape/spirit as buildRiskSnapshot()
// in expiry-risk.js, but flags risk purely on MOS vs. OVERSTOCK_MOS_THRESHOLD_MO
// — expiry date is irrelevant here. Reuses buildExpiryMap() only to source a
// consistent unit valuation (Value of Unrestricted Stock ÷ Unrestricted Stock).
function buildOverstockSnapshot(typeFilter, searchQ, plantFilter) {
  if (typeof mosMerged === "undefined" || !mosMerged.length) return [];

  const sohMap    = buildMosSohMap();
  const valMap    = (typeof buildExpiryMap === "function") ? buildExpiryMap() : new Map();

  let rows = (typeof getMosFilteredRows === "function")
    ? getMosFilteredRows(typeFilter || "", searchQ || "")
    : mosMerged.filter(r => (!typeFilter || r.type === typeFilter));

  const out = [];
  for (const r of rows) {
    const plantMos = computeRowMOS(r, sohMap);
    for (const pm of plantMos) {
      if (plantFilter && pm.plant !== plantFilter) continue;
      if (pm.amc === null) continue;
      if (!pm.soh || pm.soh <= 0) continue;
      if (pm.mos === null || pm.mos === Infinity) continue;

      const overstocked = pm.mos > OVERSTOCK_MOS_THRESHOLD_MO;
      const excessQty    = overstocked ? Math.max(0, pm.soh - OVERSTOCK_MOS_THRESHOLD_MO * pm.amc) : 0;
      const valEntry     = valMap.get(r.code)?.[pm.plant] || null;
      const unitVal       = (typeof unitValueFor === "function") ? unitValueFor(valEntry) : 0;
      const excessVal     = excessQty * unitVal;

      out.push({
        code: r.code, desc: r.desc, type: r.type,
        isMerged: r.isMerged, origCodes: r.origCodes,
        plant: pm.plant, isHub: pm.isHub,
        soh: pm.soh, amc: pm.amc, mos: pm.mos,
        overstocked, excessQty, excessVal,
      });
    }
  }
  return out;
}

function populatePlantSelect(id) {
  const el = document.getElementById(id);
  if (el && el.options.length <= 1 && typeof mosPlants !== "undefined") {
    mosPlants.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p; opt.text = p === HUB_PLANT ? `${p} (Hub)` : p;
      el.appendChild(opt);
    });
  }
  return el ? el.value.trim() : "";
}

// ═════════════════════════════════════════════════════════════════════════
// SECTION 1 — EXPIRY RISK EXPOSURE
// ═════════════════════════════════════════════════════════════════════════
// Expiry windows the "at-risk" pool is split into. Each gets its own
// independent Top 10 Plants / Top 10 Items per Plant / Top 10 Items
// Nationally report set (see renderExpiryWindowSection below).
const EXPIRY_WINDOWS = [
  { key: "lt3",    label: "< 3 Months",   test: m => m < 3,               color: "#f85149" },
  { key: "m3to6",  label: "3 – 6 Months", test: m => m >= 3 && m < 6,     color: "#ffa657" },
  { key: "m6to12", label: "6 – 12 Months",test: m => m >= 6 && m < 12,    color: "#d29922" },
];

function renderExpiryRiskExposure() {
  const snapshot = buildRiskSnapshot("", "", "");
  const atRisk   = snapshot.filter(r => r.atRisk && r.atRiskQty > 0);

  const totalVal       = atRisk.reduce((s, r) => s + r.atRiskVal, 0);
  const plantsAffected = new Set(atRisk.map(r => r.plant)).size;
  const itemsAffected  = new Set(atRisk.map(r => r.code)).size;
  const plantsTotal    = (typeof mosPlants !== "undefined" && mosPlants.length) ? mosPlants.length : plantsAffected;
  setKpis("topperf-kpis", [
    ["Total At-Risk Value", fmtETB(totalVal), "Network-wide exposure", "red"],
    ["Plants Affected", plantsAffected.toLocaleString(), `of ${plantsTotal} plants`, "amber"],
    ["Materials Affected", itemsAffected.toLocaleString(), "distinct at-risk SKUs", "amber"],
  ]);

  for (const w of EXPIRY_WINDOWS) {
    const rows = atRisk.filter(r => w.test(r.shelfLeftMo));
    renderExpiryWindowSection(w, rows);
  }
}

// Renders one expiry window's report: Top 10 Plants by total at-risk
// value, scoped to `rows` (the at-risk rows already filtered down to
// this window).
function renderExpiryWindowSection(w, rows) {
  const chartId       = `chart-topperf-plants-${w.key}`;
  const plantsTableId = `topperf-table-plants-${w.key}`;

  if (!rows.length) {
    const el = document.getElementById(plantsTableId);
    if (el) el.innerHTML = `<div class="alert-info">No at-risk items in this window.</div>`;
    const chartEl = document.getElementById(chartId);
    if (chartEl) chartEl.innerHTML = "";
    return;
  }

  // Top 10 plants by total at-risk value (this window)
  const plantMap = new Map();
  for (const r of rows) {
    if (!plantMap.has(r.plant)) plantMap.set(r.plant, { plant: r.plant, isHub: r.isHub, atRiskVal: 0, atRiskQty: 0, items: new Set() });
    const e = plantMap.get(r.plant);
    e.atRiskVal += r.atRiskVal;
    e.atRiskQty += r.atRiskQty;
    e.items.add(r.code);
  }
  const topPlants = [...plantMap.values()]
    .map(e => ({ plant: e.plant, isHub: e.isHub, atRiskVal: e.atRiskVal, atRiskQty: e.atRiskQty, itemCount: e.items.size }))
    .sort((a, b) => b.atRiskVal - a.atRiskVal)
    .slice(0, 10);

  const chartEl = document.getElementById(chartId);
  if (topPlants.length && chartEl) {
    const chartRows = [...topPlants].reverse();
    Plotly.newPlot(chartId, [{
      type: "bar", orientation: "h",
      x: chartRows.map(p => p.atRiskVal),
      y: chartRows.map(p => p.isHub ? `${p.plant} (Hub)` : p.plant),
      marker: { color: w.color },
      hovertemplate: "<b>%{y}</b><br>At-risk value: ETB %{x:,.0f}<extra></extra>",
    }], {
      ...PLOTLY_LAYOUT, height: 360,
      margin: { l: 90, r: 30, t: 20, b: 40 },
      xaxis: { title: "At-Risk Value (ETB)", tickformat: "~s" },
    }, PLOTLY_CONFIG);
  } else if (chartEl) {
    chartEl.innerHTML = "";
  }

  const plantRankRows = topPlants.map((p, i) => ({ ...p, rank: i + 1 }));
  const plantCols = [
    { key: "rank", label: "#" },
    { key: "plant", label: "Plant",
      fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v), raw: true },
    { key: "itemCount", label: "At-Risk Materials" },
    { key: "atRiskQty", label: "At-Risk Qty", fmt: fmtQty },
    { key: "atRiskVal", label: "At-Risk Value", fmt: v => `<b style="color:var(--red)">${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById(plantsTableId).innerHTML = buildTable(
    plantRankRows, plantCols, () => "", "", { id: `topperf-plants-export-${w.key}`, title: "" }
  );
  wireTableExport(`topperf-plants-export-${w.key}`, plantRankRows, plantCols.map(c => ({ key: c.key, label: c.label })), `top10_plants_at_risk_${w.key}`);
}

// ═════════════════════════════════════════════════════════════════════════
// SECTION 2 — OVERSTOCK RISK EXPOSURE
// ═════════════════════════════════════════════════════════════════════════
function renderOverstockRiskExposure() {
  const plantVal = populatePlantSelect("topperf-ov-plant");

  const snapshot  = buildOverstockSnapshot("", "", "");
  const overstock = snapshot.filter(r => r.overstocked && r.excessQty > 0);

  const totalVal       = overstock.reduce((s, r) => s + r.excessVal, 0);
  const plantsAffected = new Set(overstock.map(r => r.plant)).size;
  const itemsAffected  = new Set(overstock.map(r => r.code)).size;
  const plantsTotal    = (typeof mosPlants !== "undefined" && mosPlants.length) ? mosPlants.length : plantsAffected;
  setKpis("topperf-ov-kpis", [
    ["Total Overstock Value", fmtETB(totalVal), `Excess above ${OVERSTOCK_MOS_THRESHOLD_MO} months of stock`, "red"],
    ["Plants Affected", plantsAffected.toLocaleString(), `of ${plantsTotal} plants`, "amber"],
    ["Materials Affected", itemsAffected.toLocaleString(), "distinct overstocked SKUs", "amber"],
  ]);

  if (!overstock.length) {
    ["topperf-ov-table-plants", "topperf-ov-table-plant-items", "topperf-ov-table-national", "topperf-ov-table-timing"]
      .forEach(id => { document.getElementById(id).innerHTML = `<div class="alert-info">No overstocked items found.</div>`; });
    document.getElementById("chart-topperf-ov-plants").innerHTML = "";
    return;
  }

  // 1. Top 10 plants by total overstock value
  const plantMap = new Map();
  for (const r of overstock) {
    if (!plantMap.has(r.plant)) plantMap.set(r.plant, { plant: r.plant, isHub: r.isHub, excessVal: 0, excessQty: 0, items: new Set() });
    const e = plantMap.get(r.plant);
    e.excessVal += r.excessVal;
    e.excessQty += r.excessQty;
    e.items.add(r.code);
  }
  const topPlants = [...plantMap.values()]
    .map(e => ({ plant: e.plant, isHub: e.isHub, excessVal: e.excessVal, excessQty: e.excessQty, itemCount: e.items.size }))
    .sort((a, b) => b.excessVal - a.excessVal)
    .slice(0, 10);

  if (topPlants.length) {
    const chartRows = [...topPlants].reverse();
    Plotly.newPlot("chart-topperf-ov-plants", [{
      type: "bar", orientation: "h",
      x: chartRows.map(p => p.excessVal),
      y: chartRows.map(p => p.isHub ? `${p.plant} (Hub)` : p.plant),
      marker: { color: "#ffa657" },
      hovertemplate: "<b>%{y}</b><br>Overstock value: ETB %{x:,.0f}<extra></extra>",
    }], {
      ...PLOTLY_LAYOUT, height: 360,
      margin: { l: 90, r: 30, t: 20, b: 40 },
      xaxis: { title: "Overstock Value (ETB)", tickformat: "~s" },
    }, PLOTLY_CONFIG);
  } else {
    document.getElementById("chart-topperf-ov-plants").innerHTML = "";
  }

  const plantRankRows = topPlants.map((p, i) => ({ ...p, rank: i + 1 }));
  const plantCols = [
    { key: "rank", label: "#" },
    { key: "plant", label: "Plant",
      fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v), raw: true },
    { key: "itemCount", label: "Overstocked Materials" },
    { key: "excessQty", label: "Excess Qty", fmt: fmtQty },
    { key: "excessVal", label: "Overstock Value", fmt: v => `<b style="color:var(--amber)">${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("topperf-ov-table-plants").innerHTML = buildTable(
    plantRankRows, plantCols, () => "", "", { id: "topperf-ov-plants-export", title: "" }
  );
  wireTableExport("topperf-ov-plants-export", plantRankRows, plantCols.map(c => ({ key: c.key, label: c.label })), "top10_plants_overstock");

  // 2. Top 10 overstocked items at a chosen plant (or all combined)
  const perPlantPool = plantVal ? overstock.filter(r => r.plant === plantVal) : overstock;
  const topItemsPerPlant = [...perPlantPool].sort((a, b) => b.excessVal - a.excessVal).slice(0, 10);
  const itemRankRows = topItemsPerPlant.map((r, i) => ({ ...r, rank: i + 1 }));
  const itemCols = [
    { key: "rank", label: "#" },
    { key: "code", label: "Material Code" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "plant", label: "Plant",
      fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v), raw: true },
    { key: "mos", label: "MOS", fmt: v => `<b>${v.toFixed(1)}</b> mo`, raw: true },
    { key: "excessQty", label: "Excess Qty", fmt: fmtQty },
    { key: "excessVal", label: "Overstock Value", fmt: v => `<b style="color:var(--amber)">${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("topperf-ov-table-plant-items").innerHTML = itemRankRows.length
    ? buildTable(itemRankRows, itemCols, () => "", "", { id: "topperf-ov-planitems-export", title: "" })
    : `<div class="alert-info">No overstocked items ${plantVal ? `at <b>${escHtml(plantVal)}</b>` : ""}.</div>`;
  if (itemRankRows.length) {
    wireTableExport("topperf-ov-planitems-export", itemRankRows, itemCols.map(c => ({ key: c.key, label: c.label })), `top10_overstock_items_${plantVal || "all_plants"}`);
  }

  // 3. Top 10 overstocked items nationally
  const itemMap = new Map();
  for (const r of overstock) {
    if (!itemMap.has(r.code)) itemMap.set(r.code, { code: r.code, desc: r.desc, excessVal: 0, excessQty: 0, plants: new Set(), mos: r.mos });
    const e = itemMap.get(r.code);
    e.excessVal += r.excessVal;
    e.excessQty += r.excessQty;
    e.plants.add(r.plant);
    if (r.mos !== null && (e.mos === null || r.mos > e.mos)) e.mos = r.mos; // worst (highest) MOS drives the risk
  }
  const topNational = [...itemMap.values()]
    .map(e => ({ code: e.code, desc: e.desc, excessVal: e.excessVal, excessQty: e.excessQty, plantCount: e.plants.size, mos: e.mos }))
    .sort((a, b) => b.excessVal - a.excessVal)
    .slice(0, 10);
  const natRankRows = topNational.map((r, i) => ({ ...r, rank: i + 1 }));
  const natCols = [
    { key: "rank", label: "#" },
    { key: "code", label: "Material Code" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "plantCount", label: "Plants Affected" },
    { key: "mos", label: "Worst MOS", fmt: v => `<b>${v.toFixed(1)}</b> mo`, raw: true },
    { key: "excessQty", label: "Excess Qty", fmt: fmtQty },
    { key: "excessVal", label: "Overstock Value", fmt: v => `<b style="color:var(--amber)">${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("topperf-ov-table-national").innerHTML = buildTable(
    natRankRows, natCols, () => "", "", { id: "topperf-ov-national-export", title: "" }
  );
  wireTableExport("topperf-ov-national-export", natRankRows, natCols.map(c => ({ key: c.key, label: c.label })), "top10_overstock_items_nationally");

  // 4. Overstock severity by plant — 6-9mo / 9-12mo / 12mo+, ascending by total value
  const timingMap = new Map();
  for (const r of overstock) {
    if (!timingMap.has(r.plant)) timingMap.set(r.plant, { plant: r.plant, isHub: r.isHub, m6to9: 0, m9to12: 0, m12plus: 0, total: 0 });
    const e = timingMap.get(r.plant);
    const m = r.mos;
    if (m < 9) e.m6to9 += r.excessVal;
    else if (m < 12) e.m9to12 += r.excessVal;
    else e.m12plus += r.excessVal;
    e.total += r.excessVal;
  }
  const timingRows = [...timingMap.values()].sort((a, b) => a.total - b.total);
  const timingCols = [
    { key: "plant", label: "Plant",
      fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v), raw: true },
    { key: "m6to9", label: "6 – 9 Months", fmt: v => v > 0 ? fmtETB(v) : "—", raw: true },
    { key: "m9to12", label: "9 – 12 Months", fmt: v => v > 0 ? `<b style="color:var(--amber)">${fmtETB(v)}</b>` : "—", raw: true },
    { key: "m12plus", label: "12+ Months", fmt: v => v > 0 ? `<b style="color:var(--red)">${fmtETB(v)}</b>` : "—", raw: true },
    { key: "total", label: "Total Overstock Value", fmt: v => `<b>${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("topperf-ov-table-timing").innerHTML = buildTable(
    timingRows, timingCols, () => "", "", { id: "topperf-ov-timing-export", title: "" }
  );
  wireTableExport("topperf-ov-timing-export", timingRows, timingCols.map(c => ({ key: c.key, label: c.label })), "overstock_severity_by_plant");
}

// ═════════════════════════════════════════════════════════════════════════
// TOP-LEVEL: renders both sections
// ═════════════════════════════════════════════════════════════════════════
async function renderTopPerformers() {
  await waitForPlotly();

  const hasInventory = typeof rawDf !== "undefined" && rawDf.length > 0;
  const hasAmc        = typeof mosMerged !== "undefined" && mosMerged.length > 0;
  const hasRiskFn      = typeof buildRiskSnapshot === "function";

  if (!hasInventory || !hasAmc || !hasRiskFn) {
    document.getElementById("topperf-no-data").style.display = "block";
    document.getElementById("topperf-content").style.display = "none";
    return;
  }
  document.getElementById("topperf-no-data").style.display  = "none";
  document.getElementById("topperf-content").style.display = "block";

  try { renderExpiryRiskExposure(); } catch (e) { console.error("[top-performers] expiry section failed:", e); }
  try { renderOverstockRiskExposure(); } catch (e) { console.error("[top-performers] overstock section failed:", e); }
}

// ── WIRE INTO PAGE_RENDERERS AND EVENT LISTENERS ──────────────────────────────
(function wireTopPerformersModule() {
  function extend() {
    if (typeof PAGE_RENDERERS !== "undefined") {
      PAGE_RENDERERS["top-performers"] = renderTopPerformers;
    }

    const _origRenderPage = window.renderPage;
    window.renderPage = function (id) {
      if (id === "top-performers") {
        currentPage = id;
        document.getElementById("landingView").style.display = "none";
        document.querySelectorAll(".page").forEach(el => { el.style.display = "none"; });
        const pg = document.getElementById("page-top-performers");
        if (pg) pg.style.display = "block";
        document.querySelectorAll(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.page === id));
        try { renderTopPerformers(); } catch (e) { console.error(e); }
        return;
      }
      _origRenderPage(id);
    };

    const filterMap = {
      "topperf-ov-apply": renderOverstockRiskExposure,
      "topperf-ov-clear": () => {
        const p = document.getElementById("topperf-ov-plant"); if (p) p.value = "";
        renderOverstockRiskExposure();
      },
    };

    document.body.addEventListener("click", (e) => {
      const btn = e.target.closest("button[id]");
      if (!btn) return;
      const fn = filterMap[btn.id];
      if (fn) { e.stopPropagation(); fn(); }
    }, true);

    // Re-render if currently on this page and either source file changes
    const fileInput   = document.getElementById("fileInput");
    const mosAmcInput = document.getElementById("mosAmcFileInput");
    [fileInput, mosAmcInput].forEach(inp => {
      if (!inp) return;
      inp.addEventListener("change", () => {
        setTimeout(() => { if (currentPage === "top-performers") renderTopPerformers(); }, 350);
      });
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", extend);
  } else {
    extend();
  }
})();
