// =============================================================================
// PharmaTrack v2 — top-performers.js
// "🏆 Top Performers — Expiry Risk" — ranks plants and materials by at-risk
// exposure, and breaks down risk timing by urgency window.
//
//   1. Top 10 Plants by total at-risk value
//   2. Top 10 At-Risk Items at a chosen plant (defaults to all plants combined)
//   3. Top 10 At-Risk Items nationally (summed across all plants)
//   4. Risk Timing by Plant — < 3mo / 3–6mo / 6–12mo at-risk value, shown as
//      side-by-side columns, rows sorted ascending by total at-risk value
//      (lowest-risk plants first)
//
// Requires: script.js (rawDf, fmtQty, fmtETB, escHtml, buildTable,
//           wireTableExport, setKpis, PLOTLY_LAYOUT, PLOTLY_CONFIG,
//           waitForPlotly, PAGE_RENDERERS, renderPage)
//           mos.js (mosMerged, mosPlants, HUB_PLANT)
//           expiry-risk.js (buildRiskSnapshot)
// Must be loaded AFTER script.js, mos.js, AND expiry-risk.js.
// =============================================================================

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

  const plantEl = document.getElementById("topperf-plant");
  if (plantEl && plantEl.options.length <= 1 && typeof mosPlants !== "undefined") {
    mosPlants.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p; opt.text = p === HUB_PLANT ? `${p} (Hub)` : p;
      plantEl.appendChild(opt);
    });
  }
  const plantVal = plantEl ? plantEl.value.trim() : "";

  // ── Full at-risk snapshot, network-wide (respects the global person filter
  // the same way expiry-risk.js does, via buildRiskSnapshot → getMosFilteredRows) ──
  const snapshot = buildRiskSnapshot("", "", "");
  const atRisk   = snapshot.filter(r => r.atRisk && r.atRiskQty > 0);

  // ── KPIs ────────────────────────────────────────────────────────────────
  const totalVal      = atRisk.reduce((s, r) => s + r.atRiskVal, 0);
  const plantsAffected = new Set(atRisk.map(r => r.plant)).size;
  const itemsAffected  = new Set(atRisk.map(r => r.code)).size;
  const plantsTotal    = (typeof mosPlants !== "undefined" && mosPlants.length) ? mosPlants.length : plantsAffected;
  setKpis("topperf-kpis", [
    ["Total At-Risk Value", fmtETB(totalVal), "Network-wide exposure", "red"],
    ["Plants Affected", plantsAffected.toLocaleString(), `of ${plantsTotal} plants`, "amber"],
    ["Materials Affected", itemsAffected.toLocaleString(), "distinct at-risk SKUs", "amber"],
  ]);

  if (!atRisk.length) {
    ["topperf-table-plants", "topperf-table-plant-items", "topperf-table-national", "topperf-table-timing"]
      .forEach(id => { document.getElementById(id).innerHTML = `<div class="alert-info">No at-risk items found.</div>`; });
    document.getElementById("chart-topperf-plants").innerHTML = "";
    return;
  }

  // ── 1. TOP 10 PLANTS BY TOTAL AT-RISK VALUE ───────────────────────────────
  const plantMap = new Map();
  for (const r of atRisk) {
    if (!plantMap.has(r.plant)) {
      plantMap.set(r.plant, { plant: r.plant, isHub: r.isHub, atRiskVal: 0, atRiskQty: 0, items: new Set() });
    }
    const e = plantMap.get(r.plant);
    e.atRiskVal += r.atRiskVal;
    e.atRiskQty += r.atRiskQty;
    e.items.add(r.code);
  }
  const topPlants = [...plantMap.values()]
    .map(e => ({ plant: e.plant, isHub: e.isHub, atRiskVal: e.atRiskVal, atRiskQty: e.atRiskQty, itemCount: e.items.size }))
    .sort((a, b) => b.atRiskVal - a.atRiskVal)
    .slice(0, 10);

  if (topPlants.length) {
    const chartRows = [...topPlants].reverse(); // horizontal bar: highest at top
    Plotly.newPlot("chart-topperf-plants", [{
      type: "bar",
      orientation: "h",
      x: chartRows.map(p => p.atRiskVal),
      y: chartRows.map(p => p.isHub ? `${p.plant} (Hub)` : p.plant),
      marker: { color: "#f85149" },
      hovertemplate: "<b>%{y}</b><br>At-risk value: ETB %{x:,.0f}<extra></extra>",
    }], {
      ...PLOTLY_LAYOUT,
      height: 360,
      margin: { l: 90, r: 30, t: 20, b: 40 },
      xaxis: { title: "At-Risk Value (ETB)", tickformat: "~s" },
    }, PLOTLY_CONFIG);
  } else {
    document.getElementById("chart-topperf-plants").innerHTML = "";
  }

  const plantRankRows = topPlants.map((p, i) => ({ ...p, rank: i + 1 }));
  const plantCols = [
    { key: "rank", label: "#" },
    { key: "plant", label: "Plant",
      fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v),
      raw: true },
    { key: "itemCount", label: "At-Risk Materials" },
    { key: "atRiskQty", label: "At-Risk Qty", fmt: fmtQty },
    { key: "atRiskVal", label: "At-Risk Value", fmt: v => `<b style="color:var(--red)">${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("topperf-table-plants").innerHTML = buildTable(
    plantRankRows, plantCols, () => "", "", { id: "topperf-plants-export", title: "" }
  );
  wireTableExport("topperf-plants-export", plantRankRows,
    plantCols.map(c => ({ key: c.key, label: c.label })), "top10_plants_at_risk");

  // ── 2. TOP 10 AT-RISK ITEMS AT A PLANT (or all plants if none selected) ──
  const perPlantPool = plantVal ? atRisk.filter(r => r.plant === plantVal) : atRisk;
  const topItemsPerPlant = [...perPlantPool].sort((a, b) => b.atRiskVal - a.atRiskVal).slice(0, 10);
  const itemRankRows = topItemsPerPlant.map((r, i) => ({ ...r, rank: i + 1 }));
  const itemCols = [
    { key: "rank", label: "#" },
    { key: "code", label: "Material Code" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "plant", label: "Plant",
      fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v),
      raw: true },
    { key: "shelfLeftMo", label: "Shelf Life Left", fmt: v => v < 0 ? `<b style="color:var(--red)">EXPIRED</b>` : `<b>${v.toFixed(1)}</b> mo`, raw: true },
    { key: "atRiskQty", label: "At-Risk Qty", fmt: fmtQty },
    { key: "atRiskVal", label: "At-Risk Value", fmt: v => `<b style="color:var(--red)">${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("topperf-table-plant-items").innerHTML = itemRankRows.length
    ? buildTable(itemRankRows, itemCols, () => "", "", { id: "topperf-planitems-export", title: "" })
    : `<div class="alert-info">No at-risk items ${plantVal ? `at <b>${escHtml(plantVal)}</b>` : ""}.</div>`;
  if (itemRankRows.length) {
    wireTableExport("topperf-planitems-export", itemRankRows,
      itemCols.map(c => ({ key: c.key, label: c.label })), `top10_items_${plantVal || "all_plants"}`);
  }

  // ── 3. TOP 10 AT-RISK ITEMS NATIONALLY (summed across all plants) ────────
  const itemMap = new Map();
  for (const r of atRisk) {
    if (!itemMap.has(r.code)) {
      itemMap.set(r.code, { code: r.code, desc: r.desc, atRiskVal: 0, atRiskQty: 0, plants: new Set(), shelfLeftMo: r.shelfLeftMo });
    }
    const e = itemMap.get(r.code);
    e.atRiskVal += r.atRiskVal;
    e.atRiskQty += r.atRiskQty;
    e.plants.add(r.plant);
    // Worst-case (soonest) shelf life across all plants carrying this item
    if (r.shelfLeftMo !== null && (e.shelfLeftMo === null || r.shelfLeftMo < e.shelfLeftMo)) e.shelfLeftMo = r.shelfLeftMo;
  }
  const topNational = [...itemMap.values()]
    .map(e => ({ code: e.code, desc: e.desc, atRiskVal: e.atRiskVal, atRiskQty: e.atRiskQty, plantCount: e.plants.size, shelfLeftMo: e.shelfLeftMo }))
    .sort((a, b) => b.atRiskVal - a.atRiskVal)
    .slice(0, 10);
  const natRankRows = topNational.map((r, i) => ({ ...r, rank: i + 1 }));
  const natCols = [
    { key: "rank", label: "#" },
    { key: "code", label: "Material Code" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "plantCount", label: "Plants Affected" },
    { key: "shelfLeftMo", label: "Worst Shelf Life Left", fmt: v => v < 0 ? `<b style="color:var(--red)">EXPIRED</b>` : `<b>${v.toFixed(1)}</b> mo`, raw: true },
    { key: "atRiskQty", label: "At-Risk Qty", fmt: fmtQty },
    { key: "atRiskVal", label: "At-Risk Value", fmt: v => `<b style="color:var(--red)">${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("topperf-table-national").innerHTML = buildTable(
    natRankRows, natCols, () => "", "", { id: "topperf-national-export", title: "" }
  );
  wireTableExport("topperf-national-export", natRankRows,
    natCols.map(c => ({ key: c.key, label: c.label })), "top10_items_nationally");

  // ── 4. RISK TIMING BY PLANT ────────────────────────────────────────────────
  // < 3mo / 3–6mo / 6–12mo at-risk value shown as side-by-side columns.
  // Rows sorted ascending by TOTAL at-risk value (lowest-risk plants first).
  const timingMap = new Map();
  for (const r of atRisk) {
    if (!timingMap.has(r.plant)) {
      timingMap.set(r.plant, { plant: r.plant, isHub: r.isHub, lt3: 0, m3to6: 0, m6to12: 0, total: 0 });
    }
    const e = timingMap.get(r.plant);
    const m = r.shelfLeftMo;
    if (m < 3) e.lt3 += r.atRiskVal;
    else if (m < 6) e.m3to6 += r.atRiskVal;
    else if (m < 12) e.m6to12 += r.atRiskVal;
    // items with 12+ months of shelf life left still count toward the plant's
    // grand total, just not toward any of the three named windows.
    e.total += r.atRiskVal;
  }
  const timingRows = [...timingMap.values()].sort((a, b) => a.total - b.total);
  const timingCols = [
    { key: "plant", label: "Plant",
      fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v),
      raw: true },
    { key: "lt3", label: "< 3 Months", fmt: v => v > 0 ? `<b style="color:var(--red)">${fmtETB(v)}</b>` : "—", raw: true },
    { key: "m3to6", label: "3 – 6 Months", fmt: v => v > 0 ? `<b style="color:var(--amber)">${fmtETB(v)}</b>` : "—", raw: true },
    { key: "m6to12", label: "6 – 12 Months", fmt: v => v > 0 ? fmtETB(v) : "—", raw: true },
    { key: "total", label: "Total At-Risk Value", fmt: v => `<b>${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("topperf-table-timing").innerHTML = buildTable(
    timingRows, timingCols, () => "", "", { id: "topperf-timing-export", title: "" }
  );
  wireTableExport("topperf-timing-export", timingRows,
    timingCols.map(c => ({ key: c.key, label: c.label })), "at_risk_timing_by_plant");
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
      "topperf-apply": renderTopPerformers,
      "topperf-clear": () => {
        const p = document.getElementById("topperf-plant"); if (p) p.value = "";
        renderTopPerformers();
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
