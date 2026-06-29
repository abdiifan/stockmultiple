// =============================================================================
// PharmaTrack v2 — amc.js
// AMC Analysis Module: HO01 Distribution, Coverage Map, Stock Imbalance,
// Redistribution Suggestion.
//
// Requires: script.js (fmtETB, fmtQty, escHtml, buildTable, downloadCSV,
//           downloadExcel, mappingTable, PLOTLY_LAYOUT, PLOTLY_CONFIG, waitForPlotly)
// Must be loaded AFTER script.js.
// =============================================================================

// ── AMC STATE ─────────────────────────────────────────────────────────────────
let amcRaw      = [];    // parsed rows from AMC.xlsx
let amcPlants   = [];    // ordered plant code list
let amcMerged   = [];    // deduplicated rows (after mapping)
let amcPersonMap = new Map(); // materialCode → PERSON name (from Sheet1)

// ── AMC PLANT LABELS (friendly names where known) ─────────────────────────────
const AMC_PLANT_LABELS = {
  AA01:"AA01", AA02:"AA02", AD01:"AD01", AR01:"AR01", AS01:"AS01",
  BD01:"BD01", DE01:"DE01", DI01:"DI01", GA01:"GA01", GO01:"GO01",
  HA01:"HA01", JI01:"JI01", JJ01:"JJ01", KD01:"KD01", MK01:"MK01",
  NB01:"NB01", NK01:"NK01", SE01:"SE01", SH01:"SH01",
};

// ── AMC FILE LOADER ───────────────────────────────────────────────────────────
function loadAMCFile(file) {
  const statusEl = document.getElementById("amcFileStatus");
  const btnEl    = document.getElementById("amcUploadBtnText");
  if (statusEl) statusEl.innerHTML = '<div class="status-loading">⏳ Parsing…</div>';

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb   = XLSX.read(e.target.result, { type: "array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

      if (!rows.length) throw new Error("AMC file is empty.");

      // Parse Sheet1 for PERSON lookup (materialCode → person name)
      amcPersonMap = new Map();
      const sheet1Name = wb.SheetNames.find(n => n.toLowerCase() === "sheet1");
      if (sheet1Name) {
        const ws1   = wb.Sheets[sheet1Name];
        const rows1 = XLSX.utils.sheet_to_json(ws1, { defval: null });
        for (const r1 of rows1) {
          const code1   = String(r1["Material Code"] || "").trim();
          const person1 = String(r1["PERSON"] || "").trim().toUpperCase();
          if (code1 && person1) amcPersonMap.set(code1, person1);
        }
      }

      // Detect plant columns (everything except meta cols + PERSON)
      const META = ["Material Code", "Description", "Material Type Code", "PERSON"];
      const firstRow = rows[0];
      const detectedPlants = Object.keys(firstRow).filter(k => !META.includes(k));
      if (!detectedPlants.length) throw new Error("No plant columns found in AMC file.");

      amcPlants = detectedPlants;
      amcRaw    = rows.map(r => ({
        code:   String(r["Material Code"] || "").trim(),
        desc:   String(r["Description"]   || "").trim(),
        type:   String(r["Material Type Code"] || "").trim().toUpperCase(),
        person: amcPersonMap.get(String(r["Material Code"] || "").trim()) || "",
        amcs:   Object.fromEntries(
          detectedPlants.map(p => [p, (r[p] == null || r[p] === "" || typeof r[p] === "string") ? null : Number(r[p])])
        ),
      }));

      amcMerged = buildAmcMerged();

      const count = amcMerged.length;
      if (statusEl) statusEl.innerHTML = `<div class="status-ok">✓ LOADED</div><div class="status-name">${escHtml(file.name)}</div><div class="status-name" style="color:var(--green)">${count} items · ${detectedPlants.length} plants</div>`;
      if (btnEl)    btnEl.textContent = "✓ " + file.name;

      // Show content on all AMC pages
      ["amc-no-file","amc-cov-no-file","amc-imb-no-file","amc-redist-no-file"].forEach(id => {
        const el = document.getElementById(id); if (el) el.style.display = "none";
      });
      ["amc-dist-content","amc-cov-content","amc-imb-content","amc-redist-content"].forEach(id => {
        const el = document.getElementById(id); if (el) el.style.display = "block";
      });

      // Re-render current AMC page if already on one
      const amcPages = ["amc-distribution","amc-coverage","amc-imbalance","amc-redistribution"];
      if (amcPages.includes(currentPage)) renderPage(currentPage);

    } catch (err) {
      console.error("AMC load error:", err);
      if (statusEl) statusEl.innerHTML = `<div class="status-error">⚠️ ${escHtml(err.message)}</div>`;
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── AMC DEDUPLICATION (mapping-aware) ─────────────────────────────────────────
// If a mapping file is loaded, collapse multiple AMC codes that share the same
// canonical target code. The AMC values are summed per plant across duplicates.
function buildAmcMerged() {
  if (!amcRaw.length) return [];

  // Build canonical code lookup: amcCode → { targetCode, targetDesc }
  // We check if an AMC material code appears as a SOURCE in mappingTable.
  const merged = new Map(); // canonicalCode → mergedRow

  for (const row of amcRaw) {
    let canonical = row.code;
    let canonDesc = row.desc;

    if (mappingTable && mappingTable.size > 0) {
      // Try exact match
      const entry = mappingTable.get(row.code);
      if (entry) {
        canonical = entry.targetCode;
        canonDesc = entry.targetDesc || row.desc;
      }
    }

    if (!merged.has(canonical)) {
      merged.set(canonical, {
        code:     canonical,
        origCode: row.code,
        origCodes: new Set([row.code]),
        desc:     canonDesc,
        type:     row.type,
        person:   row.person || "",
        amcs:     Object.fromEntries(amcPlants.map(p => [p, null])),
        isMerged: false,
      });
    }
    const m = merged.get(canonical);

    // Track all source codes for tooltip
    m.origCodes.add(row.code);
    if (m.origCodes.size > 1) m.isMerged = true;

    // Sum AMC values per plant (null + value = value; null + null = null)
    for (const p of amcPlants) {
      const v = row.amcs[p];
      if (v !== null && v !== undefined) {
        m.amcs[p] = (m.amcs[p] || 0) + v;
      }
    }
  }

  return Array.from(merged.values()).map(m => ({
    ...m,
    origCodes: [...m.origCodes].join(", "),
  }));
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function amcFmtVal(v) {
  if (v === null || v === undefined) return '<span style="color:var(--muted);font-size:0.8em">N/A</span>';
  return fmtQty(v);
}

function amcStockedCount(row) {
  return amcPlants.filter(p => row.amcs[p] !== null).length;
}

function amcTotalAMC(row) {
  return amcPlants.reduce((s, p) => s + (row.amcs[p] || 0), 0);
}

function amcGini(row) {
  const vals = amcPlants.map(p => row.amcs[p] || 0).filter((_, i) => row.amcs[amcPlants[i]] !== null);
  if (!vals.length) return 0;
  const n = vals.length;
  const sorted = [...vals].sort((a, b) => a - b);
  const mean = sorted.reduce((s, v) => s + v, 0) / n;
  if (!mean) return 0;
  let num = 0;
  for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) num += Math.abs(sorted[i] - sorted[j]);
  return num / (2 * n * n * mean);
}

function getAmcFilteredRows(typeFilter, personFilter) {
  if (!amcMerged.length) return [];
  let rows = amcMerged;
  if (typeFilter)   rows = rows.filter(r => r.type === typeFilter);
  if (personFilter) rows = rows.filter(r => r.person === personFilter.toUpperCase());
  return rows;
}

/** Build sorted list of unique persons from amcMerged */
function getAmcPersons() {
  const set = new Set(amcMerged.map(r => r.person).filter(Boolean));
  return [...set].sort();
}

/** Populate a <select> element with person options (idempotent) */
function populatePersonSelect(selectId) {
  const el = document.getElementById(selectId);
  if (!el) return;
  // Keep current value if possible
  const current = el.value;
  // Remove old person options (keep first blank option)
  while (el.options.length > 1) el.remove(1);
  for (const p of getAmcPersons()) {
    const opt = document.createElement("option");
    opt.value = p; opt.text = "👤 " + p;
    el.appendChild(opt);
  }
  if (current) el.value = current;
}

function amcKpiCard(label, value, sub, color) {
  return `<div class="kpi-card"><div class="kpi-label">${escHtml(label)}</div><div class="kpi-value" style="color:var(--${color||'blue'})">${value}</div>${sub ? `<div class="kpi-sub">${sub}</div>` : ""}</div>`;
}

function amcKpiRow(id, cards) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = cards.join("");
}

// ── 1. HO01 DISTRIBUTION TRACKER ─────────────────────────────────────────────
async function renderAmcDistribution() {
  await waitForPlotly();
  if (!amcMerged.length) return;

  const typeEl   = document.getElementById("amc-dist-type");
  const searchEl = document.getElementById("amc-dist-search");
  const personEl = document.getElementById("amc-dist-person");
  const typeVal  = typeEl   ? typeEl.value.trim()   : "";
  const searchQ  = searchEl ? searchEl.value.trim().toLowerCase() : "";
  const personVal = personEl ? personEl.value.trim() : "";
  populatePersonSelect("amc-dist-person");

  let rows = getAmcFilteredRows(typeVal, personVal);
  if (searchQ) rows = rows.filter(r =>
    r.code.toLowerCase().includes(searchQ) || r.desc.toLowerCase().includes(searchQ)
  );

  // KPIs
  const totalItems   = rows.length;
  const stockedInAll = rows.filter(r => amcStockedCount(r) === amcPlants.length).length;
  const naAll        = rows.filter(r => amcStockedCount(r) === 0).length;
  const totalAMCVal  = rows.reduce((s, r) => s + amcTotalAMC(r), 0);
  amcKpiRow("amc-dist-kpis", [
    amcKpiCard("Total Items", totalItems.toLocaleString(), `${typeVal || "All Types"}`, "blue"),
    amcKpiCard("Stocked in All Plants", stockedInAll.toLocaleString(), `${amcPlants.length} plants`, "green"),
    amcKpiCard("N/A in All Plants", naAll.toLocaleString(), "No coverage anywhere", "red"),
    amcKpiCard("Total AMC Value", fmtQty(totalAMCVal), "Across all plants", "purple"),
  ]);

  // Chart: items stocked per plant (bar chart of coverage)
  const plantCoverage = amcPlants.map(p => ({
    plant: p,
    stocked:  rows.filter(r => r.amcs[p] !== null).length,
    totalAMC: rows.reduce((s, r) => s + (r.amcs[p] || 0), 0),
  }));

  const COLORS = ["#3a8fd4","#2e9e5a","#c47f17","#8763cc","#5cbfdb","#4db87a",
    "#e09b2d","#e86060","#59b8f5","#70ce94","#3a8fd4","#2e9e5a","#c47f17",
    "#d94040","#8763cc","#5cbfdb","#4db87a","#e09b2d","#e86060"];

  const pl = () => ({
    ...PLOTLY_LAYOUT,
    paper_bgcolor:"rgba(0,0,0,0)", plot_bgcolor:"rgba(0,0,0,0)",
    font: { family:"Inter,sans-serif", color:"var(--text)", size:11 },
  });

  Plotly.newPlot("chart-amc-dist", [
    {
      type:"bar", name:"Items Stocked",
      x: plantCoverage.map(p => p.plant),
      y: plantCoverage.map(p => p.stocked),
      marker: { color: COLORS, opacity:0.9 },
      hovertemplate: "<b>%{x}</b><br>Items stocked: %{y}<extra></extra>",
      yaxis:"y",
    },
    {
      type:"scatter", mode:"lines+markers", name:"Total AMC Value (Qty)",
      x: plantCoverage.map(p => p.plant),
      y: plantCoverage.map(p => p.totalAMC),
      line: { color:"#ffa657", width:2 },
      marker: { size:6 },
      hovertemplate: "<b>%{x}</b><br>AMC: %{y:,.0f} units<extra></extra>",
      yaxis:"y2",
    }
  ], {
    ...pl(),
    height: 340,
    barmode:"group",
    xaxis: { title:"Plant", tickfont:{size:10} },
    yaxis: { title:"Items Stocked", side:"left" },
    yaxis2:{ title:"AMC Value (Qty)", overlaying:"y", side:"right", showgrid:false },
    legend:{ orientation:"h", y:-0.2 },
    margin:{ l:60, r:60, t:30, b:60 },
  }, PLOTLY_CONFIG);

  // Table
  const cols = [
    { key:"code",  label:"Material Code",
      fmt:(v,r) => r.isMerged
        ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-mapped-badge" title="Merged from: ${escHtml(r.origCodes)}">MERGED</span>`
        : `<span class="col-mat-code">${escHtml(v)}</span>`,
      raw:true, cellClass:"col-mat-code-wrap" },
    { key:"desc",  label:"Description",  cellClass:"col-mat-desc-wrap" },
    { key:"type",  label:"Type" },
    { key:"_stocked", label:"Plants Stocked",
      fmt:(v) => `<span style="font-weight:600;color:${v===amcPlants.length?'var(--green)':v===0?'var(--red)':'var(--text)'}">${v} / ${amcPlants.length}</span>`,
      raw:true },
    { key:"_total", label:"Total AMC (Qty)", fmt:fmtQty },
    ...amcPlants.map(p => ({
      key:`_p_${p}`, label:p,
      fmt:(v) => v===null ? '<span style="color:var(--muted);font-size:0.8em">N/A</span>' : fmtQty(v),
      raw:true,
    })),
  ];

  const tableRows = rows.map(r => ({
    ...r,
    _stocked: amcStockedCount(r),
    _total:   amcTotalAMC(r),
    ...Object.fromEntries(amcPlants.map(p => [`_p_${p}`, r.amcs[p]])),
  })).sort((a,b) => b._total - a._total);

  document.getElementById("amc-dist-table").innerHTML = buildTable(
    tableRows, cols, (row) => row._stocked === 0 ? "row-critical" : ""
  );

  // Export buttons
  const exportCols = [
    { key:"code", label:"Material Code" }, { key:"desc", label:"Description" },
    { key:"type", label:"Type" }, { key:"_stocked", label:"Plants Stocked" },
    { key:"_total", label:"Total AMC (Qty)" },
    ...amcPlants.map(p => ({ key:`_p_${p}`, label:p,
      fmt:(v)=>v===null?"N/A":String(Number(v).toFixed(2)) })),
  ];
  const dlRow = document.getElementById("amc-dist-dl-row");
  if (dlRow) {
    dlRow.innerHTML = '<button class="dl-btn">⬇ CSV</button><button class="dl-btn">⬇ Excel</button>';
    dlRow.querySelectorAll(".dl-btn")[0].onclick = () => downloadCSV(tableRows,   exportCols, "amc_distribution.csv");
    dlRow.querySelectorAll(".dl-btn")[1].onclick = () => downloadExcel(tableRows, exportCols, "amc_distribution.xlsx");
  }
}

// ── 2. AMC COVERAGE MAP ───────────────────────────────────────────────────────
async function renderAmcCoverage() {
  await waitForPlotly();
  if (!amcMerged.length) return;

  const typeEl   = document.getElementById("amc-cov-type");
  const plantEl  = document.getElementById("amc-cov-plant");
  const viewEl   = document.getElementById("amc-cov-view");
  const personEl = document.getElementById("amc-cov-person");
  const typeVal   = typeEl   ? typeEl.value.trim()  : "";
  const plantVal  = plantEl  ? plantEl.value.trim() : "";
  const viewMode  = viewEl   ? viewEl.value         : "all";
  const personVal = personEl ? personEl.value.trim() : "";
  populatePersonSelect("amc-cov-person");

  // Populate plant dropdown once
  if (plantEl && plantEl.options.length <= 1) {
    amcPlants.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p; opt.text = p;
      plantEl.appendChild(opt);
    });
  }

  let rows = getAmcFilteredRows(typeVal, personVal);

  // KPIs
  const totalItems   = rows.length;
  const totalCells   = rows.length * amcPlants.length;
  const naCount      = rows.reduce((s,r) => s + amcPlants.filter(p=>r.amcs[p]===null).length, 0);
  const covPct       = totalCells > 0 ? ((totalCells - naCount) / totalCells * 100).toFixed(1) : 0;
  const fullyStocked = rows.filter(r => amcPlants.every(p=>r.amcs[p]!==null)).length;

  amcKpiRow("amc-cov-kpis", [
    amcKpiCard("Total Items",          totalItems.toLocaleString(), typeVal||"All types", "blue"),
    amcKpiCard("Overall Coverage",     `${covPct}%`,               `${totalCells-naCount} of ${totalCells} cells`, "green"),
    amcKpiCard("N/A Gaps",             naCount.toLocaleString(),   "Missing coverage", "red"),
    amcKpiCard("Stocked Everywhere",   fullyStocked.toLocaleString(), `All ${amcPlants.length} plants`, "purple"),
  ]);

  // Heatmap: plants × items (top 40 items by coverage gap)
  let heatRows = [...rows].sort((a,b) => amcStockedCount(a) - amcStockedCount(b));
  // If plant filter applied, only show that plant's coverage
  const displayPlants = plantVal ? [plantVal] : amcPlants;
  const top = heatRows.slice(0, 50);

  // Z matrix: 1=stocked, 0=N/A
  const zMatrix = displayPlants.map(p => top.map(r => r.amcs[p] !== null ? 1 : 0));
  const labels  = top.map(r => r.desc.length > 35 ? r.desc.substring(0,35)+"…" : r.desc);

  Plotly.newPlot("chart-amc-coverage", [{
    type: "heatmap",
    z:    zMatrix,
    x:    labels,
    y:    displayPlants,
    colorscale: [[0,"#f85149"],[1,"#2e9e5a"]],
    zmin:0, zmax:1,
    showscale: true,
    colorbar: { tickvals:[0,1], ticktext:["N/A","Stocked"], len:0.4 },
    hovertemplate: "<b>%{y}</b> × <b>%{x}</b><br>%{z==1?'✅ Stocked':'❌ N/A'}<extra></extra>",
  }], {
    ...PLOTLY_LAYOUT,
    height: Math.max(300, displayPlants.length * 22 + 80),
    margin: { l:70, r:30, t:30, b:160 },
    xaxis: { tickangle:-45, tickfont:{size:9} },
    yaxis: { tickfont:{size:10} },
    paper_bgcolor:"rgba(0,0,0,0)", plot_bgcolor:"rgba(0,0,0,0)",
  }, PLOTLY_CONFIG);

  // Coverage detail table
  const covCols = [
    { key:"code", label:"Material Code",
      fmt:(v,r)=>r.isMerged
        ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-mapped-badge" title="Merged: ${escHtml(r.origCodes)}">MERGED</span>`
        : `<span class="col-mat-code">${escHtml(v)}</span>`,
      raw:true, cellClass:"col-mat-code-wrap" },
    { key:"desc", label:"Description", cellClass:"col-mat-desc-wrap" },
    { key:"type", label:"Type" },
    { key:"_stocked",   label:"Plants w/ Stock",
      fmt:(v) => `<span style="font-weight:600;color:${v===amcPlants.length?'var(--green)':v===0?'var(--red)':'var(--orange)'}">${v} / ${amcPlants.length}</span>`, raw:true },
    { key:"_naPlants",  label:"N/A Plants",
      fmt:(v,r) => v ? `<span style="font-size:0.78em;color:var(--red)">${escHtml(v)}</span>` : '<span style="color:var(--green)">None</span>', raw:true },
    { key:"_totalAMC",  label:"Total AMC (Qty)", fmt:fmtQty },
  ];

  let tableRows = rows.map(r => ({
    ...r,
    _stocked:  amcStockedCount(r),
    _naPlants: amcPlants.filter(p=>r.amcs[p]===null).join(", "),
    _totalAMC: amcTotalAMC(r),
  }));

  if (plantVal) tableRows = tableRows.filter(r => {
    const stocked = r.amcs[plantVal] !== null;
    if (viewMode === "gaps")    return !stocked;
    if (viewMode === "covered") return stocked;
    return true;
  }); else {
    if (viewMode === "gaps")    tableRows = tableRows.filter(r => r._stocked < amcPlants.length);
    if (viewMode === "covered") tableRows = tableRows.filter(r => r._stocked === amcPlants.length);
  }

  tableRows.sort((a,b) => a._stocked - b._stocked);
  document.getElementById("amc-cov-table").innerHTML = buildTable(
    tableRows, covCols, (row) => row._stocked === 0 ? "row-critical" : row._stocked < amcPlants.length/2 ? "row-warning" : ""
  );

  // Export
  const exportCols = [
    {key:"code",label:"Material Code"},{key:"desc",label:"Description"},{key:"type",label:"Type"},
    {key:"_stocked",label:"Plants w/ Stock"},{key:"_naPlants",label:"N/A Plants"},{key:"_totalAMC",label:"Total AMC (Qty)"},
  ];
  const dlRow = document.getElementById("amc-cov-dl-row");
  if (dlRow) {
    dlRow.innerHTML = '<button class="dl-btn">⬇ CSV</button><button class="dl-btn">⬇ Excel</button>';
    dlRow.querySelectorAll(".dl-btn")[0].onclick = () => downloadCSV(tableRows,   exportCols, "amc_coverage.csv");
    dlRow.querySelectorAll(".dl-btn")[1].onclick = () => downloadExcel(tableRows, exportCols, "amc_coverage.xlsx");
  }
}

// ── 3. STOCK IMBALANCE ALERT ─────────────────────────────────────────────────
async function renderAmcImbalance() {
  await waitForPlotly();
  if (!amcMerged.length) return;

  const threshEl = document.getElementById("amc-imb-threshold");
  const typeEl   = document.getElementById("amc-imb-type");
  const personEl = document.getElementById("amc-imb-person");
  const threshold = threshEl ? Number(threshEl.value) : 10;
  const typeVal   = typeEl   ? typeEl.value.trim()    : "";
  const personVal = personEl ? personEl.value.trim()  : "";
  populatePersonSelect("amc-imb-person");

  let rows = getAmcFilteredRows(typeVal, personVal);

  // Compute imbalance metrics per row
  const scored = rows.map(r => {
    const vals = amcPlants.map(p => r.amcs[p]);
    const stocked = vals.filter(v => v !== null);
    if (stocked.length < 2) return null;

    const max  = Math.max(...stocked);
    const mean = stocked.reduce((s,v)=>s+v,0) / stocked.length;
    const zerosAmongStocked = stocked.filter(v=>v===0).length;
    const naCount = vals.filter(v=>v===null).length;

    // Gini coefficient
    const gini = amcGini(r);
    const ratio = mean > 0 ? max / mean : 0;

    if (ratio < threshold) return null;

    // Find top plant and zero-stock plants
    let topPlant = "", topVal = -Infinity;
    const zeroPlants = [];
    for (const p of amcPlants) {
      const v = r.amcs[p];
      if (v !== null && v > topVal) { topVal = v; topPlant = p; }
      if (v === 0) zeroPlants.push(p);
    }

    return {
      ...r,
      _stocked:     stocked.length,
      _naCount:     naCount,
      _max:         max,
      _mean:        mean,
      _ratio:       ratio,
      _gini:        gini,
      _topPlant:    topPlant,
      _topVal:      topVal,
      _zeroPlants:  zeroPlants.join(", "),
      _zeroCount:   zeroPlants.length,
      _totalAMC:    amcTotalAMC(r),
    };
  }).filter(Boolean).sort((a,b) => b._gini - a._gini);

  // KPIs
  amcKpiRow("amc-imb-kpis", [
    amcKpiCard("Imbalanced Items", scored.length.toLocaleString(), `≥ ${threshold}× avg`, "red"),
    amcKpiCard("Avg Gini Coefficient", scored.length ? (scored.reduce((s,r)=>s+r._gini,0)/scored.length).toFixed(3) : "—", "0=equal · 1=monopoly", "orange"),
    amcKpiCard("Plants at Zero", [...new Set(scored.flatMap(r=>r._zeroPlants.split(", ").filter(Boolean)))].length.toLocaleString(), "unique plants w/ zero stock", "purple"),
    amcKpiCard("Screened Items", rows.length.toLocaleString(), `${typeVal||"All types"}`, "blue"),
  ]);

  // Chart: top 20 by Gini
  const top20 = scored.slice(0, 20);
  Plotly.newPlot("chart-amc-imbalance", [{
    type: "bar",
    orientation: "h",
    x: top20.map(r => r._gini).reverse(),
    y: top20.map(r => (r.desc.length > 40 ? r.desc.substring(0,40)+"…" : r.desc)).reverse(),
    marker: {
      color: top20.map(r => r._gini > 0.7 ? "#f85149" : r._gini > 0.5 ? "#ffa657" : "#d29922").reverse(),
    },
    hovertemplate: "<b>%{y}</b><br>Gini: %{x:.3f}<extra></extra>",
    text: top20.map(r => `Gini: ${r._gini.toFixed(3)}`).reverse(),
    textposition: "outside",
    textfont: { size: 9 },
  }], {
    ...PLOTLY_LAYOUT,
    height: Math.max(300, top20.length * 22 + 80),
    margin: { l:220, r:80, t:20, b:40 },
    xaxis: { title:"Gini Coefficient (higher = more concentrated)", range:[0,1.05] },
    yaxis: { tickfont:{size:10} },
    paper_bgcolor:"rgba(0,0,0,0)", plot_bgcolor:"rgba(0,0,0,0)",
  }, PLOTLY_CONFIG);

  // Table
  const cols = [
    { key:"code", label:"Material Code",
      fmt:(v,r)=>r.isMerged
        ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-mapped-badge">MERGED</span>`
        : `<span class="col-mat-code">${escHtml(v)}</span>`,
      raw:true, cellClass:"col-mat-code-wrap" },
    { key:"desc", label:"Description", cellClass:"col-mat-desc-wrap" },
    { key:"type", label:"Type" },
    { key:"_topPlant", label:"Top Plant", fmt:(v,r)=>`<b>${escHtml(v)}</b> (${fmtQty(r._topVal)})`, raw:true },
    { key:"_ratio",    label:"Max/Avg Ratio", fmt:v=>`<b style="color:var(--red)">${Number(v).toFixed(1)}×</b>`, raw:true },
    { key:"_gini",     label:"Gini",    fmt:v=>`${Number(v).toFixed(3)}`, raw:true },
    { key:"_zeroCount",label:"Zero-Stock Plants", fmt:(v,r)=>v ? `<span style="color:var(--red);font-weight:600">${v}</span> <span style="font-size:0.77em;color:var(--muted)">(${escHtml(r._zeroPlants)})</span>` : '—', raw:true },
    { key:"_totalAMC", label:"Total AMC (Qty)", fmt:fmtQty },
  ];

  document.getElementById("amc-imb-table").innerHTML = buildTable(
    scored, cols, (row) => row._gini > 0.7 ? "row-critical" : row._gini > 0.5 ? "row-warning" : ""
  );

  const exportCols = [
    {key:"code",label:"Material Code"},{key:"desc",label:"Description"},{key:"type",label:"Type"},
    {key:"_topPlant",label:"Top Plant"},{key:"_topVal",label:"Top Plant AMC"},
    {key:"_ratio",label:"Max/Avg Ratio"},{key:"_gini",label:"Gini Coefficient"},
    {key:"_zeroPlants",label:"Zero-Stock Plants"},{key:"_totalAMC",label:"Total AMC (Qty)"},
  ];
  const dlRow = document.getElementById("amc-imb-dl-row");
  if (dlRow) {
    dlRow.innerHTML = '<button class="dl-btn">⬇ CSV</button><button class="dl-btn">⬇ Excel</button>';
    dlRow.querySelectorAll(".dl-btn")[0].onclick = () => downloadCSV(scored,   exportCols, "amc_imbalance.csv");
    dlRow.querySelectorAll(".dl-btn")[1].onclick = () => downloadExcel(scored, exportCols, "amc_imbalance.xlsx");
  }
}

// ── MOS HELPERS ───────────────────────────────────────────────────────────────
/**
 * Builds a nested lookup: materialCode → plantCode → SOH (unrestricted qty).
 * Uses the global rawDf (inventory snapshot) when available.
 * Returns an empty Map if rawDf is not loaded.
 */
function buildSohMap() {
  const map = new Map(); // materialCode → { plantCode: qty }
  if (typeof rawDf === "undefined" || !rawDf.length) return map;
  for (const row of rawDf) {
    const mat  = String(row._mappedMaterial || row["Material"] || "").trim();
    const plt  = String(row["Plant"] || "").trim().toUpperCase();
    const qty  = Number(row["Unrestricted Stock"] || 0);
    if (!mat || !plt) continue;
    if (!map.has(mat)) map.set(mat, {});
    map.get(mat)[plt] = (map.get(mat)[plt] || 0) + qty;
  }
  return map;
}

/**
 * Given SOH and AMC for a plant, return Months of Stock (MOS).
 * Returns null when AMC is zero/null (can't divide by zero).
 * Returns Infinity when SOH > 0 but AMC is 0 (infinite coverage).
 */
function calcMOS(soh, amc) {
  if (amc === null || amc === undefined) return null;
  if (amc === 0) return soh > 0 ? Infinity : null;
  return soh / amc;
}

/** Format MOS for display */
function fmtMOS(mos) {
  if (mos === null || mos === undefined) return '<span style="color:var(--muted);font-size:0.8em">N/A</span>';
  if (mos === Infinity) return '<span style="color:var(--orange)">∞</span>';
  return `<b>${Number(mos).toFixed(1)}</b> mo`;
}

/** Color a MOS value for badge display */
function mosBadgeStyle(mos) {
  if (mos === null) return "color:var(--muted)";
  if (mos === Infinity || mos > 6) return "color:var(--red);font-weight:700";
  if (mos >= 3) return "color:var(--green);font-weight:600";
  if (mos >= 1) return "color:var(--orange);font-weight:600";
  return "color:var(--red);font-weight:700";
}

// ── 4. REDISTRIBUTION SUGGESTION (MOS-BASED) ─────────────────────────────────
//
// TWO-TIER LOGIC (as Supply Specialist):
//
// TIER 1 — SURPLUS TRANSFER (MOS > 6 months at source):
//   Find all plants where MOS > 6.  For each such plant+item, suggest transfer
//   to any plant of the same item where MOS < 6.  Transfer qty = bring source
//   down to exactly 6 months; distribute among deficit plants proportionally.
//
// TIER 2 — ISOLATION FLAG (MOS 1–6, item absent elsewhere):
//   Items where every plant has 1 < MOS < 6 AND the item does not appear in
//   any other plant at all.  These need redistribution planning even though no
//   plant is in acute surplus — they are isolated and at future risk.
//
async function renderAmcRedistribution() {
  await waitForPlotly();
  if (!amcMerged.length) return;

  const sourceEl   = document.getElementById("amc-redist-source");
  const typeEl     = document.getElementById("amc-redist-type");
  const personEl   = document.getElementById("amc-redist-person");
  const mosThrEl   = document.getElementById("amc-redist-mos-threshold");
  const sourceMode = sourceEl  ? sourceEl.value       : "any";
  const typeVal    = typeEl    ? typeEl.value.trim()   : "";
  const personVal  = personEl  ? personEl.value.trim() : "";
  const MOS_HIGH   = mosThrEl  ? Number(mosThrEl.value) || 6 : 6; // surplus threshold
  const MOS_LOW    = 1;  // below this = critical shortage, not redistribution target
  populatePersonSelect("amc-redist-person");

  let rows = getAmcFilteredRows(typeVal, personVal);

  // ── Build SOH lookup from inventory file (rawDf) ─────────────────────────
  const sohMap     = buildSohMap();
  const hasSoh     = sohMap.size > 0;

  // ── Plant code → plant name mapping from rawDf ────────────────────────────
  const plantNames = {}; // plantCode → Plant Name
  if (hasSoh && typeof rawDf !== "undefined") {
    for (const row of rawDf) {
      const code = String(row["Plant"] || "").trim().toUpperCase();
      const name = String(row["Plant Name"] || "").trim();
      if (code && name && !plantNames[code]) plantNames[code] = name;
    }
  }
  const pLabel = p => plantNames[p] ? `${p} (${plantNames[p]})` : p;

  // ── TIER 1: SURPLUS TRANSFER ──────────────────────────────────────────────
  const surplusSuggestions = [];

  for (const r of rows) {
    const stockedPlants = amcPlants.filter(p => r.amcs[p] !== null && r.amcs[p] > 0);
    if (!stockedPlants.length) continue;

    // Per-plant MOS
    const plantMOS = stockedPlants.map(p => {
      const amc = r.amcs[p];           // monthly consumption at plant
      const soh = hasSoh
        ? (sohMap.get(r.code)?.[p] ?? sohMap.get(r.origCode)?.[p] ?? 0)
        : null;
      const mos = hasSoh ? calcMOS(soh, amc) : null;
      return { plant: p, amc, soh, mos };
    });

    // When no SOH file — fall back to AMC-only logic (original behaviour):
    // treat plants with AMC > 2× mean as "surplus" and AMC = 0 plants as "deficit"
    if (!hasSoh) {
      const vals = stockedPlants.map(p => ({ plant:p, amc: r.amcs[p] || 0 }));
      const mean = vals.reduce((s,v)=>s+v.amc,0) / vals.length;
      if (!mean) continue;
      const zeroPlantsInAMC = amcPlants.filter(p => r.amcs[p] === 0).map(p => p);
      const surplusPlants   = vals.filter(v => v.amc > mean * 2);
      if (!surplusPlants.length || !zeroPlantsInAMC.length) continue;
      if (sourceMode === "HO01") {
        const ho01 = surplusPlants.find(v => v.plant === "HO01");
        if (!ho01) continue;
        surplusPlants.length = 0; surplusPlants.push(ho01);
      }
      const topSource = surplusPlants.sort((a,b)=>b.amc-a.amc)[0];
      const excess    = topSource.amc - mean * 1.5;
      if (excess <= 0) continue;
      const perRecipient = excess / zeroPlantsInAMC.length;
      surplusSuggestions.push({
        _tier:        1,
        code:         r.code,
        desc:         r.desc,
        type:         r.type,
        person:       r.person,
        isMerged:     r.isMerged,
        origCodes:    r.origCodes,
        _sourceP:     topSource.plant,
        _sourceMOS:   null,
        _sourceSOH:   null,
        _sourceAMC:   topSource.amc,
        _transferQty: null,
        _transferVal: excess,
        _targets:     zeroPlantsInAMC.join(", "),
        _targetCount: zeroPlantsInAMC.length,
        _perTarget:   perRecipient,
        _action:      "TRANSFER",
        _priority:    excess / mean,
        _mosDetail:   "",
      });
      continue;
    }

    // ── SOH-aware MOS path ────────────────────────────────────────────────
    // Source = plants where MOS > MOS_HIGH
    const surplusEntries = plantMOS.filter(e =>
      e.mos !== null && e.mos !== Infinity && e.mos > MOS_HIGH &&
      (sourceMode !== "HO01" || e.plant === "HO01")
    );
    // Deficit = plants of same item with MOS < MOS_HIGH (but > 0 so they consume)
    const deficitEntries = plantMOS.filter(e =>
      e.mos !== null && e.mos < MOS_HIGH && e.mos >= 0 && e.amc > 0
    );

    if (!surplusEntries.length || !deficitEntries.length) continue;

    // Pick the most surplus source
    surplusEntries.sort((a,b) => b.mos - a.mos);
    const src = surplusEntries[0];

    // Transfer qty = bring source down to exactly MOS_HIGH months
    const transferQty = Math.max(0, src.soh - MOS_HIGH * src.amc);
    if (transferQty <= 0) continue;

    // Distribute proportionally to deficit plants by their AMC (higher AMC = more urgent)
    const totalDeficitAMC = deficitEntries.reduce((s,e)=>s+e.amc, 0);
    const targetLines = deficitEntries.map(e => ({
      plant:    e.plant,
      curMOS:   e.mos,
      amc:      e.amc,
      soh:      e.soh,
      allocQty: totalDeficitAMC > 0 ? (e.amc / totalDeficitAMC) * transferQty : transferQty / deficitEntries.length,
    }));

    // Transfer value proxy: qty × monthly consumption (quantity-based, not monetary)
    const mosDetailParts = plantMOS.map(e =>
      `${e.plant}: ${e.mos === null ? "N/A" : e.mos === Infinity ? "∞" : e.mos.toFixed(1)}mo`
    );

    surplusSuggestions.push({
      _tier:        1,
      code:         r.code,
      desc:         r.desc,
      type:         r.type,
      person:       r.person,
      isMerged:     r.isMerged,
      origCodes:    r.origCodes,
      _sourceP:     src.plant,
      _sourceMOS:   src.mos,
      _sourceSOH:   src.soh,
      _sourceAMC:   src.amc,
      _transferQty: transferQty,
      _transferVal: transferQty * src.amc, // proxy: qty × monthly consumption
      _targets:     deficitEntries.map(e=>e.plant).join(", "),
      _targetLines: targetLines,
      _targetCount: deficitEntries.length,
      _perTarget:   transferQty / deficitEntries.length,
      _action:      "TRANSFER",
      _priority:    src.mos / MOS_HIGH, // how many times over the threshold
      _mosDetail:   mosDetailParts.join(" | "),
    });
  }

  surplusSuggestions.sort((a,b) => b._priority - a._priority);

  // ── TIER 2: ISOLATION FLAG (1 < MOS < 6, item absent from other plants) ──
  const isolatedSuggestions = [];

  for (const r of rows) {
    // Count plants where this item exists in AMC (has non-null AMC > 0)
    const presentIn = amcPlants.filter(p => r.amcs[p] !== null && r.amcs[p] > 0);
    if (presentIn.length === 0) continue;

    // Skip if already flagged in Tier 1 (has a surplus source)
    const alreadyFlagged = surplusSuggestions.some(s => s.code === r.code);
    if (alreadyFlagged) continue;

    // All present plants must have MOS between MOS_LOW and MOS_HIGH
    if (!hasSoh) continue; // requires SOH data for MOS

    const plantMOS = presentIn.map(p => {
      const amc = r.amcs[p];
      const soh = sohMap.get(r.code)?.[p] ?? sohMap.get(r.origCode)?.[p] ?? 0;
      const mos = calcMOS(soh, amc);
      return { plant: p, amc, soh, mos };
    });

    // All plants must be in the 1–6 range (not critical, not surplus)
    const inRange = plantMOS.every(e =>
      e.mos !== null && e.mos > MOS_LOW && e.mos <= MOS_HIGH
    );
    if (!inRange) continue;

    // Item must NOT appear in any other plant (isolated — not distributed)
    if (presentIn.length > 1) continue; // already multi-plant, skip isolation flag

    const onlyPlant = plantMOS[0];
    const mosDetailParts = plantMOS.map(e =>
      `${e.plant}: ${e.mos === null ? "N/A" : e.mos.toFixed(1)}mo`
    );

    isolatedSuggestions.push({
      _tier:        2,
      code:         r.code,
      desc:         r.desc,
      type:         r.type,
      person:       r.person,
      isMerged:     r.isMerged,
      origCodes:    r.origCodes,
      _sourceP:     onlyPlant.plant,
      _sourceMOS:   onlyPlant.mos,
      _sourceSOH:   onlyPlant.soh,
      _sourceAMC:   onlyPlant.amc,
      _transferQty: null,
      _transferVal: null,
      _targets:     "(No other plant carries this item)",
      _targetLines: [],
      _targetCount: 0,
      _perTarget:   0,
      _action:      "PLAN REDISTRIBUTION",
      _priority:    0,
      _mosDetail:   mosDetailParts.join(" | "),
    });
  }

  // ── COMBINED TABLE ────────────────────────────────────────────────────────
  const allSuggestions = [...surplusSuggestions, ...isolatedSuggestions];

  // ── KPIs ──────────────────────────────────────────────────────────────────
  const totalTransferQty = surplusSuggestions.reduce((s,r)=>s+(r._transferQty||0),0);
  const totalTransferVal = surplusSuggestions.reduce((s,r)=>s+(r._transferVal||0),0);
  amcKpiRow("amc-redist-kpis", [
    amcKpiCard(
      "Surplus Transfer Items",
      surplusSuggestions.length.toLocaleString(),
      `MOS > ${MOS_HIGH} mo at source plant`,
      "red"
    ),
    amcKpiCard(
      "Est. Transfer Qty",
      hasSoh ? fmtQty(totalTransferQty) : "—",
      hasSoh ? "Units to move from surplus plants" : "Load inventory file for qty",
      "orange"
    ),
    amcKpiCard(
      "Isolated Items (Plan Needed)",
      isolatedSuggestions.length.toLocaleString(),
      `1–${MOS_HIGH} mo, single-plant only`,
      "purple"
    ),
    amcKpiCard(
      "Unique Source Plants",
      [...new Set(surplusSuggestions.map(r=>r._sourceP))].length.toLocaleString(),
      "Plants with >6 mo stock",
      "blue"
    ),
  ]);

  // ── CHART: MOS Distribution of surplus sources (Tier 1 only) ─────────────
  const top20 = surplusSuggestions.slice(0, 20);
  if (top20.length) {
    Plotly.newPlot("chart-amc-redist", [
      {
        type: "bar",
        orientation: "h",
        name: "Source MOS (months)",
        x: top20.map(r => hasSoh ? r._sourceMOS : r._sourceAMC).reverse(),
        y: top20.map(r => (r.desc.length > 38 ? r.desc.substring(0,38)+"…" : r.desc)).reverse(),
        marker: {
          color: top20.map(r => {
            const v = hasSoh ? r._sourceMOS : r._priority;
            return v > 12 ? "#f85149" : v > 9 ? "#ffa657" : "#3a8fd4";
          }).reverse(),
        },
        hovertemplate: hasSoh
          ? "<b>%{y}</b><br>Source MOS: %{x:.1f} months<extra></extra>"
          : "<b>%{y}</b><br>Excess AMC: %{x:,.0f} units<extra></extra>",
        text: top20.map(r => hasSoh
          ? `${r._sourceMOS.toFixed(1)} mo @ ${r._sourceP}`
          : fmtQty(r._sourceAMC)
        ).reverse(),
        textposition: "outside",
        textfont:{ size:9 },
      },
    ], {
      ...PLOTLY_LAYOUT,
      height: Math.max(300, top20.length * 24 + 80),
      margin: { l:240, r:120, t:20, b:40 },
      xaxis: { title: hasSoh ? "Source Plant MOS (months)" : "Excess AMC Value (Qty)" },
      yaxis: { tickfont:{size:10} },
      shapes: hasSoh ? [{
        type:"line", x0: MOS_HIGH, x1: MOS_HIGH,
        y0:-0.5, y1: top20.length - 0.5,
        line:{ color:"#f85149", width:2, dash:"dot" },
      }] : [],
      annotations: hasSoh ? [{
        x: MOS_HIGH, y: top20.length - 0.5, xanchor:"left",
        text: `${MOS_HIGH}mo threshold`, showarrow:false,
        font:{ color:"#f85149", size:10 },
      }] : [],
      paper_bgcolor:"rgba(0,0,0,0)", plot_bgcolor:"rgba(0,0,0,0)",
    }, PLOTLY_CONFIG);
  } else {
    document.getElementById("chart-amc-redist").innerHTML =
      '<div class="alert-info" style="margin:1rem 0">✓ No surplus transfers required for current filters.</div>';
  }

  // ── TIER 1 TABLE ──────────────────────────────────────────────────────────
  const tier1Cols = [
    { key:"code", label:"Material Code",
      fmt:(v,r)=>r.isMerged
        ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-mapped-badge">MERGED</span>`
        : `<span class="col-mat-code">${escHtml(v)}</span>`,
      raw:true, cellClass:"col-mat-code-wrap" },
    { key:"desc",        label:"Description", cellClass:"col-mat-desc-wrap" },
    { key:"type",        label:"Type" },
    { key:"person",      label:"Specialist",
      fmt:v=>v?`<span style="font-size:0.78em;color:var(--muted)">👤 ${escHtml(v)}</span>`:"",
      raw:true },
    { key:"_sourceP",    label:"Source Plant",
      fmt:(v,r)=>{
        const mosStr = r._sourceMOS !== null ? `<span style="color:var(--red);font-weight:700">${r._sourceMOS.toFixed(1)} mo</span>` : "";
        const sohStr = r._sourceSOH !== null ? `<span style="color:var(--muted);font-size:0.78em"> · SOH: ${fmtQty(r._sourceSOH)}</span>` : "";
        const amcStr = `<span style="color:var(--muted);font-size:0.78em"> · AMC: ${fmtQty(r._sourceAMC)} units</span>`;
        return `<b style="color:var(--orange)">${escHtml(v)}</b> ${mosStr}${sohStr}${amcStr}`;
      }, raw:true },
    { key:"_transferQty", label:"Transfer Qty",
      fmt:(v,r)=> v !== null
        ? `<b style="color:var(--blue)">${fmtQty(v)}</b><span style="color:var(--muted);font-size:0.78em"> units</span>`
        : `<span style="color:var(--muted);font-size:0.78em">${fmtQty(r._transferVal)} units AMC-equiv</span>`,
      raw:true },
    { key:"_targets",    label:"Recipient Plants",
      fmt:(v,r)=>{
        if (!r._targetLines || !r._targetLines.length) {
          return `<span style="color:var(--green);font-size:0.8em">${escHtml(v)}</span>`;
        }
        const lines = r._targetLines.map(t =>
          `<span style="display:inline-block;margin:1px 4px 1px 0;padding:1px 6px;border-radius:4px;background:var(--card-bg);border:1px solid var(--border);font-size:0.75em">
            <b>${escHtml(t.plant)}</b>
            <span style="color:var(--orange)">${t.curMOS !== null ? t.curMOS.toFixed(1)+"mo" : ""}</span>
            <span style="color:var(--blue)">→ +${fmtQty(t.allocQty)}</span>
          </span>`
        ).join("");
        return lines;
      }, raw:true },
    { key:"_mosDetail", label:"All-Plant MOS",
      fmt:v=>v ? `<span style="font-size:0.73em;color:var(--muted)">${escHtml(v)}</span>` : "", raw:true },
  ];

  const tier1El = document.getElementById("amc-redist-table-t1");
  if (tier1El) {
    tier1El.innerHTML = surplusSuggestions.length
      ? buildTable(surplusSuggestions, tier1Cols, row =>
          row._sourceMOS !== null && row._sourceMOS > 12 ? "row-critical"
          : row._sourceMOS !== null && row._sourceMOS > 9  ? "row-warning"
          : row._priority > 5 ? "row-critical"
          : row._priority > 3 ? "row-warning"
          : ""
        )
      : '<div class="alert-info" style="margin:0.5rem 0">✓ No surplus transfers flagged.</div>';
  }

  // ── TIER 2 TABLE ──────────────────────────────────────────────────────────
  const tier2Cols = [
    { key:"code", label:"Material Code",
      fmt:(v,r)=>r.isMerged
        ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-mapped-badge">MERGED</span>`
        : `<span class="col-mat-code">${escHtml(v)}</span>`,
      raw:true, cellClass:"col-mat-code-wrap" },
    { key:"desc",       label:"Description", cellClass:"col-mat-desc-wrap" },
    { key:"type",       label:"Type" },
    { key:"person",     label:"Specialist",
      fmt:v=>v?`<span style="font-size:0.78em;color:var(--muted)">👤 ${escHtml(v)}</span>`:"", raw:true },
    { key:"_sourceP",   label:"Only Plant",
      fmt:(v,r)=>{
        const mosStr = r._sourceMOS !== null
          ? `<span style="color:var(--orange);font-weight:600">${r._sourceMOS.toFixed(1)} mo</span>`
          : "";
        const sohStr = r._sourceSOH !== null
          ? `<span style="color:var(--muted);font-size:0.78em"> · SOH: ${fmtQty(r._sourceSOH)}</span>` : "";
        return `<b style="color:var(--blue)">${escHtml(v)}</b> ${mosStr}${sohStr}`;
      }, raw:true },
    { key:"_sourceAMC", label:"AMC (Qty)", fmt:fmtQty },
    { key:"_action",    label:"Recommendation",
      fmt:()=>`<span style="padding:2px 8px;border-radius:4px;background:rgba(139,99,204,0.15);color:var(--purple);font-size:0.78em;font-weight:600">PLAN REDISTRIBUTION</span>`,
      raw:true },
    { key:"_mosDetail", label:"MOS Detail",
      fmt:v=>v?`<span style="font-size:0.73em;color:var(--muted)">${escHtml(v)}</span>`:"", raw:true },
  ];

  const tier2El = document.getElementById("amc-redist-table-t2");
  if (tier2El) {
    tier2El.innerHTML = hasSoh
      ? (isolatedSuggestions.length
          ? buildTable(isolatedSuggestions, tier2Cols, ()=>"row-warning")
          : '<div class="alert-info" style="margin:0.5rem 0">✓ No isolated single-plant items found in the 1–6 month MOS range.</div>'
        )
      : '<div class="alert-info" style="margin:0.5rem 0">⚠️ Load the <b>Inventory snapshot</b> file to enable MOS-based isolation detection.</div>';
  }

  // ── EXPORT ────────────────────────────────────────────────────────────────
  const exportRows = allSuggestions.map(r => ({
    ...r,
    _targetLines: r._targetLines
      ? r._targetLines.map(t=>`${t.plant}(MOS:${t.curMOS!==null?t.curMOS.toFixed(1):"N/A"} +${t.allocQty.toFixed(0)}u)`).join("; ")
      : "",
  }));
  const exportCols = [
    {key:"_tier",        label:"Tier"},
    {key:"code",         label:"Material Code"},
    {key:"desc",         label:"Description"},
    {key:"type",         label:"Type"},
    {key:"person",       label:"Specialist"},
    {key:"_action",      label:"Action"},
    {key:"_sourceP",     label:"Source Plant"},
    {key:"_sourceMOS",   label:"Source MOS (months)", fmt:v=>v!==null&&v!==undefined?Number(v).toFixed(1):"N/A"},
    {key:"_sourceSOH",   label:"Source SOH (units)",  fmt:v=>v!==null&&v!==undefined?Number(v).toFixed(0):"N/A"},
    {key:"_sourceAMC",   label:"Source AMC (Qty)",    fmt:v=>v!==null&&v!==undefined?Number(v).toFixed(0):"N/A"},
    {key:"_transferQty", label:"Transfer Qty (units)", fmt:v=>v!==null&&v!==undefined?Number(v).toFixed(0):""},
    {key:"_targets",     label:"Recipient Plants"},
    {key:"_targetLines", label:"Per-Plant Allocation"},
    {key:"_mosDetail",   label:"All-Plant MOS Detail"},
  ];

  const dlRow = document.getElementById("amc-redist-dl-row");
  if (dlRow) {
    dlRow.innerHTML = '<button class="dl-btn">⬇ CSV</button><button class="dl-btn">⬇ Excel</button>';
    dlRow.querySelectorAll(".dl-btn")[0].onclick = () => downloadCSV(exportRows,   exportCols, "amc_redistribution.csv");
    dlRow.querySelectorAll(".dl-btn")[1].onclick = () => downloadExcel(exportRows, exportCols, "amc_redistribution.xlsx");
  }
}

// ── WIRE INTO PAGE_RENDERERS AND EVENT LISTENERS ───────────────────────────────
// Called after DOM is ready — extend existing PAGE_RENDERERS and event listeners.
(function wireAmcModule() {
  function extend() {
    // 1. Extend PAGE_RENDERERS
    if (typeof PAGE_RENDERERS !== "undefined") {
      PAGE_RENDERERS["amc-distribution"]  = renderAmcDistribution;
      PAGE_RENDERERS["amc-coverage"]      = renderAmcCoverage;
      PAGE_RENDERERS["amc-imbalance"]     = renderAmcImbalance;
      PAGE_RENDERERS["amc-redistribution"]= renderAmcRedistribution;
    }

    // 2. AMC nav buttons — override renderPage to allow AMC pages without rawDf
    const _origRenderPage = window.renderPage;
    window.renderPage = function(id) {
      const amcPages = ["amc-distribution","amc-coverage","amc-imbalance","amc-redistribution"];
      if (amcPages.includes(id)) {
        currentPage = id;
        document.getElementById("landingView").style.display = "none";
        document.querySelectorAll(".page").forEach(el => { el.style.display = "none"; });
        const pg = document.getElementById(`page-${id}`);
        if (pg) pg.style.display = "block";
        document.querySelectorAll(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.page === id));
        if (amcMerged.length) {
          try { PAGE_RENDERERS[id]?.(); } catch(e) { console.error(e); }
        }
        return;
      }
      _origRenderPage(id);
    };

    // 3. AMC file upload
    const amcInput = document.getElementById("amcFileInput");
    if (amcInput) {
      amcInput.addEventListener("change", e => {
        const f = e.target.files[0]; if (f) loadAMCFile(f);
        e.target.value = "";
      });
    }

    // 4. AMC filter buttons
    const amcFilterMap = {
      "amc-dist-apply":   renderAmcDistribution,
      "amc-dist-clear":   () => {
        const s = document.getElementById("amc-dist-search"); if (s) s.value="";
        const t = document.getElementById("amc-dist-type");   if (t) t.value="";
        const p = document.getElementById("amc-dist-person"); if (p) p.value="";
        renderAmcDistribution();
      },
      "amc-cov-apply":    renderAmcCoverage,
      "amc-imb-apply":    renderAmcImbalance,
      "amc-redist-apply": renderAmcRedistribution,
    };

    document.body.addEventListener("click", (e) => {
      const btn = e.target.closest("button[id]");
      if (!btn || !amcMerged.length) return;
      const fn = amcFilterMap[btn.id];
      if (fn) { e.stopPropagation(); fn(); }
    }, true); // capture phase so it fires before other listeners

    // 5. Rebuild amcMerged when mapping file reloads
    // Hook: after loadMappingFile sets mappingTable, rebuild AMC merged data
    const _origApplyMapping = window.applyMaterialMapping;
    if (_origApplyMapping) {
      window.applyMaterialMapping = function() {
        _origApplyMapping.apply(this, arguments);
        if (amcRaw.length) {
          amcMerged = buildAmcMerged();
          const amcPages = ["amc-distribution","amc-coverage","amc-imbalance","amc-redistribution"];
          if (amcPages.includes(currentPage)) {
            try { PAGE_RENDERERS[currentPage]?.(); } catch(e) {}
          }
        }
      };
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", extend);
  } else {
    // DOMContentLoaded already fired (script loaded deferred after DOM ready)
    extend();
  }
})();
