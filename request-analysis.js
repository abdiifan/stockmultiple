// =============================================================================
// PharmaTrack v2 — stockout-risk.js
// "📉 Stockout Risk" — NATIONAL-LEVEL ONLY (no per-plant breakdown).
//
// RULE (per product decision)
// ----------------------------
//   Every material in scope is classified into ONE of five MOS bands, where
//   National MOS is EXACTLY mos.js's existing computeNationalMOS():
//       National MOS = (SOH at every plant, including HO01)
//                     ÷ (AMC at every BRANCH plant, excluding HO01)
//
//   status = "out"       → National MOS < 1 month        → STOCKOUT
//                           (already exhausted or nearly so — not a future
//                           "risk," this needs emergency action now)
//   status = "high"      → 1 ≤ National MOS < 3 month     → HIGH RISK
//   status = "medium"    → 3 ≤ National MOS < 6 month     → MEDIUM RISK
//   status = "optimal"   → 6 ≤ National MOS < 12 month    → OPTIMAL
//   status = "overstock" → National MOS ≥ 12 month        → OVERSTOCK
//
//   atRisk (boolean) = status is "out" or "high", i.e. MOS < 3 — kept for
//   backward compatibility with filtering/sorting logic that predates the
//   5-way split (this is the same "needs attention soon" boundary the old
//   single at-risk threshold used to represent).
//
// SCOPE DECISIONS (confirmed)
// ---------------------------
//   - Thresholds are fixed constants (see STOCKOUT_*_THRESHOLD below), not UI
//     inputs.
//   - Respects the global sidebar Person Filter (rows are still scoped by it,
//     same as every other MOS-derived page) — but the Person column itself is
//     NOT displayed on this page or in its export.
//   - ONLY these three material types are ever in scope, regardless of the
//     Type dropdown selection: ZME, ZMS, ZLC. Any other type (e.g. ZMD) is
//     excluded from this page entirely, even under "All Types."
//   - Materials with National MOS = null (no branch committed at all) or
//     = Infinity (stock but zero branch demand) are EXCLUDED from this page —
//     there's no finite MOS to place in a band.
//   - No Material Group filter — removed per product decision.
//   - NO ETB value anywhere on this page, NO chart — KPIs + table only.
//   - KPI row shows a per-type at-risk breakdown (ZME / ZMS / ZLC counts,
//     Stockout + High Risk only) instead of an average-MOS figure.
//
// Requires: script.js (fmtQty, escHtml, buildTable, wireTableExport,
//           downloadCSV, downloadExcel, PAGE_RENDERERS, renderPage,
//           currentPage, personFilter, rawDf)
//           mos.js (HUB_PLANT, mosMerged, mosPlants, buildMosSohMap,
//           computeNationalMOS, getMosFilteredRows, fmtMosVal)
// Must be loaded AFTER both script.js and mos.js.
// =============================================================================

const STOCKOUT_OUT_THRESHOLD     = 1;  // months — below this: STOCKOUT
const STOCKOUT_HIGH_THRESHOLD    = 3;  // months — below this (and >= out): HIGH RISK; also the "at risk" / expiry cross-check ceiling
const STOCKOUT_MEDIUM_THRESHOLD  = 6;  // months — below this (and >= high): MEDIUM RISK
const STOCKOUT_OPTIMAL_THRESHOLD = 12; // months — below this (and >= medium): OPTIMAL; at/above this: OVERSTOCK
const STOCKOUT_ALLOWED_TYPES = new Set(["ZME", "ZMS", "ZLC"]); // page scope is fixed to these three

// Classify a finite National MOS value into one of the five status bands.
function stkoClassifyStatus(mos) {
  if (mos < STOCKOUT_OUT_THRESHOLD)     return "out";
  if (mos < STOCKOUT_HIGH_THRESHOLD)    return "high";
  if (mos < STOCKOUT_MEDIUM_THRESHOLD)  return "medium";
  if (mos < STOCKOUT_OPTIMAL_THRESHOLD) return "optimal";
  return "overstock";
}

// ── EXPIRY-ADJUSTED RISK (extra cross-check signal, does NOT alter the core
//    National MOS rule or "confirmed" status thresholds above) ───────────────
// National MOS is pure quantity ÷ consumption — it has no idea how much of
// that SOH is about to expire. A material can show MOS >= 4 ("ok") today and
// still be quietly heading toward a stockout the moment a big batch expires.
//
// This reuses buildExpiryMap()/monthsUntil() from expiry-risk.js — the SAME
// earliest-batch-per-plant expiry basis used on the Overstock & Expiry Risk
// page — so the two pages agree on what "expiring soon" means.
//
//   expiringQty  = national SOH (all plants incl. HO01) whose earliest-
//                  expiring batch falls within the next STOCKOUT_HIGH_THRESHOLD
//                  months (already-expired batches count too — they're gone).
//   adjustedMos  = (totalSoh - expiringQty) ÷ totalAmc
//                  → what National MOS becomes once that soon-to-expire
//                  stock is excluded from the count.
//   exprAdjustedStatus = adjustedMos run through the SAME 5-tier classifier
//                  (stkoClassifyStatus) as the headline status — "Expiry-
//                  Adjusted MOS" always shows the real band, not a raw
//                  threshold check.
//   exprAdjustedRisk = true ONLY when a material is currently NOT already
//                  atRisk (status is "medium", "optimal", or "overstock")
//                  but exprAdjustedStatus is "out" or "high" — i.e. it looks
//                  safe today (Medium Risk or better) but would drop into
//                  High Risk / Stockout territory once that batch expires.
function buildNationalExpiringQtyMap(thresholdMonths) {
  const map = new Map(); // code -> national qty expiring within thresholdMonths
  if (typeof buildExpiryMap !== "function") return map; // expiry-risk.js not loaded
  const expiryMap = buildExpiryMap(); // from expiry-risk.js: code -> plant -> {expiry, qtySum, valSum}
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (const [code, plantMap] of expiryMap.entries()) {
    let expiringQty = 0;
    for (const plant in plantMap) {
      const entry = plantMap[plant];
      if (!entry.expiry) continue; // no expiry data at this plant — can't judge, skip
      const left = monthsUntil(entry.expiry, today); // from expiry-risk.js
      if (left !== null && left < thresholdMonths) expiringQty += entry.qtySum;
    }
    if (expiringQty > 0) map.set(code, expiringQty);
  }
  return map;
}

// ── BUILD THE NATIONAL STOCKOUT-RISK SNAPSHOT ─────────────────────────────────
// Returns an array of { code, desc, type, totalSoh, totalAmc, mos, atRisk, status }
// status: "out"       → MOS < 1        → STOCKOUT
//         "high"      → 1 ≤ MOS < 3    → HIGH RISK
//         "medium"    → 3 ≤ MOS < 6    → MEDIUM RISK
//         "optimal"   → 6 ≤ MOS < 12   → OPTIMAL
//         "overstock" → MOS ≥ 12       → OVERSTOCK
// Only ZME/ZMS/ZLC types are ever included, and only materials where National
// MOS is a real, finite number (null/Infinity dropped).
function buildStockoutSnapshot(typeFilter, searchQ) {
  if (typeof mosMerged === "undefined" || !mosMerged.length) return [];

  const sohMap = buildMosSohMap();          // from mos.js

  // getMosFilteredRows() already applies the global personFilter before
  // type/search, exactly like MOS by Plant / Expiry Risk.
  let rows = (typeof getMosFilteredRows === "function")
    ? getMosFilteredRows(typeFilter || "", searchQ || "")
    : mosMerged.filter(r => (!typeFilter || r.type === typeFilter));

  // Hard scope: only ZME/ZMS/ZLC, regardless of the dropdown value —
  // "All Types" on this page still means "all of ZME/ZMS/ZLC," never ZMD
  // or anything else.
  rows = rows.filter(r => STOCKOUT_ALLOWED_TYPES.has(r.type));

  const expiringQtyMap = buildNationalExpiringQtyMap(STOCKOUT_HIGH_THRESHOLD);

  const out = [];
  for (const r of rows) {
    const nat = computeNationalMOS(r, sohMap); // from mos.js
    if (nat.mos === null || nat.mos === Infinity) continue; // no basis / no real demand

    const status = stkoClassifyStatus(nat.mos);
    const atRisk = status === "out" || status === "high"; // MOS < 3

    // Expiry-adjusted cross-check — doesn't affect status/atRisk above.
    // The adjusted MOS is run through the SAME 5-tier classifier as the
    // headline status (stkoClassifyStatus), so "Expiry-Adjusted MOS" always
    // shows the real band that value falls into — never a bare threshold
    // check — and stays in lockstep with the Stockout/High/Medium/Optimal/
    // Overstock definitions if those bands ever change.
    const rawExpiringQty = expiringQtyMap.get(r.code) || 0;
    const expiringQty     = Math.min(rawExpiringQty, nat.totalSoh); // guard vs basis mismatch
    const adjustedMos      = (expiringQty > 0 && nat.totalAmc > 0)
      ? (nat.totalSoh - expiringQty) / nat.totalAmc
      : null;
    const exprAdjustedStatus = adjustedMos !== null ? stkoClassifyStatus(adjustedMos) : null;
    // Flag only the "looked safe today, expiry pushes it into Stockout/High
    // Risk" case — a material already atRisk on pure MOS doesn't need a
    // second flag here.
    const exprAdjustedRisk = !atRisk && (exprAdjustedStatus === "out" || exprAdjustedStatus === "high");

    out.push({
      code: r.code, desc: r.desc, type: r.type,
      isMerged: r.isMerged, origCodes: r.origCodes,
      totalSoh: nat.totalSoh, totalAmc: nat.totalAmc, mos: nat.mos,
      atRisk, status,
      expiringQty, adjustedMos, exprAdjustedStatus, exprAdjustedRisk,
    });
  }
  return out;
}

// ── CLICKABLE KPI CARDS ────────────────────────────────────────────────────────
// Clicking a KPI card on this page filters the table below to just the
// materials behind that number. Clicking the same (active) card again clears
// the filter. filterKey values: "out" | "high" | "medium" | "optimal" |
// "overstock" | "ZME" | "ZMS" | "ZLC" | "exprAdj" | "all" (the "Materials
// Screened" card, which always resets).
let stkoCardFilter = null;

// ── FORMATTING HELPERS ────────────────────────────────────────────────────────
function stkoKpiCard(label, value, sub, color, filterKey) {
  if (!filterKey) {
    return `<div class="kpi-card"><div class="kpi-label">${escHtml(label)}</div><div class="kpi-value" style="color:var(--${color||'blue'})">${value}</div>${sub ? `<div class="kpi-sub">${sub}</div>` : ""}</div>`;
  }
  const isActive = stkoCardFilter === filterKey;
  const activeStyle = isActive ? `border-color:var(--${color||'blue'});box-shadow:0 0 0 1px var(--${color||'blue'})` : "";
  return `<div class="kpi-card" data-stko-filter="${escHtml(filterKey)}" role="button" tabindex="0"
      style="cursor:pointer;${activeStyle}" title="Click to show these items in the table below">
    <div class="kpi-label">${escHtml(label)}</div>
    <div class="kpi-value" style="color:var(--${color||'blue'})">${value}</div>
    ${sub ? `<div class="kpi-sub">${sub}</div>` : ""}
  </div>`;
}
function stkoKpiRow(cards) {
  const el = document.getElementById("stko-kpis");
  if (el) el.innerHTML = cards.join("");
}
// MOS cell text color: "out" (<1mo) gets the strongest red, "high" (1–3mo)
// standard red, "medium" (3–6mo) amber, "optimal" (6–12mo) green,
// "overstock" (>=12mo) blue.
function stkoMosCellStyle(status) {
  if (status === "out")       return "color:var(--red);font-weight:800";
  if (status === "high")      return "color:var(--red);font-weight:700";
  if (status === "medium")    return "color:var(--amber);font-weight:700";
  if (status === "optimal")   return "color:var(--green);font-weight:600";
  if (status === "overstock") return "color:var(--blue);font-weight:700";
  return "color:var(--text)";
}
function stkoStatusBadge(status) {
  if (status === "out")       return '<span class="stko-badge stko-badge-out">STOCKOUT</span>';
  if (status === "high")      return '<span class="stko-badge stko-badge-risk">HIGH RISK</span>';
  if (status === "medium")    return '<span class="stko-badge stko-badge-medium">MEDIUM RISK</span>';
  if (status === "optimal")   return '<span class="stko-badge stko-badge-ok">OPTIMAL</span>';
  if (status === "overstock") return '<span class="stko-badge stko-badge-overstock">OVERSTOCK</span>';
  return '<span class="stko-badge">—</span>';
}
function stkoStatusLabel(status) {
  if (status === "out")       return "Stockout";
  if (status === "high")      return "High Risk";
  if (status === "medium")    return "Medium Risk";
  if (status === "optimal")   return "Optimal";
  if (status === "overstock") return "Overstock";
  return "—";
}
// Expiry-adjusted MOS cell: shows "—" when there's no expiry basis to judge.
// Otherwise the adjusted value is colored/weighted per its OWN classified
// band (exprAdjustedStatus) — same 5-tier styling as the main MOS column —
// with the matching status badge. An extra amber "⚠ EXPIRY-ADJUSTED" flare
// is appended only for the "looked safe today, expiry pushes it into risk"
// case (exprAdjustedRisk).
function stkoExprAdjCell(r) {
  if (r.adjustedMos === null) return '<span style="color:var(--muted)">—</span>';
  const style = stkoMosCellStyle(r.exprAdjustedStatus);
  const statusBadge = stkoStatusBadge(r.exprAdjustedStatus);
  const flare = r.exprAdjustedRisk
    ? ` <span class="stko-badge stko-badge-expadj" title="${fmtQty(r.expiringQty)} units expire within ${STOCKOUT_HIGH_THRESHOLD}mo nationally">⚠ EXPIRY-ADJUSTED</span>`
    : "";
  return `<span style="${style}">${fmtMosVal(r.adjustedMos)}</span> ${statusBadge}${flare}`;
}


// ── MAIN RENDER ────────────────────────────────────────────────────────────────
function renderStockoutRisk() {
  const hasInventory = typeof rawDf !== "undefined" && rawDf.length > 0;
  const hasAmc        = typeof mosMerged !== "undefined" && mosMerged.length > 0;

  if (!hasInventory || !hasAmc) {
    const noDataEl = document.getElementById("stko-no-data");
    if (noDataEl) {
      if (!hasAmc && !hasInventory) {
        noDataEl.innerHTML = 'Upload the <b>main inventory Excel</b> and <b>AMC.xlsx</b> (📐 AMC upload in the sidebar) to enable this analysis.';
      } else if (!hasAmc) {
        noDataEl.innerHTML = 'Upload <b>AMC.xlsx</b> (📐 AMC upload in the sidebar) to enable this analysis — the main inventory file is already loaded.';
      } else {
        noDataEl.innerHTML = 'Upload the <b>main inventory Excel</b> to enable this analysis — AMC.xlsx is already loaded.';
      }
      noDataEl.style.display = "block";
    }
    document.getElementById("stko-content").style.display = "none";
    return;
  }
  document.getElementById("stko-no-data").style.display  = "none";
  document.getElementById("stko-content").style.display  = "block";

  const searchEl     = document.getElementById("stko-search");
  const typeEl       = document.getElementById("stko-type");
  const atRiskOnly   = document.getElementById("stko-at-risk-only");
  const exprAdjOnly  = document.getElementById("stko-expiry-adjusted");

  const searchQ    = searchEl ? searchEl.value.trim() : "";
  const typeVal    = typeEl   ? typeEl.value.trim()   : "";
  const riskOnly   = atRiskOnly  ? atRiskOnly.checked  : true;
  const showExprAdj = exprAdjOnly ? exprAdjOnly.checked : false;

  const snapshot = buildStockoutSnapshot(typeVal, searchQ);

  const screenedCount = snapshot.length;
  const atRiskRows     = snapshot.filter(r => r.atRisk);                // MOS < 3 (out + high combined)
  const outRows        = snapshot.filter(r => r.status === "out");       // MOS < 1
  const highRows       = snapshot.filter(r => r.status === "high");      // 1 ≤ MOS < 3
  const mediumRows     = snapshot.filter(r => r.status === "medium");    // 3 ≤ MOS < 6
  const optimalRows    = snapshot.filter(r => r.status === "optimal");   // 6 ≤ MOS < 12
  const overstockRows  = snapshot.filter(r => r.status === "overstock"); // MOS ≥ 12
  const exprAdjRows    = snapshot.filter(r => r.exprAdjustedRisk);       // looks safe today, but not once near-expiry stock excluded

  // ── Per-type breakdown (ZME / ZMS / ZLC), split by status ──────────────────
  const countByType = {
    ZME: { out: 0, high: 0 },
    ZMS: { out: 0, high: 0 },
    ZLC: { out: 0, high: 0 },
  };
  atRiskRows.forEach(r => {
    const t = countByType[r.type];
    if (!t) return;
    if (r.status === "out") t.out++;
    else if (r.status === "high") t.high++;
  });
  const totalByType = {
    ZME: countByType.ZME.out + countByType.ZME.high,
    ZMS: countByType.ZMS.out + countByType.ZMS.high,
    ZLC: countByType.ZLC.out + countByType.ZLC.high,
  };
  const TYPE_LABELS = { ZME: "Medicines", ZMS: "Medical Supplies", ZLC: "ZLC" };
  const typeSub = (t) => `${TYPE_LABELS[t]} · ${countByType[t].out.toLocaleString()} stockout · ${countByType[t].high.toLocaleString()} high risk`;

  // ── KPIs ──────────────────────────────────────────────────────────────────
  stkoKpiRow([
    stkoKpiCard("Materials Screened", screenedCount.toLocaleString(), "ZME · ZMS · ZLC · National MOS only", "blue", "all"),
    stkoKpiCard(`Stockout (<${STOCKOUT_OUT_THRESHOLD}mo)`, outRows.length.toLocaleString(), "Needs emergency action now", "red", "out"),
    stkoKpiCard(`High Risk (${STOCKOUT_OUT_THRESHOLD}–${STOCKOUT_HIGH_THRESHOLD}mo)`, highRows.length.toLocaleString(), "Window to act before it runs out", "red", "high"),
    stkoKpiCard(`Medium Risk (${STOCKOUT_HIGH_THRESHOLD}–${STOCKOUT_MEDIUM_THRESHOLD}mo)`, mediumRows.length.toLocaleString(), "Worth monitoring", "amber", "medium"),
    stkoKpiCard(`Optimal (${STOCKOUT_MEDIUM_THRESHOLD}–${STOCKOUT_OPTIMAL_THRESHOLD}mo)`, optimalRows.length.toLocaleString(), "Healthy coverage band", "green", "optimal"),
    stkoKpiCard(`Overstock (≥${STOCKOUT_OPTIMAL_THRESHOLD}mo)`, overstockRows.length.toLocaleString(), "More than a year of cover on hand", "blue", "overstock"),
    stkoKpiCard("ZME Flagged", totalByType.ZME.toLocaleString(), typeSub("ZME"), "amber", "ZME"),
    stkoKpiCard("ZMS Flagged", totalByType.ZMS.toLocaleString(), typeSub("ZMS"), "purple", "ZMS"),
    stkoKpiCard("ZLC Flagged", totalByType.ZLC.toLocaleString(), typeSub("ZLC"), "blue", "ZLC"),
    stkoKpiCard("⚠ Expiry-Adjusted Risk", exprAdjRows.length.toLocaleString(), `MOS ≥ ${STOCKOUT_HIGH_THRESHOLD}mo today, but drops below once near-expiry stock is excluded`, "amber", "exprAdj"),
  ]);

  // ── TABLE ──────────────────────────────────────────────────────────────────
  // "At-risk only" scopes to MOS < 3 (Stockout + High Risk) as before. When
  // "Include Expiry-Adjusted Risk" is also checked, materials that pass the
  // pure-MOS cutoff today but are flagged by the expiry cross-check are pulled
  // into the view too (they're MOS >= 3 so atRiskOnly alone would otherwise
  // hide them). With at-risk-only unchecked, the full snapshot already
  // includes them — the checkbox has no extra effect.
  const baseRows = riskOnly
    ? (showExprAdj ? snapshot.filter(r => r.atRisk || r.exprAdjustedRisk) : atRiskRows)
    : snapshot;

  // ── Apply the active KPI-card filter (if any) on top of the above ──────────
  // "all" (Materials Screened) always resets to the full snapshot, regardless
  // of the at-risk-only checkbox, since it represents everything screened.
  let cardFilteredRows = baseRows;
  let cardFilterLabel = null;
  if (stkoCardFilter === "all") {
    cardFilteredRows = snapshot;
  } else if (stkoCardFilter === "out") {
    cardFilteredRows = outRows;
    cardFilterLabel = `Stockout (<${STOCKOUT_OUT_THRESHOLD}mo)`;
  } else if (stkoCardFilter === "high") {
    cardFilteredRows = highRows;
    cardFilterLabel = `High Risk (${STOCKOUT_OUT_THRESHOLD}–${STOCKOUT_HIGH_THRESHOLD}mo)`;
  } else if (stkoCardFilter === "medium") {
    cardFilteredRows = mediumRows;
    cardFilterLabel = `Medium Risk (${STOCKOUT_HIGH_THRESHOLD}–${STOCKOUT_MEDIUM_THRESHOLD}mo)`;
  } else if (stkoCardFilter === "optimal") {
    cardFilteredRows = optimalRows;
    cardFilterLabel = `Optimal (${STOCKOUT_MEDIUM_THRESHOLD}–${STOCKOUT_OPTIMAL_THRESHOLD}mo)`;
  } else if (stkoCardFilter === "overstock") {
    cardFilteredRows = overstockRows;
    cardFilterLabel = `Overstock (≥${STOCKOUT_OPTIMAL_THRESHOLD}mo)`;
  } else if (stkoCardFilter === "ZME" || stkoCardFilter === "ZMS" || stkoCardFilter === "ZLC") {
    cardFilteredRows = atRiskRows.filter(r => r.type === stkoCardFilter);
    cardFilterLabel = `${stkoCardFilter} Flagged (stockout + high risk)`;
  } else if (stkoCardFilter === "exprAdj") {
    // Pull straight from the full snapshot — these rows may not be in
    // baseRows unless the expiry-adjusted checkbox is on.
    cardFilteredRows = exprAdjRows;
    cardFilterLabel = "Expiry-Adjusted Risk";
  }

  const tableRows = cardFilteredRows.slice().sort((a, b) => (a.desc || "").localeCompare(b.desc || "")); // alphabetical by description, default

  const cardFilterBanner = document.getElementById("stko-card-filter-banner");
  if (cardFilterBanner) {
    cardFilterBanner.innerHTML = cardFilterLabel
      ? `<div class="alert-info" style="margin:0 0 0.5rem;display:flex;align-items:center;justify-content:space-between;gap:0.6rem">
           <span>Showing <b>${tableRows.length.toLocaleString()}</b> item${tableRows.length === 1 ? "" : "s"} for: <b>${escHtml(cardFilterLabel)}</b></span>
           <button type="button" id="stko-clear-card-filter" class="apply-btn secondary small" style="padding:0.2rem 0.6rem">✕ Clear</button>
         </div>`
      : "";
  }

  const cols = [
    { key: "code", label: "Material Code",
      fmt: (v, r) => r.isMerged
        ? `<span class="col-mat-code mat-code-clickable" data-drill-mat="${escHtml(v)}" title="Click to see Stock Concentration for this material">${escHtml(v)}</span><span class="mat-mapped-badge" title="Merged from: ${escHtml(r.origCodes)}">MERGED</span>`
        : `<span class="col-mat-code mat-code-clickable" data-drill-mat="${escHtml(v)}" title="Click to see Stock Concentration for this material">${escHtml(v)}</span>`,
      raw: true, cellClass: "col-mat-code-wrap" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "type", label: "Type" },
    { key: "totalSoh", label: "National SOH", fmt: fmtQty },
    { key: "totalAmc", label: "National AMC", fmt: fmtQty },
    { key: "mos", label: "National MOS",
      fmt: (v, r) => `<span style="${stkoMosCellStyle(r.status)}">${fmtMosVal(v)}</span>`, raw: true },
    { key: "adjustedMos", label: "Expiry-Adjusted MOS",
      fmt: (v, r) => stkoExprAdjCell(r), raw: true },
    { key: "status", label: "Status", fmt: (v) => stkoStatusBadge(v), raw: true },
  ];

  const stkoRowClass = (row) => {
    if (row.status === "out") return "row-stocked-out";
    if (row.status === "high") return "row-critical";
    if (row.exprAdjustedRisk) return "row-expiry-adjusted";
    if (row.status === "overstock") return "row-overstock";
    return "";
  };
  document.getElementById("stko-table").innerHTML = tableRows.length
    ? buildTable(tableRows, cols, stkoRowClass)
    : '<div class="alert-info" style="margin:0.5rem 0">✓ No materials match the current filters at national stockout risk.</div>';

  // ── EXPORT ────────────────────────────────────────────────────────────────
  // Export columns mirror the on-screen `cols` above. The on-screen cell
  // packs the adjusted MOS value + its classified band into one cell
  // (stkoExprAdjCell); the export keeps them as two plain columns instead,
  // since the export helper's fmt callback only ever receives the raw cell
  // value here (not the row), matching this file's existing export pattern.
  const exportCols = [
    { key: "code", label: "Material Code" },
    { key: "desc", label: "Description" },
    { key: "type", label: "Type" },
    { key: "totalSoh", label: "National SOH", fmt: v => Number(v || 0).toFixed(2) },
    { key: "totalAmc", label: "National AMC", fmt: v => Number(v || 0).toFixed(2) },
    { key: "mos", label: "National MOS", fmt: v => Number(v).toFixed(2) },
    { key: "adjustedMos", label: "Expiry-Adjusted MOS", fmt: v => v === null ? "" : Number(v).toFixed(2) },
    { key: "exprAdjustedStatus", label: "Expiry-Adjusted Status", fmt: v => v === null ? "" : stkoStatusLabel(v) },
    { key: "status", label: "Status", fmt: v => stkoStatusLabel(v) },
  ];

  const dlRow = document.getElementById("stko-dl-row");
  if (dlRow) {
    dlRow.innerHTML = '<button class="dl-btn">⬇ CSV</button><button class="dl-btn">⬇ Excel</button>';
    dlRow.querySelectorAll(".dl-btn")[0].onclick = () => downloadCSV(tableRows,   exportCols, "national_stockout_risk.csv");
    dlRow.querySelectorAll(".dl-btn")[1].onclick = () => downloadExcel(tableRows, exportCols, "national_stockout_risk.xlsx");
  }
}

// ── MATERIAL CODE SEARCH SUGGESTIONS ────────────────────────────────────────
// Lightweight autocomplete dropdown for #stko-search, styled to match the
// existing "Who's Responsible?" search (reuses its who-resp-* CSS classes —
// see index.html). Self-contained: does not depend on who-responsible.js,
// so it works even if that page/module isn't loaded.
let stkoSuggestActiveIdx = -1;
let stkoSuggestItems = [];

// Suggestion source respects the page's fixed type scope (ZME/ZMS/ZLC only),
// same as the table itself — no point suggesting a code you can't see here.
function stkoSuggestionSource() {
  if (typeof mosMerged === "undefined" || !mosMerged.length) return [];
  const seen = new Set();
  const out = [];
  for (const r of mosMerged) {
    if (!STOCKOUT_ALLOWED_TYPES.has(r.type)) continue;
    if (seen.has(r.code)) continue;
    seen.add(r.code);
    out.push({ code: r.code, desc: r.desc || "", type: r.type });
  }
  return out;
}

function stkoHighlight(text, q) {
  const s = String(text || "");
  if (!q) return escHtml(s);
  const idx = s.toLowerCase().indexOf(q.toLowerCase());
  if (idx === -1) return escHtml(s);
  return escHtml(s.slice(0, idx)) + "<mark>" + escHtml(s.slice(idx, idx + q.length)) + "</mark>" + escHtml(s.slice(idx + q.length));
}

// Suggestions box is position:fixed (see CSS), so it needs manual placement
// under the input, re-run on open/scroll/resize.
function stkoPositionSuggestions(input, box) {
  const rect = input.getBoundingClientRect();
  box.style.left  = rect.left + "px";
  box.style.top   = (rect.bottom + 4) + "px";
  box.style.width = rect.width + "px";
}

function stkoCloseSuggestions() {
  const box = document.getElementById("stko-search-suggestions");
  if (box) { box.classList.remove("open"); box.innerHTML = ""; }
  stkoSuggestActiveIdx = -1;
  stkoSuggestItems = [];
}

function stkoRenderSuggestions(query) {
  const input = document.getElementById("stko-search");
  const box   = document.getElementById("stko-search-suggestions");
  if (!input || !box) return;

  const q = query.trim();
  if (!q) { stkoCloseSuggestions(); return; }

  const ql = q.toLowerCase();
  const matches = stkoSuggestionSource().filter(m =>
    m.code.toLowerCase().includes(ql) || m.desc.toLowerCase().includes(ql)
  );
  // Codes starting with the query rank first, then description matches, each alphabetical.
  matches.sort((a, b) => {
    const aStarts = a.code.toLowerCase().startsWith(ql) ? 0 : 1;
    const bStarts = b.code.toLowerCase().startsWith(ql) ? 0 : 1;
    if (aStarts !== bStarts) return aStarts - bStarts;
    return a.code.localeCompare(b.code);
  });

  stkoSuggestItems = matches.slice(0, 25);
  stkoSuggestActiveIdx = -1;

  box.innerHTML = stkoSuggestItems.length
    ? stkoSuggestItems.map((m, i) =>
        `<div class="who-resp-item" data-idx="${i}">
           <div class="who-resp-item-code">${stkoHighlight(m.code, q)}</div>
           <div class="who-resp-item-desc">${stkoHighlight(m.desc, q)} · ${escHtml(m.type)}</div>
         </div>`
      ).join("")
    : '<div class="who-resp-empty">No matching materials</div>';

  stkoPositionSuggestions(input, box);
  box.classList.add("open");
}

function stkoSelectSuggestion(idx) {
  const item = stkoSuggestItems[idx];
  if (!item) return;
  const input = document.getElementById("stko-search");
  if (input) input.value = item.code;
  stkoCloseSuggestions();
  renderStockoutRisk();
}

function stkoSetActiveSuggestion(idx) {
  const box = document.getElementById("stko-search-suggestions");
  if (!box) return;
  const items = box.querySelectorAll(".who-resp-item");
  items.forEach(el => el.classList.remove("who-resp-active"));
  if (idx >= 0 && items[idx]) {
    items[idx].classList.add("who-resp-active");
    items[idx].scrollIntoView({ block: "nearest" });
  }
  stkoSuggestActiveIdx = idx;
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
        const s = document.getElementById("stko-search");            if (s) s.value = "";
        const t = document.getElementById("stko-type");              if (t) t.value = "";
        const c = document.getElementById("stko-at-risk-only");      if (c) c.checked = true;
        const e = document.getElementById("stko-expiry-adjusted");   if (e) e.checked = false;
        stkoCardFilter = null;
        renderStockoutRisk();
      },
    };

    document.body.addEventListener("click", (e) => {
      const btn = e.target.closest("button[id]");
      if (!btn) return;
      const fn = filterMap[btn.id];
      if (fn) { e.stopPropagation(); fn(); }
    }, true);

    // ── Click on a Stockout Risk KPI card → filter the table to those items ──
    // Clicking the already-active card (or its "✕ Clear" banner button) resets
    // the filter back to the normal search/type/checkbox view.
    function applyStkoCardFilter(el) {
      const key = el.dataset.stkoFilter;
      if (!key) return;
      stkoCardFilter = (stkoCardFilter === key) ? null : key;
      renderStockoutRisk();
      const table = document.getElementById("stko-table");
      if (table) table.scrollIntoView({ behavior: "smooth", block: "start" });
    }
    document.body.addEventListener("click", (e) => {
      if (e.target.closest("#stko-clear-card-filter")) {
        stkoCardFilter = null;
        renderStockoutRisk();
        return;
      }
      const card = e.target.closest("#stko-kpis [data-stko-filter]");
      if (card) applyStkoCardFilter(card);
    });
    document.body.addEventListener("keydown", (e) => {
      if (e.key !== "Enter" && e.key !== " ") return;
      const card = e.target.closest("#stko-kpis [data-stko-filter]");
      if (card) { e.preventDefault(); applyStkoCardFilter(card); }
    });

    const atRiskToggle = document.getElementById("stko-at-risk-only");
    if (atRiskToggle) atRiskToggle.addEventListener("change", () => { if (mosMerged.length) renderStockoutRisk(); });

    const exprAdjToggle = document.getElementById("stko-expiry-adjusted");
    if (exprAdjToggle) exprAdjToggle.addEventListener("change", () => { if (mosMerged.length) renderStockoutRisk(); });

    // Material code search: autocomplete suggestions + Enter-to-apply
    const searchInput = document.getElementById("stko-search");
    const suggestBox  = document.getElementById("stko-search-suggestions");
    if (searchInput) {
      searchInput.addEventListener("input", () => stkoRenderSuggestions(searchInput.value));
      searchInput.addEventListener("focus", () => { if (searchInput.value.trim()) stkoRenderSuggestions(searchInput.value); });
      searchInput.addEventListener("keydown", (e) => {
        const open = !!(suggestBox && suggestBox.classList.contains("open") && stkoSuggestItems.length);
        if (e.key === "ArrowDown" && open) {
          e.preventDefault();
          stkoSetActiveSuggestion(Math.min(stkoSuggestActiveIdx + 1, stkoSuggestItems.length - 1));
        } else if (e.key === "ArrowUp" && open) {
          e.preventDefault();
          stkoSetActiveSuggestion(Math.max(stkoSuggestActiveIdx - 1, 0));
        } else if (e.key === "Enter") {
          if (open && stkoSuggestActiveIdx >= 0) {
            e.preventDefault();
            stkoSelectSuggestion(stkoSuggestActiveIdx);
          } else {
            stkoCloseSuggestions();
            renderStockoutRisk();
          }
        } else if (e.key === "Escape" && open) {
          stkoCloseSuggestions();
        }
      });
    }
    if (suggestBox) {
      suggestBox.addEventListener("click", (e) => {
        const item = e.target.closest(".who-resp-item[data-idx]");
        if (item) stkoSelectSuggestion(Number(item.dataset.idx));
      });
    }
    document.addEventListener("click", (e) => {
      if (!e.target.closest(".stko-search-wrap")) stkoCloseSuggestions();
    });
    window.addEventListener("resize", () => {
      if (suggestBox && suggestBox.classList.contains("open") && searchInput) stkoPositionSuggestions(searchInput, suggestBox);
    });
    window.addEventListener("scroll", () => {
      if (suggestBox && suggestBox.classList.contains("open") && searchInput) stkoPositionSuggestions(searchInput, suggestBox);
    }, true);

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
