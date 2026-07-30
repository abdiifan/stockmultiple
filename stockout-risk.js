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
//
//   Within that <4mo band, rows are split by urgency (status field):
//       status = "out"  → National MOS < 1 month  → CURRENTLY STOCKED OUT
//                          (already exhausted or nearly so — not a future
//                          "risk," this needs emergency action now)
//       status = "risk" → 1 ≤ National MOS < 4     → AT RISK (forward-looking
//                          window to act before it becomes a stockout)
//       status = "ok"   → National MOS ≥ 4         → not flagged
//   atRisk (boolean) = status is "out" or "risk", i.e. MOS < 4 — kept for
//   backward compatibility with filtering/sorting logic that predates the split.
//
// SCOPE DECISIONS (confirmed)
// ---------------------------
//   - Threshold is a fixed constant (STOCKOUT_MOS_THRESHOLD = 4), not a UI input.
//   - Respects the global sidebar Person Filter (rows are still scoped by it,
//     same as every other MOS-derived page) — but the Person column itself is
//     NOT displayed on this page or in its export.
//   - ONLY these three material types are ever in scope, regardless of the
//     Type dropdown selection: ZME, ZMS, ZLC. Any other type (e.g. ZMD) is
//     excluded from this page entirely, even under "All Types."
//   - Materials with National MOS = null (no branch committed at all) or
//     = Infinity (stock but zero branch demand) are EXCLUDED from this page.
//   - No Material Group filter — removed per product decision.
//   - NO ETB value anywhere on this page, NO chart — KPIs + table only.
//   - KPI row shows a per-type at-risk breakdown (ZME / ZMS / ZLC counts)
//     instead of an average-MOS figure.
//
// Requires: script.js (fmtQty, escHtml, buildTable, wireTableExport,
//           downloadCSV, downloadExcel, PAGE_RENDERERS, renderPage,
//           currentPage, personFilter, rawDf)
//           mos.js (HUB_PLANT, mosMerged, mosPlants, buildMosSohMap,
//           computeNationalMOS, getMosFilteredRows, fmtMosVal)
// Must be loaded AFTER both script.js and mos.js.
// =============================================================================

const STOCKOUT_MOS_THRESHOLD = 4; // months — "at risk" ceiling, fixed per product decision, network-wide only
const STOCKOUT_OUT_THRESHOLD = 1; // months — below this, treated as CURRENTLY STOCKED OUT, not merely "at risk"
const STOCKOUT_ALLOWED_TYPES = new Set(["ZME", "ZMS", "ZLC"]); // page scope is fixed to these three

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
//                  expiring batch falls within the next STOCKOUT_MOS_THRESHOLD
//                  months (already-expired batches count too — they're gone).
//   adjustedMos  = (totalSoh - expiringQty) ÷ totalAmc
//                  → what National MOS becomes once that soon-to-expire
//                  stock is excluded from the count.
//   exprAdjustedRisk = true ONLY when a material is currently "ok"
//                  (MOS >= 4) but adjustedMos < STOCKOUT_MOS_THRESHOLD —
//                  i.e. looks safe today, won't be once that batch expires.
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
// status: "out"  → MOS < 1   (currently stocked out, not merely "at risk")
//         "risk" → 1 ≤ MOS < 4 (at risk of stocking out)
//         "ok"   → MOS ≥ 4   (not flagged; only ever appears when "at-risk only" is unchecked)
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

  const expiringQtyMap = buildNationalExpiringQtyMap(STOCKOUT_MOS_THRESHOLD);

  const out = [];
  for (const r of rows) {
    const nat = computeNationalMOS(r, sohMap); // from mos.js
    if (nat.mos === null || nat.mos === Infinity) continue; // no basis / no real demand

    const status = nat.mos < STOCKOUT_OUT_THRESHOLD ? "out"
                  : nat.mos < STOCKOUT_MOS_THRESHOLD ? "risk"
                  : "ok";

    // Expiry-adjusted cross-check — doesn't affect status/atRisk above.
    const rawExpiringQty = expiringQtyMap.get(r.code) || 0;
    const expiringQty     = Math.min(rawExpiringQty, nat.totalSoh); // guard vs basis mismatch
    const adjustedMos      = (expiringQty > 0 && nat.totalAmc > 0)
      ? (nat.totalSoh - expiringQty) / nat.totalAmc
      : null;
    const exprAdjustedRisk = status === "ok" && adjustedMos !== null && adjustedMos < STOCKOUT_MOS_THRESHOLD;

    out.push({
      code: r.code, desc: r.desc, type: r.type,
      isMerged: r.isMerged, origCodes: r.origCodes,
      totalSoh: nat.totalSoh, totalAmc: nat.totalAmc, mos: nat.mos,
      atRisk: nat.mos < STOCKOUT_MOS_THRESHOLD, status,
      expiringQty, adjustedMos, exprAdjustedRisk,
    });
  }
  return out;
}

// ── CLICKABLE KPI CARDS ────────────────────────────────────────────────────────
// Clicking a KPI card on this page filters the table below to just the
// materials behind that number. Clicking the same (active) card again clears
// the filter. filterKey values: "out" | "risk" | "ZME" | "ZMS" | "ZLC" |
// "exprAdj" | "all" (the "Materials Screened" card, which always resets).
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
// MOS cell text color: "out" (< 1mo) gets the strongest red, "risk" (1–4mo)
// gets standard red, "ok" (≥4mo) is left neutral.
function stkoMosCellStyle(status) {
  if (status === "out")  return "color:var(--red);font-weight:800";
  if (status === "risk") return "color:var(--red);font-weight:700";
  return "color:var(--text)";
}
function stkoStatusBadge(status) {
  if (status === "out")  return '<span class="stko-badge stko-badge-out">STOCKED OUT</span>';
  if (status === "risk") return '<span class="stko-badge stko-badge-risk">AT RISK</span>';
  return '<span class="stko-badge stko-badge-ok">OK</span>';
}
function stkoStatusLabel(status) {
  return status === "out" ? "Stocked Out" : status === "risk" ? "At Risk" : "OK";
}
// Expiry-adjusted MOS cell: shows "—" when there's no expiry basis to judge,
// the adjusted figure in neutral text when it's informational only, or a
// flagged amber badge when a currently-"ok" material would drop below the
// threshold once its soon-to-expire stock is excluded.
function stkoExprAdjCell(r) {
  if (r.adjustedMos === null) return '<span style="color:var(--muted)">—</span>';
  const style = r.exprAdjustedRisk ? "color:var(--amber);font-weight:700" : "color:var(--text)";
  const badge = r.exprAdjustedRisk
    ? ` <span class="stko-badge stko-badge-expadj" title="${fmtQty(r.expiringQty)} units expire within ${STOCKOUT_MOS_THRESHOLD}mo nationally">⚠ EXPIRY-ADJUSTED</span>`
    : "";
  return `<span style="${style}">${fmtMosVal(r.adjustedMos)}</span>${badge}`;
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
  const stockOutOnlyEl = document.getElementById("stko-stockout-only");

  const searchQ    = searchEl ? searchEl.value.trim() : "";
  const typeVal    = typeEl   ? typeEl.value.trim()   : "";
  const riskOnly   = atRiskOnly  ? atRiskOnly.checked  : true;
  const stockOutOnly = stockOutOnlyEl ? stockOutOnlyEl.checked : false;

  const snapshot = buildStockoutSnapshot(typeVal, searchQ);

  const screenedCount = snapshot.length;
  const atRiskRows     = snapshot.filter(r => r.atRisk);           // MOS < 4 (out + risk combined)
  const outRows        = snapshot.filter(r => r.status === "out"); // MOS < 1
  const riskOnlyRows   = snapshot.filter(r => r.status === "risk"); // 1 ≤ MOS < 4
  const exprAdjRows    = snapshot.filter(r => r.exprAdjustedRisk);  // "ok" today, but not once near-expiry stock excluded

  // ── Per-type breakdown (ZME / ZMS / ZLC), split by status ──────────────────
  const countByType = {
    ZME: { out: 0, risk: 0 },
    ZMS: { out: 0, risk: 0 },
    ZLC: { out: 0, risk: 0 },
  };
  atRiskRows.forEach(r => {
    const t = countByType[r.type];
    if (!t) return;
    if (r.status === "out") t.out++;
    else if (r.status === "risk") t.risk++;
  });
  const totalByType = {
    ZME: countByType.ZME.out + countByType.ZME.risk,
    ZMS: countByType.ZMS.out + countByType.ZMS.risk,
    ZLC: countByType.ZLC.out + countByType.ZLC.risk,
  };
  const TYPE_LABELS = { ZME: "Medicines", ZMS: "Medical Supplies", ZLC: "ZLC" };
  const typeSub = (t) => `${TYPE_LABELS[t]} · ${countByType[t].out.toLocaleString()} out · ${countByType[t].risk.toLocaleString()} at risk`;

  // ── KPIs ──────────────────────────────────────────────────────────────────
  stkoKpiRow([
    stkoKpiCard("Materials Screened", screenedCount.toLocaleString(), "ZME · ZMS · ZLC · National MOS only", "blue", "all"),
    stkoKpiCard(`Currently Stocked Out (<${STOCKOUT_OUT_THRESHOLD}mo)`, outRows.length.toLocaleString(), "Needs emergency action now", "red", "out"),
    stkoKpiCard(`At Risk (${STOCKOUT_OUT_THRESHOLD}–${STOCKOUT_MOS_THRESHOLD}mo)`, riskOnlyRows.length.toLocaleString(), "Window to act before it runs out", "amber", "risk"),
    stkoKpiCard("ZME Flagged", totalByType.ZME.toLocaleString(), typeSub("ZME"), "amber", "ZME"),
    stkoKpiCard("ZMS Flagged", totalByType.ZMS.toLocaleString(), typeSub("ZMS"), "purple", "ZMS"),
    stkoKpiCard("ZLC Flagged", totalByType.ZLC.toLocaleString(), typeSub("ZLC"), "blue", "ZLC"),
    stkoKpiCard("⚠ Expiry-Adjusted Risk", exprAdjRows.length.toLocaleString(), `MOS ≥ ${STOCKOUT_MOS_THRESHOLD}mo today, but drops below once near-expiry stock is excluded`, "amber", "exprAdj"),
  ]);

  // ── TABLE ──────────────────────────────────────────────────────────────────
  // "At-risk only" and "Stock out only" are independent filters over the two
  // distinct status bands, not a nested "narrow further" pair:
  //   riskOnly     → status "risk" only  (1 ≤ MOS < 4)
  //   stockOutOnly → status "out"  only  (MOS < 1)
  // Checking both shows the union of the two bands (equivalent to the old
  // atRisk flag, MOS < 4). Checking neither shows everything, "ok" included.
  let baseRows;
  if (riskOnly && stockOutOnly) {
    baseRows = atRiskRows;
  } else if (riskOnly) {
    baseRows = riskOnlyRows;
  } else if (stockOutOnly) {
    baseRows = outRows;
  } else {
    baseRows = snapshot;
  }

  // ── Apply the active KPI-card filter (if any) on top of the above ──────────
  // "all" (Materials Screened) always resets to the full snapshot, regardless
  // of the at-risk-only checkbox, since it represents everything screened.
  let cardFilteredRows = baseRows;
  let cardFilterLabel = null;
  if (stkoCardFilter === "all") {
    cardFilteredRows = snapshot;
  } else if (stkoCardFilter === "out") {
    cardFilteredRows = outRows;
    cardFilterLabel = `Currently Stocked Out (<${STOCKOUT_OUT_THRESHOLD}mo)`;
  } else if (stkoCardFilter === "risk") {
    cardFilteredRows = riskOnlyRows;
    cardFilterLabel = `At Risk (${STOCKOUT_OUT_THRESHOLD}–${STOCKOUT_MOS_THRESHOLD}mo)`;
  } else if (stkoCardFilter === "ZME" || stkoCardFilter === "ZMS" || stkoCardFilter === "ZLC") {
    cardFilteredRows = atRiskRows.filter(r => r.type === stkoCardFilter);
    cardFilterLabel = `${stkoCardFilter} Flagged (stocked out + at risk)`;
  } else if (stkoCardFilter === "exprAdj") {
    // Pull straight from the full snapshot — these rows are "ok" by status
    // and may not be in baseRows unless the expiry-adjusted checkbox is on.
    cardFilteredRows = exprAdjRows;
    cardFilterLabel = "Expiry-Adjusted Risk";
  }

  const tableRows = cardFilteredRows.slice().sort((a, b) => a.mos - b.mos); // most urgent first

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
        ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-mapped-badge" title="Merged from: ${escHtml(r.origCodes)}">MERGED</span>`
        : `<span class="col-mat-code">${escHtml(v)}</span>`,
      raw: true, cellClass: "col-mat-code-wrap" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "type", label: "Type" },
    { key: "totalSoh", label: "National SOH", fmt: fmtQty },
    { key: "totalAmc", label: "National AMC (branches)", fmt: fmtQty },
    { key: "mos", label: "National MOS",
      fmt: (v, r) => `<span style="${stkoMosCellStyle(r.status)}">${fmtMosVal(v)}</span>`, raw: true },
    { key: "adjustedMos", label: "Expiry-Adjusted MOS",
      fmt: (v, r) => stkoExprAdjCell(r), raw: true },
    { key: "status", label: "Status", fmt: (v) => stkoStatusBadge(v), raw: true },
  ];

  document.getElementById("stko-table").innerHTML = tableRows.length
    ? buildTable(tableRows, cols, (row) => row.status === "out" ? "row-stocked-out" : row.atRisk ? "row-critical" : row.exprAdjustedRisk ? "row-expiry-adjusted" : "")
    : '<div class="alert-info" style="margin:0.5rem 0">✓ No materials match the current filters at national stockout risk.</div>';

  // ── EXPORT ────────────────────────────────────────────────────────────────
  // Export columns are kept in exact 1:1 lockstep with the on-screen `cols`
  // above — same set, same order, same labels — so the download always
  // matches what the user is looking at on the page. Nothing extra
  // (no expiringQty, exprAdjustedRisk flag, or atRisk flag).
  const exportCols = [
    { key: "code", label: "Material Code" },
    { key: "desc", label: "Description" },
    { key: "type", label: "Type" },
    { key: "totalSoh", label: "National SOH", fmt: v => Number(v || 0).toFixed(2) },
    { key: "totalAmc", label: "National AMC (branches)", fmt: v => Number(v || 0).toFixed(2) },
    { key: "mos", label: "National MOS", fmt: v => Number(v).toFixed(2) },
    { key: "adjustedMos", label: "Expiry-Adjusted MOS", fmt: v => v === null ? "" : Number(v).toFixed(2) },
    { key: "status", label: "Status", fmt: v => stkoStatusLabel(v) },
  ];

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
        const s = document.getElementById("stko-search");            if (s) s.value = "";
        const t = document.getElementById("stko-type");              if (t) t.value = "";
        const c = document.getElementById("stko-at-risk-only");      if (c) c.checked = true;
        const e = document.getElementById("stko-stockout-only");     if (e) e.checked = false;
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

    const stockOutOnlyToggle = document.getElementById("stko-stockout-only");
    if (stockOutOnlyToggle) stockOutOnlyToggle.addEventListener("change", () => { if (mosMerged.length) renderStockoutRisk(); });

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
