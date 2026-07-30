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

  const out = [];
  for (const r of rows) {
    const nat = computeNationalMOS(r, sohMap); // from mos.js
    if (nat.mos === null || nat.mos === Infinity) continue; // no basis / no real demand

    const status = nat.mos < STOCKOUT_OUT_THRESHOLD ? "out"
                  : nat.mos < STOCKOUT_MOS_THRESHOLD ? "risk"
                  : "ok";

    out.push({
      code: r.code, desc: r.desc, type: r.type,
      isMerged: r.isMerged, origCodes: r.origCodes,
      totalSoh: nat.totalSoh, totalAmc: nat.totalAmc, mos: nat.mos,
      atRisk: nat.mos < STOCKOUT_MOS_THRESHOLD, status,
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

  const searchEl   = document.getElementById("stko-search");
  const typeEl     = document.getElementById("stko-type");
  const atRiskOnly = document.getElementById("stko-at-risk-only");

  const searchQ  = searchEl ? searchEl.value.trim() : "";
  const typeVal  = typeEl   ? typeEl.value.trim()   : "";
  const riskOnly = atRiskOnly ? atRiskOnly.checked : true;

  const snapshot = buildStockoutSnapshot(typeVal, searchQ);

  const screenedCount = snapshot.length;
  const atRiskRows     = snapshot.filter(r => r.atRisk);           // MOS < 4 (out + risk combined)
  const outRows        = snapshot.filter(r => r.status === "out"); // MOS < 1
  const riskOnlyRows   = snapshot.filter(r => r.status === "risk"); // 1 ≤ MOS < 4

  // ── Per-type breakdown (ZME / ZMS / ZLC), split by status ──────────────────
  const countByType = { ZME: 0, ZMS: 0, ZLC: 0 };
  atRiskRows.forEach(r => { if (countByType[r.type] !== undefined) countByType[r.type]++; });

  // ── KPIs ──────────────────────────────────────────────────────────────────
  stkoKpiRow([
    stkoKpiCard("Materials Screened", screenedCount.toLocaleString(), "ZME · ZMS · ZLC · National MOS only", "blue"),
    stkoKpiCard(`Currently Stocked Out (<${STOCKOUT_OUT_THRESHOLD}mo)`, outRows.length.toLocaleString(), "Needs emergency action now", "red"),
    stkoKpiCard(`At Risk (${STOCKOUT_OUT_THRESHOLD}–${STOCKOUT_MOS_THRESHOLD}mo)`, riskOnlyRows.length.toLocaleString(), "Window to act before it runs out", "amber"),
    stkoKpiCard("ZME Flagged", countByType.ZME.toLocaleString(), "Medicines · stocked out + at risk", "amber"),
    stkoKpiCard("ZMS Flagged", countByType.ZMS.toLocaleString(), "Medical Supplies · stocked out + at risk", "purple"),
    stkoKpiCard("ZLC Flagged", countByType.ZLC.toLocaleString(), "ZLC · stocked out + at risk", "blue"),
  ]);

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
    { key: "totalSoh", label: "National SOH", fmt: fmtQty },
    { key: "totalAmc", label: "National AMC (branches)", fmt: fmtQty },
    { key: "mos", label: "National MOS",
      fmt: (v, r) => `<span style="${stkoMosCellStyle(r.status)}">${fmtMosVal(v)}</span>`, raw: true },
    { key: "status", label: "Status", fmt: (v) => stkoStatusBadge(v), raw: true },
  ];

  document.getElementById("stko-table").innerHTML = tableRows.length
    ? buildTable(tableRows, cols, (row) => row.status === "out" ? "row-stocked-out" : row.atRisk ? "row-critical" : "", "", { id: "stko-export", title: "" })
    : '<div class="alert-info" style="margin:0.5rem 0">✓ No materials match the current filters at national stockout risk.</div>';

  // ── EXPORT ────────────────────────────────────────────────────────────────
  const exportCols = [
    { key: "code", label: "Material Code" }, { key: "desc", label: "Description" }, { key: "type", label: "Type" },
    { key: "totalSoh", label: "National SOH", fmt: v => Number(v || 0).toFixed(2) },
    { key: "totalAmc", label: "National AMC (branches only)", fmt: v => Number(v || 0).toFixed(2) },
    { key: "mos", label: "National MOS (months)", fmt: v => Number(v).toFixed(2) },
    { key: "status", label: "Status", fmt: v => stkoStatusLabel(v) },
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
