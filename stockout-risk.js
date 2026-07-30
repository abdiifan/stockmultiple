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
//     entirely — neither case is "at risk of running out."
//   - NO ETB value anywhere on this page (no value column, no value KPI, no
//     min-value filter) — per product decision, this screen is purely about
//     months-of-stock, not monetary exposure (that's Expiry Risk's job).
//   - NO chart — KPIs + table only, per product decision.
//
// Requires: script.js (fmtQty, escHtml, buildTable, wireTableExport,
//           downloadCSV, downloadExcel, PAGE_RENDERERS, renderPage,
//           currentPage, personFilter)
//           mos.js (HUB_PLANT, mosMerged, mosPlants, buildMosSohMap,
//           computeNationalMOS, getMosFilteredRows, fmtMosVal)
// Must be loaded AFTER both script.js and mos.js.
// =============================================================================

const STOCKOUT_MOS_THRESHOLD = 4; // months — fixed per product decision, network-wide only

// ── BUILD THE NATIONAL STOCKOUT-RISK SNAPSHOT ─────────────────────────────────
// Returns an array of { code, desc, type, person, totalSoh, totalAmc, mos, atRisk }
// Only includes materials where National MOS is a real, finite number — null
// (nothing committed anywhere) and Infinity (demand-free) are dropped, per
// the confirmed scope decision above.
function buildStockoutSnapshot(typeFilter, searchQ) {
  if (typeof mosMerged === "undefined" || !mosMerged.length) return [];

  const sohMap = buildMosSohMap(); // from mos.js

  // getMosFilteredRows() already applies the global personFilter before
  // type/search, exactly like MOS by Plant / Expiry Risk.
  const rows = (typeof getMosFilteredRows === "function")
    ? getMosFilteredRows(typeFilter || "", searchQ || "")
    : mosMerged.filter(r => (!typeFilter || r.type === typeFilter));

  const out = [];
  for (const r of rows) {
    const nat = computeNationalMOS(r, sohMap); // from mos.js
    if (nat.mos === null || nat.mos === Infinity) continue; // no basis / no real demand

    out.push({
      code: r.code, desc: r.desc, type: r.type, person: r.person || "",
      isMerged: r.isMerged, origCodes: r.origCodes,
      totalSoh: nat.totalSoh, totalAmc: nat.totalAmc, mos: nat.mos,
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
function renderStockoutRisk() {
  const hasInventory = typeof rawDf !== "undefined" && rawDf.length > 0;
  const hasAmc        = typeof mosMerged !== "undefined" && mosMerged.length > 0;

  if (!hasInventory || !hasAmc) {
    // Specific message about whichever one is actually missing, instead of a
    // generic "upload both" line that doesn't tell you which is the problem.
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
  const atRiskRows    = snapshot.filter(r => r.atRisk);
  const avgAtRiskMos  = atRiskRows.length
    ? atRiskRows.reduce((s, r) => s + r.mos, 0) / atRiskRows.length
    : null;

  // ── KPIs ──────────────────────────────────────────────────────────────────
  stkoKpiRow([
    stkoKpiCard("Materials Screened", screenedCount.toLocaleString(), typeVal || "All types · National MOS only", "blue"),
    stkoKpiCard(`At Risk of Stockout (<${STOCKOUT_MOS_THRESHOLD}mo)`, atRiskRows.length.toLocaleString(), `of ${screenedCount.toLocaleString()} screened nationally`, "red"),
    stkoKpiCard("Avg. MOS (At-Risk Only)", avgAtRiskMos !== null ? `${avgAtRiskMos.toFixed(1)} mo` : "—", `Threshold: ${STOCKOUT_MOS_THRESHOLD} mo`, "purple"),
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
    { key: "person", label: "Person Assigned", fmt: v => v || "—" },
    { key: "totalSoh", label: "National SOH", fmt: fmtQty },
    { key: "totalAmc", label: "National AMC (branches)", fmt: fmtQty },
    { key: "mos", label: "National MOS",
      fmt: v => `<span style="${stkoMosCellStyle(v)}">${fmtMosVal(v)}</span>`, raw: true },
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
