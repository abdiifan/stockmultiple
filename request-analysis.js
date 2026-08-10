// =============================================================================
// PharmaTrack v2 — request-analysis.js
// "🧾 Request Analysis" — self-serve sidebar tool. Any user (not just admins)
// uploads their OWN Transfer Requests Excel (Purchase Req Num, Poste, Material,
// Short Text, Requested Quantity, Stock on hand, Delivery date) and instantly
// sees it reconciled against the currently-loaded HO01 (hub) stock. Nothing is
// saved to a shared database — the uploaded file lives only in this browser
// tab's memory, exactly like the person who uploaded it intended.
//
// WHAT THIS ANALYSIS SHOWS
// -------------------------
// 1. Request vs Stock (side-by-side) — every request line, with the request
//    file's OWN "Stock on hand" column shown next to a LIVE recomputed HO01
//    stock figure (from the currently loaded main inventory), so any mismatch
//    between what the requester's system said and what HO01 actually has
//    right now is visible at a glance.
// 2. Suggested Code Corrections — request lines whose material code doesn't
//    match a SAP code directly, but resolves through the app's existing
//    Material Standardization mapping file to a target code that DOES carry
//    stock at HO01. These are "the stock is there, but requested under a
//    different code" cases — we surface the code that should actually be
//    requested.
// 3. HO01 Stockout but Requested — request lines whose resolved material has
//    ZERO stock at HO01 right now.
// 4. HO01 Stock Not Requested — every material with stock at HO01 that does
//    NOT appear anywhere in the uploaded request file at all.
//
// MATERIAL CODE MATCHING
// -----------------------
// Request file material codes (e.g. "115-ZOLE-0301-01") are NOT SAP codes.
// We resolve them the same way the rest of the app reconciles codes: via the
// existing Material Standardization mapping table (mappingTable, loaded via
// the sidebar's "⚗️ Material Standardization" upload). If a mapping entry
// exists, its target code is the canonical SAP code used for stock lookups.
// If no mapping entry exists, we fall back to trying the raw code as-is
// (covers cases where a request already used the real SAP code).
//
// Requires: script.js (rawDf, mappingTable, escHtml, fmtQty, kpiCard, buildTable,
//           wireTableExport, downloadCSV, downloadExcel, parseExpiryDate,
//           fmtLocalDate, getReconciledBase, PAGE_RENDERERS, renderPage, currentPage)
//           mos.js (HUB_PLANT, buildMosSohMap)
// Must be loaded AFTER both script.js and mos.js.
// =============================================================================

(function requestAnalysisModule() {

  // ── STATE ──────────────────────────────────────────────────────────────────
  // Lives only in memory for this browser tab/session — never written to any
  // shared store. Re-uploading replaces it; closing the tab discards it.
  let reqRows   = [];   // parsed request lines
  let reqFileName = "";

  const REQUIRED_COLS = [
    "Purchase Req Num", "Poste", "Material", "Short Text",
    "Requested Quantity", "Stock on hand", "Delivery date",
  ];

  // ── FILE PARSING ───────────────────────────────────────────────────────────
  function loadRequestFile(file) {
    const statusEl = document.getElementById("reqan-file-status");
    const btnEl    = document.getElementById("reqan-upload-btn-text");
    if (statusEl) {
      statusEl.style.display = "block";
      statusEl.innerHTML = `<div class="status-ok">⏳ LOADING…</div><div class="status-name">Parsing ${escHtml(file.name)}</div>`;
    }

    const reader = new FileReader();
    reader.onload = e => {
      setTimeout(() => {
        try {
          const wb   = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
          const ws   = wb.Sheets[wb.SheetNames[0]];
          const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
          if (!data.length) { showReqError("The uploaded file contains no data."); return; }

          const trimmed = data.map(row => {
            const r = {};
            for (const [k, v] of Object.entries(row)) r[String(k).trim()] = v;
            return r;
          });

          const colMap = {};
          Object.keys(trimmed[0]).forEach(k => { colMap[k.toLowerCase()] = k; });
          const missing = REQUIRED_COLS.filter(c => !colMap[c.toLowerCase()]);
          if (missing.length) {
            showReqError(`Missing columns: ${missing.join(", ")}. Found: ${Object.keys(trimmed[0]).join(", ")}`);
            return;
          }
          const get = (row, name) => row[colMap[name.toLowerCase()]];

          const parsed = trimmed
            .map(row => ({
              prNum:    String(get(row, "Purchase Req Num") ?? "").trim(),
              poste:    String(get(row, "Poste") ?? "").trim(),
              material: String(get(row, "Material") ?? "").trim(),
              shortText:String(get(row, "Short Text") ?? "").trim(),
              reqQty:   parseFloat(get(row, "Requested Quantity")) || 0,
              reqSoh:   parseFloat(get(row, "Stock on hand")) || 0,
              deliveryDate: (typeof parseExpiryDate === "function") ? parseExpiryDate(get(row, "Delivery date")) : null,
            }))
            .filter(r => r.material);

          if (!parsed.length) { showReqError("No valid rows with a Material code were found."); return; }

          reqRows = parsed;
          reqFileName = file.name;

          if (statusEl) {
            statusEl.innerHTML =
              `<div class="status-ok">✓ FILE LOADED</div>` +
              `<div class="status-name">${escHtml(file.name)} (${parsed.length.toLocaleString()} lines)</div>`;
          }
          if (btnEl) btnEl.textContent = "📥 Change Request File";

          const clearBtn = document.getElementById("reqan-clear-file");
          if (clearBtn) clearBtn.style.display = "inline-flex";

          if (typeof renderPage === "function") renderPage("request-analysis");
        } catch (err) {
          showReqError(`Could not read Excel file: ${err.message}`);
        }
      }, 30);
    };
    reader.readAsArrayBuffer(file);
  }

  function showReqError(msg) {
    const statusEl = document.getElementById("reqan-file-status");
    if (statusEl) {
      statusEl.style.display = "block";
      statusEl.innerHTML = `<div class="status-ok" style="color:var(--red)">✗ ${escHtml(msg)}</div>`;
    }
  }

  function clearRequestFile() {
    reqRows = [];
    reqFileName = "";
    const statusEl = document.getElementById("reqan-file-status");
    if (statusEl) { statusEl.style.display = "none"; statusEl.innerHTML = ""; }
    const btnEl = document.getElementById("reqan-upload-btn-text");
    if (btnEl) btnEl.textContent = "📥 Upload Request Excel";
    const clearBtn = document.getElementById("reqan-clear-file");
    if (clearBtn) clearBtn.style.display = "none";
    if (typeof renderPage === "function") renderPage("request-analysis");
  }

  // ── CODE RESOLUTION ────────────────────────────────────────────────────────
  // Resolves a request line's raw material code to a canonical SAP code using
  // the existing Material Standardization mapping table. Falls back to the
  // raw code (uppercased) when no mapping entry exists, in case the request
  // already used a real SAP code.
  function resolveRequestMaterial(rawCode) {
    const raw = String(rawCode || "").trim();
    if (!raw) return { canonical: "", desc: "", viaMapping: false, raw };
    const rawUpper = raw.toUpperCase();
    if (typeof mappingTable !== "undefined" && mappingTable.has(rawUpper)) {
      const entry = mappingTable.get(rawUpper);
      return { canonical: entry.targetCode, desc: entry.targetDesc || "", viaMapping: true, raw };
    }
    return { canonical: raw, desc: "", viaMapping: false, raw };
  }

  // Description lookup for canonical codes, sourced from the currently loaded
  // inventory (mapping-reconciled), so materials with no mapping-file target
  // description still show something readable.
  function buildCanonicalDescMap() {
    const out = new Map();
    const base = (typeof getReconciledBase === "function") ? getReconciledBase() : (typeof rawDf !== "undefined" ? rawDf : []);
    base.forEach(row => {
      const code = String(row._mappedMaterial || row["Material"] || "").trim();
      if (!code || out.has(code)) return;
      const desc = String(row._mappedDesc || row["Material Description"] || "").trim();
      if (desc) out.set(code, desc);
    });
    return out;
  }

  // ── CORE ANALYSIS ──────────────────────────────────────────────────────────
  function buildRequestAnalysis() {
    const hub    = (typeof HUB_PLANT === "function" || typeof HUB_PLANT !== "undefined") ? HUB_PLANT : "HO01";
    const sohMap = (typeof buildMosSohMap === "function") ? buildMosSohMap() : new Map();
    const descMap = buildCanonicalDescMap();

    const requestedCanonical = new Set();

    const rows = reqRows.map(r => {
      const resolved   = resolveRequestMaterial(r.material);
      const canonical  = resolved.canonical;
      const inInventory = !!canonical && sohMap.has(canonical);
      const liveHo01   = inInventory ? (sohMap.get(canonical)[hub] || 0) : 0;
      const desc       = r.shortText || resolved.desc || descMap.get(canonical) || "";

      let status;
      if (!canonical || !inInventory) status = "no-match";
      else if (liveHo01 <= 0)          status = "stockout";
      else                             status = "ok";

      const hasSuggestion = resolved.viaMapping && canonical && canonical.toUpperCase() !== resolved.raw.toUpperCase();
      const sohMismatch   = inInventory && Math.abs(liveHo01 - r.reqSoh) > 0.001;

      if (canonical && inInventory) requestedCanonical.add(canonical);

      return {
        ...r,
        canonical, desc, status,
        liveHo01, sohMismatch,
        hasSuggestion,
        suggestedCode: hasSuggestion ? canonical : null,
        suggestedDesc: hasSuggestion ? (resolved.desc || descMap.get(canonical) || "") : null,
      };
    });

    // HO01 stock that never shows up (under its canonical code) in the request at all
    const ho01NotRequested = [];
    sohMap.forEach((plantMap, code) => {
      const qty = plantMap[hub] || 0;
      if (qty > 0 && !requestedCanonical.has(code)) {
        ho01NotRequested.push({ code, desc: descMap.get(code) || "", qty });
      }
    });
    ho01NotRequested.sort((a, b) => b.qty - a.qty);

    return { rows, ho01NotRequested };
  }

  // ── SMALL HELPERS ──────────────────────────────────────────────────────────
  function reqStatusBadge(status) {
    const M = {
      "ok":        { bg: "rgba(48,168,95,0.14)",  color: "var(--green,#30a85f)", label: "✓ In Stock at HO01" },
      "stockout":  { bg: "rgba(220,38,38,0.14)",  color: "var(--red)",           label: "🚫 HO01 Stockout" },
      "no-match":  { bg: "rgba(120,120,120,0.14)",color: "var(--muted)",         label: "❓ No SAP Match" },
    };
    const s = M[status] || M["no-match"];
    return `<span style="display:inline-block;padding:0.15rem 0.55rem;border-radius:999px;font-size:0.72rem;font-weight:700;white-space:nowrap;background:${s.bg};color:${s.color}">${s.label}</span>`;
  }

  function fmtReqDate(d) {
    if (typeof fmtLocalDate === "function" && d instanceof Date && !isNaN(d)) return fmtLocalDate(d);
    return d instanceof Date && !isNaN(d) ? d.toLocaleDateString() : "—";
  }

  function reqKpi(label, value, sub, color) {
    return (typeof kpiCard === "function")
      ? kpiCard(label, value, sub, color)
      : `<div class="kpi-card ${color}"><div class="kpi-label">${escHtml(label)}</div><div class="kpi-value">${escHtml(String(value))}</div><div class="kpi-sub">${escHtml(sub)}</div></div>`;
  }

  // ── RENDER ─────────────────────────────────────────────────────────────────
  function renderRequestAnalysis() {
    const emptyEl   = document.getElementById("reqan-empty");
    const noInvEl   = document.getElementById("reqan-no-inventory");
    const contentEl = document.getElementById("reqan-content");
    if (!emptyEl || !contentEl) return;

    if (!reqRows.length) {
      emptyEl.style.display   = "block";
      noInvEl.style.display   = "none";
      contentEl.style.display = "none";
      return;
    }
    emptyEl.style.display = "none";

    if (typeof rawDf === "undefined" || !rawDf.length) {
      noInvEl.style.display   = "block";
      contentEl.style.display = "none";
      return;
    }
    noInvEl.style.display   = "none";
    contentEl.style.display = "block";

    const searchEl = document.getElementById("reqan-search");
    const statusEl = document.getElementById("reqan-status-filter");
    const searchQ  = searchEl ? searchEl.value.trim().toLowerCase() : "";
    const statusF  = statusEl ? statusEl.value : "";

    const { rows, ho01NotRequested } = buildRequestAnalysis();

    const matches = r => {
      if (!searchQ) return true;
      return r.material.toLowerCase().includes(searchQ)
          || (r.canonical || "").toLowerCase().includes(searchQ)
          || (r.desc || "").toLowerCase().includes(searchQ);
    };

    let filteredRows = rows.filter(matches);
    if (statusF) filteredRows = filteredRows.filter(r => r.status === statusF);

    const suggestionRows = rows.filter(r => r.hasSuggestion && matches(r));
    const stockoutRows   = rows.filter(r => r.status === "stockout" && matches(r));
    const notRequested   = ho01NotRequested.filter(r =>
      !searchQ || r.code.toLowerCase().includes(searchQ) || (r.desc || "").toLowerCase().includes(searchQ)
    );

    // ── KPIs ─────────────────────────────────────────────────────────────────
    const matchedCount  = rows.filter(r => r.status !== "no-match").length;
    document.getElementById("reqan-kpis").innerHTML = [
      reqKpi("Request Lines Uploaded", rows.length.toLocaleString(), reqFileName || "", "blue"),
      reqKpi("Matched to SAP Stock", `${matchedCount.toLocaleString()} / ${rows.length.toLocaleString()}`, "Resolved via SAP code or mapping", "green"),
      reqKpi("HO01 Stockout (Requested)", rows.filter(r => r.status === "stockout").length.toLocaleString(), "Zero HO01 stock right now", "red"),
      reqKpi("Suggested Code Corrections", rows.filter(r => r.hasSuggestion).length.toLocaleString(), "Stock exists under a different code", "amber"),
      reqKpi("HO01 Stock Not Requested", ho01NotRequested.length.toLocaleString(), "Materials at HO01, absent from this request", "purple"),
    ].join("");

    if (!(typeof mappingTable !== "undefined" && mappingTable.size > 0)) {
      document.getElementById("reqan-mapping-banner").innerHTML =
        `<div class="alert-warning" style="margin-bottom:0.8rem;font-size:0.78rem">⚠️ No Material Standardization mapping file is loaded — code-mismatch suggestions can't be computed, and any request material code that isn't already an exact SAP code will show as "No SAP Match".</div>`;
    } else {
      document.getElementById("reqan-mapping-banner").innerHTML = "";
    }

    // ── TAB 1: Request vs Stock (side-by-side) ─────────────────────────────
    const cols1 = [
      { key: "prNum", label: "PR Num" },
      { key: "poste", label: "Poste" },
      { key: "material", label: "Requested Code",
        fmt: (v, r) => r.hasSuggestion
          ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-desc-badge" title="A different standardized code carries this stock — see Suggested Code Corrections tab">≠ CODE</span>`
          : `<span class="col-mat-code">${escHtml(v)}</span>`,
        raw: true, cellClass: "col-mat-code-wrap" },
      { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
      { key: "reqQty", label: "Requested Qty", fmt: v => fmtQty(v), cellClass: "col-qty" },
      { key: "reqSoh", label: "SOH (per Request File)", fmt: v => fmtQty(v), cellClass: "col-qty" },
      { key: "liveHo01", label: "SOH (Live, HO01)",
        fmt: (v, r) => r.sohMismatch ? `<b style="color:var(--amber)">${fmtQty(v)}</b>` : fmtQty(v),
        raw: true, cellClass: "col-qty" },
      { key: "deliveryDate", label: "Delivery Date", fmt: v => fmtReqDate(v) },
      { key: "status", label: "Status", fmt: v => reqStatusBadge(v), raw: true },
    ];
    document.getElementById("reqan-table-all").innerHTML = buildTable(
      filteredRows, cols1,
      (row) => row.status === "stockout" ? "row-red" : (row.hasSuggestion ? "row-amber" : ""),
      "", { id: "reqan-export-all", title: "" }
    );
    wireTableExport("reqan-export-all", filteredRows.map(r => ({
      prNum: r.prNum, poste: r.poste, material: r.material, canonical: r.canonical, desc: r.desc,
      reqQty: r.reqQty, reqSoh: r.reqSoh, liveHo01: r.liveHo01,
      deliveryDate: fmtReqDate(r.deliveryDate), status: r.status,
    })), [
      { key: "prNum", label: "Purchase Req Num" }, { key: "poste", label: "Poste" },
      { key: "material", label: "Requested Code" }, { key: "canonical", label: "Resolved SAP Code" },
      { key: "desc", label: "Description" }, { key: "reqQty", label: "Requested Quantity" },
      { key: "reqSoh", label: "Stock on Hand (Request File)" }, { key: "liveHo01", label: "Stock on Hand (Live, HO01)" },
      { key: "deliveryDate", label: "Delivery Date" }, { key: "status", label: "Status" },
    ], "request_analysis_all_lines");

    // ── TAB 2: Suggested Code Corrections ───────────────────────────────────
    const cols2 = [
      { key: "prNum", label: "PR Num" },
      { key: "material", label: "Code As Requested", cellClass: "col-mat-code-wrap" },
      { key: "shortText", label: "Description (as requested)" },
      { key: "suggestedCode", label: "Suggested Standard Code",
        fmt: v => `<span class="col-mat-code mat-code-clickable">${escHtml(v)}</span><span class="mat-mapped-badge">STD</span>`,
        raw: true, cellClass: "col-mat-code-wrap" },
      { key: "suggestedDesc", label: "Standard Description" },
      { key: "liveHo01", label: "HO01 Stock Under Suggested Code", fmt: v => fmtQty(v), cellClass: "col-qty" },
      { key: "reqQty", label: "Requested Qty", fmt: v => fmtQty(v), cellClass: "col-qty" },
    ];
    document.getElementById("reqan-table-suggest").innerHTML = buildTable(
      suggestionRows, cols2, () => "row-amber", "", { id: "reqan-export-suggest", title: "" }
    );
    wireTableExport("reqan-export-suggest", suggestionRows.map(r => ({
      prNum: r.prNum, material: r.material, shortText: r.shortText,
      suggestedCode: r.suggestedCode, suggestedDesc: r.suggestedDesc, liveHo01: r.liveHo01, reqQty: r.reqQty,
    })), [
      { key: "prNum", label: "Purchase Req Num" }, { key: "material", label: "Code As Requested" },
      { key: "shortText", label: "Description (as requested)" }, { key: "suggestedCode", label: "Suggested Standard Code" },
      { key: "suggestedDesc", label: "Standard Description" }, { key: "liveHo01", label: "HO01 Stock Under Suggested Code" },
      { key: "reqQty", label: "Requested Qty" },
    ], "request_analysis_suggested_codes");

    // ── TAB 3: HO01 Stockout but Requested ──────────────────────────────────
    const cols3 = [
      { key: "prNum", label: "PR Num" },
      { key: "poste", label: "Poste" },
      { key: "canonical", label: "Material Code", cellClass: "col-mat-code-wrap" },
      { key: "desc", label: "Description" },
      { key: "reqQty", label: "Requested Qty", fmt: v => fmtQty(v), cellClass: "col-qty" },
      { key: "deliveryDate", label: "Delivery Date", fmt: v => fmtReqDate(v) },
    ];
    document.getElementById("reqan-table-stockout").innerHTML = buildTable(
      stockoutRows, cols3, () => "row-red", "", { id: "reqan-export-stockout", title: "" }
    );
    wireTableExport("reqan-export-stockout", stockoutRows.map(r => ({
      prNum: r.prNum, poste: r.poste, canonical: r.canonical, desc: r.desc, reqQty: r.reqQty, deliveryDate: fmtReqDate(r.deliveryDate),
    })), [
      { key: "prNum", label: "Purchase Req Num" }, { key: "poste", label: "Poste" },
      { key: "canonical", label: "Material Code" }, { key: "desc", label: "Description" },
      { key: "reqQty", label: "Requested Qty" }, { key: "deliveryDate", label: "Delivery Date" },
    ], "request_analysis_ho01_stockout");

    // ── TAB 4: HO01 Stock Not Requested ─────────────────────────────────────
    const cols4 = [
      { key: "code", label: "Material Code", cellClass: "col-mat-code-wrap" },
      { key: "desc", label: "Description" },
      { key: "qty", label: "HO01 Stock on Hand", fmt: v => fmtQty(v), cellClass: "col-qty" },
    ];
    document.getElementById("reqan-table-notreq").innerHTML = buildTable(
      notRequested, cols4, () => "", "", { id: "reqan-export-notreq", title: "" }
    );
    wireTableExport("reqan-export-notreq", notRequested, [
      { key: "code", label: "Material Code" }, { key: "desc", label: "Description" }, { key: "qty", label: "HO01 Stock on Hand" },
    ], "request_analysis_ho01_not_requested");

    // ── Tab counts (badges in tab labels) ───────────────────────────────────
    setTabCount("reqan-tab-count-all", filteredRows.length);
    setTabCount("reqan-tab-count-suggest", suggestionRows.length);
    setTabCount("reqan-tab-count-stockout", stockoutRows.length);
    setTabCount("reqan-tab-count-notreq", notRequested.length);
  }

  function setTabCount(id, n) {
    const el = document.getElementById(id);
    if (el) el.textContent = n.toLocaleString();
  }

  // ── TAB SWITCHING ──────────────────────────────────────────────────────────
  function switchTab(tab) {
    document.querySelectorAll(".reqan-tab-btn").forEach(b => {
      b.classList.toggle("active", b.dataset.tab === tab);
    });
    document.querySelectorAll(".reqan-tab-panel").forEach(p => {
      p.style.display = p.id === `reqan-tab-${tab}` ? "block" : "none";
    });
  }

  // ── WIRING ─────────────────────────────────────────────────────────────────
  function wire() {
    const fileInput = document.getElementById("requestAnalysisFileInput");
    if (fileInput) {
      fileInput.addEventListener("change", e => {
        const f = e.target.files[0];
        if (f) loadRequestFile(f);
        e.target.value = "";
      });
    }

    document.body.addEventListener("click", (e) => {
      if (e.target.closest("#reqan-clear-file")) { e.preventDefault(); clearRequestFile(); return; }

      const tabBtn = e.target.closest(".reqan-tab-btn");
      if (tabBtn) { switchTab(tabBtn.dataset.tab); return; }

      if (e.target.id === "reqan-apply") { renderRequestAnalysis(); return; }
      if (e.target.id === "reqan-clear") {
        const s = document.getElementById("reqan-search"); if (s) s.value = "";
        const st = document.getElementById("reqan-status-filter"); if (st) st.value = "";
        renderRequestAnalysis();
        return;
      }
    });

    const searchInput = document.getElementById("reqan-search");
    if (searchInput) {
      searchInput.addEventListener("keydown", (e) => { if (e.key === "Enter") renderRequestAnalysis(); });
    }

    // Re-render if already on this page when inventory or mapping data changes.
    const mainFileInput = document.getElementById("fileInput");
    if (mainFileInput) {
      mainFileInput.addEventListener("change", () => {
        setTimeout(() => { if (currentPage === "request-analysis") renderRequestAnalysis(); }, 300);
      });
    }
    const mappingInput = document.getElementById("mappingFileInput");
    if (mappingInput) {
      mappingInput.addEventListener("change", () => {
        setTimeout(() => { if (currentPage === "request-analysis") renderRequestAnalysis(); }, 300);
      });
    }

    // Register the page renderer and let it render even when no main
    // inventory has been loaded yet (renderPage() normally bails on empty
    // rawDf) — same technique used by mos.js / national-table.js.
    if (typeof PAGE_RENDERERS !== "undefined") {
      PAGE_RENDERERS["request-analysis"] = renderRequestAnalysis;
    }
    const _origRenderPage = window.renderPage;
    window.renderPage = function (id) {
      if (id === "request-analysis") {
        currentPage = id;
        document.getElementById("landingView").style.display = "none";
        document.querySelectorAll(".page").forEach(el => { el.style.display = "none"; });
        const pg = document.getElementById("page-request-analysis");
        if (pg) pg.style.display = "block";
        document.querySelectorAll(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.page === id));
        try { renderRequestAnalysis(); } catch (e) { console.error(e); }
        return;
      }
      _origRenderPage(id);
    };
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", wire);
  } else {
    wire();
  }
})();
