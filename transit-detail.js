// =============================================================================
// PharmaTrack v2 — transit-detail.js
// Adds the "📄 Detail File" and "🏭 Within Plant" tabs to the Stock in Transit
// page. Lets an admin upload a separate PO-level transit detail workbook
// (Material, Plant, Purchasing Document, Item, Supplying Plant, Quantity,
// Order Quantity, Net Order Value, …) and shows it as its own table.
//
// "Within Plant" = rows where Supplying Plant equals the receiving Plant —
// i.e. an intra-plant transfer (e.g. between storage locations) that SAP
// still books through the in-transit special-stock indicator, as opposed to
// a normal inter-plant shipment between two different plants.
//
// Requires: script.js (escHtml, fmtQty, fmtETB, buildTable, downloadCSV,
//           downloadExcel, wireTableExport, setKpis, currentPage, renderPage)
// Must be loaded AFTER script.js. Does not touch renderTransit() itself —
// wraps window.renderPage so it re-renders alongside the rest of the Transit
// page, the same pattern mos.js / national-table.js / expiry-risk.js use.
// =============================================================================

(function transitDetailModule() {
  // The workbook's header row has TWO columns literally named "Currency"
  // (one for Amt.in Loc.Cur., one for Net Order Value), so we can't parse it
  // with sheet_to_json's object mode (later "Currency" would clobber the
  // earlier one). We parse by fixed column position instead.
  const COLS = [
    "material", "materialDesc", "plant", "plantName", "po", "item",
    "supplyingPlant", "specialStock", "qty", "unit", "amtLocalCur",
    "currency", "orderQty", "orderUnit", "netOrderValue", "orderCurrency",
  ];
  // Header text we expect at each position — used only to sanity-check the
  // upload and warn if someone hands us a differently-shaped file.
  const EXPECTED_HEADERS = [
    "Material", "Material Description", "Plant", "Name 1",
    "Purchasing Document", "Item", "Supplying Plant", "Special Stock",
    "Quantity", "Base Unit of Measure", "Amt.in Loc.Cur.", "Currency",
    "Order Quantity", "Order Unit", "Net Order Value", "Currency",
  ];

  // ── STATE ──────────────────────────────────────────────────────────────
  let transitDetailRaw = []; // parsed rows, one per PO line — see COLS above
  let transitDetailExcludedCount = 0; // rows dropped because they're on the

  // Same "unverified transit" exclusion list script.js already applies to
  // rawDf everywhere else (Dashboard, main Transit page, Branch Comparison,
  // etc.) — _unverifiedLookup is built there, keyed "materialCode|plantCode".
  // We reuse it here so a material/plant flagged as unverified never shows
  // up in the Detail File / Within Plant tabs either.
  function isUnverifiedTransitRow(row) {
    if (typeof _unverifiedLookup === "undefined") return false;
    return _unverifiedLookup.has(row.material + "|" + row.plant);
  }

  // Same Special Stock Type exclusion script.js applies to rawDf everywhere
  // else (Dashboard, main Transit page, Concentration, Branch) — Q and W are
  // never real on-hand/in-transit stock and must not surface anywhere.
  function isExcludedSpecialStock(row) {
    const s = String(row.specialStock || "").trim().toUpperCase();
    return s === "Q" || s === "W";
  }

  // ── FILE LOADER ────────────────────────────────────────────────────────
  function loadTransitDetailFile(file) {
    const statusEl = document.getElementById("transitDetailFileStatus");
    const btnEl    = document.getElementById("transitDetailUploadBtnText");
    if (statusEl) statusEl.innerHTML = '<div class="status-loading">⏳ Parsing…</div>';

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        // header:1 -> array-of-arrays, so duplicate "Currency" headers don't collide.
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
        if (!aoa.length) throw new Error("Transit Detail file is empty.");

        const headerRow = aoa[0].map(h => String(h ?? "").trim());
        const mismatch = EXPECTED_HEADERS.some((h, i) => (headerRow[i] || "") !== h);
        const dataRows  = aoa.slice(1).filter(r => r && r.some(v => v !== null && v !== ""));

        transitDetailRaw = dataRows.map(r => {
          const row = {};
          COLS.forEach((key, i) => { row[key] = r[i]; });
          row.material        = String(row.material ?? "").trim();
          row.materialDesc    = String(row.materialDesc ?? "").trim();
          row.plant            = String(row.plant ?? "").trim().toUpperCase();
          row.plantName        = String(row.plantName ?? "").trim();
          row.po                = String(row.po ?? "").trim();
          row.item              = String(row.item ?? "").trim();
          row.supplyingPlant    = String(row.supplyingPlant ?? "").trim().toUpperCase();
          row.specialStock      = String(row.specialStock ?? "").trim();
          row.qty                = Number(row.qty) || 0;
          row.unit               = String(row.unit ?? "").trim();
          row.amtLocalCur        = Number(row.amtLocalCur) || 0;
          row.currency            = String(row.currency ?? "").trim();
          row.orderQty             = Number(row.orderQty) || 0;
          row.orderUnit             = String(row.orderUnit ?? "").trim();
          row.netOrderValue         = Number(row.netOrderValue) || 0;
          row.orderCurrency          = String(row.orderCurrency ?? "").trim();
          row._withinPlant = row.plant && row.supplyingPlant && row.plant === row.supplyingPlant;
          return row;
        }).filter(r => r.material && r.plant);

        // Drop rows that are on the hardcoded unverified-transit list, or that
        // carry an excluded Special Stock Type (Q/W) — same exclusions
        // applied to every other transit figure in the app.
        const preExclusionCount = transitDetailRaw.length;
        transitDetailRaw = transitDetailRaw.filter(r => !isUnverifiedTransitRow(r));
        const afterUnverifiedCount = transitDetailRaw.length;
        transitDetailRaw = transitDetailRaw.filter(r => !isExcludedSpecialStock(r));
        transitDetailExcludedCount = preExclusionCount - afterUnverifiedCount;
        const specialStockExcludedCount = afterUnverifiedCount - transitDetailRaw.length;

        const count = transitDetailRaw.length;
        const withinCount = transitDetailRaw.filter(r => r._withinPlant).length;

        if (statusEl) statusEl.innerHTML =
          `<div class="status-ok">✓ LOADED</div><div class="status-name">${escHtml(file.name)}</div>` +
          `<div class="status-name" style="color:var(--green)">${count} PO line${count === 1 ? "" : "s"}` +
          (withinCount ? ` · ${withinCount} within-plant` : "") +
          (transitDetailExcludedCount ? ` · ${transitDetailExcludedCount} excluded (unverified)` : "") +
          (specialStockExcludedCount ? ` · ${specialStockExcludedCount} excluded (Q/W stock)` : "") + `</div>` +
          (mismatch ? `<div class="status-name" style="color:var(--amber)">⚠️ Column headers differ from the expected layout — check the data lines up correctly.</div>` : "");
        if (btnEl) btnEl.textContent = "✓ " + file.name;

        if (currentPage === "transit") { renderTransitDetailTab(); renderTransitWithinPlantTab(); }

      } catch (err) {
        console.error("Transit Detail load error:", err);
        if (statusEl) statusEl.innerHTML = `<div class="status-error">⚠️ ${escHtml(err.message)}</div>`;
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ── SHARED TABLE COLUMNS ───────────────────────────────────────────────
  const detailCols = [
    { key: "material",       label: "Material Code" },
    { key: "materialDesc",   label: "Material Description" },
    { key: "plantName",      label: "Receiving Plant" },
    { key: "supplyingPlant", label: "Supplying Plant" },
    { key: "po",              label: "PO Number" },
    { key: "item",             label: "Item" },
    { key: "qty",                label: "Transit Qty", fmt: fmtQty, rawKey: "qty", cellClass: "col-qty" },
    { key: "orderQty",            label: "Order Qty",   fmt: fmtQty, rawKey: "orderQty", cellClass: "col-qty" },
    { key: "netOrderValue",        label: "Net Order Value (ETB)", fmt: fmtETB, rawKey: "netOrderValue", cellClass: "col-val" },
  ];

  function kpiCardsFor(rows) {
    const totalQty = rows.reduce((s, r) => s + r.qty, 0);
    const totalVal = rows.reduce((s, r) => s + r.netOrderValue, 0);
    const uniqMat  = new Set(rows.map(r => r.material)).size;
    return [
      ["PO Lines",            String(rows.length), "Transit detail records", "blue"],
      ["Unique Materials",    String(uniqMat),      "Distinct SKUs",          "green"],
      ["Total Transit Qty",   fmtQty(totalQty),      "Across shown lines",    "blue"],
      ["Total Net Order Value", fmtETB(totalVal),    "Across shown lines",    "amber"],
    ];
  }

  // ── DETAIL FILE TAB ────────────────────────────────────────────────────
  function renderTransitDetailTab() {
    const noFileEl  = document.getElementById("transit-detail-no-file");
    const contentEl = document.getElementById("transit-detail-content");
    if (!noFileEl || !contentEl) return;

    if (!transitDetailRaw.length) {
      noFileEl.style.display  = "block";
      contentEl.style.display = "none";
      return;
    }
    noFileEl.style.display  = "none";
    contentEl.style.display = "block";

    const rows = [...transitDetailRaw].sort((a, b) => b.netOrderValue - a.netOrderValue);
    setKpis("transit-detail-kpis", kpiCardsFor(rows));

    document.getElementById("transit-detail-table-wrap").innerHTML = rows.length
      ? buildTable(rows, detailCols, r => r._withinPlant ? "row-amber" : "", "", { id: "transit-detail-export", title: "Transit Detail" })
      : `<div class="alert-info">No transit detail rows found.</div>`;
    if (rows.length) {
      injectDlButtons("transit-detail-dl-row",
        () => downloadCSV(rows, detailCols, "transit_detail.csv"),
        () => downloadExcel(rows, detailCols, "transit_detail.xlsx"));
      wireTableExport("transit-detail-export", rows, detailCols, "transit_detail");
    } else {
      document.getElementById("transit-detail-dl-row").innerHTML = "";
    }
  }

  // ── WITHIN PLANT TAB ───────────────────────────────────────────────────
  function renderTransitWithinPlantTab() {
    const noFileEl  = document.getElementById("transit-within-no-file");
    const contentEl = document.getElementById("transit-within-content");
    if (!noFileEl || !contentEl) return;

    if (!transitDetailRaw.length) {
      noFileEl.style.display  = "block";
      contentEl.style.display = "none";
      return;
    }
    noFileEl.style.display  = "none";
    contentEl.style.display = "block";

    const rows = transitDetailRaw.filter(r => r._withinPlant).sort((a, b) => b.netOrderValue - a.netOrderValue);
    setKpis("transit-within-kpis", kpiCardsFor(rows));

    document.getElementById("transit-within-table-wrap").innerHTML = rows.length
      ? buildTable(rows, detailCols, () => "row-amber", "", { id: "transit-within-export", title: "Within-Plant Transit" })
      : `<div class="alert-info">✅ No within-plant transfers found — every transit line has a different supplying plant than its receiving plant.</div>`;
    if (rows.length) {
      injectDlButtons("transit-within-dl-row",
        () => downloadCSV(rows, detailCols, "transit_within_plant.csv"),
        () => downloadExcel(rows, detailCols, "transit_within_plant.xlsx"));
      wireTableExport("transit-within-export", rows, detailCols, "transit_within_plant");
    } else {
      document.getElementById("transit-within-dl-row").innerHTML = "";
    }
  }

  // ── WIRING ─────────────────────────────────────────────────────────────
  function wire() {
    const input = document.getElementById("transitDetailFileInput");
    if (input) {
      input.addEventListener("change", (e) => {
        const f = e.target.files[0];
        if (f) loadTransitDetailFile(f);
        e.target.value = "";
      });
    }

    // Re-render our two tabs whenever the Transit page (re)renders, same
    // pattern mos.js / national-table.js / expiry-risk.js use — wrapping
    // window.renderPage lets us hook in without touching renderTransit().
    const _origRenderPage = window.renderPage;
    window.renderPage = function (id) {
      _origRenderPage(id);
      if (id === "transit" && transitDetailRaw.length) {
        try { renderTransitDetailTab(); renderTransitWithinPlantTab(); } catch (e) { console.error(e); }
      } else if (id === "transit") {
        // still show the correct empty-state even with nothing loaded yet
        try { renderTransitDetailTab(); renderTransitWithinPlantTab(); } catch (e) { console.error(e); }
      }
    };
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", wire);
  } else {
    wire();
  }
})();
