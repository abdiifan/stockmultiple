// ════════════════════════════════════════════════════════════════
// pending-dispatch.js — "Pending Dispatch" page (Inventory Ops group)
//
// Reads the Pending Dispatch Excel (columns: Delivery, Item, Goods Issue
// Date, Material, Route, Ship-to Party, Name of the ship-to party,
// Delivery Quantity, Item Description, Purchasing Document, Created By,
// Storage Location, Special Stock) uploaded via #pendingDispatchFileInput
// (synced through storage-sync.js like every other upload slot) and
// renders:
//   1. KPI row
//   2. Top 10 deliveries with multiple line items
//   3. Pending items by storage location (Plotly bar chart)
//   4. Breakdown by branch/plant (derived from first 4 chars of
//      Ship-to Party, mapped against name)
//   5. Filterable detail table
//
// Stock type: "Q" -> Special Stock (Q). Blank/unmarked -> "RDF".
// Special Stock (Q) is ALWAYS included in totals by default — the Stock
// Type filter only narrows the view, it never silently excludes Q.
// ════════════════════════════════════════════════════════════════

(function () {
  const STATE = {
    rows: [],           // raw parsed rows
    branchMaster: {},   // plantCode(4 chars) -> branch name
    filters: { search: "", sloc: "", branch: "", stockType: "" },
  };

  // ── Helpers ──────────────────────────────────────────────────
  function excelDateToJs(val) {
    if (val instanceof Date) return val;
    if (typeof val === "number") {
      // Excel serial date -> JS Date (1900 date system)
      return new Date(Math.round((val - 25569) * 86400 * 1000));
    }
    if (typeof val === "string" && val.trim()) {
      const d = new Date(val);
      if (!isNaN(d.getTime())) return d;
    }
    return null;
  }

  function fmtDate(d) {
    if (!d) return "—";
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    return `${dd}-${mm}-${d.getFullYear()}`;
  }

  function plantCode(shipToParty) {
    return (shipToParty || "").toString().trim().slice(0, 4).toUpperCase();
  }

  function stockTypeOf(row) {
    const v = (row.specialStock || "").toString().trim().toUpperCase();
    return v === "Q" ? "Q" : "RDF";
  }

  function stockBadge(type) {
    if (type === "Q")   return `<span class="badge pd-badge-q">Special Stock (Q)</span>`;
    if (type === "MIX") return `<span class="badge pd-badge-mix">Mixed</span>`;
    return `<span class="badge pd-badge-rdf">RDF</span>`;
  }

  function escapeHtml(s) {
    return (s == null ? "" : String(s)).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  // ── Branch master: prefer a site-wide mapping if one exists
  // (window.EPSS_PLANT_MASTER / window.PLANT_MASTER, e.g. built from the
  // main Inventory or Mapping upload elsewhere in the app), otherwise
  // derive names from the most frequent "Name of the ship-to party" seen
  // per 4-char plant code in this very file. ──
  function buildBranchMaster(rows) {
    const external = window.EPSS_PLANT_MASTER || window.PLANT_MASTER || null;
    const counts = {}; // code -> { name -> count }
    rows.forEach((r) => {
      const code = plantCode(r.shipToParty);
      if (!code) return;
      const name = (r.shipToPartyName || "").toString().trim();
      if (!name) return;
      counts[code] = counts[code] || {};
      counts[code][name] = (counts[code][name] || 0) + 1;
    });
    const derived = {};
    Object.keys(counts).forEach((code) => {
      let best = null, bestN = -1;
      Object.entries(counts[code]).forEach(([name, n]) => {
        if (n > bestN) { best = name; bestN = n; }
      });
      derived[code] = best || code;
    });
    if (external && typeof external === "object") {
      // External master wins where it has an entry; fall back to derived.
      return Object.assign({}, derived, external);
    }
    return derived;
  }

  function branchName(code) {
    return STATE.branchMaster[code] || code || "Unknown";
  }

  // ── Excel parsing ────────────────────────────────────────────
  function parseSheet(workbook) {
    const sheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });

    // Tolerant header matching — same source-of-truth column names as the
    // uploaded template, but don't hard-fail on stray casing/whitespace.
    const norm = (k) => k.toString().trim().toLowerCase();
    const rows = raw.map((r) => {
      const map = {};
      Object.keys(r).forEach((k) => { map[norm(k)] = r[k]; });
      return {
        delivery:        (map["delivery"] || "").toString().trim(),
        item:             (map["item"] || "").toString().trim(),
        giDate:           excelDateToJs(map["goods issue date"]),
        material:         (map["material"] || "").toString().trim(),
        route:            (map["route"] || "").toString().trim(),
        shipToParty:      (map["ship-to party"] || "").toString().trim(),
        shipToPartyName:  (map["name of the ship-to party"] || "").toString().trim(),
        qty:              Number(map["delivery quantity"]) || 0,
        itemDescription:  (map["item description"] || "").toString().trim(),
        purchasingDoc:    (map["purchasing document"] || "").toString().trim(),
        createdBy:        (map["created by"] || "").toString().trim(),
        storageLocation:  (map["storage location"] || "").toString().trim(),
        specialStock:     (map["special stock"] || "").toString().trim(),
      };
    }).filter((r) => r.delivery); // drop blank rows

    return rows;
  }

  // ── Filtering ────────────────────────────────────────────────
  function applyFilters(rows) {
    const { search, sloc, branch, stockType } = STATE.filters;
    const s = search.trim().toLowerCase();
    return rows.filter((r) => {
      if (sloc && r.storageLocation !== sloc) return false;
      if (branch && plantCode(r.shipToParty) !== branch) return false;
      if (stockType && stockTypeOf(r) !== stockType) return false;
      if (s) {
        const hay = `${r.delivery} ${r.material} ${r.itemDescription} ${r.shipToParty} ${r.shipToPartyName} ${branchName(plantCode(r.shipToParty))}`.toLowerCase();
        if (!hay.includes(s)) return false;
      }
      return true; // NOTE: Special Stock (Q) is never excluded unless the
                   // Stock Type dropdown is explicitly set to a value.
    });
  }

  // ── KPIs ─────────────────────────────────────────────────────
  function renderKpis(rows) {
    const deliveries = new Set(rows.map((r) => r.delivery));
    const branches = new Set(rows.map((r) => plantCode(r.shipToParty)).filter(Boolean));
    const slocs = new Set(rows.map((r) => r.storageLocation).filter(Boolean));

    const cards = [
      { label: "Deliveries Pending Dispatch", value: deliveries.size, color: "blue", sub: "unique delivery numbers" },
      { label: "Pending Line Items", value: rows.length, color: "amber", sub: "individual delivery items" },
      { label: "Branches Awaiting Stock", value: branches.size, color: "green", sub: "ship-to plants" },
      { label: "Storage Locations Involved", value: slocs.size, color: "purple", sub: "source storage locations" },
    ];

    document.getElementById("pd-kpis").innerHTML = cards.map((c) => `
      <div class="kpi-card ${c.color}">
        <div class="kpi-label">${c.label}</div>
        <div class="kpi-value">${c.value.toLocaleString()}</div>
        <div class="kpi-sub">${c.sub}</div>
      </div>
    `).join("");
  }

  // ── Top 10 deliveries with multiple line items ──────────────
  function renderTop10(rows) {
    const wrap = document.getElementById("pd-top10-wrap");
    const byDeliv = {};
    rows.forEach((r) => {
      const key = r.delivery;
      if (!byDeliv[key]) {
        byDeliv[key] = {
          delivery: r.delivery,
          items: 0,
          branchCode: plantCode(r.shipToParty),
          slocs: new Set(),
          stockTypes: new Set(),
        };
      }
      const g = byDeliv[key];
      g.items += 1;
      if (r.storageLocation) g.slocs.add(r.storageLocation);
      g.stockTypes.add(stockTypeOf(r));
    });

    const list = Object.values(byDeliv)
      .filter((g) => g.items > 1) // "multiple line items"
      .sort((a, b) => b.items - a.items)
      .slice(0, 10);

    if (!list.length) {
      wrap.innerHTML = `<div class="pd-empty">No deliveries with multiple line items in the current filter.</div>`;
      return;
    }

    wrap.innerHTML = `<div class="pd-rank-list">` + list.map((g, i) => {
      const slocLabel = g.slocs.size > 1 ? `${[...g.slocs].join(", ")}` : ([...g.slocs][0] || "—");
      const typeLabel = g.stockTypes.size > 1 ? "MIX" : [...g.stockTypes][0];
      return `
        <div class="pd-rank-item">
          <div class="pd-rank-num">${i + 1}</div>
          <div class="pd-rank-body">
            <div class="pd-rank-deliv">${escapeHtml(g.delivery)} ${stockBadge(typeLabel)}</div>
            <div class="pd-rank-meta">${escapeHtml(branchName(g.branchCode))} (${escapeHtml(g.branchCode)}) &nbsp;·&nbsp; SLoc: ${escapeHtml(slocLabel)}</div>
          </div>
          <div class="pd-rank-count">
            ${g.items}
            <div class="pd-rank-count-label">items</div>
          </div>
        </div>
      `;
    }).join("") + `</div>`;
  }

  // ── Storage location chart ──────────────────────────────────
  function renderSlocChart(rows) {
    const el = document.getElementById("chart-pd-sloc");
    const counts = {};
    rows.forEach((r) => {
      const k = r.storageLocation || "—";
      counts[k] = (counts[k] || 0) + 1;
    });
    const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    if (!entries.length || typeof Plotly === "undefined") {
      el.innerHTML = `<div class="pd-empty">No data to chart.</div>`;
      return;
    }
    const x = entries.map((e) => e[0]);
    const y = entries.map((e) => e[1]);
    const isLight = document.documentElement.getAttribute("data-theme") === "light";
    Plotly.newPlot(el, [{
      x, y, type: "bar",
      marker: { color: "#3d94e0" },
      hovertemplate: "%{x}<br>%{y} pending items<extra></extra>",
    }], {
      margin: { t: 10, l: 40, r: 10, b: 60 },
      paper_bgcolor: "rgba(0,0,0,0)",
      plot_bgcolor: "rgba(0,0,0,0)",
      font: { color: isLight ? "#18253a" : "#dce8f5", size: 11 },
      xaxis: { tickangle: -30 },
      yaxis: { title: "Pending items" },
    }, { displayModeBar: false, responsive: true });
  }

  // ── Branch breakdown table ──────────────────────────────────
  function renderBranchTable(rows) {
    const wrap = document.getElementById("pd-branch-wrap");
    const byBranch = {};
    rows.forEach((r) => {
      const code = plantCode(r.shipToParty);
      if (!code) return;
      if (!byBranch[code]) byBranch[code] = { code, deliveries: new Set(), items: 0, qty: 0 };
      byBranch[code].deliveries.add(r.delivery);
      byBranch[code].items += 1;
      byBranch[code].qty += r.qty || 0;
    });
    const list = Object.values(byBranch).sort((a, b) => b.items - a.items);
    if (!list.length) {
      wrap.innerHTML = `<div class="pd-empty">No branch data in the current filter.</div>`;
      return;
    }
    wrap.innerHTML = `
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>#</th><th>Branch</th><th>Plant Code</th>
            <th>Pending Deliveries</th><th>Pending Items</th><th>Total Qty</th>
          </tr></thead>
          <tbody>
            ${list.map((b, i) => `
              <tr>
                <td>${i + 1}</td>
                <td>${escapeHtml(branchName(b.code))}</td>
                <td style="font-family:'IBM Plex Mono',monospace">${escapeHtml(b.code)}</td>
                <td>${b.deliveries.size.toLocaleString()}</td>
                <td>${b.items.toLocaleString()}</td>
                <td>${b.qty.toLocaleString()}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  // ── Detail table ─────────────────────────────────────────────
  function renderDetailTable(rows) {
    const wrap = document.getElementById("pd-detail-wrap");

    if (!rows.length) {
      wrap.innerHTML = `<div class="pd-empty">No pending line items match the current filters.</div>`;
      return;
    }

    const sorted = [...rows].sort((a, b) => (a.giDate && b.giDate ? a.giDate - b.giDate : 0));
    const MAX_ROWS = 500; // render cap for performance; full set still exports via CSV/Excel
    const shown = sorted.slice(0, MAX_ROWS);

    wrap.innerHTML = `
      <div class="tbl-wrap tbl-wrap-freeze">
        <table class="freeze-header">
          <thead><tr>
            <th>Delivery</th><th>Item</th><th>GI Date</th><th>Material</th>
            <th>Description</th><th>Branch</th><th>Storage Loc.</th>
            <th>Qty</th><th>Stock Type</th>
          </tr></thead>
          <tbody>
            ${shown.map((r) => {
              const code = plantCode(r.shipToParty);
              const type = stockTypeOf(r);
              return `
                <tr>
                  <td style="font-family:'IBM Plex Mono',monospace">${escapeHtml(r.delivery)}</td>
                  <td>${escapeHtml(r.item)}</td>
                  <td>${fmtDate(r.giDate)}</td>
                  <td class="col-mat-code">${escapeHtml(r.material)}</td>
                  <td class="col-mat-desc" style="white-space:normal;max-width:260px">${escapeHtml(r.itemDescription)}</td>
                  <td>${escapeHtml(branchName(code))}</td>
                  <td>${escapeHtml(r.storageLocation)}</td>
                  <td class="col-qty">${(r.qty || 0).toLocaleString()}</td>
                  <td>${stockBadge(type)}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      </div>
      ${rows.length > MAX_ROWS ? `<div class="pd-empty">Showing first ${MAX_ROWS.toLocaleString()} of ${rows.length.toLocaleString()} rows — use ⬇ CSV / ⬇ Excel to export the full filtered set.</div>` : ""}
    `;
  }

  // ── Filter dropdown population ──────────────────────────────
  function populateFilterOptions() {
    const slocSel = document.getElementById("pd-filter-sloc");
    const branchSel = document.getElementById("pd-filter-branch");

    const slocs = [...new Set(STATE.rows.map((r) => r.storageLocation).filter(Boolean))].sort();
    const prevSloc = slocSel.value;
    slocSel.innerHTML = `<option value="">🏬 All Storage Locations</option>` +
      slocs.map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join("");
    slocSel.value = slocs.includes(prevSloc) ? prevSloc : "";

    const branchCodes = [...new Set(STATE.rows.map((r) => plantCode(r.shipToParty)).filter(Boolean))]
      .sort((a, b) => branchName(a).localeCompare(branchName(b)));
    const prevBranch = branchSel.value;
    branchSel.innerHTML = `<option value="">🏢 All Branches</option>` +
      branchCodes.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(branchName(c))} (${escapeHtml(c)})</option>`).join("");
    branchSel.value = branchCodes.includes(prevBranch) ? prevBranch : "";
  }

  // ── Export ───────────────────────────────────────────────────
  function currentFilteredRows() {
    return applyFilters(STATE.rows);
  }

  function exportRows(rows, kind) {
    const data = rows.map((r) => ({
      "Delivery": r.delivery,
      "Item": r.item,
      "Goods Issue Date": r.giDate ? fmtDate(r.giDate) : "",
      "Material": r.material,
      "Item Description": r.itemDescription,
      "Branch": branchName(plantCode(r.shipToParty)),
      "Plant Code": plantCode(r.shipToParty),
      "Storage Location": r.storageLocation,
      "Delivery Quantity": r.qty,
      "Stock Type": stockTypeOf(r) === "Q" ? "Special Stock (Q)" : "RDF",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    if (kind === "csv") {
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      downloadBlob(blob, "pending_dispatch.csv");
    } else {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Pending Dispatch");
      XLSX.writeFile(wb, "pending_dispatch.xlsx");
    }
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  }

  // ── Tab bar (mirrors the site's existing reqan-tab-bar pattern) ──
  const TABS = ["top10", "sloc", "branch", "detail"];
  let activeTab = "top10";

  function setTabCounts(filtered) {
    const deliveries = {};
    filtered.forEach((r) => { deliveries[r.delivery] = (deliveries[r.delivery] || 0) + 1; });
    const multiCount = Object.values(deliveries).filter((n) => n > 1).length;
    const slocCount = new Set(filtered.map((r) => r.storageLocation).filter(Boolean)).size;
    const branchCount = new Set(filtered.map((r) => plantCode(r.shipToParty)).filter(Boolean)).size;

    document.getElementById("pd-tab-count-top10").textContent  = Math.min(multiCount, 10).toLocaleString();
    document.getElementById("pd-tab-count-sloc").textContent   = slocCount.toLocaleString();
    document.getElementById("pd-tab-count-branch").textContent = branchCount.toLocaleString();
    document.getElementById("pd-tab-count-detail").textContent = filtered.length.toLocaleString();
  }

  function showTab(tab) {
    activeTab = tab;
    TABS.forEach((t) => {
      document.getElementById(`pd-tab-${t}`).style.display = t === tab ? "block" : "none";
    });
    document.querySelectorAll(".pd-tab-btn").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.tab === tab);
    });
    // Plotly needs a visible container to size correctly, so (re)draw the
    // chart the moment its tab becomes active.
    if (tab === "sloc") renderSlocChart(currentFilteredRows());
  }

  function wireTabs() {
    document.querySelectorAll(".pd-tab-btn").forEach((btn) => {
      btn.addEventListener("click", () => showTab(btn.dataset.tab));
    });
  }

  // ── Master render ────────────────────────────────────────────
  function renderAll() {
    const filtered = currentFilteredRows();
    renderKpis(filtered);
    setTabCounts(filtered);
    renderTop10(filtered);
    renderBranchTable(filtered);
    renderDetailTable(filtered);
    if (activeTab === "sloc") renderSlocChart(filtered); // only draw while visible
  }

  function showHasData() {
    document.getElementById("pd-no-data").style.display = "none";
    document.getElementById("pd-content").style.display = "block";
  }
  function showNoData() {
    document.getElementById("pd-no-data").style.display = "block";
    document.getElementById("pd-content").style.display = "none";
  }

  // ── Wire filter controls ────────────────────────────────────
  function wireFilters() {
    const search = document.getElementById("pd-search");
    const sloc = document.getElementById("pd-filter-sloc");
    const branch = document.getElementById("pd-filter-branch");
    const stockType = document.getElementById("pd-filter-stocktype");
    const clearBtn = document.getElementById("pd-filter-clear");

    let debounce;
    search.addEventListener("input", () => {
      clearTimeout(debounce);
      debounce = setTimeout(() => {
        STATE.filters.search = search.value;
        renderAll();
      }, 150);
    });
    sloc.addEventListener("change", () => { STATE.filters.sloc = sloc.value; renderAll(); });
    branch.addEventListener("change", () => { STATE.filters.branch = branch.value; renderAll(); });
    stockType.addEventListener("change", () => { STATE.filters.stockType = stockType.value; renderAll(); });
    clearBtn.addEventListener("click", () => {
      STATE.filters = { search: "", sloc: "", branch: "", stockType: "" };
      search.value = ""; sloc.value = ""; branch.value = ""; stockType.value = "";
      renderAll();
    });

    document.getElementById("pd-dl-csv").addEventListener("click", () => exportRows(currentFilteredRows(), "csv"));
    document.getElementById("pd-dl-xlsx").addEventListener("click", () => exportRows(currentFilteredRows(), "xlsx"));
  }

  // ── File input handler ──────────────────────────────────────
  function wireFileInput() {
    const input = document.getElementById("pendingDispatchFileInput");
    const statusEl = document.getElementById("pendingDispatchFileStatus");
    if (!input) return;

    input.addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      if (statusEl) statusEl.innerHTML = `⏳ Parsing ${escapeHtml(file.name)}…`;

      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const data = new Uint8Array(ev.target.result);
          const workbook = XLSX.read(data, { type: "array", cellDates: true });
          const rows = parseSheet(workbook);
          if (!rows.length) throw new Error("No usable rows found — check column headers match the template.");

          STATE.rows = rows;
          STATE.branchMaster = buildBranchMaster(rows);
          populateFilterOptions();
          showHasData();
          renderAll();

          if (statusEl) statusEl.innerHTML = `✓ ${rows.length.toLocaleString()} rows loaded · ${escapeHtml(file.name)}`;
        } catch (err) {
          console.error("[pending-dispatch] Parse failed:", err);
          if (statusEl) statusEl.innerHTML = `✗ Failed to parse: ${escapeHtml(err.message || "unknown error")}`;
        } finally {
          input.value = ""; // matches the reset pattern used by the app's other loaders
        }
      };
      reader.onerror = () => {
        if (statusEl) statusEl.innerHTML = `✗ Could not read file`;
      };
      reader.readAsArrayBuffer(file);
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    wireFilters();
    wireTabs();
    wireFileInput();
    showNoData();
  });
})();
