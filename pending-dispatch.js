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
    detailFilters: { material: "", delivery: "", createdBy: "" }, // scoped only to the "All Pending Items" table
    localUploadedAt: null,   // when THIS browser last parsed a file (fallback)
    sourceUploadedAt: null,  // authoritative "uploaded to Supabase" time, if known
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

  // ── "Top 10s" tab: storage-location & plant leaderboards, plus
  //    the original top 10 multi-item-deliveries list, presented as
  //    sub-tabs styled like the site's main tab bar (icon + title +
  //    count, underline on active) ─────────────────────────────────
  let activeTop10Sub = "sloc-deliv"; // persists across re-renders/filter changes

  const TOP10_SUBS = [
    { key: "sloc-deliv",  icon: "🗄️", label: "Storage Locations — by Deliveries" },
    { key: "sloc-items",  icon: "🗄️", label: "Storage Locations — by Line Items" },
    { key: "plant-deliv", icon: "🏭", label: "Plants — by Deliveries" },
    { key: "plant-items", icon: "🏭", label: "Plants — by Line Items" },
    { key: "multi",       icon: "🏆", label: "Multi-Item Deliveries" },
  ];

  function renderTop10(rows) {
    const wrap = document.getElementById("pd-top10-wrap");
    if (!rows.length) {
      wrap.innerHTML = `<div class="pd-empty">No data in the current filter.</div>`;
      return;
    }

    // ── aggregate by storage location ──
    const bySloc = {};
    rows.forEach((r) => {
      const k = r.storageLocation || "—";
      if (!bySloc[k]) bySloc[k] = { key: k, deliveries: new Set(), items: 0 };
      bySloc[k].deliveries.add(r.delivery);
      bySloc[k].items += 1;
    });
    const slocList = Object.values(bySloc);
    const slocByDeliv = [...slocList].sort((a, b) => b.deliveries.size - a.deliveries.size).slice(0, 10);
    const slocByItems = [...slocList].sort((a, b) => b.items - a.items).slice(0, 10);
    const topSlocOutbound = slocByItems[0];

    // ── aggregate by plant/branch ──
    const byPlant = {};
    rows.forEach((r) => {
      const code = plantCode(r.shipToParty);
      if (!code) return;
      if (!byPlant[code]) byPlant[code] = { code, deliveries: new Set(), items: 0 };
      byPlant[code].deliveries.add(r.delivery);
      byPlant[code].items += 1;
    });
    const plantList = Object.values(byPlant);
    const plantByDeliv = [...plantList].sort((a, b) => b.deliveries.size - a.deliveries.size).slice(0, 10);
    const plantByItems = [...plantList].sort((a, b) => b.items - a.items).slice(0, 10);
    const topPlantsAtRisk = [...plantList].sort((a, b) => b.items - a.items).slice(0, 3);

    // ── original: top 10 deliveries with multiple line items ──
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
    const multiList = Object.values(byDeliv)
      .filter((g) => g.items > 1)
      .sort((a, b) => b.items - a.items)
      .slice(0, 10);

    // ── counts shown in the sub-tab badges ──
    const counts = {
      "sloc-deliv":  Math.min(slocList.length, 10),
      "sloc-items":  Math.min(slocList.length, 10),
      "plant-deliv": Math.min(plantList.length, 10),
      "plant-items": Math.min(plantList.length, 10),
      "multi":       multiList.length,
    };

    // ── small helper to build a rank table ──
    const rankTable = (list, keyLabel, valueKey, valueLabel, nameFn) => `
      <div class="tbl-wrap">
        <table>
          <thead><tr><th>#</th><th>${keyLabel}</th><th>${valueLabel}</th></tr></thead>
          <tbody>
            ${list.map((g, i) => `
              <tr>
                <td>${i + 1}</td>
                <td>${nameFn ? escapeHtml(nameFn(g)) : escapeHtml(g.key)}</td>
                <td>${(valueKey === "deliveries" ? g.deliveries.size : g.items).toLocaleString()}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;

    const calloutStyle = "margin:0 0 12px; padding:10px 14px; border-radius:8px; background:rgba(61,148,224,0.12); border:1px solid rgba(61,148,224,0.3); font-size:13px; line-height:1.5;";
    const callout = (html) => `<div class="pd-callout" style="${calloutStyle}">${html}</div>`;

    // ── build the sub-tab bar (mirrors the site's main tab bar look) ──
    const subTabBar = `
      <div class="pd-subtab-bar" style="display:flex; flex-wrap:wrap; gap:4px; border-bottom:1px solid rgba(120,140,160,0.25); margin-bottom:18px;">
        ${TOP10_SUBS.map((s) => {
          const active = s.key === activeTop10Sub;
          return `
            <button type="button" class="pd-subtab-btn" data-subtab="${s.key}" style="
              display:flex; align-items:center; gap:8px;
              background:transparent; border:none; cursor:pointer;
              padding:10px 14px; font-size:13px; font-weight:600;
              color:${active ? "#3d94e0" : "inherit"}; opacity:${active ? "1" : "0.65"};
              border-bottom:2px solid ${active ? "#3d94e0" : "transparent"};
              margin-bottom:-1px;
            ">
              <span>${s.icon}</span>
              <span>${s.label}</span>
              <span style="
                background:rgba(120,140,160,0.18); border-radius:10px;
                padding:1px 8px; font-size:12px; font-weight:600;
                color:inherit; opacity:0.85;
              ">${counts[s.key].toLocaleString()}</span>
            </button>
          `;
        }).join("")}
      </div>
    `;

    // ── build the active panel ──
    let panel = "";
    if (activeTop10Sub === "sloc-deliv") {
      panel = rankTable(slocByDeliv, "Storage Location", "deliveries", "Pending Deliveries");
    } else if (activeTop10Sub === "sloc-items") {
      if (topSlocOutbound) {
        panel += callout(`📦 Most outbound activity is at <strong>${escapeHtml(topSlocOutbound.key)}</strong>
          with ${topSlocOutbound.items.toLocaleString()} pending line items across ${topSlocOutbound.deliveries.size.toLocaleString()} deliveries.`);
      }
      panel += rankTable(slocByItems, "Storage Location", "items", "Pending Line Items");
    } else if (activeTop10Sub === "plant-deliv") {
      panel = rankTable(plantByDeliv, "Plant", "deliveries", "Pending Deliveries", (g) => `${branchName(g.code)} (${g.code})`);
    } else if (activeTop10Sub === "plant-items") {
      if (topPlantsAtRisk.length) {
        panel += callout(`⚠️ If dispatch is delayed, <strong>${escapeHtml(topPlantsAtRisk.map((p) => `${branchName(p.code)} (${p.code})`).join(", "))}</strong>
          will be the most affected, with the highest volume of pending line items waiting on stock.`);
      }
      panel += rankTable(plantByItems, "Plant", "items", "Pending Line Items", (g) => `${branchName(g.code)} (${g.code})`);
    } else if (activeTop10Sub === "multi") {
      if (!multiList.length) {
        panel = `<div class="pd-empty">No deliveries with multiple line items in the current filter.</div>`;
      } else {
        panel = `<div class="pd-rank-list">` + multiList.map((g, i) => {
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
    }

    wrap.innerHTML = subTabBar + tableExportButtonsHtml("top10:" + activeTop10Sub) + `<div class="pd-subtab-panel">${panel}</div>`;

    wrap.querySelectorAll(".pd-subtab-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        activeTop10Sub = btn.dataset.subtab;
        renderTop10(currentFilteredRows());
      });
    });
  }

  // ── Storage location: table only (chart removed) ─────────────
  function renderSlocChart(rows) {
    const el = document.getElementById("chart-pd-sloc");
    if (el) el.style.display = "none"; // chart no longer shown, keep element hidden
    renderSlocTable(rows);
  }

  // ── Storage location table (deliveries / items / RDF / Q) ────
  function renderSlocTable(rows) {
    // Ensure a wrap element exists, even if the hosting HTML page hasn't
    // been updated with a dedicated container. Placed where the chart
    // used to be (chart element is now hidden).
    let wrap = document.getElementById("pd-sloc-table-wrap");
    if (!wrap) {
      const chartEl = document.getElementById("chart-pd-sloc");
      if (!chartEl || !chartEl.parentNode) return;
      wrap = document.createElement("div");
      wrap.id = "pd-sloc-table-wrap";
      wrap.style.marginTop = "0";
      chartEl.parentNode.insertBefore(wrap, chartEl.nextSibling);
    }

    const bySloc = {};
    rows.forEach((r) => {
      const k = r.storageLocation || "—";
      if (!bySloc[k]) bySloc[k] = { sloc: k, deliveries: new Set(), items: 0, rdf: 0, q: 0 };
      const g = bySloc[k];
      g.deliveries.add(r.delivery);
      g.items += 1;
      if (stockTypeOf(r) === "Q") g.q += 1; else g.rdf += 1;
    });

    const list = Object.values(bySloc).sort((a, b) => b.items - a.items);
    if (!list.length) {
      wrap.innerHTML = `<div class="pd-empty">No storage location data in the current filter.</div>`;
      return;
    }

    wrap.innerHTML = tableExportButtonsHtml("sloc") + `
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>#</th><th>Storage Location</th>
            <th># Deliveries</th><th># Line Items</th><th>RDF</th><th>Q</th>
          </tr></thead>
          <tbody>
            ${list.map((s, i) => `
              <tr>
                <td>${i + 1}</td>
                <td style="font-family:'IBM Plex Mono',monospace">${escapeHtml(s.sloc)}</td>
                <td>${s.deliveries.size.toLocaleString()}</td>
                <td>${s.items.toLocaleString()}</td>
                <td>${s.rdf.toLocaleString()}</td>
                <td>${s.q.toLocaleString()}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  // ── Branch breakdown table ──────────────────────────────────
  function renderBranchTable(rows) {
    const wrap = document.getElementById("pd-branch-wrap");
    const byBranch = {};
    rows.forEach((r) => {
      const code = plantCode(r.shipToParty);
      if (!code) return;
      if (!byBranch[code]) byBranch[code] = { code, deliveries: new Set(), items: 0 };
      byBranch[code].deliveries.add(r.delivery);
      byBranch[code].items += 1;
    });
    const list = Object.values(byBranch).sort((a, b) => b.items - a.items);
    if (!list.length) {
      wrap.innerHTML = `<div class="pd-empty">No branch data in the current filter.</div>`;
      return;
    }
    wrap.innerHTML = tableExportButtonsHtml("branch") + `
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>#</th><th>Branch</th><th>Plant Code</th>
            <th>Pending Deliveries</th><th>Pending Items</th>
          </tr></thead>
          <tbody>
            ${list.map((b, i) => `
              <tr>
                <td>${i + 1}</td>
                <td>${escapeHtml(branchName(b.code))}</td>
                <td style="font-family:'IBM Plex Mono',monospace">${escapeHtml(b.code)}</td>
                <td>${b.deliveries.size.toLocaleString()}</td>
                <td>${b.items.toLocaleString()}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  // ── Heat-scale background for matrix cells (green → yellow → red,
  //    same 3-color-scale idea as an Excel conditional format) ─────
  function heatBg(value, max) {
    if (!value) return "";
    const stops = [
      [99, 190, 123],   // green  (low)
      [255, 235, 132],  // yellow (mid)
      [248, 105, 107],  // red    (high)
    ];
    const t = Math.max(0, Math.min(1, value / max));
    const scaled = t * (stops.length - 1);
    const i = Math.min(stops.length - 2, Math.floor(scaled));
    const localT = scaled - i;
    const [r1, g1, b1] = stops[i];
    const [r2, g2, b2] = stops[i + 1];
    const r = Math.round(r1 + (r2 - r1) * localT);
    const g = Math.round(g1 + (g2 - g1) * localT);
    const b = Math.round(b1 + (b2 - b1) * localT);
    return `background:rgb(${r},${g},${b})`;
  }

  // ── "Branch × Storage Location" cross-tab (unique DELIVERY counts,
  //    not line-item counts — a delivery is counted once per branch×sloc
  //    cell it appears in, regardless of how many line items it has) ────
  function renderMatrixTable(rows) {
    const wrap = document.getElementById("pd-matrix-wrap");
    if (!wrap) return;

    const matrix = {};        // matrix[branch][sloc] = Set(delivery)
    const branchTotals = {};  // branch -> Set(delivery)
    const slocTotals = {};    // sloc -> Set(delivery)
    const grandDeliveries = new Set();
    const slocSeen = [];      // first-seen order, used to build column order below

    rows.forEach((r) => {
      const branch = branchName(plantCode(r.shipToParty)) || "—";
      const sloc = r.storageLocation || "—";
      if (!matrix[branch]) matrix[branch] = {};
      if (!matrix[branch][sloc]) matrix[branch][sloc] = new Set();
      matrix[branch][sloc].add(r.delivery);
      if (!branchTotals[branch]) branchTotals[branch] = new Set();
      branchTotals[branch].add(r.delivery);
      if (!slocTotals[sloc]) slocTotals[sloc] = new Set();
      slocTotals[sloc].add(r.delivery);
      if (!slocSeen.includes(sloc)) slocSeen.push(sloc);
      grandDeliveries.add(r.delivery);
    });

    const grandTotal = grandDeliveries.size;
    if (!grandTotal) {
      wrap.innerHTML = `<div class="pd-empty">No data in the current filter.</div>`;
      return;
    }

    // Column order: named locations first (as first seen), then any
    // "Main-N" locations sorted numerically — mirrors the source layout.
    const mainCols = slocSeen
      .filter((s) => /^Main-/i.test(s))
      .sort((a, b) => {
        const na = parseFloat(a.replace(/^Main-/i, "").replace("/", "."));
        const nb = parseFloat(b.replace(/^Main-/i, "").replace("/", "."));
        return na - nb;
      });
    const namedCols = slocSeen.filter((s) => !/^Main-/i.test(s));
    const cols = [...namedCols, ...mainCols];

    const branches = Object.keys(branchTotals).sort((a, b) => a.localeCompare(b));

    const cellCount = (b, c) => matrix[b]?.[c]?.size || 0;
    const rowTotal = (b) => branchTotals[b]?.size || 0;
    const colTotal = (c) => slocTotals[c]?.size || 0;

    const maxCell = Math.max(1, ...branches.flatMap((b) => cols.map((c) => cellCount(b, c))));
    const maxRowTotal = Math.max(1, ...branches.map((b) => rowTotal(b)));
    const maxColTotal = Math.max(1, ...cols.map((c) => colTotal(c)));

    const headerCells = cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("");

    const bodyRows = branches.map((b, i) => {
      const cells = cols.map((c) => {
        const v = cellCount(b, c);
        return `<td style="${heatBg(v, maxCell)}">${v ? v.toLocaleString() : ""}</td>`;
      }).join("");
      const rt = rowTotal(b);
      const rowPct = grandTotal ? Math.round((rt / grandTotal) * 100) : 0;
      return `
        <tr>
          <td>${i + 1}</td>
          <td style="text-align:left">${escapeHtml(b)}</td>
          ${cells}
          <td style="${heatBg(rt, maxRowTotal)};font-weight:700">${rt.toLocaleString()}</td>
          <td style="background:#dbe9fb;font-weight:600">${rowPct}%</td>
        </tr>`;
    }).join("");

    const totalCells = cols.map((c) => `<td style="${heatBg(colTotal(c), maxColTotal)};font-weight:700">${colTotal(c).toLocaleString()}</td>`).join("");
    const pctCells = cols.map((c) => {
      const pct = grandTotal ? Math.round((colTotal(c) / grandTotal) * 100) : 0;
      return `<td style="background:#dbe9fb;font-weight:600">${pct}%</td>`;
    }).join("");

    wrap.innerHTML = tableExportButtonsHtml("matrix") + `
      <div class="tbl-wrap">
        <table class="pd-matrix-table">
          <thead>
            <tr>
              <th>SN</th><th style="text-align:left">Branch</th>${headerCells}<th>Total</th><th>Total %</th>
            </tr>
          </thead>
          <tbody>
            ${bodyRows}
            <tr class="pd-matrix-total-row">
              <td colspan="2">Total</td>
              ${totalCells}
              <td style="font-weight:700">${grandTotal.toLocaleString()}</td>
              <td style="font-weight:700">100%</td>
            </tr>
            <tr class="pd-matrix-pct-row">
              <td colspan="2"></td>
              ${pctCells}
              <td></td>
              <td></td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  }


  // ── Days late vs. planned Goods Issue date ────────────────────
  // today - giDate: negative => not yet due (on track / "Good"),
  // positive => days overdue past the planned GI date.
  function daysLate(giDate) {
    if (!giDate) return null;
    const today = new Date();
    const t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const g0 = new Date(giDate.getFullYear(), giDate.getMonth(), giDate.getDate());
    return Math.round((t0 - g0) / 86400000);
  }

  function daysLateBadge(giDate) {
    const d = daysLate(giDate);
    if (d === null) return `<span class="badge">—</span>`;
    if (d <= 0) {
      return `<span class="badge pd-badge-good" style="background:rgba(46,175,110,0.15);color:#2eaf6e;border:1px solid rgba(46,175,110,0.35);">Good</span>`;
    }
    return `<span class="badge pd-badge-late" style="background:rgba(224,77,61,0.15);color:#e04d3d;border:1px solid rgba(224,77,61,0.35);">${d} day${d === 1 ? "" : "s"} late</span>`;
  }

  // ── Detail-table-only filters (Material / Delivery / Created By) —
  //    real filters, not a search box, scoped specifically to the
  //    "All Pending Items" table (on top of the page-level filters).
  //    Suggestions are a custom dropdown (not the native <datalist>
  //    popup) so it always opens BELOW the field, never upward over
  //    the page content. ──
  function applyDetailFilters(rows) {
    const { material, delivery, createdBy } = STATE.detailFilters;
    return rows.filter((r) => {
      if (material && r.material !== material) return false;
      if (delivery && r.delivery !== delivery) return false;
      if (createdBy && r.createdBy !== createdBy) return false;
      return true;
    });
  }

  function detailFilterBarHtml(optionRows) {
    const { material, delivery, createdBy } = STATE.detailFilters;
    const inputStyle = "font-size:12px; padding:6px 8px; border-radius:6px; border:1px solid rgba(120,140,160,0.35); background:transparent; min-width:170px; width:100%; box-sizing:border-box;";

    const field = (id, label, value) => `
      <div class="pd-detail-filter-field" style="position:relative; min-width:170px;">
        <input type="text" id="pd-detail-filter-${id}" placeholder="Filter by ${label}…"
          value="${escapeHtml(value)}" style="${inputStyle}" autocomplete="off">
        <div id="pd-detail-suggest-${id}" class="pd-detail-suggest" style="
          display:none; position:absolute; top:100%; left:0; right:0; margin-top:4px;
          max-height:220px; overflow-y:auto; background:#fff; color:#1a1a1a;
          border:1px solid rgba(120,140,160,0.35); border-radius:6px;
          box-shadow:0 6px 16px rgba(0,0,0,0.15); z-index:40;
        "></div>
      </div>
    `;

    return `
      <div class="pd-detail-filter-bar" style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:10px; align-items:flex-start;">
        ${field("material", "Material", material)}
        ${field("delivery", "Delivery", delivery)}
        ${field("createdby", "Created By", createdBy)}
        ${(material || delivery || createdBy) ? `<button type="button" id="pd-detail-filter-clear" style="${inputStyle} width:auto; cursor:pointer;">✕ Clear</button>` : ""}
      </div>
    `;
  }

  function wireDetailFilterBar(wrap, optionRows) {
    const fieldDefs = [
      { id: "material",  prop: "material" },
      { id: "delivery",  prop: "delivery" },
      { id: "createdby", prop: "createdBy" },
    ];

    fieldDefs.forEach(({ id, prop }) => {
      const input = wrap.querySelector(`#pd-detail-filter-${id}`);
      const suggestBox = wrap.querySelector(`#pd-detail-suggest-${id}`);
      if (!input || !suggestBox) return;

      const values = [...new Set(optionRows.map((r) => r[prop]).filter(Boolean))].sort();

      const commit = (val) => {
        STATE.detailFilters[prop] = val;
        renderDetailTable(currentFilteredRows());
      };

      const showSuggestions = () => {
        const q = input.value.trim().toLowerCase();
        const matches = (q ? values.filter((v) => v.toLowerCase().includes(q)) : values).slice(0, 50);
        if (!matches.length) {
          suggestBox.innerHTML = `<div style="padding:8px 10px; font-size:12px; opacity:0.6;">No matches</div>`;
        } else {
          suggestBox.innerHTML = matches.map((v) => `<div class="pd-suggest-item" data-value="${escapeHtml(v)}" style="padding:7px 10px; font-size:12px; cursor:pointer;">${escapeHtml(v)}</div>`).join("");
          suggestBox.querySelectorAll(".pd-suggest-item").forEach((item) => {
            // mousedown (not click) + preventDefault so the input never
            // blurs before we read the picked value — keeps the dropdown
            // reliably below the field with no flicker.
            item.addEventListener("mousedown", (e) => {
              e.preventDefault();
              input.value = item.dataset.value;
              suggestBox.style.display = "none";
              commit(item.dataset.value);
            });
            item.addEventListener("mouseenter", () => { item.style.background = "rgba(61,148,224,0.12)"; });
            item.addEventListener("mouseleave", () => { item.style.background = "transparent"; });
          });
        }
        suggestBox.style.display = "block";
      };

      input.addEventListener("focus", showSuggestions);
      input.addEventListener("input", showSuggestions);
      input.addEventListener("blur", () => {
        suggestBox.style.display = "none";
        // Typed-and-tabbed-away exact match still applies as a filter;
        // anything else (partial text with no selection) is discarded.
        const typed = input.value.trim();
        if (typed !== STATE.detailFilters[prop] && values.includes(typed)) {
          commit(typed);
        } else if (typed !== STATE.detailFilters[prop] && !typed) {
          commit("");
        } else if (typed !== STATE.detailFilters[prop]) {
          input.value = STATE.detailFilters[prop]; // reject non-matching free text
        }
      });
      input.addEventListener("keydown", (e) => {
        if (e.key === "Escape") { suggestBox.style.display = "none"; input.blur(); }
      });
    });

    const clearBtn = wrap.querySelector("#pd-detail-filter-clear");
    if (clearBtn) clearBtn.addEventListener("click", () => {
      STATE.detailFilters = { material: "", delivery: "", createdBy: "" };
      renderDetailTable(currentFilteredRows());
    });
  }

  function renderDetailTable(rows) {
    const wrap = document.getElementById("pd-detail-wrap");

    if (!rows.length) {
      wrap.innerHTML = `<div class="pd-empty">No pending line items match the current filters.</div>`;
      return;
    }

    const filterBar = detailFilterBarHtml(rows);
    const exportBtns = tableExportButtonsHtml("detail");
    const filtered = applyDetailFilters(rows);

    if (!filtered.length) {
      wrap.innerHTML = filterBar + exportBtns + `<div class="pd-empty">No pending line items match the Material / Delivery / Created By filter.</div>`;
      wireDetailFilterBar(wrap, rows);
      return;
    }

    const sorted = [...filtered].sort((a, b) => (a.giDate && b.giDate ? a.giDate - b.giDate : 0));
    const shown = sorted; // no render cap — full filtered set is shown

    wrap.innerHTML = filterBar + exportBtns + `
      <div class="tbl-wrap tbl-wrap-freeze">
        <table class="freeze-header">
          <thead><tr>
            <th>Delivery</th><th>GI Planned Date</th><th>Days Late</th><th>Material</th>
            <th>Description</th><th>Purchasing Document</th><th>Branch</th><th>Storage Loc.</th>
            <th>Qty</th><th>Stock Type</th><th>Created By</th>
          </tr></thead>
          <tbody>
            ${shown.map((r) => {
              const code = plantCode(r.shipToParty);
              const type = stockTypeOf(r);
              return `
                <tr>
                  <td style="font-family:'IBM Plex Mono',monospace">${escapeHtml(r.delivery)}</td>
                  <td>${fmtDate(r.giDate)}</td>
                  <td>${daysLateBadge(r.giDate)}</td>
                  <td class="col-mat-code">${escapeHtml(r.material)}</td>
                  <td class="col-mat-desc" style="white-space:normal;max-width:260px">${escapeHtml(r.itemDescription)}</td>
                  <td style="font-family:'IBM Plex Mono',monospace">${escapeHtml(r.purchasingDoc)}</td>
                  <td>${escapeHtml(branchName(code))}</td>
                  <td>${escapeHtml(r.storageLocation)}</td>
                  <td class="col-qty">${(r.qty || 0).toLocaleString()}</td>
                  <td>${stockBadge(type)}</td>
                  <td>${escapeHtml(r.createdBy)}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      </div>
    `;
    wireDetailFilterBar(wrap, rows);
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

  // ── Row-level "All Pending Items" data (same shape as before) ──
  function buildDetailExportData(rows) {
    return rows.map((r) => ({
      "Delivery": r.delivery,
      "Item": r.item,
      "GI Planned Date": r.giDate ? fmtDate(r.giDate) : "",
      "Days Late": (() => { const d = daysLate(r.giDate); return d === null ? "" : (d <= 0 ? "Good" : d); })(),
      "Material": r.material,
      "Item Description": r.itemDescription,
      "Purchasing Document": r.purchasingDoc,
      "Branch": branchName(plantCode(r.shipToParty)),
      "Plant Code": plantCode(r.shipToParty),
      "Storage Location": r.storageLocation,
      "Delivery Quantity": r.qty,
      "Stock Type": stockTypeOf(r) === "Q" ? "Special Stock (Q)" : "RDF",
    }));
  }

  // ── "By Storage Location" data (mirrors renderSlocTable) ───────
  function buildSlocExportData(rows) {
    const bySloc = {};
    rows.forEach((r) => {
      const k = r.storageLocation || "—";
      if (!bySloc[k]) bySloc[k] = { sloc: k, deliveries: new Set(), items: 0, rdf: 0, q: 0 };
      const g = bySloc[k];
      g.deliveries.add(r.delivery);
      g.items += 1;
      if (stockTypeOf(r) === "Q") g.q += 1; else g.rdf += 1;
    });
    return Object.values(bySloc)
      .sort((a, b) => b.items - a.items)
      .map((s) => ({
        "Storage Location": s.sloc,
        "# Deliveries": s.deliveries.size,
        "# Line Items": s.items,
        "RDF": s.rdf,
        "Q": s.q,
      }));
  }

  // ── "By Branch / Plant" data (mirrors renderBranchTable) ───────
  function buildBranchExportData(rows) {
    const byBranch = {};
    rows.forEach((r) => {
      const code = plantCode(r.shipToParty);
      if (!code) return;
      if (!byBranch[code]) byBranch[code] = { code, deliveries: new Set(), items: 0 };
      byBranch[code].deliveries.add(r.delivery);
      byBranch[code].items += 1;
    });
    return Object.values(byBranch)
      .sort((a, b) => b.items - a.items)
      .map((b, i) => ({
        "#": i + 1,
        "Branch": branchName(b.code),
        "Plant Code": b.code,
        "Pending Deliveries": b.deliveries.size,
        "Pending Items": b.items,
      }));
  }

  // ── "Top 10s" data — top 10 deliveries with multiple line items
  //    (mirrors the "multi" sub-tab in renderTop10) ────────────────
  function buildTop10sExportData(rows) {
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
    return Object.values(byDeliv)
      .filter((g) => g.items > 1)
      .sort((a, b) => b.items - a.items)
      .slice(0, 10)
      .map((g, i) => ({
        "#": i + 1,
        "Delivery": g.delivery,
        "Branch": branchName(g.branchCode),
        "Plant Code": g.branchCode,
        "Storage Location(s)": [...g.slocs].join(", "),
        "Stock Type": g.stockTypes.size > 1 ? "MIX" : (g.stockTypes.values().next().value === "Q" ? "Special Stock (Q)" : "RDF"),
        "# Line Items": g.items,
      }));
  }

  // Excel sheet names can't contain / \ ? * [ ] : and are capped at 31 chars.
  function safeSheetName(name) {
    return name.replace(/[\\/?*[\]:]/g, "-").slice(0, 31);
  }

  // The timestamp this data was actually uploaded (source-of-truth if known,
  // else the local parse time) — shown on the page banner and now baked
  // into every export too, since a report can easily be opened well after
  // the underlying SAP data has moved on.
  function dataAsOfString() {
    const ts = STATE.sourceUploadedAt || STATE.localUploadedAt;
    return ts ? ts.toLocaleString() : "Unknown";
  }

  // ── Per-table export buttons ────────────────────────────────
  // Every table gets its own small CSV/Excel export pair that exports
  // ONLY the data in that specific table — never the other tables' data.
  function tableExportButtonsHtml(tableKey) {
    return `
      <div class="pd-table-export" style="display:flex; justify-content:flex-end; gap:6px; margin-bottom:8px;">
        <button type="button" class="pd-export-btn" data-export-table="${tableKey}" data-export-kind="csv"
          style="font-size:12px; font-weight:600; padding:5px 10px; border-radius:6px; border:1px solid rgba(120,140,160,0.35); background:transparent; cursor:pointer;">⬇ CSV</button>
        <button type="button" class="pd-export-btn" data-export-table="${tableKey}" data-export-kind="xlsx"
          style="font-size:12px; font-weight:600; padding:5px 10px; border-radius:6px; border:1px solid rgba(120,140,160,0.35); background:transparent; cursor:pointer;">⬇ Excel</button>
      </div>
    `;
  }

  // ── "Branch × Storage Location" matrix export (mirrors renderMatrixTable)
  //    — counts UNIQUE DELIVERIES per cell, not line items ──
  function buildMatrixExportData(rows) {
    const matrix = {};        // matrix[branch][sloc] = Set(delivery)
    const branchTotals = {};  // branch -> Set(delivery)
    const slocTotals = {};    // sloc -> Set(delivery)
    const grandDeliveries = new Set();
    const slocSeen = [];

    rows.forEach((r) => {
      const branch = branchName(plantCode(r.shipToParty)) || "—";
      const sloc = r.storageLocation || "—";
      if (!matrix[branch]) matrix[branch] = {};
      if (!matrix[branch][sloc]) matrix[branch][sloc] = new Set();
      matrix[branch][sloc].add(r.delivery);
      if (!branchTotals[branch]) branchTotals[branch] = new Set();
      branchTotals[branch].add(r.delivery);
      if (!slocTotals[sloc]) slocTotals[sloc] = new Set();
      slocTotals[sloc].add(r.delivery);
      if (!slocSeen.includes(sloc)) slocSeen.push(sloc);
      grandDeliveries.add(r.delivery);
    });

    const grandTotal = grandDeliveries.size;
    if (!grandTotal) return [];

    const mainCols = slocSeen
      .filter((s) => /^Main-/i.test(s))
      .sort((a, b) => {
        const na = parseFloat(a.replace(/^Main-/i, "").replace("/", "."));
        const nb = parseFloat(b.replace(/^Main-/i, "").replace("/", "."));
        return na - nb;
      });
    const namedCols = slocSeen.filter((s) => !/^Main-/i.test(s));
    const cols = [...namedCols, ...mainCols];
    const branches = Object.keys(branchTotals).sort((a, b) => a.localeCompare(b));

    const rowsOut = branches.map((b) => {
      const row = { "Branch": b };
      cols.forEach((c) => { row[c] = matrix[b]?.[c]?.size || 0; });
      const rt = branchTotals[b].size;
      row["Total"] = rt;
      row["Total %"] = grandTotal ? Math.round((rt / grandTotal) * 100) + "%" : "0%";
      return row;
    });

    const totalRow = { "Branch": "Total" };
    cols.forEach((c) => { totalRow[c] = slocTotals[c]?.size || 0; });
    totalRow["Total"] = grandTotal;
    totalRow["Total %"] = "100%";
    rowsOut.push(totalRow);

    return rowsOut;
  }

  // ── "Top 10s" sub-tab export data — mirrors each panel in
  //    renderTop10 exactly, keyed by the sub-tab it's exported from ──
  function buildTop10SubExportData(rows, subKey) {
    const bySloc = {};
    rows.forEach((r) => {
      const k = r.storageLocation || "—";
      if (!bySloc[k]) bySloc[k] = { key: k, deliveries: new Set(), items: 0 };
      bySloc[k].deliveries.add(r.delivery);
      bySloc[k].items += 1;
    });
    const byPlant = {};
    rows.forEach((r) => {
      const code = plantCode(r.shipToParty);
      if (!code) return;
      if (!byPlant[code]) byPlant[code] = { code, deliveries: new Set(), items: 0 };
      byPlant[code].deliveries.add(r.delivery);
      byPlant[code].items += 1;
    });

    if (subKey === "sloc-deliv") {
      return Object.values(bySloc)
        .sort((a, b) => b.deliveries.size - a.deliveries.size)
        .slice(0, 10)
        .map((s, i) => ({ "#": i + 1, "Storage Location": s.key, "Pending Deliveries": s.deliveries.size }));
    }
    if (subKey === "sloc-items") {
      return Object.values(bySloc)
        .sort((a, b) => b.items - a.items)
        .slice(0, 10)
        .map((s, i) => ({ "#": i + 1, "Storage Location": s.key, "Pending Line Items": s.items }));
    }
    if (subKey === "plant-deliv") {
      return Object.values(byPlant)
        .sort((a, b) => b.deliveries.size - a.deliveries.size)
        .slice(0, 10)
        .map((p, i) => ({ "#": i + 1, "Branch": branchName(p.code), "Plant Code": p.code, "Pending Deliveries": p.deliveries.size }));
    }
    if (subKey === "plant-items") {
      return Object.values(byPlant)
        .sort((a, b) => b.items - a.items)
        .slice(0, 10)
        .map((p, i) => ({ "#": i + 1, "Branch": branchName(p.code), "Plant Code": p.code, "Pending Line Items": p.items }));
    }
    // "multi" — top 10 deliveries with multiple line items
    return buildTop10sExportData(rows);
  }

  // ── Single-table export dispatcher — resolves a tableKey (as set on
  //    the button that was clicked) to its data + a filename, then
  //    writes ONLY that table, never the other tables on the page. ──
  function exportSingleTable(tableKey, kind) {
    const rows = currentFilteredRows();
    let data, filename, sheetName;

    if (tableKey === "detail") {
      data = buildDetailExportData(applyDetailFilters(rows)); filename = "pending_dispatch_all_items"; sheetName = "All Pending Items";
    } else if (tableKey === "sloc") {
      data = buildSlocExportData(rows); filename = "pending_dispatch_by_storage_location"; sheetName = "By Storage Location";
    } else if (tableKey === "branch") {
      data = buildBranchExportData(rows); filename = "pending_dispatch_by_branch"; sheetName = "By Branch or Plant";
    } else if (tableKey === "matrix") {
      data = buildMatrixExportData(rows); filename = "pending_dispatch_branch_sloc_matrix"; sheetName = "Branch x SLoc Matrix";
    } else if (tableKey.startsWith("top10:")) {
      const subKey = tableKey.slice("top10:".length);
      data = buildTop10SubExportData(rows, subKey);
      filename = "pending_dispatch_top10_" + subKey.replace(/-/g, "_");
      sheetName = safeSheetName("Top 10 - " + subKey);
    } else {
      return;
    }

    if (!data || !data.length) {
      data = [{ "Note": "No data in the current filter." }];
    }

    exportSingleSheet(data, kind, filename, sheetName);
  }

  // Writes exactly one table's data as CSV or as a single-sheet XLSX
  // (with the same "data as of / exported at" metadata as the full
  // report export), never bundling in any other table.
  function exportSingleSheet(data, kind, filenameBase, sheetName) {
    const asOf = dataAsOfString();
    const exportedAt = new Date().toLocaleString();

    if (kind === "csv") {
      const ws = XLSX.utils.json_to_sheet(data);
      const meta = `Data as of,${asOf}\nExported at,${exportedAt}\n\n`;
      const csv = meta + XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      downloadBlob(blob, `${filenameBase}.csv`);
      return;
    }

    const wb = XLSX.utils.book_new();
    const infoData = [
      { "Field": "Report", "Value": "Pending Dispatch — " + sheetName },
      { "Field": "Data as of", "Value": asOf },
      { "Field": "Exported at", "Value": exportedAt },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(infoData), safeSheetName("Info"));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), safeSheetName(sheetName));
    XLSX.writeFile(wb, `${filenameBase}.xlsx`);
  }

  // Delegated click handler — works for buttons injected into any
  // table wrap, including ones re-rendered after filtering/tab switches.
  function wireTableExportButtons() {
    document.addEventListener("click", (e) => {
      const btn = e.target.closest(".pd-export-btn");
      if (!btn) return;
      const tableKey = btn.dataset.exportTable;
      const kind = btn.dataset.exportKind;
      if (!tableKey || !kind) return;
      exportSingleTable(tableKey, kind);
    });
  }

  function exportRows(rows, kind) {
    const asOf = dataAsOfString();
    const exportedAt = new Date().toLocaleString();

    if (kind === "csv") {
      // CSV has no concept of multiple sheets — keep it as the flat
      // "All Pending Items" detail export, with the timestamps as a
      // couple of leading rows before the real header/data rows.
      const ws = XLSX.utils.json_to_sheet(buildDetailExportData(rows));
      const meta = `Data as of,${asOf}\nExported at,${exportedAt}\n\n`;
      const csv = meta + XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      downloadBlob(blob, "pending_dispatch.csv");
      return;
    }

    // XLSX export: one sheet per table, in the same order as the tabs —
    // All Pending Items, By Storage Location, By Branch / Plant, Top 10s —
    // plus a leading "Info" sheet noting when the data is as of.
    const wb = XLSX.utils.book_new();
    const infoData = [
      { "Field": "Report", "Value": "Pending Dispatch" },
      { "Field": "Data as of", "Value": asOf },
      { "Field": "Exported at", "Value": exportedAt },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(infoData), safeSheetName("Info"));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDetailExportData(applyDetailFilters(rows))), safeSheetName("All Pending Items"));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildSlocExportData(rows)), safeSheetName("By Storage Location"));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildBranchExportData(rows)), safeSheetName("By Branch or Plant"));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildTop10sExportData(rows)), safeSheetName("Top 10s"));
    XLSX.writeFile(wb, "pending_dispatch.xlsx");
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  }

  // ── Tab bar (mirrors the site's existing reqan-tab-bar pattern) ──
  const TABS = ["detail", "sloc", "branch", "top10", "matrix"];
  let activeTab = "detail";

  function setTabCounts(filtered) {
    const deliveries = {};
    filtered.forEach((r) => { deliveries[r.delivery] = (deliveries[r.delivery] || 0) + 1; });
    const multiCount = Object.values(deliveries).filter((n) => n > 1).length;
    const slocCount = new Set(filtered.map((r) => r.storageLocation).filter(Boolean)).size;
    const branchCount = new Set(filtered.map((r) => plantCode(r.shipToParty)).filter(Boolean)).size;

    document.getElementById("pd-tab-count-top10").style.display  = "none";
    document.getElementById("pd-tab-count-sloc").style.display   = "none";
    document.getElementById("pd-tab-count-branch").style.display = "none";
    document.getElementById("pd-tab-count-detail").style.display = "none";
    document.getElementById("pd-tab-count-matrix").style.display = "none";
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
    // The detail table can be large (no row cap) — only build/render it
    // when its tab is actually visible, instead of on every filter change.
    if (tab === "detail") renderDetailTable(currentFilteredRows());
    if (tab === "matrix") renderMatrixTable(currentFilteredRows());
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
    if (activeTab === "sloc") renderSlocChart(filtered);   // only draw while visible
    if (activeTab === "detail") renderDetailTable(filtered); // heavy — only build while visible
    if (activeTab === "matrix") renderMatrixTable(filtered); // heavy — only build while visible
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
      STATE.detailFilters = { material: "", delivery: "", createdBy: "" };
      search.value = ""; sloc.value = ""; branch.value = ""; stockType.value = "";
      renderAll();
    });

    // The page-level "download everything" CSV/Excel buttons are now
    // redundant with the per-table export buttons that sit directly
    // above each table, so they're hidden rather than wired up.
    const globalCsvBtn = document.getElementById("pd-dl-csv");
    const globalXlsxBtn = document.getElementById("pd-dl-xlsx");
    if (globalCsvBtn) globalCsvBtn.style.display = "none";
    if (globalXlsxBtn) globalXlsxBtn.style.display = "none";
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
          // Fallback "as of" timestamp — used unless/until the more
          // authoritative Supabase upload time arrives via pd-source-meta
          // (see wireSourceMetaListener below).
          STATE.localUploadedAt = new Date();
          populateFilterOptions();
          showHasData();
          renderAsOfBanner();

          // Let the browser paint the "Parsing…" status before doing the
          // heavy synchronous render work (esp. the uncapped detail table).
          requestAnimationFrame(() => {
            try {
              renderAll();
              if (statusEl) statusEl.innerHTML = `✓ ${rows.length.toLocaleString()} rows loaded · ${escapeHtml(file.name)}`;
            } catch (err) {
              console.error("[pending-dispatch] Render failed:", err);
              if (statusEl) statusEl.innerHTML = `✗ Failed to render: ${escapeHtml(err.message || "unknown error")}`;
            }
          });
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

  // ── "Data as of" banner ──────────────────────────────────────
  // Prefers the authoritative Supabase upload time (STATE.sourceUploadedAt,
  // set via the pd-source-meta event from storage-sync.js) over the local
  // "when this browser parsed it" fallback — the source time is what
  // actually matters, since SAP data may have moved on since then.
  function renderAsOfBanner() {
    const el = document.getElementById("pd-data-asof");
    if (!el) return;
    const ts = STATE.sourceUploadedAt || STATE.localUploadedAt;
    if (!ts) { el.textContent = ""; return; }
    el.textContent = `🕒 Data as of ${ts.toLocaleString()} — items may have since shipped or changed in SAP.`;
    el.title = ts.toISOString();
  }

  // storage-sync.js dispatches this whenever it learns/updates the
  // Pending Dispatch slot's Supabase upload metadata (on pull, push, or
  // manual refresh). It's optional — pages without storage-sync.js loaded
  // simply never receive it, and fall back to the local parse time above.
  function wireSourceMetaListener() {
    document.addEventListener("pd-source-meta", (e) => {
      const meta = e.detail;
      if (meta && meta.uploadedAt) {
        STATE.sourceUploadedAt = meta.uploadedAt instanceof Date ? meta.uploadedAt : new Date(meta.uploadedAt);
        renderAsOfBanner();
      }
    });
  }

  document.addEventListener("DOMContentLoaded", () => {
    wireFilters();
    wireTabs();
    wireFileInput();
    wireSourceMetaListener();
    wireTableExportButtons();
    showNoData();
  });
})();
