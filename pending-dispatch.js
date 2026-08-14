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

    wrap.innerHTML = subTabBar + `<div class="pd-subtab-panel">${panel}</div>`;

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

    wrap.innerHTML = `
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
    wrap.innerHTML = `
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

  // ── "Branch × Storage Location" cross-tab (line-item counts) ────
  function renderMatrixTable(rows) {
    const wrap = document.getElementById("pd-matrix-wrap");
    if (!wrap) return;

    const matrix = {};       // matrix[branch][sloc] = count
    const branchTotals = {};
    const slocTotals = {};
    const slocSeen = [];     // first-seen order, used to build column order below
    let grandTotal = 0;

    rows.forEach((r) => {
      const branch = branchName(plantCode(r.shipToParty)) || "—";
      const sloc = r.storageLocation || "—";
      if (!matrix[branch]) matrix[branch] = {};
      matrix[branch][sloc] = (matrix[branch][sloc] || 0) + 1;
      branchTotals[branch] = (branchTotals[branch] || 0) + 1;
      slocTotals[sloc] = (slocTotals[sloc] || 0) + 1;
      if (!slocSeen.includes(sloc)) slocSeen.push(sloc);
      grandTotal += 1;
    });

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

    const maxCell = Math.max(1, ...branches.flatMap((b) => cols.map((c) => matrix[b]?.[c] || 0)));
    const maxRowTotal = Math.max(1, ...branches.map((b) => branchTotals[b] || 0));
    const maxColTotal = Math.max(1, ...cols.map((c) => slocTotals[c] || 0));

    const headerCells = cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("");

    const bodyRows = branches.map((b, i) => {
      const cells = cols.map((c) => {
        const v = matrix[b]?.[c] || 0;
        return `<td style="${heatBg(v, maxCell)}">${v ? v.toLocaleString() : ""}</td>`;
      }).join("");
      return `
        <tr>
          <td>${i + 1}</td>
          <td style="text-align:left">${escapeHtml(b)}</td>
          ${cells}
          <td style="${heatBg(branchTotals[b], maxRowTotal)};font-weight:700">${branchTotals[b].toLocaleString()}</td>
        </tr>`;
    }).join("");

    const totalCells = cols.map((c) => `<td style="${heatBg(slocTotals[c] || 0, maxColTotal)};font-weight:700">${(slocTotals[c] || 0).toLocaleString()}</td>`).join("");
    const pctCells = cols.map((c) => {
      const pct = grandTotal ? Math.round(((slocTotals[c] || 0) / grandTotal) * 100) : 0;
      return `<td style="background:#dbe9fb;font-weight:600">${pct}%</td>`;
    }).join("");

    wrap.innerHTML = `
      <div class="tbl-wrap">
        <table class="pd-matrix-table">
          <thead>
            <tr>
              <th>SN</th><th style="text-align:left">Branch</th>${headerCells}<th>Total</th>
            </tr>
          </thead>
          <tbody>
            ${bodyRows}
            <tr class="pd-matrix-total-row">
              <td colspan="2">Total</td>
              ${totalCells}
              <td style="font-weight:700">${grandTotal.toLocaleString()}</td>
            </tr>
            <tr class="pd-matrix-pct-row">
              <td colspan="2"></td>
              ${pctCells}
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

  function renderDetailTable(rows) {
    const wrap = document.getElementById("pd-detail-wrap");

    if (!rows.length) {
      wrap.innerHTML = `<div class="pd-empty">No pending line items match the current filters.</div>`;
      return;
    }

    const sorted = [...rows].sort((a, b) => (a.giDate && b.giDate ? a.giDate - b.giDate : 0));
    const shown = sorted; // no render cap — full filtered set is shown

    wrap.innerHTML = `
      <div class="tbl-wrap tbl-wrap-freeze">
        <table class="freeze-header">
          <thead><tr>
            <th>Delivery</th><th>GI Date</th><th>Days Late</th><th>Material</th>
            <th>Description</th><th>Branch</th><th>Storage Loc.</th>
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
      "Goods Issue Date": r.giDate ? fmtDate(r.giDate) : "",
      "Days Late": (() => { const d = daysLate(r.giDate); return d === null ? "" : (d <= 0 ? "Good" : d); })(),
      "Material": r.material,
      "Item Description": r.itemDescription,
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

  function exportRows(rows, kind) {
    if (kind === "csv") {
      // CSV has no concept of multiple sheets — keep it as the flat
      // "All Pending Items" detail export.
      const ws = XLSX.utils.json_to_sheet(buildDetailExportData(rows));
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      downloadBlob(blob, "pending_dispatch.csv");
      return;
    }

    // XLSX export: one sheet per table, in the same order as the tabs —
    // All Pending Items, By Storage Location, By Branch / Plant, Top 10s.
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDetailExportData(rows)), safeSheetName("All Pending Items"));
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

    document.getElementById("pd-tab-count-top10").textContent  = Math.min(multiCount, 10).toLocaleString();
    document.getElementById("pd-tab-count-sloc").textContent   = slocCount.toLocaleString();
    document.getElementById("pd-tab-count-branch").textContent = branchCount.toLocaleString();
    document.getElementById("pd-tab-count-detail").textContent = filtered.length.toLocaleString();
    document.getElementById("pd-tab-count-matrix").textContent = branchCount.toLocaleString();
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

  document.addEventListener("DOMContentLoaded", () => {
    wireFilters();
    wireTabs();
    wireFileInput();
    showNoData();
  });
})();
