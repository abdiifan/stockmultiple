// =============================================================================
// PharmaTrack v2 — expiry-risk.js
// Overstock & Expiry Risk Analysis.
//
// CORE IDEA
// ---------
// For every plant + item, we know:
//   - SOH            stock-on-hand right now (Unrestricted Stock)
//   - AMC            average monthly consumption at that plant
//                     (HO01 uses total branch demand — see mos.js HUB rule)
//   - MOS            = SOH ÷ AMC   → months of stock at current pace
//   - shelfLeftMo    months remaining until the earliest-expiring batch
//                     at that plant expires
//
// An item is AT RISK at a plant when:  MOS > shelfLeftMo
//   (there's more stock than can possibly be consumed before it expires)
//
// AT-RISK QUANTITY is only the part that can't be saved by normal consumption:
//   atRiskQty = max(0, SOH - shelfLeftMo * AMC)
//   atRiskVal = atRiskQty * unitValue   (unitValue = Value of Unrestricted
//               Stock ÷ Unrestricted Stock, from the inventory file)
//
// REDISTRIBUTION (per item, independently — see design discussion):
//   Source  = any plant (including HO01) with atRiskQty > 0.
//   Recipient = any OTHER plant (HO01 excluded — it never receives) that is
//               NOT itself at risk (its own MOS <= its own shelfLeftMo).
//   Recipient headroom = max(0, shelfLeftMo_recipient * AMC_recipient - SOH_recipient)
//               → the most that plant could absorb without becoming at-risk.
//   Source's atRiskQty is split across eligible recipients PROPORTIONALLY by
//   recipient AMC, each allocation capped at that recipient's headroom.
//   Whatever can't be placed (no eligible recipients, or headroom exhausted)
//   becomes RESIDUAL RISK — the number that goes to the marketing director.
//
// Requires: script.js (rawDf, mappingTable, fmtETB, fmtQty, escHtml, buildTable,
//           downloadCSV, downloadExcel, PLOTLY_LAYOUT, PLOTLY_CONFIG, waitForPlotly,
//           PAGE_RENDERERS, renderPage, currentPage)
//           mos.js (HUB_PLANT, mosMerged, mosPlants, buildMosSohMap)
// Must be loaded AFTER both script.js and mos.js.
// =============================================================================

const MS_PER_DAY   = 24 * 60 * 60 * 1000;
const DAYS_PER_MO  = 30.44; // average month length, consistent with rest of app's date math

// ── BUILD EXPIRY LOOKUP (earliest batch expiry per material+plant) ───────────
// materialCode → plantCode → { expiry: Date|null, unitVal: number }
function buildExpiryMap() {
  const map = new Map();
  // Use getReconciledBase() so the person filter (and mapping) applies here too
  const base = (typeof getReconciledBase === "function") ? getReconciledBase() : (typeof rawDf !== "undefined" ? rawDf : []);
  if (!base.length) return map;

  for (const row of base) {
    const mat = String(row._mappedMaterial || row["Material"] || "").trim();
    const plt = String(row["Plant"] || "").trim().toUpperCase();
    const qty = Number(row["Unrestricted Stock"] || 0);
    const val = Number(row["Value of Unrestricted Stock"] || 0);
    if (!mat || !plt || qty <= 0) continue;

    if (!map.has(mat)) map.set(mat, {});
    const plantMap = map.get(mat);
    if (!plantMap[plt]) plantMap[plt] = { expiry: null, valSum: 0, qtySum: 0 };

    const entry = plantMap[plt];
    entry.valSum += val;
    entry.qtySum += qty;

    // Earliest-expiring batch wins (pharma best practice, same rule used
    // elsewhere in this app when collapsing batches).
    const exp = row._expiry instanceof Date && !isNaN(row._expiry) ? row._expiry : null;
    if (exp && (!entry.expiry || exp < entry.expiry)) entry.expiry = exp;
  }

  return map;
}

function monthsUntil(date, today) {
  if (!(date instanceof Date) || isNaN(date)) return null;
  const days = (date.getTime() - today.getTime()) / MS_PER_DAY;
  return days / DAYS_PER_MO; // can be negative if already expired
}

function unitValueFor(entry) {
  if (!entry || !entry.qtySum) return 0;
  return entry.valSum / entry.qtySum;
}

// ── BUILD THE FULL PLANT × ITEM RISK SNAPSHOT ─────────────────────────────────
// Returns an array of { code, desc, type, plant, isHub, soh, amc, mos,
//                        shelfLeftMo, unitVal, atRisk, atRiskQty, atRiskVal }
function buildRiskSnapshot(typeFilter, searchQ, plantFilter) {
  if (typeof mosMerged === "undefined" || !mosMerged.length) return [];

  const sohMap    = buildMosSohMap();   // from mos.js — materialCode → plant → SOH
  const expiryMap = buildExpiryMap();
  const today     = new Date();
  today.setHours(0, 0, 0, 0);

  // getMosFilteredRows already applies the global personFilter before type/search
  let rows = (typeof getMosFilteredRows === "function")
    ? getMosFilteredRows(typeFilter || "", searchQ || "")
    : mosMerged.filter(r => (!typeFilter || r.type === typeFilter));

  const out = [];
  for (const r of rows) {
    const plantMos = computeRowMOS(r, sohMap); // from mos.js — per-plant {plant,soh,amc,mos,isHub}

    for (const pm of plantMos) {
      if (plantFilter && pm.plant !== plantFilter) continue;
      // Include ghost-stock rows (amc null, soh > 0) — they carry real product
      // that WILL expire. Without a consumption plan the entire SOH is at risk.
      // Only skip when there is genuinely no stock at this plant.
      if (!pm.soh || pm.soh <= 0) continue;

      const expEntry    = expiryMap.get(r.code)?.[pm.plant] || null;
      const shelfLeftMo  = expEntry ? monthsUntil(expEntry.expiry, today) : null;
      const unitVal      = unitValueFor(expEntry);

      // Need a shelf-life date to judge expiry risk.
      // For ghost-stock (amc null): mos = Infinity → the whole SOH is at risk.
      if (shelfLeftMo === null) continue;

      const isGhost   = pm.amc === null;           // no AMC commitment
      const effectiveMos = isGhost ? Infinity : pm.mos;

      // Skip if mos is finite and well within shelf life (no risk)
      if (effectiveMos !== null && effectiveMos !== Infinity && effectiveMos <= shelfLeftMo) {
        // Still push so it appears in the "safe" pool for redistribution headroom calc
        const safeQty = Math.max(0, shelfLeftMo) * (pm.amc || 0);
        out.push({
          code: r.code, desc: r.desc, type: r.type,
          isMerged: r.isMerged, origCodes: r.origCodes,
          plant: pm.plant, isHub: pm.isHub, isGhost,
          soh: pm.soh, amc: pm.amc, mos: effectiveMos,
          shelfLeftMo, unitVal,
          atRisk: false, atRiskQty: 0, atRiskVal: 0,
          headroom: Math.max(0, safeQty - pm.soh),
        });
        continue;
      }

      // Item is at risk (mos > shelfLeftMo, or mos === Infinity i.e. ghost stock)
      const safeQty   = isGhost ? 0 : Math.max(0, shelfLeftMo) * (pm.amc || 0); // ghost: nothing consumable
      const atRiskQty = Math.max(0, pm.soh - safeQty);
      const atRiskVal = atRiskQty * unitVal;

      out.push({
        code: r.code, desc: r.desc, type: r.type,
        isMerged: r.isMerged, origCodes: r.origCodes,
        plant: pm.plant, isHub: pm.isHub, isGhost,
        soh: pm.soh, amc: pm.amc, mos: effectiveMos,
        shelfLeftMo, unitVal,
        atRisk: atRiskQty > 0, atRiskQty, atRiskVal,
        headroom: 0,
      });
    }
  }
  return out;
}

// ── REDISTRIBUTION ENGINE ─────────────────────────────────────────────────────
// Works per material code: sources = at-risk rows (any plant incl. HO01),
// recipients = non-at-risk rows at OTHER plants for the SAME material,
// excluding HO01 as a recipient.
function computeRedistribution(snapshot) {
  const byCode = new Map();
  for (const row of snapshot) {
    if (!byCode.has(row.code)) byCode.set(row.code, []);
    byCode.get(row.code).push(row);
  }

  const transfers = [];          // individual source→recipient moves
  const residualByKey = new Map(); // `${code}|${plant}` → remaining unplaced qty/val

  for (const [code, rows] of byCode) {
    const sources    = rows.filter(r => r.atRisk && r.atRiskQty > 0)
                            .sort((a, b) => b.atRiskVal - a.atRiskVal); // highest ETB exposure claims headroom first
    const recipients = rows.filter(r => !r.isHub && !r.atRisk && r.headroom > 0);

    for (const src of sources) {
      // Recipients must be a DIFFERENT plant than the source (can't redistribute to self)
      const eligible = recipients.filter(rc => rc.plant !== src.plant);
      let remaining = src.atRiskQty;

      if (eligible.length && remaining > 0) {
        // Iterative proportional-by-AMC allocation, capped at each recipient's
        // remaining headroom. A single pass can strand usable headroom when one
        // recipient's cap is hit early (its unused share doesn't automatically
        // flow to recipients who still have room) — so we keep re-allocating
        // the leftover among recipients that still have headroom until either
        // the source's excess is fully placed or no recipient has room left.
        let pool = eligible.filter(rc => rc.headroom > 0);
        let toPlace = remaining;

        while (pool.length && toPlace > 1e-9) {
          const totalAmc = pool.reduce((s, rc) => s + rc.amc, 0);
          let placedThisRound = 0;

          for (const rc of pool) {
            const share = totalAmc > 0 ? (rc.amc / totalAmc) * toPlace : toPlace / pool.length;
            const alloc = Math.min(share, rc.headroom);
            if (alloc <= 0) continue;

            transfers.push({
              code, desc: src.desc, type: src.type,
              fromPlant: src.plant, fromIsHub: src.isHub,
              toPlant: rc.plant,
              qty: alloc, val: alloc * src.unitVal,
              toMosAfter: rc.amc > 0 ? (rc.soh + alloc) / rc.amc : null,
              toShelfLeftMo: rc.shelfLeftMo,
            });
            rc.headroom -= alloc;
            placedThisRound += alloc;
          }

          toPlace -= placedThisRound;
          pool = pool.filter(rc => rc.headroom > 1e-9);
          // Safety valve: if a round places nothing (shouldn't happen given the
          // headroom>0 filter, but guards against float edge cases), stop.
          if (placedThisRound <= 1e-9) break;
        }

        remaining = Math.max(0, toPlace);
      }

      remaining = Math.max(0, remaining);
      if (remaining > 0) {
        const key = `${code}|${src.plant}`;
        residualByKey.set(key, {
          code, desc: src.desc, type: src.type,
          plant: src.plant, isHub: src.isHub,
          qty: remaining, val: remaining * src.unitVal,
          unitVal: src.unitVal,
        });
      }
    }
  }

  return { transfers, residual: [...residualByKey.values()] };
}

// ── FORMATTING / FILTER HELPERS ────────────────────────────────────────────────
function exprKpiCard(label, value, sub, color) {
  return `<div class="kpi-card"><div class="kpi-label">${escHtml(label)}</div><div class="kpi-value" style="color:var(--${color||'blue'})">${value}</div>${sub ? `<div class="kpi-sub">${sub}</div>` : ""}</div>`;
}
function exprKpiRow(id, cards) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = cards.join("");
}

// ── MAIN RENDER ────────────────────────────────────────────────────────────────
async function renderExpiryRisk() {
  await waitForPlotly();

  const hasInventory = typeof rawDf !== "undefined" && rawDf.length > 0;
  const hasAmc        = typeof mosMerged !== "undefined" && mosMerged.length > 0;

  if (!hasInventory || !hasAmc) {
    document.getElementById("exprisk-no-data").style.display = "block";
    document.getElementById("exprisk-content").style.display = "none";
    return;
  }
  document.getElementById("exprisk-no-data").style.display  = "none";
  document.getElementById("exprisk-content").style.display = "block";

  const searchEl = document.getElementById("exprisk-search");
  const plantEl  = document.getElementById("exprisk-plant");
  const typeEl   = document.getElementById("exprisk-type");
  const searchQ  = searchEl ? searchEl.value.trim() : "";
  const plantVal = plantEl  ? plantEl.value.trim()  : "";
  const typeVal  = typeEl   ? typeEl.value.trim()   : "";

  if (plantEl && plantEl.options.length <= 1 && typeof mosPlants !== "undefined") {
    mosPlants.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p; opt.text = p === HUB_PLANT ? `${p} (Hub)` : p;
      plantEl.appendChild(opt);
    });
  }

  // ── Build snapshot (unfiltered by plant, so redistribution can see all plants
  // for each item) then apply plant filter only to the BEFORE view ──────────────
  const fullSnapshot = buildRiskSnapshot(typeVal, searchQ, "");
  const beforeRows   = plantVal ? fullSnapshot.filter(r => r.plant === plantVal) : fullSnapshot;
  const atRiskBefore = beforeRows.filter(r => r.atRisk && r.atRiskQty > 0);

  // ── KPIs: BEFORE ──────────────────────────────────────────────────────────────
  const totalAtRiskQtyBefore = atRiskBefore.reduce((s, r) => s + r.atRiskQty, 0);
  const totalAtRiskValBefore = atRiskBefore.reduce((s, r) => s + r.atRiskVal, 0);
  const hubAtRiskBefore      = atRiskBefore.filter(r => r.isHub);
  exprKpiRow("exprisk-kpis-before", [
    exprKpiCard("Plant-Item Pairs At Risk", atRiskBefore.length.toLocaleString(), `MOS > shelf-life remaining`, "red"),
    exprKpiCard("At-Risk Quantity", fmtQty(totalAtRiskQtyBefore), "units that may expire unused", "orange"),
    exprKpiCard("At-Risk Value", fmtETB(totalAtRiskValBefore), "Ethiopian Birr exposure", "red"),
    exprKpiCard(`${HUB_PLANT} Share`, fmtQty(hubAtRiskBefore.reduce((s,r)=>s+r.atRiskQty,0)), `${hubAtRiskBefore.length} hub item(s) at risk`, "purple"),
  ]);

  // ── CHART: BEFORE — items at risk aggregated across all plants (line chart) ──
  // Collapse plant-level rows into one entry per material (sum qty & val, keep
  // earliest shelf-life and worst MOS so the line reflects the true item-level risk).
  const itemRiskMap = new Map();
  for (const r of atRiskBefore) {
    if (!itemRiskMap.has(r.code)) {
      itemRiskMap.set(r.code, { code: r.code, desc: r.desc, atRiskQty: 0, atRiskVal: 0,
        shelfLeftMo: r.shelfLeftMo, mos: r.mos });
    }
    const e = itemRiskMap.get(r.code);
    e.atRiskQty  += r.atRiskQty;
    e.atRiskVal  += r.atRiskVal;
    // Worst-case shelf life (minimum across plants)
    if (r.shelfLeftMo !== null && (e.shelfLeftMo === null || r.shelfLeftMo < e.shelfLeftMo))
      e.shelfLeftMo = r.shelfLeftMo;
    // Highest MOS (most overstocked plant drives the risk score)
    if (r.mos !== null && (e.mos === null || r.mos > e.mos)) e.mos = r.mos;
  }
  const itemRiskArr = [...itemRiskMap.values()]
    .sort((a, b) => b.atRiskVal - a.atRiskVal)
    .slice(0, 30);

  if (itemRiskArr.length) {
    const labels = itemRiskArr.map(r => r.desc.length > 36 ? r.desc.slice(0, 36) + "…" : r.desc);
    Plotly.newPlot("chart-exprisk-before", [
      {
        // At-risk VALUE line (primary axis)
        type: "scatter", mode: "lines+markers",
        name: "At-Risk Value (ETB)",
        x: labels,
        y: itemRiskArr.map(r => r.atRiskVal),
        line: { color: "#f85149", width: 2.5 },
        marker: { color: "#f85149", size: 7 },
        hovertemplate: "<b>%{x}</b><br>At-risk value: ETB %{y:,.0f}<extra></extra>",
        yaxis: "y",
      },
      {
        // At-risk QTY line (secondary axis)
        type: "scatter", mode: "lines+markers",
        name: "At-Risk Qty (units)",
        x: labels,
        y: itemRiskArr.map(r => r.atRiskQty),
        line: { color: "#ffa657", width: 2, dash: "dot" },
        marker: { color: "#ffa657", size: 6 },
        hovertemplate: "<b>%{x}</b><br>At-risk qty: %{y:,.0f} units<extra></extra>",
        yaxis: "y2",
      },
    ], {
      ...PLOTLY_LAYOUT,
      height: 360,
      margin: { l: 60, r: 70, t: 24, b: 130 },
      xaxis: {
        tickangle: -38,
        tickfont: { size: 9.5 },
        showgrid: false,
      },
      yaxis:  { title: "At-Risk Value (ETB)", titlefont: { color: "#f85149" }, tickfont: { color: "#f85149" } },
      yaxis2: { title: "At-Risk Qty", titlefont: { color: "#ffa657" }, tickfont: { color: "#ffa657" },
                overlaying: "y", side: "right", showgrid: false },
      legend: { orientation: "h", y: 1.12, x: 0 },
      paper_bgcolor: "rgba(0,0,0,0)", plot_bgcolor: "rgba(0,0,0,0)",
    }, PLOTLY_CONFIG);
  } else {
    document.getElementById("chart-exprisk-before").innerHTML =
      '<div class="alert-info" style="margin:1rem 0">✓ No items currently at risk of expiring before they can be consumed.</div>';
  }

  // ── TABLE: BEFORE ──────────────────────────────────────────────────────────────
  const beforeCols = [
    { key: "code", label: "Material Code",
      fmt: (v, r) => r.isMerged
        ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-mapped-badge" title="Merged: ${escHtml(r.origCodes)}">MERGED</span>`
        : `<span class="col-mat-code">${escHtml(v)}</span>`,
      raw: true, cellClass: "col-mat-code-wrap" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "type", label: "Type" },
    { key: "plant", label: "Plant", fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v), raw: true },
    { key: "soh", label: "SOH", fmt: fmtQty },
    { key: "amc", label: "AMC",
      fmt: (v, r) => {
        if (r.isGhost || v === null) return `<span class="amc-soh-only-badge" title="No AMC commitment — entire SOH is at risk of expiry">No AMC</span>`;
        return `${fmtQty(v)}${r.isHub ? ' <span style="font-size:0.7em;color:var(--muted)">(Σ branch)</span>' : ""}`;
      }, raw: true },
    { key: "mos", label: "MOS",
      fmt: (v, r) => (r.isGhost || v === Infinity)
        ? `<span style="color:var(--amber);font-weight:700" title="No consumption plan — stock will not be drawn down">∞ mo</span>`
        : `<b>${Number(v).toFixed(1)}</b> mo`,
      raw: true },
    { key: "shelfLeftMo", label: "Shelf Life Left", fmt: v => v < 0 ? `<b style="color:var(--red)">EXPIRED</b>` : `<b>${v.toFixed(1)}</b> mo`, raw: true },
    { key: "atRiskQty", label: "At-Risk Qty", fmt: (v, r) => r.isGhost
        ? `<b style="color:var(--red)">${fmtQty(r.soh)}</b> <span style="font-size:0.7em;color:var(--amber)">(all SOH)</span>`
        : `<b style="color:var(--red)">${fmtQty(v)}</b>`, raw: true },
    { key: "atRiskVal", label: "At-Risk Value", fmt: v => `<b style="color:var(--red)">${fmtETB(v)}</b>`, raw: true },
  ];
  document.getElementById("exprisk-table-before").innerHTML = buildTable(
    [...atRiskBefore].sort((a,b)=>b.atRiskVal-a.atRiskVal), beforeCols, () => ""
  );

  // ── REDISTRIBUTION (always computed on the FULL unfiltered snapshot, so the
  //    plan is correct regardless of the plant filter applied to the view) ──────
  const { transfers, residual } = computeRedistribution(fullSnapshot);
  const visTransfers = plantVal ? transfers.filter(t => t.fromPlant === plantVal || t.toPlant === plantVal) : transfers;

  const redistCols = [
    { key: "code", label: "Material Code", cellClass: "col-mat-code-wrap" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "fromPlant", label: "From",
      fmt: (v, r) => r.fromIsHub ? `<b style="color:var(--purple)">${escHtml(v)} (Hub)</b>` : `<b style="color:var(--orange)">${escHtml(v)}</b>`,
      raw: true },
    { key: "toPlant", label: "To", fmt: v => `<b style="color:var(--blue)">${escHtml(v)}</b>`, raw: true },
    { key: "qty", label: "Transfer Qty", fmt: fmtQty },
    { key: "val", label: "Transfer Value", fmt: fmtETB },
    { key: "toMosAfter", label: "Recipient MOS After", fmt: v => v===null ? "—" : `${v.toFixed(1)} mo`, raw: true },
    { key: "toShelfLeftMo", label: "Recipient Shelf Life", fmt: v => `${v.toFixed(1)} mo`, raw: true },
  ];
  document.getElementById("exprisk-redist-table").innerHTML = visTransfers.length
    ? buildTable([...visTransfers].sort((a,b)=>b.val-a.val), redistCols, () => "")
    : '<div class="alert-info" style="margin:0.5rem 0">No eligible transfers found — either nothing is at risk, or no recipient plant has safe headroom for the at-risk items.</div>';

  // ── RESIDUAL (AFTER redistribution) — for marketing director ──────────────────
  // Exclude rows where Residual Qty is zero (or negligibly small due to float
  // arithmetic) — nothing actionable to show the marketing director.
  const visResidual = (plantVal ? residual.filter(r => r.plant === plantVal) : residual)
    .filter(r => r.qty > 1e-9);
  const totalResidualQty = visResidual.reduce((s, r) => s + r.qty, 0);
  const totalResidualVal = visResidual.reduce((s, r) => s + r.val, 0);
  const hubResidual = visResidual.filter(r => r.isHub);

  const recoveredQty = totalAtRiskQtyBefore - residual.reduce((s,r)=>s+r.qty,0); // network-wide, for context
  const recoveredPct = totalAtRiskQtyBefore > 0
    ? (((totalAtRiskQtyBefore - residual.reduce((s,r)=>s+r.qty,0)) / totalAtRiskQtyBefore) * 100).toFixed(1)
    : "0.0";

  exprKpiRow("exprisk-kpis-after", [
    exprKpiCard("Residual At-Risk Qty", fmtQty(totalResidualQty), "Could not be placed anywhere safely", "red"),
    exprKpiCard("Residual At-Risk Value", fmtETB(totalResidualVal), "Recommend for private sale / discount channel", "red"),
    exprKpiCard("Recovered by Redistribution", `${recoveredPct}%`, "of network-wide at-risk qty resolved by transfer", "green"),
    exprKpiCard(`From ${HUB_PLANT}`, fmtQty(hubResidual.reduce((s,r)=>s+r.qty,0)), `${hubResidual.length} item(s) from the hub`, "purple"),
  ]);

  const afterCols = [
    { key: "code", label: "Material Code", cellClass: "col-mat-code-wrap" },
    { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
    { key: "type", label: "Type" },
    { key: "plant", label: "Plant", fmt: (v, r) => r.isHub ? `<b>${escHtml(v)}</b> <span style="font-size:0.75em;color:var(--purple)">(Hub)</span>` : escHtml(v), raw: true },
    { key: "qty", label: "Residual Qty", fmt: v => `<b style="color:var(--red)">${fmtQty(v)}</b>`, raw: true },
    { key: "val", label: "Residual Value", fmt: v => `<b style="color:var(--red)">${fmtETB(v)}</b>`, raw: true },
    { key: "unitVal", label: "Unit Value", fmt: v => fmtETB(v) },
  ];
  document.getElementById("exprisk-table-after").innerHTML = visResidual.length
    ? buildTable([...visResidual].sort((a,b)=>b.val-a.val), afterCols, () => "")
    : '<div class="alert-info" style="margin:0.5rem 0">✓ Nothing left over — redistribution fully resolves the at-risk stock for the current filters.</div>';

  // ── EXPORT (export the marketing-director residual list, the most actionable one) ──
  const exportCols = [
    { key: "code", label: "Material Code" }, { key: "desc", label: "Description" }, { key: "type", label: "Type" },
    { key: "plant", label: "Plant" }, { key: "isHub", label: "Hub Plant?", fmt: v => v ? "Yes" : "No" },
    { key: "qty", label: "Residual Qty (units)", fmt: v => Number(v).toFixed(2) },
    { key: "val", label: "Residual Value (ETB)", fmt: v => Number(v).toFixed(2) },
    { key: "unitVal", label: "Unit Value (ETB)", fmt: v => Number(v).toFixed(2) },
  ];
  const dlRow = document.getElementById("exprisk-dl-row");
  if (dlRow) {
    dlRow.innerHTML = '<button class="dl-btn">⬇ CSV</button><button class="dl-btn">⬇ Excel</button>';
    dlRow.querySelectorAll(".dl-btn")[0].onclick = () => downloadCSV(visResidual,   exportCols, "expiry_risk_residual_for_marketing.csv");
    dlRow.querySelectorAll(".dl-btn")[1].onclick = () => downloadExcel(visResidual, exportCols, "expiry_risk_residual_for_marketing.xlsx");
  }
}

// ── WIRE INTO PAGE_RENDERERS AND EVENT LISTENERS ──────────────────────────────
(function wireExpiryRiskModule() {
  function extend() {
    if (typeof PAGE_RENDERERS !== "undefined") {
      PAGE_RENDERERS["expiry-risk"] = renderExpiryRisk;
    }

    const _origRenderPage = window.renderPage;
    window.renderPage = function (id) {
      if (id === "expiry-risk") {
        currentPage = id;
        document.getElementById("landingView").style.display = "none";
        document.querySelectorAll(".page").forEach(el => { el.style.display = "none"; });
        const pg = document.getElementById("page-expiry-risk");
        if (pg) pg.style.display = "block";
        document.querySelectorAll(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.page === id));
        try { renderExpiryRisk(); } catch (e) { console.error(e); }
        return;
      }
      _origRenderPage(id);
    };

    const filterMap = {
      "exprisk-apply": renderExpiryRisk,
      "exprisk-clear": () => {
        const s = document.getElementById("exprisk-search"); if (s) s.value = "";
        const p = document.getElementById("exprisk-plant");  if (p) p.value = "";
        const t = document.getElementById("exprisk-type");   if (t) t.value = "";
        renderExpiryRisk();
      },
    };

    document.body.addEventListener("click", (e) => {
      const btn = e.target.closest("button[id]");
      if (!btn) return;
      const fn = filterMap[btn.id];
      if (fn) { e.stopPropagation(); fn(); }
    }, true);

    // Re-render if currently on this page and either source file changes
    const fileInput    = document.getElementById("fileInput");
    const mosAmcInput   = document.getElementById("mosAmcFileInput");
    [fileInput, mosAmcInput].forEach(inp => {
      if (!inp) return;
      inp.addEventListener("change", () => {
        setTimeout(() => { if (currentPage === "expiry-risk") renderExpiryRisk(); }, 350);
      });
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", extend);
  } else {
    extend();
  }
})();
