// =============================================================================
// PharmaTrack v2 — request-analysis.js
// "🧾 Request Analysis" — self-serve sidebar tool. Any user (not just admins)
// uploads their OWN Transfer Requests Excel (Purchase Req Num, Poste, Material,
// Short Text, Requested Quantity, Stock on hand, Delivery date, Created By,
// Location, Plant) and instantly sees it reconciled against the currently-loaded
// HO01 (hub) stock. Nothing is saved to a shared database — the uploaded file
// lives only in this browser tab's memory, exactly like the person who
// uploaded it intended.
//
// LOCATION MISMATCH CHECK
// ------------------------
// The request file's "Location" column is the Storage Location the requester
// typed for that line — this is where the REQUESTING plant/branch itself
// keeps the item, NOT HO01. It's compared against the real Storage
// Location(s) that material actually sits in AT THE REQUESTING PLANT
// (row["Storage Location"] on the main inventory file, restricted to
// reqPlant — the same plant this whole file is scoped to). If the
// requester's material resolves fine and storage-location data exists for
// it at their own plant, but the typed Location isn't one of that plant's
// actual storage locations for the material, the line is flagged
// "⚠ check location" — this catches requests sent against the wrong
// storage location even when the material/code itself is correct. If the
// requesting plant has no storage-location data for the material at all,
// it's left unverified (shown as "—") rather than guessed at as a mismatch.
//
// PLANT SCOPING
// -------------
// The uploaded file is always for ONE requesting plant at a time (a branch
// pasting its own transfer requests). The "Plant" column is used to:
//   1. Label the analysis ("Requesting Plant: GO01") for clarity.
//   2. Scope Tab 4 ("HO01 Stock Not Requested") criticality so it only flags
//      HO01 stock as critical when it's THIS branch (not some other branch)
//      that's running low — otherwise it would surface items that are fine
//      for the branch that actually uploaded the file.
// If a file somehow contains more than one distinct Plant value, the most
// frequent one is used for scoping and a warning is shown.
// "Created By" is carried through purely for visibility (shown as a column
// in the main Request vs Stock table) — it has no effect on the analysis.
//
// WHAT THIS ANALYSIS SHOWS
// -------------------------
// 1. Request vs Stock (side-by-side) — every request line, with the request
//    file's OWN "Stock on hand" column shown next to a LIVE recomputed HO01
//    stock figure (from the currently loaded main inventory), so any mismatch
//    between what the requester's system said and what HO01 actually has
//    right now is visible at a glance.
// 2. Suggested Code Corrections — request lines whose canonical material has
//    live stock at HO01, but NOT (only) under the exact code the requester
//    typed. The mapping file's target/"standard" code is used ONLY to figure
//    out which raw SAP codes belong to the same material — it is never
//    itself the suggestion, because it may not be a real, orderable SAP code
//    with its own stock record. The suggestion is always the actual raw SAP
//    code(s) that currently carry stock at HO01 for that material, each with
//    its own live quantity (in that code's own native unit, exactly as SAP
//    shows it — NOT the standardized/converted number used for totals). If
//    more than one raw code carries stock, ALL of them are surfaced together
//    (e.g. "115-ZOLE-0301-01 (120) or 115-ZOLE-0301-02 (15)") so a requester
//    isn't steered toward a nearly-empty code when a fuller one exists.
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
// exists, its target code identifies which OTHER raw SAP codes are the same
// material (for grouping/aggregation only — see point 2 above). If no
// mapping entry exists, we fall back to trying the raw code as-is (covers
// cases where a request already used the real SAP code).
//
// TOTALS VS. SUGGESTIONS — two different quantities, on purpose
// ----------------------------------------------------------------
// - "Live HO01 stock" for status/totals (Tab 1, Tab 3, Tab 4, KPIs) is always
//   the STANDARDIZED total across every raw code that maps to the canonical
//   material (sohMap from buildMosSohMap(), same converted numbers the rest
//   of the app uses) — this answers "is there stock, in total, right now."
// - "Suggested code" quantities (Tab 2) are the RAW, unconverted stock under
//   each individual SAP code (Unrestricted + verified Transit + QC, in that
//   code's own unit) — this answers "which exact code do I type into SAP,
//   and how much is really under it." Only codes with a live inventory row
//   and stock > 0 are ever suggested — nothing from the mapping table alone.
//
// Requires: script.js (rawDf, mappingTable, escHtml, fmtQty, kpiCard, buildTable,
//           wireTableExport, downloadCSV, downloadExcel, parseExpiryDate,
//           fmtLocalDate, getReconciledBase, PAGE_RENDERERS, renderPage, currentPage,
//           buildMultiSelect)
//           mos.js (HUB_PLANT, buildMosSohMap)
// Must be loaded AFTER both script.js and mos.js.
// =============================================================================

(function requestAnalysisModule() {

  // ── STATE ──────────────────────────────────────────────────────────────────
  // Lives only in memory for this browser tab/session — never written to any
  // shared store. Re-uploading replaces it; closing the tab discards it.
  let reqRows   = [];   // parsed request lines
  let reqFileName = "";
  let reqPlant  = "";   // the (single) requesting plant this file is for
  let reqPlantMismatch = false; // true if the file had more than one distinct Plant value

  // Material Type filter (e.g. ZME, ZMS…) — multi-select. Empty set = no
  // filter applied (show everything). Populated from the "Material Type"
  // column on the main inventory data (rawDf), keyed by canonical code, and
  // applies across all 4 tabs.
  let reqMatTypeFilter = new Set();

  // Material Group filter — same pattern as Material Type above, but sourced
  // from the literal "Material Group Name" column on the main inventory data
  // (rawDf), not a helper function. Multi-select; empty set = no filter.
  let reqMatGroupFilter = new Set();

  // FEAT-COPY-CODES: click-to-select on the "Requested Code" cells in Tab 1
  // (Request vs Stock table) so users can grab several material codes at
  // once without click-dragging across the whole table (which also grabs
  // Description/Qty/SOH text from other columns). Selecting persists across
  // re-renders (filter changes, etc.) by code value, not DOM node.
  let reqCodeCopySelection = new Set();


  const REQUIRED_COLS = [
    "Purchase Req Num", "Poste", "Material", "Short Text",
    "Requested Quantity", "Stock on hand", "Delivery date",
    "Created By", "Plant", "Location",
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
              createdBy: String(get(row, "Created By") ?? "").trim(),
              plant:     String(get(row, "Plant") ?? "").trim().toUpperCase(),
              location:  String(get(row, "Location") ?? "").trim().toUpperCase(),
            }))
            .filter(r => r.material);

          if (!parsed.length) { showReqError("No valid rows with a Material code were found."); return; }

          // This file is expected to be from ONE requesting plant. Take the
          // most frequent Plant value as the file's plant; flag if mixed.
          const plantCounts = new Map();
          parsed.forEach(r => { if (r.plant) plantCounts.set(r.plant, (plantCounts.get(r.plant) || 0) + 1); });
          const plantEntries = [...plantCounts.entries()].sort((a, b) => b[1] - a[1]);
          reqPlant = plantEntries.length ? plantEntries[0][0] : "";
          reqPlantMismatch = plantEntries.length > 1;

          if (!reqPlant) { showReqError("No Plant value found — every row is missing a Plant."); return; }

          reqRows = parsed;
          reqFileName = file.name;

          if (statusEl) {
            const mismatchNote = reqPlantMismatch
              ? `<div class="status-name" style="color:var(--amber,#d97706)">⚠ Multiple Plant values found — using ${escHtml(reqPlant)} (most common) for scoping</div>`
              : "";
            statusEl.innerHTML =
              `<div class="status-ok">✓ FILE LOADED</div>` +
              `<div class="status-name">${escHtml(file.name)} (${parsed.length.toLocaleString()} lines) · Plant ${escHtml(reqPlant)}</div>` +
              mismatchNote;
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
    reqPlant = "";
    reqPlantMismatch = false;
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

  // ── HO01 STOCK BY RAW SAP CODE (for Suggested Code Corrections) ────────────
  // Returns Map<canonicalCode, Array<{code, qty}>> — for every canonical
  // material, the individual raw SAP codes that currently carry stock at
  // HO01 and how much (RAW/native units, NOT converted), sorted highest
  // quantity first. Codes with zero or no stock are never included, so
  // anything in this map is, by construction, "actually in SAP right now."
  //
  // Deliberately mirrors buildMosSohMap()'s Total Quantity definition
  // (Unrestricted + verified Transit + QC) so the numbers agree with the
  // rest of the app — the only difference is grouping by the RAW code
  // (row["Material"]) instead of the canonical code, and using the raw
  // (pre-conversion) fields instead of the standardized _cv* fields, since
  // this needs to reflect exactly what's under that one specific SAP code.
  function buildHo01RawCodeMap(hub) {
    const map = new Map(); // canonical -> Map<rawCode, qty>
    const base = (typeof getReconciledBase === "function") ? getReconciledBase() : (typeof rawDf !== "undefined" ? rawDf : []);
    base.forEach(row => {
      const plt = String(row["Plant"] || "").trim().toUpperCase();
      if (plt !== hub) return;

      const canonical = String(row._mappedMaterial || row["Material"] || "").trim();
      const rawCode   = String(row["Material"] || "").trim();
      if (!canonical || !rawCode) return;

      const unrestricted   = Number(row["Unrestricted Stock"] || 0);
      const rawTransit     = Number(row["Stock in Transit"] || 0);
      const phantomTransit = Number(row._phantomTransitQty || 0);
      const verifiedTransit = Math.max(0, rawTransit - phantomTransit);
      const qc              = Number(row["Stock in Quality Inspection"] || 0);
      const qty = unrestricted + verifiedTransit + qc;
      if (qty <= 0) return;

      if (!map.has(canonical)) map.set(canonical, new Map());
      const inner = map.get(canonical);
      inner.set(rawCode, (inner.get(rawCode) || 0) + qty);
    });

    const out = new Map();
    map.forEach((inner, canonical) => {
      const list = [...inner.entries()]
        .map(([code, qty]) => ({ code, qty }))
        .sort((a, b) => b.qty - a.qty);
      out.set(canonical, list);
    });
    return out;
  }

  // Canonical code -> total stock currently sitting "Stock in Quality
  // Inspection" at the hub, summed raw (native units) across every raw SAP
  // code that maps to that canonical material. This is stock that's
  // physically at HO01 but not yet released/usable — useful context next to
  // the live SOH figure, since it's already counted inside that total.
  function buildCanonicalQcMap(hub) {
    const out = new Map();
    const base = (typeof getReconciledBase === "function") ? getReconciledBase() : (typeof rawDf !== "undefined" ? rawDf : []);
    base.forEach(row => {
      const plt = String(row["Plant"] || "").trim().toUpperCase();
      if (plt !== hub) return;
      const canonical = String(row._mappedMaterial || row["Material"] || "").trim();
      if (!canonical) return;
      const qc = Number(row["Stock in Quality Inspection"] || 0);
      if (!qc) return;
      out.set(canonical, (out.get(canonical) || 0) + qc);
    });
    return out;
  }

  // Canonical code -> Material Type (e.g. "ZME", "ZMS"…), sourced the same
  // way the rest of the app derives it (Dashboard/Transit/Expiry/QC filter
  // bars all use getValuationType(row) — see script.js — NOT a literal
  // "Material Type" column, which doesn't exist under that name in the SAP
  // export). This previously read row["Material Type"] directly, which was
  // always blank/undefined, so this map came back empty regardless of what
  // was loaded — the filter looked "connected" but could never resolve a
  // single type. Keyed the same way as buildCanonicalDescMap (first
  // non-blank value wins). Used to power the Material Type filter bar.
  function buildMaterialTypeMap() {
    const out = new Map();
    const base = (typeof getReconciledBase === "function") ? getReconciledBase() : (typeof rawDf !== "undefined" ? rawDf : []);
    base.forEach(row => {
      const code = String(row._mappedMaterial || row["Material"] || "").trim();
      if (!code || out.has(code)) return;
      const type = (typeof getValuationType === "function" ? String(getValuationType(row) || "") : "").trim().toUpperCase();
      if (type && type !== "(NONE)") out.set(code, type);
    });
    return out;
  }

  // Canonical code -> Material Group, sourced directly from the literal
  // "Material Group Name" column on the main inventory data (rawDf) — this
  // is the real SAP field name used throughout script.js (Dashboard,
  // Branch Comparison, Expiry Risk, etc. all read row["Material Group Name"],
  // NOT "Material Group" — that key doesn't exist, which is why this filter
  // showed "Unavailable" even with data loaded). getReconciledBase() already
  // excludes non-medical groups (isNonMedicalGroup) before we ever see it,
  // same as every other Material Group control in the app. Keyed the same
  // way as buildCanonicalDescMap/buildMaterialTypeMap (first non-blank value
  // wins). Used to power the Material Group filter bar.
  function buildMaterialGroupMap() {
    const out = new Map();
    const base = (typeof getReconciledBase === "function") ? getReconciledBase() : (typeof rawDf !== "undefined" ? rawDf : []);
    base.forEach(row => {
      const code = String(row._mappedMaterial || row["Material"] || "").trim();
      if (!code || out.has(code)) return;
      const group = String(row["Material Group Name"] || "").trim();
      if (group) out.set(code, group);
    });
    return out;
  }

  // Canonical code -> Set of Storage Locations that carry it at a given
  // plant, sourced from the literal "Storage Location" column on the main
  // inventory data — the same field script.js's own Storage Location
  // filters/columns use. Used to catch requests where the requester typed
  // the WRONG storage location for an otherwise-correct material — checked
  // against the REQUESTING plant's own storage locations (e.g. plant HA01's
  // real locations), NOT HO01's, since the "Location" column in the request
  // file is where the requesting branch itself keeps the item, not where
  // HO01 keeps it. Includes every storage location seen for the material at
  // that plant, not just ones with current stock > 0, since a location can
  // still be the "right" one to request against even if it's momentarily
  // empty.
  function buildPlantStorageLocationMap(plant) {
    const out = new Map(); // canonical -> Set<string>
    if (!plant) return out;
    const base = (typeof getReconciledBase === "function") ? getReconciledBase() : (typeof rawDf !== "undefined" ? rawDf : []);
    base.forEach(row => {
      const plt = String(row["Plant"] || "").trim().toUpperCase();
      if (plt !== plant) return;
      const canonical = String(row._mappedMaterial || row["Material"] || "").trim();
      const loc = String(row["Storage Location"] || "").trim().toUpperCase();
      if (!canonical || !loc) return;
      if (!out.has(canonical)) out.set(canonical, new Set());
      out.get(canonical).add(loc);
    });
    return out;
  }

  // Canonical code -> responsible person, sourced from mosMerged (same
  // "r.person" field who-responsible.js / the global sidebar person filter
  // use) so "assigned to" here means the same thing it means everywhere
  // else in the app — NOT the request file's "Created By" column.
  function buildPersonMap() {
    const out = new Map();
    if (typeof mosMerged !== "undefined" && mosMerged.length) {
      mosMerged.forEach(r => { if (r.code) out.set(r.code, r.person || ""); });
    }
    return out;
  }

  // ── CORE ANALYSIS ──────────────────────────────────────────────────────────
  function buildRequestAnalysis() {
    const hub    = (typeof HUB_PLANT === "function" || typeof HUB_PLANT !== "undefined") ? HUB_PLANT : "HO01";
    const sohMap = (typeof buildMosSohMap === "function") ? buildMosSohMap() : new Map();
    const descMap = buildCanonicalDescMap();
    const rawCodeMap = buildHo01RawCodeMap(hub); // canonical -> [{code, qty}], live SAP codes only
    const qcMap = buildCanonicalQcMap(hub); // canonical -> stock currently in Quality Inspection at hub
    const personMap = buildPersonMap();
    const matTypeMap = buildMaterialTypeMap(); // canonical -> "ZME"/"ZMS"/…
    const matGroupMap = buildMaterialGroupMap(); // canonical -> "Material Group" value
    const storageLocMap = buildPlantStorageLocationMap(reqPlant); // canonical -> Set of Storage Locations at the REQUESTING plant

    const requestedCanonical = new Set();

    const rows = reqRows.map(r => {
      const resolved   = resolveRequestMaterial(r.material);
      const canonical  = resolved.canonical;
      const inInventory = !!canonical && sohMap.has(canonical);
      // "Total" / status figures stay STANDARDIZED (converted, summed across
      // every raw code for this material) — this is unchanged and intentional.
      const liveHo01   = inInventory ? (sohMap.get(canonical)[hub] || 0) : 0;
      const desc       = r.shortText || resolved.desc || descMap.get(canonical) || "";
      // Description of the MAPPED/canonical material the live HO01 figure was
      // actually pulled for — separate from `desc` above, which prefers the
      // requester's own free-text Short Text. This always reflects the
      // canonical code's description (mapping table target, or live-inventory
      // lookup), so it stays accurate even when the requester's Short Text
      // doesn't match the resolved material.
      const mappedDesc = resolved.desc || descMap.get(canonical) || "";
      const qcHo01     = canonical ? (qcMap.get(canonical) || 0) : 0;
      // Live current stock at the REQUESTING plant itself (not HO01) — same
      // sohMap used for the HO01 figure, just keyed by the branch's own plant
      // code instead of the hub, so it reflects what's actually there right
      // now for comparison against what the request file claims.
      const liveReqPlant = (canonical && reqPlant && sohMap.has(canonical)) ? (sohMap.get(canonical)[reqPlant] || 0) : 0;
      const person     = canonical ? (personMap.get(canonical) || "") : "";
      const materialType = canonical ? (matTypeMap.get(canonical) || "") : "";
      const materialGroup = canonical ? (matGroupMap.get(canonical) || "") : "";

      // Storage Location check: is the location the requester typed one of
      // the actual Storage Location(s) this material sits in at HO01? Only
      // flagged when we HAVE location data for this material at HO01 (an
      // empty/unknown set means "can't verify" — not a mismatch), the
      // requester actually typed a location, and it's a real, resolvable
      // material — same "don't guess" principle used for AMC criticality.
      const actualLocations = canonical ? [...(storageLocMap.get(canonical) || [])].sort() : [];
      const locationKnown = actualLocations.length > 0;
      const locationMismatch = !!canonical && !!r.location && locationKnown && !actualLocations.includes(r.location);

      let status;
      if (!canonical || !inInventory) status = "no-match";
      else if (liveHo01 <= 0)          status = "stockout";
      else                             status = "ok";

      // The SUGGESTION is a different thing: which actual, live SAP code(s)
      // carry that stock right now, in their own native quantities.
      //
      // A suggestion is only shown when the code the requester typed can NOT
      // fully cover the requested quantity by itself:
      //   - If the typed code alone already has enough stock to cover the
      //     requested qty, NO suggestion is shown — even if other codes also
      //     happen to carry stock for the same item (nothing to fix).
      //   - If the typed code has stock but not enough (partial) AND another
      //     live code exists for the same item, the suggestion combo
      //     includes the typed code itself plus the other code(s), so the
      //     requester sees the full combination needed to fulfill.
      //   - If the typed code has zero live stock but another code exists,
      //     that other code (or codes) is suggested — even if their combined
      //     total still doesn't fully cover the requested qty, since partial
      //     stock elsewhere is still useful to know.
      //   - If there's no OTHER live code at all (only the typed code, or
      //     nothing), there's nothing to suggest either way.
      const rawCandidates = canonical ? (rawCodeMap.get(canonical) || []) : [];
      const typedCode = String(r.material || "").trim().toUpperCase();
      const typedEntry = rawCandidates.find(c => c.code.toUpperCase() === typedCode);
      const typedQty = typedEntry ? typedEntry.qty : 0;
      const otherCandidates = rawCandidates.filter(c => c.code.toUpperCase() !== typedCode);

      const typedFullyCovers = typedQty > 0 && typedQty >= r.reqQty;
      const hasSuggestion = !typedFullyCovers && otherCandidates.length > 0;
      const suggestionCandidates = hasSuggestion
        ? (typedEntry ? [typedEntry, ...otherCandidates] : otherCandidates)
        : [];

      const sohMismatch = inInventory && Math.abs(liveHo01 - r.reqSoh) > 0.001;

      if (canonical && inInventory) requestedCanonical.add(canonical);

      return {
        ...r,
        canonical, desc, status, person, materialType, materialGroup,
        liveHo01, mappedDesc, qcHo01, liveReqPlant, sohMismatch,
        actualLocations, locationKnown, locationMismatch,
        hasSuggestion,
        suggestedCode: hasSuggestion
          ? suggestionCandidates.map(c => `${c.code} (${fmtQty(c.qty)})`).join(" or ")
          : null,
        suggestedDesc: hasSuggestion ? (resolved.desc || descMap.get(canonical) || "") : null,
        suggestedTotal: hasSuggestion ? suggestionCandidates.reduce((s, c) => s + c.qty, 0) : 0,
      };
    });

    // ── DOUBLE REQUEST DETECTION ─────────────────────────────────────────────
    // Same physical item requested more than once in THIS file — same or
    // different raw code. Grouping key is the canonical code when resolvable;
    // when it isn't (no mapping match), we still group by the raw code as
    // typed, so exact-duplicate lines are caught even with no mapping loaded.
    const dupGroups = new Map(); // key -> array of row indices
    rows.forEach((r, i) => {
      const key = r.canonical || `__raw__${r.material.toUpperCase()}`;
      if (!dupGroups.has(key)) dupGroups.set(key, []);
      dupGroups.get(key).push(i);
    });
    rows.forEach((r, i) => {
      const key = r.canonical || `__raw__${r.material.toUpperCase()}`;
      const group = dupGroups.get(key);
      r.isDuplicate = group.length > 1;
      if (r.isDuplicate) {
        const siblings = group.filter(j => j !== i).map(j => rows[j]);
        r.duplicateCount = group.length;
        r.duplicateTotalQty = group.reduce((s, j) => s + (rows[j].reqQty || 0), 0);
        r.duplicateSiblingsLabel = siblings
          .map(s => `${s.material} · PR ${s.prNum}${s.poste ? "/" + s.poste : ""} (${fmtQty(s.reqQty)})`)
          .join("; ");
      } else {
        r.duplicateCount = 1;
        r.duplicateTotalQty = r.reqQty;
        r.duplicateSiblingsLabel = "";
      }
    });

    // HO01 stock that never shows up (under its canonical code) in the request at all
    const ho01NotRequestedAll = [];
    sohMap.forEach((plantMap, code) => {
      const qty = plantMap[hub] || 0;
      if (qty > 0 && !requestedCanonical.has(code)) {
        ho01NotRequestedAll.push({ code, desc: descMap.get(code) || "", qty, person: personMap.get(code) || "", materialType: matTypeMap.get(code) || "", materialGroup: matGroupMap.get(code) || "" });
      }
    });

    // Per clarified requirement: only surface items where THIS request file's
    // own requesting plant (reqPlant) — never HO01 itself, the hub has no
    // consumption of its own — is running critical (MOS < 1 month). This is
    // now scoped to reqPlant specifically (not "any branch"), since the
    // analysis is always for one requesting plant vs HO01. Requires the AMC
    // file (MOS by Plant page) to be loaded — mosMerged/computeRowMOS/
    // isMosCritical come from mos.js.
    const mosDataLoaded = typeof mosMerged !== "undefined" && mosMerged.length > 0
      && typeof computeRowMOS === "function" && typeof isMosCritical === "function";

    let ho01NotRequested = [];
    if (mosDataLoaded) {
      ho01NotRequested = ho01NotRequestedAll
        .map(r => {
          const amcRow = mosMerged.find(m => m.code === r.code);
          // No AMC commitment data at all for this material -> can't confirm
          // it's critical anywhere, so don't flag it (avoids false positives).
          const criticalBranches = amcRow
            ? computeRowMOS(amcRow, sohMap).filter(p =>
                !p.isHub &&
                String(p.plant || "").trim().toUpperCase() === reqPlant &&
                isMosCritical(p.mos))
            : [];
          return { ...r, criticalBranches };
        })
        .filter(r => r.criticalBranches.length > 0)
        // Most urgent (lowest MOS among its critical branches) first.
        .sort((a, b) => Math.min(...a.criticalBranches.map(c => c.mos)) - Math.min(...b.criticalBranches.map(c => c.mos)));
    }

    // All distinct Material Types present in this analysis (request lines +
    // HO01 stock not requested), sorted alphabetically — powers the filter
    // bar's option list. Blank/unknown types are excluded from the list
    // itself (there's nothing meaningful to filter on for them), but their
    // rows remain visible whenever the filter is inactive.
    const availableMatTypes = [...new Set([
      ...rows.map(r => r.materialType),
      ...ho01NotRequestedAll.map(r => r.materialType),
    ].filter(Boolean))].sort();

    // FIX-MATTYPE-EMPTY (corrected): an earlier version of this fallback
    // quietly offered ZME/ZMS/ZLC/ZMD as clickable options whenever
    // availableMatTypes came back empty — but every row's materialType is
    // ALSO "" in that exact situation (matTypeMap has nothing to key off
    // of, see buildMaterialTypeMap()), so those options matched zero rows
    // and made the filter look broken ("I picked ZME and everything
    // disappeared"). Surface the real reason instead: whether Material Type
    // could be resolved for ANY material at all (matTypeMap wasn't empty).
    const matTypeDataAvailable = matTypeMap.size > 0;

    // Same idea as availableMatTypes/matTypeDataAvailable above, for Material
    // Group — sourced straight from the literal column, so "data available"
    // just means at least one row had that column populated.
    const availableMatGroups = [...new Set([
      ...rows.map(r => r.materialGroup),
      ...ho01NotRequestedAll.map(r => r.materialGroup),
    ].filter(Boolean))].sort();
    const matGroupDataAvailable = matGroupMap.size > 0;

    return {
      rows, ho01NotRequested, ho01NotRequestedAllCount: ho01NotRequestedAll.length, mosDataLoaded,
      availableMatTypes, matTypeDataAvailable,
      availableMatGroups, matGroupDataAvailable,
    };
  }

  // ── MATERIAL TYPE FILTER BAR ────────────────────────────────────────────────
  // Multi-select dropdown injected inline next to the existing status filter
  // (reqan-status-filter). Same searchable checkbox-dropdown control used for
  // Material Type / Material Group / Plant everywhere else in the app (see
  // script.js's buildMultiSelect() + the .ms-wrap/.ms-btn/.ms-dropdown markup
  // on the Material tab) rather than a bespoke panel, so it looks and behaves
  // consistently. Rebuilt on every render so its option list stays in sync
  // with whatever Material Types are actually present in the currently
  // loaded data; checked state is preserved via reqMatTypeFilter.
  function renderMatTypeFilterBar(types, dataAvailable) {
    const statusEl = document.getElementById("reqan-status-filter");
    if (!statusEl || !statusEl.parentElement) return;

    // FIX-MATTYPE-LOOK: match the same labeled-box pattern the Material tab
    // and Branch Comparison use for their multi-selects (a small "nav-label"
    // caption sitting above the .ms-wrap button), instead of a bare unlabeled
    // button — that's what made this control look out of place next to the
    // rest of the filter bar.
    let outer = document.getElementById("reqan-mattype-outer");
    if (!outer) {
      outer = document.createElement("div");
      outer.id = "reqan-mattype-outer";
      outer.style.cssText =
        "display:inline-flex;flex-direction:column;gap:5px;margin-left:0.5rem;vertical-align:bottom;min-width:170px;";
      outer.innerHTML =
        `<div class="nav-label" style="font-size:0.65rem">Material Type</div>` +
        `<div class="ms-wrap" id="reqan-mattype-wrap" style="min-width:0;width:100%">` +
          `<button class="ms-btn" type="button" style="width:100%">All Material Types <span class="ms-arrow">▾</span></button>` +
          `<div class="ms-dropdown" id="reqan-mattype-dd"></div>` +
        `</div>`;
      statusEl.parentElement.insertBefore(outer, statusEl.nextSibling);
    }
    const wrap = document.getElementById("reqan-mattype-wrap");
    const btn  = wrap ? wrap.querySelector(".ms-btn") : null;

    // FIX-MATTYPE-NO-DATA: don't show ZME/ZMS/ZLC/ZMD (or any options) as
    // if they'll filter something when they can't — that's what caused
    // "I picked an item and everything disappeared." Material Type can only
    // ever be resolved from the main inventory data's "Material Type"
    // column (buildMaterialTypeMap()); if that data isn't loaded/reconciled
    // this session, disable the control entirely and say why, rather than
    // pretending it works.
    let note = document.getElementById("reqan-mattype-note");
    if (!dataAvailable) {
      if (btn) {
        btn.disabled = true;
        btn.textContent = "";
        btn.innerHTML = "Unavailable <span class=\"ms-arrow\">▾</span>";
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
      }
      if (!note) {
        note = document.createElement("div");
        note.id = "reqan-mattype-note";
        note.style.cssText = "font-size:0.65rem;color:var(--dim);max-width:220px;line-height:1.3;";
        note.textContent = "Load the main inventory file to enable filtering by Material Type.";
        outer.appendChild(note);
      }
      return; // nothing to wire up — leave any previously-checked filter as is
    }
    if (note) note.remove();
    if (btn) { btn.disabled = false; btn.style.opacity = ""; btn.style.cursor = ""; }

    // buildMultiSelect() fully rebuilds the search box + checkbox list each
    // call, so we re-seed the checked state from reqMatTypeFilter afterward
    // (this control isn't tied to the pageFilters store buildMultiSelect
    // normally reads its initial selection from).
    buildMultiSelect("reqan-mattype-wrap", "reqan-mattype-dd", types, "All Material Types");
    const dd = document.getElementById("reqan-mattype-dd");
    if (dd) {
      dd.querySelectorAll(".ms-item input").forEach(cb => {
        cb.checked = reqMatTypeFilter.has(cb.value);
      });
    }
    // Re-render from the checked state we just restored (also refreshes the
    // button label / selected-count badge).
    if (wrap._refreshOptions) wrap._refreshOptions();
  }

  // ── MATERIAL GROUP FILTER BAR ───────────────────────────────────────────────
  // Identical control/pattern to renderMatTypeFilterBar() above, just sourced
  // from the literal "Material Group Name" column instead of getValuationType().
  // Anchored right after the Material Type filter bar so the two sit
  // together in the filter row.
  function renderMatGroupFilterBar(groups, dataAvailable) {
    const mtOuter = document.getElementById("reqan-mattype-outer");
    const statusEl = document.getElementById("reqan-status-filter");
    const anchor = mtOuter || statusEl;
    if (!anchor || !anchor.parentElement) return;

    let outer = document.getElementById("reqan-matgroup-outer");
    if (!outer) {
      outer = document.createElement("div");
      outer.id = "reqan-matgroup-outer";
      outer.style.cssText =
        "display:inline-flex;flex-direction:column;gap:5px;margin-left:0.5rem;vertical-align:bottom;min-width:170px;";
      outer.innerHTML =
        `<div class="nav-label" style="font-size:0.65rem">Material Group</div>` +
        `<div class="ms-wrap" id="reqan-matgroup-wrap" style="min-width:0;width:100%">` +
          `<button class="ms-btn" type="button" style="width:100%">All Material Groups <span class="ms-arrow">▾</span></button>` +
          `<div class="ms-dropdown" id="reqan-matgroup-dd"></div>` +
        `</div>`;
      anchor.parentElement.insertBefore(outer, anchor.nextSibling);
    }
    const wrap = document.getElementById("reqan-matgroup-wrap");
    const btn  = wrap ? wrap.querySelector(".ms-btn") : null;

    let note = document.getElementById("reqan-matgroup-note");
    if (!dataAvailable) {
      if (btn) {
        btn.disabled = true;
        btn.textContent = "";
        btn.innerHTML = "Unavailable <span class=\"ms-arrow\">▾</span>";
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
      }
      if (!note) {
        note = document.createElement("div");
        note.id = "reqan-matgroup-note";
        note.style.cssText = "font-size:0.65rem;color:var(--dim);max-width:220px;line-height:1.3;";
        note.textContent = "Load the main inventory file to enable filtering by Material Group.";
        outer.appendChild(note);
      }
      return;
    }
    if (note) note.remove();
    if (btn) { btn.disabled = false; btn.style.opacity = ""; btn.style.cursor = ""; }

    buildMultiSelect("reqan-matgroup-wrap", "reqan-matgroup-dd", groups, "All Material Groups");
    const dd = document.getElementById("reqan-matgroup-dd");
    if (dd) {
      dd.querySelectorAll(".ms-item input").forEach(cb => {
        cb.checked = reqMatGroupFilter.has(cb.value);
      });
    }
    if (wrap._refreshOptions) wrap._refreshOptions();
  }

  // ── COPY SELECTED CODES TOOLBAR ─────────────────────────────────────────────
  // Lets users click individual material-code cells — across ALL 4 tabs
  // (Request vs Stock, Suggested Code Corrections, HO01 Stockout, HO01 Not
  // Requested) — to build up a multi-code selection, then copy just those
  // codes (one per line) to the clipboard. Clicking is scoped to the
  // .col-mat-code-wrap cell (or the specific .col-mat-code span, for cells
  // that hold more than one code, like Suggested Code Corrections' "X or Y"
  // list), so it never grabs Description/Qty/SOH text from other columns the
  // way click-dragging across a row would. Selection is shared across tabs
  // (tracked in reqCodeCopySelection by code text, not DOM identity or tab),
  // so you can pick codes from more than one tab and copy them together —
  // every tab's toolbar shows the same live count.
  const REQ_CODE_TABLE_IDS = ["reqan-table-all", "reqan-table-suggest", "reqan-table-stockout", "reqan-table-notreq"];

  function renderCopyCodesToolbars() {
    REQ_CODE_TABLE_IDS.forEach(id => {
      const tableEl = document.getElementById(id);
      if (!tableEl || !tableEl.parentElement) return;
      let bar = document.getElementById(id + "-copybar");
      if (!bar) {
        bar = document.createElement("div");
        bar.id = id + "-copybar";
        bar.className = "reqan-copycode-bar";
        bar.style.cssText =
          "display:none;align-items:center;gap:0.6rem;margin-bottom:0.6rem;" +
          "padding:0.5rem 0.75rem;background:var(--surface2);border:1px solid var(--border2);" +
          "border-radius:var(--radius-sm);font-size:0.78rem;";
        bar.innerHTML =
          `<span class="reqan-copycode-count" style="color:var(--text);font-weight:600"></span>` +
          `<button class="reqan-copycode-btn dl-btn" type="button" style="padding:3px 12px">⧉ Copy Codes</button>` +
          `<button class="reqan-copycode-clear" type="button" style="background:none;border:none;color:var(--blue);cursor:pointer;font-size:0.76rem;padding:0">Clear selection</button>` +
          `<span style="color:var(--dim);font-size:0.72rem;margin-left:auto">Tip: click a code to select it — click again to deselect. Selection carries across tabs.</span>`;
        tableEl.parentElement.insertBefore(bar, tableEl);
      }
    });
    updateCopyCodesToolbars();
  }

  function updateCopyCodesToolbars() {
    const n = reqCodeCopySelection.size;
    document.querySelectorAll(".reqan-copycode-bar").forEach(bar => {
      bar.style.display = n > 0 ? "flex" : "none";
      const countEl = bar.querySelector(".reqan-copycode-count");
      if (countEl) countEl.textContent = `${n} code${n === 1 ? "" : "s"} selected`;
    });
  }

  // Re-applies the "picked" highlight to whichever code cells/spans (by
  // text, not DOM identity) are currently in reqCodeCopySelection — needed
  // every time buildTable() replaces a table's innerHTML. Highlights the
  // specific .col-mat-code span when a cell holds more than one code (e.g.
  // Suggested Code Corrections' "X or Y" cells), otherwise the whole cell.
  function applyCopySelectionHighlight() {
    const sel = REQ_CODE_TABLE_IDS.map(id => `#${id} td.col-mat-code-wrap`).join(", ");
    document.querySelectorAll(sel).forEach(td => {
      td.style.cursor = "pointer";
      const spans = td.querySelectorAll(".col-mat-code");
      const targets = spans.length ? [...spans] : [td];
      targets.forEach(el => {
        const code = el.textContent.trim();
        const picked = reqCodeCopySelection.has(code);
        el.style.background = picked ? "var(--accent-glow)" : "";
        el.style.outline = picked ? "1px solid var(--blue)" : "";
        el.style.borderRadius = picked ? "3px" : "";
        el.title = picked ? "Click to deselect" : "Click to select for copying";
      });
    });
  }

  // Copies text to the clipboard, falling back to a hidden textarea +
  // execCommand for browsers/contexts where navigator.clipboard is
  // unavailable (e.g. non-HTTPS).
  function copyTextToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(() => fallbackCopyToClipboard(text));
    } else {
      fallbackCopyToClipboard(text);
    }
  }
  function fallbackCopyToClipboard(text) {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); } catch (err) { /* no-op */ }
    document.body.removeChild(ta);
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

    const {
      rows, ho01NotRequested, ho01NotRequestedAllCount, mosDataLoaded,
      availableMatTypes, matTypeDataAvailable,
      availableMatGroups, matGroupDataAvailable,
    } = buildRequestAnalysis();

    renderMatTypeFilterBar(availableMatTypes, matTypeDataAvailable);
    renderMatGroupFilterBar(availableMatGroups, matGroupDataAvailable);

    const matches = r => {
      if (!searchQ) return true;
      return r.material.toLowerCase().includes(searchQ)
          || (r.canonical || "").toLowerCase().includes(searchQ)
          || (r.desc || "").toLowerCase().includes(searchQ);
    };

    // Material Type filter (ZME, ZMS…) — multi-select, applies to all 4
    // tabs. Empty selection = no filtering.
    const matTypeActive = reqMatTypeFilter.size > 0;
    const matTypeMatches = r => !matTypeActive || reqMatTypeFilter.has(r.materialType);

    // Material Group filter — same shape as Material Type, applies to all 4 tabs.
    const matGroupActive = reqMatGroupFilter.size > 0;
    const matGroupMatches = r => !matGroupActive || reqMatGroupFilter.has(r.materialGroup);

    // "Assigned to" = the same global sidebar person filter used everywhere
    // else in the app (who-responsible.js, dashboard, expiry-risk, etc.) —
    // NOT the request file's "Created By" column. Applies to every tab here,
    // since it's a property of the MATERIAL, not of the request line.
    const personActive = typeof personFilter !== "undefined" && personFilter.size > 0;
    const personMatches = r => !personActive || (r.person && personFilter.has(r.person));

    let filteredRows = rows.filter(r => matches(r) && personMatches(r) && matTypeMatches(r) && matGroupMatches(r));
    if (statusF) filteredRows = filteredRows.filter(r => r.status === statusF);

    const suggestionRows = rows.filter(r => r.hasSuggestion && matches(r) && personMatches(r) && matTypeMatches(r) && matGroupMatches(r));
    const stockoutRows   = rows.filter(r => r.status === "stockout" && matches(r) && personMatches(r) && matTypeMatches(r) && matGroupMatches(r));
    const notRequested   = ho01NotRequested.filter(r =>
      (!searchQ || r.code.toLowerCase().includes(searchQ) || (r.desc || "").toLowerCase().includes(searchQ))
      && personMatches(r) && matTypeMatches(r) && matGroupMatches(r)
    );

    // ── KPIs ─────────────────────────────────────────────────────────────────
    const matchedCount  = rows.filter(r => r.status !== "no-match").length;
    const dupLineCount   = rows.filter(r => r.isDuplicate).length;
    const dupGroupCount  = new Set(rows.filter(r => r.isDuplicate).map(r => r.canonical || `__raw__${r.material.toUpperCase()}`)).size;
    const locMismatchCount = rows.filter(r => r.locationMismatch).length;
    document.getElementById("reqan-kpis").innerHTML = [
      reqKpi("Request Lines Uploaded", rows.length.toLocaleString(), reqFileName ? `${reqFileName} · Plant ${reqPlant}` : "", "blue"),
      reqKpi("Matched to SAP Stock", `${matchedCount.toLocaleString()} / ${rows.length.toLocaleString()}`, "Resolved via SAP code or mapping", "green"),
      reqKpi("HO01 Stockout (Requested)", rows.filter(r => r.status === "stockout").length.toLocaleString(), "Zero HO01 stock right now", "red"),
      reqKpi("Suggested Code Corrections", rows.filter(r => r.hasSuggestion).length.toLocaleString(), "Stock exists under a different code", "amber"),
      reqKpi("Possible Double Requests", `${dupLineCount.toLocaleString()} lines / ${dupGroupCount.toLocaleString()} items`, "Same item requested more than once — same or different code", "amber"),
      reqKpi("Location Mismatches", locMismatchCount.toLocaleString(), `Requested Location ≠ ${reqPlant || "requesting plant"}'s actual Storage Location`, "red"),
      reqKpi("Critical & Not Requested", ho01NotRequested.length.toLocaleString(),
        mosDataLoaded
          ? `Branch MOS < 1mo, absent from this request (${ho01NotRequestedAllCount.toLocaleString()} idle at HO01 in total)`
          : "Load an AMC file (MOS by Plant page) to compute this",
        "purple"),
    ].join("");

    if (!(typeof mappingTable !== "undefined" && mappingTable.size > 0)) {
      document.getElementById("reqan-mapping-banner").innerHTML =
        `<div class="alert-warning" style="margin-bottom:0.8rem;font-size:0.78rem">⚠️ No Material Standardization mapping file is loaded — code-mismatch suggestions can't be computed, and any request material code that isn't already an exact SAP code will show as "No SAP Match".</div>`;
    } else {
      document.getElementById("reqan-mapping-banner").innerHTML = "";
    }

    if (personActive) {
      const names = [...personFilter].join(", ");
      const banner = document.getElementById("reqan-mapping-banner");
      if (banner) {
        banner.innerHTML += `<div class="alert-info" style="margin-bottom:0.8rem;font-size:0.78rem">👤 Filtered to items assigned to <b>${escHtml(names)}</b> (sidebar person filter) — every tab on this page reflects this.</div>`;
      }
    }

    // ── TAB 1: Request vs Stock (side-by-side) ─────────────────────────────
    const cols1 = [
      { key: "prNum", label: "PR Num" },
      { key: "poste", label: "Poste" },
      { key: "createdBy", label: "Created By" },
      { key: "material", label: "Requested Code",
        fmt: (v, r) => r.hasSuggestion
          ? `<span class="col-mat-code">${escHtml(v)}</span><span class="mat-desc-badge" title="This stock currently sits under a different live SAP code — see Suggested Code Corrections tab">≠ CODE</span>`
          : `<span class="col-mat-code">${escHtml(v)}</span>`,
        raw: true, cellClass: "col-mat-code-wrap" },
      { key: "desc", label: "Description", cellClass: "col-mat-desc-wrap" },
      { key: "reqQty", label: "Requested Qty", fmt: v => fmtQty(v), cellClass: "col-qty" },
      { key: "reqSoh", label: "SOH (per Request File)", fmt: v => fmtQty(v), cellClass: "col-qty" },
      { key: "liveHo01", label: "SOH (HO01)",
        fmt: (v, r) => r.sohMismatch ? `<b style="color:var(--amber)">${fmtQty(v)}</b>` : fmtQty(v),
        raw: true, cellClass: "col-qty" },
      { key: "liveReqPlant", label: `SOH (${reqPlant || "Requesting Plant"})`, fmt: v => fmtQty(v), cellClass: "col-qty" },
      { key: "location", label: "Requested Location",
        fmt: (v, r) => {
          if (!v) return "—";
          if (r.locationMismatch) {
            const actual = r.actualLocations.join(", ");
            return `<span style="display:inline-block;padding:0.15rem 0.55rem;border-radius:999px;font-size:0.72rem;font-weight:700;white-space:nowrap;background:rgba(220,38,38,0.14);color:var(--red)" title="${escHtml(reqPlant || 'The requesting plant')} actually holds this material in: ${escHtml(actual)}">⚠ ${escHtml(v)} — check location</span>`;
          }
          return escHtml(v);
        },
        raw: true },
      { key: "actualLocations", label: `${reqPlant || "Requesting Plant"} Storage Location(s)`,
        fmt: (v, r) => r.locationKnown ? escHtml(r.actualLocations.join(", ")) : "—",
        raw: true, cellClass: "col-qty" },
      { key: "mappedDesc", label: "Description (mapped, HO01)", cellClass: "col-mat-desc-wrap" },
      { key: "qcHo01", label: "Under Quality Inspection (HO01)", fmt: v => v > 0 ? fmtQty(v) : "—", cellClass: "col-qty" },
      { key: "status", label: "Status", fmt: v => reqStatusBadge(v), raw: true },
      { key: "isDuplicate", label: "Duplicate Check",
        fmt: (v, r) => v
          ? `<span style="display:inline-block;padding:0.15rem 0.55rem;border-radius:999px;font-size:0.72rem;font-weight:700;white-space:nowrap;background:rgba(217,119,6,0.14);color:var(--amber,#d97706)" title="${escHtml(r.duplicateSiblingsLabel)}">⚠ Requested ${r.duplicateCount}× (combined ${fmtQty(r.duplicateTotalQty)})</span>`
          : "",
        raw: true },
    ];
    document.getElementById("reqan-table-all").innerHTML = buildTable(
      filteredRows, cols1,
      (row) => row.status === "stockout" ? "row-red" : (row.hasSuggestion ? "row-amber" : ""),
      "", { id: "reqan-export-all", title: "" }
    );
    wireTableExport("reqan-export-all", filteredRows.map(r => ({
      prNum: r.prNum, poste: r.poste, createdBy: r.createdBy, material: r.material, canonical: r.canonical, desc: r.desc,
      reqQty: r.reqQty, reqSoh: r.reqSoh, liveHo01: r.liveHo01, liveReqPlant: r.liveReqPlant,
      location: r.location, actualLocations: r.locationKnown ? r.actualLocations.join(", ") : "",
      locationMismatch: r.locationMismatch ? "Yes" : "No",
      mappedDesc: r.mappedDesc, qcHo01: r.qcHo01,
      deliveryDate: fmtReqDate(r.deliveryDate), status: r.status,
      isDuplicate: r.isDuplicate ? "Yes" : "No", duplicateCount: r.duplicateCount,
      duplicateTotalQty: r.duplicateTotalQty, duplicateSiblingsLabel: r.duplicateSiblingsLabel,
    })), [
      { key: "prNum", label: "Purchase Req Num" }, { key: "poste", label: "Poste" },
      { key: "createdBy", label: "Created By" },
      { key: "material", label: "Requested Code" }, { key: "canonical", label: "Resolved SAP Code" },
      { key: "desc", label: "Description" }, { key: "reqQty", label: "Requested Quantity" },
      { key: "reqSoh", label: "Stock on Hand (Request File)" }, { key: "liveHo01", label: "Stock on Hand (HO01)" },
      { key: "liveReqPlant", label: `Stock on Hand (${reqPlant || "Requesting Plant"})` },
      { key: "location", label: "Requested Location" }, { key: "actualLocations", label: `${reqPlant || "Requesting Plant"} Storage Location(s)` },
      { key: "locationMismatch", label: "Location Mismatch?" },
      { key: "mappedDesc", label: "Description (mapped, HO01)" },
      { key: "qcHo01", label: "Under Quality Inspection (HO01)" },
      { key: "deliveryDate", label: "Delivery Date" }, { key: "status", label: "Status" },
      { key: "isDuplicate", label: "Possible Duplicate?" }, { key: "duplicateCount", label: "Times Requested" },
      { key: "duplicateTotalQty", label: "Combined Requested Qty" }, { key: "duplicateSiblingsLabel", label: "Other Lines (Same Item)" },
    ], "request_analysis_all_lines");

    // ── TAB 2: Suggested Code Corrections ───────────────────────────────────
    const cols2 = [
      { key: "prNum", label: "PR Num" },
      { key: "material", label: "Code As Requested", cellClass: "col-mat-code-wrap" },
      { key: "shortText", label: "Description (as requested)" },
      { key: "suggestedCode", label: "Request Under This SAP Code Instead",
        // suggestedCode is a string like "115-ZOLE-0301-01 (120) or 115-ZOLE-0301-02 (15)" —
        // each is a real, live SAP code with its OWN native-unit quantity, not the
        // standardized/mapped code. Split it back apart just for per-code styling.
        fmt: v => String(v || "").split(" or ").map(part => {
          const m = part.match(/^(.*)\s\((.*)\)$/);
          const code = m ? m[1] : part;
          const qty  = m ? m[2] : "";
          return `<span class="col-mat-code mat-code-clickable">${escHtml(code)}</span>` +
                 (qty ? `<span class="mat-mapped-badge" title="Live HO01 stock under this exact SAP code">${escHtml(qty)}</span>` : "");
        }).join(' <span style="opacity:0.6">or</span> '),
        raw: true, cellClass: "col-mat-code-wrap" },
      { key: "suggestedDesc", label: "Description" },
      { key: "suggestedTotal", label: "Combined HO01 Stock (Suggested Codes)", fmt: v => fmtQty(v), cellClass: "col-qty" },
      { key: "reqQty", label: "Requested Qty", fmt: v => fmtQty(v), cellClass: "col-qty" },
    ];
    document.getElementById("reqan-table-suggest").innerHTML = buildTable(
      suggestionRows, cols2, () => "row-amber", "", { id: "reqan-export-suggest", title: "" }
    );
    wireTableExport("reqan-export-suggest", suggestionRows.map(r => ({
      prNum: r.prNum, material: r.material, shortText: r.shortText,
      suggestedCode: r.suggestedCode, suggestedDesc: r.suggestedDesc, suggestedTotal: r.suggestedTotal, reqQty: r.reqQty,
    })), [
      { key: "prNum", label: "Purchase Req Num" }, { key: "material", label: "Code As Requested" },
      { key: "shortText", label: "Description (as requested)" }, { key: "suggestedCode", label: "Request Under This SAP Code Instead" },
      { key: "suggestedDesc", label: "Description" }, { key: "suggestedTotal", label: "Combined HO01 Stock (Suggested Codes)" },
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

    // ── TAB 4: HO01 Stock Not Requested (branch-critical only) ─────────────
    if (!mosDataLoaded) {
      document.getElementById("reqan-table-notreq").innerHTML =
        `<div class="alert-warning" style="margin:0.8rem 0;font-size:0.8rem">⚠️ No AMC file is loaded, so branch consumption (MOS) can't be computed. This list only shows HO01 stock that's absent from the request AND critical (branch MOS &lt; 1 month) — load an AMC file on the "📐 MOS by Plant" page, then come back here.</div>`;
    } else {
      const cols4 = [
        { key: "code", label: "Material Code", cellClass: "col-mat-code-wrap" },
        { key: "desc", label: "Description" },
        { key: "qty", label: "HO01 Stock on Hand", fmt: v => fmtQty(v), cellClass: "col-qty" },
        { key: "criticalBranches", label: "Critical At (Branch MOS < 1mo)",
          fmt: v => v.map(c => `<span style="display:inline-block;margin:1px 3px 1px 0;padding:0.1rem 0.4rem;border-radius:999px;font-size:0.7rem;font-weight:700;background:rgba(220,38,38,0.14);color:var(--red)">${escHtml(c.plant)} · ${c.mos === Infinity ? "∞" : Number(c.mos).toFixed(1)}mo</span>`).join(""),
          raw: true },
      ];
      document.getElementById("reqan-table-notreq").innerHTML = buildTable(
        notRequested, cols4, () => "row-red", "", { id: "reqan-export-notreq", title: "" }
      );
      wireTableExport("reqan-export-notreq", notRequested.map(r => ({
        code: r.code, desc: r.desc, qty: r.qty,
        criticalBranches: r.criticalBranches.map(c => `${c.plant} (${c.mos === Infinity ? "Infinite" : Number(c.mos).toFixed(2)}mo)`).join("; "),
      })), [
        { key: "code", label: "Material Code" }, { key: "desc", label: "Description" }, { key: "qty", label: "HO01 Stock on Hand" },
        { key: "criticalBranches", label: "Critical At (Branch MOS < 1mo)" },
      ], "request_analysis_ho01_not_requested");
    }

    // FEAT-COPY-CODES: (re)build the toolbar for every tab that has a code
    // column and re-apply highlighting, since buildTable() just replaced
    // each table's DOM. Done once here (not per-tab above) since all 4
    // table elements exist by this point in the render.
    renderCopyCodesToolbars();
    applyCopySelectionHighlight();

    // ── Tab counts (badges in tab labels) ───────────────────────────────────
    setTabCount("reqan-tab-count-all", filteredRows.length);
    setTabCount("reqan-tab-count-suggest", suggestionRows.length);
    setTabCount("reqan-tab-count-stockout", stockoutRows.length);
    setTabCount("reqan-tab-count-notreq", mosDataLoaded ? notRequested.length : 0);
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
        reqMatTypeFilter.clear();
        const mtWrap = document.getElementById("reqan-mattype-wrap");
        if (mtWrap && mtWrap._clearSelected) mtWrap._clearSelected();
        reqMatGroupFilter.clear();
        const mgWrap = document.getElementById("reqan-matgroup-wrap");
        if (mgWrap && mgWrap._clearSelected) mgWrap._clearSelected();
        renderRequestAnalysis();
        return;
      }
      // Open/close and outside-click-to-close for the Material Type dropdown
      // are handled by buildMultiSelect()'s own listeners (same as every
      // other .ms-wrap control in the app) — nothing to wire here.

      // FEAT-COPY-CODES: click a material-code cell (any of the 4 tabs) to
      // toggle it into the copy selection — scoped to just the code
      // cell/span, so clicking never grabs Description/Qty/SOH from other
      // columns the way a click-drag text selection across the row would.
      // For cells holding more than one code (Suggested Code Corrections'
      // "X or Y" list), the specific .col-mat-code span clicked is used so
      // the two codes in that cell can be selected independently.
      const codeCell = e.target.closest(REQ_CODE_TABLE_IDS.map(id => `#${id} td.col-mat-code-wrap`).join(", "));
      if (codeCell) {
        const codeSpan = e.target.closest(".col-mat-code") || codeCell.querySelector(".col-mat-code");
        const code = (codeSpan ? codeSpan.textContent : codeCell.textContent).trim();
        if (code) {
          if (reqCodeCopySelection.has(code)) reqCodeCopySelection.delete(code);
          else reqCodeCopySelection.add(code);
          applyCopySelectionHighlight();
          updateCopyCodesToolbars();
        }
        return;
      }
      if (e.target.classList && e.target.classList.contains("reqan-copycode-btn")) {
        copyTextToClipboard([...reqCodeCopySelection].join("\n"));
        const btn = e.target;
        const original = btn.textContent;
        btn.textContent = "✓ Copied";
        setTimeout(() => { btn.textContent = original; }, 1200);
        return;
      }
      if (e.target.classList && e.target.classList.contains("reqan-copycode-clear")) {
        reqCodeCopySelection.clear();
        applyCopySelectionHighlight();
        updateCopyCodesToolbars();
        return;
      }
    });

    document.body.addEventListener("change", (e) => {
      // Material Type filter — checkbox lives inside the shared .ms-dropdown
      // control built by buildMultiSelect(); sync our Set from whatever's
      // currently checked and re-render.
      if (e.target.closest && e.target.closest("#reqan-mattype-dd") && e.target.type === "checkbox") {
        const wrap = document.getElementById("reqan-mattype-wrap");
        const selected = wrap && wrap._getSelected ? wrap._getSelected() : [];
        reqMatTypeFilter = new Set(selected);
        renderRequestAnalysis();
      }
      if (e.target.closest && e.target.closest("#reqan-matgroup-dd") && e.target.type === "checkbox") {
        const wrap = document.getElementById("reqan-matgroup-wrap");
        const selected = wrap && wrap._getSelected ? wrap._getSelected() : [];
        reqMatGroupFilter = new Set(selected);
        renderRequestAnalysis();
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
    // Re-render when the global "assigned to" sidebar person filter changes
    // (same dropdown who-responsible.js's "View all of X's items" button
    // drives), so this page's tabs stay scoped to whoever is selected there.
    const personFilterEl = document.getElementById("global-person-filter");
    if (personFilterEl) {
      personFilterEl.addEventListener("change", () => {
        if (currentPage === "request-analysis") renderRequestAnalysis();
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
