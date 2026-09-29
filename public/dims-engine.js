// DIMs calculator engine — shared by the browser (window.DimsEngine) and Node
// (require('./public/dims-engine')) so the server validates config with the
// same rules the UI uses, and jest can run the dock-measured validation loads
// as regression tests.
//
// All reference data lives in an editable config (see lib/dims-seed.json).
// Every measurement is stored in the unit it was stated in, as
// { v: number|null, u: "in"|"mm"|"cm"|"ft"|"kg"|"lb"|"g", est?: true }, and is
// converted exactly once here. Figures flagged est:true surface as warnings
// until someone replaces them with a measured value.
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.DimsEngine = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const LB_TO_KG = 0.45359237;
  const IN_TO_MM = 25.4;
  const TO_IN = { in: 1, ft: 12, mm: 1 / IN_TO_MM, cm: 10 / IN_TO_MM };
  const TO_KG = { kg: 1, g: 0.001, lb: LB_TO_KG };
  const LENGTH_UNITS = Object.keys(TO_IN);
  const WEIGHT_UNITS = Object.keys(TO_KG);

  const has = m => !!m && typeof m.v === "number" && isFinite(m.v);
  const toIn = m => (has(m) && TO_IN[m.u] != null ? m.v * TO_IN[m.u] : null);
  const toKg = m => (has(m) && TO_KG[m.u] != null ? m.v * TO_KG[m.u] : null);
  const byId = (arr, id) => (Array.isArray(arr) ? arr.find(x => x && x.id === id) : null) || null;
  const posInt = v => {
    const n = Number(v);
    return Number.isInteger(n) && n > 0 ? n : null;
  };
  const round = (x, dp) => {
    const f = Math.pow(10, dp);
    return Math.round(x * f) / f;
  };
  // "case" → "cases", "box" → "boxes", "bag" → "bags"
  const plural = w => (/(s|x|z|ch|sh)$/i.test(w) ? w + "es" : w + "s");

  // "48x40" style key used for container floor-position lookups. Larger side
  // first so a 40×48 and a 48×40 pallet share a key.
  function footprintKey(pallet) {
    const L = toIn(pallet && pallet.L), W = toIn(pallet && pallet.W);
    if (L == null || W == null) return null;
    const a = Math.round(Math.max(L, W)), b = Math.round(Math.min(L, W));
    return a + "x" + b;
  }

  function patternFor(pack, palletId) {
    return (pack && Array.isArray(pack.patterns) ? pack.patterns : []).find(p => p && p.pallet === palletId) || null;
  }

  // Calculator defaults for a pack (and optionally a chosen pallet). The UI
  // pre-fills its inputs from this and lets the user override any of it.
  function defaultsFor(config, packId, palletId) {
    const pack = byId(config.packs, packId);
    if (!pack) return null;
    const firstPattern = (pack.patterns || [])[0];
    const pid = palletId || pack.defaultPallet || (firstPattern && firstPattern.pallet) || null;
    const pattern = pid ? patternFor(pack, pid) : null;
    const pallet = byId(config.pallets, pid);
    return {
      palletId: pid,
      perLayer: pattern ? pattern.perLayer : null,
      layers: pattern ? pattern.layers : null,
      slipSheets: !!pack.slipSheets,
      bulgeIn: toIn(pack.bulgePerLayer) || 0,
      doubleStack: false,
      units: pallet && pallet.defaultUnits === "imperial" ? "imperial" : "metric",
      qtyMode: qtyModeFor(pack),
    };
  }

  // How the quantity is entered by default: kg of net product (most packs —
  // production asks are in kg) or whole units (e.g. PFS, counted in cases).
  function qtyModeFor(pack) {
    return pack && pack.qtyEntry === "units" ? "units" : "kg";
  }

  // Whole units needed to hold `kg` of net product. Rounds up (a part-filled
  // unit still ships), with a small tolerance so lb-denominated nets don't turn
  // an exact 54 cases into 55 through floating-point noise.
  function unitsForKg(kg, netKgPerUnit) {
    const exact = kg / netKgPerUnit;
    const units = Math.max(1, Math.ceil(exact - 1e-9));
    return { exact, units, roundedUp: units - exact > 1e-6 };
  }

  // input: { packId, quantity, qtyMode: "units"|"kg", palletId, perLayer,
  //          layers, slipSheets, bulgeIn, doubleStack, containerId }
  // With qtyMode "kg", quantity is kg of net product and is converted to
  // whole units first.
  function calculate(config, input) {
    const errors = [];
    const warnings = [];
    const estimates = [];
    const noteEst = (m, what) => {
      if (m && m.est && has(m)) estimates.push(what + " (" + m.v + " " + m.u + ")");
    };

    const pack = byId(config.packs, input.packId);
    if (!pack) return { ok: false, errors: ["Pick a pack format."], warnings, estimates };
    const pallet = byId(config.pallets, input.palletId);
    const unitWord = pack.unitName || "unit";

    const netKg = toKg(pack.net);
    let qty = null;
    let conversion = null;
    if (input.qtyMode === "kg") {
      const kg = Number(input.quantity);
      if (!(kg > 0) || !isFinite(kg)) errors.push("Enter a quantity in kg above 0.");
      else if (netKg) {
        const c = unitsForKg(kg, netKg);
        qty = c.units;
        conversion = { enteredKg: kg, netKgPerUnit: netKg, unitsExact: c.exact, units: c.units, roundedUp: c.roundedUp };
      }
    } else {
      qty = Number(input.quantity);
      if (!Number.isInteger(qty) || qty < 1) errors.push("Enter a whole number of " + plural(unitWord) + " (1 or more).");
    }
    if (!pallet) errors.push("Pick a pallet type.");
    const perLayer = posInt(input.perLayer);
    const layers = posInt(input.layers);
    if (!perLayer) errors.push(pack.label + ": " + plural(unitWord) + " per layer not set for this pallet.");
    if (!layers) errors.push(pack.label + ": layers per pallet not set for this pallet.");

    const tareKg = toKg(pack.tare);
    const unitH = toIn(pack.H);
    if (netKg == null) errors.push(pack.label + ": net weight per " + unitWord + " not on file.");
    if (tareKg == null) errors.push(pack.label + ": " + unitWord + " tare not on file.");
    if (unitH == null) errors.push(pack.label + ": outside height not on file.");
    const isCyl = pack.shape === "cylinder";
    const unitL = isCyl ? toIn(pack.dia) : toIn(pack.L);
    const unitW = isCyl ? toIn(pack.dia) : toIn(pack.W);
    if (unitL == null || unitW == null) {
      errors.push(pack.label + ": outside " + (isCyl ? "diameter" : "length and width") + " not on file.");
    }
    noteEst(pack.net, pack.label + " net weight");
    noteEst(pack.tare, pack.label + " tare");
    noteEst(pack.H, pack.label + " height");
    if (isCyl) noteEst(pack.dia, pack.label + " diameter");
    else { noteEst(pack.L, pack.label + " length"); noteEst(pack.W, pack.label + " width"); }

    let linerKg = 0;
    if (pack.liner) {
      const liner = byId(config.components, pack.liner);
      const w = liner ? toKg(liner.wt) : null;
      if (w == null) warnings.push("Liner weight (" + (liner ? liner.label : pack.liner) + ") not on file; excluded from gross.");
      else { linerKg = w; noteEst(liner.wt, liner.label + " weight"); }
    }

    let palletL = null, palletW = null, palletH = 0, palletTareKg = 0, heightExcludesPallet = false;
    if (pallet) {
      palletL = toIn(pallet.L);
      palletW = toIn(pallet.W);
      if (palletL == null || palletW == null) errors.push(pallet.label + ": footprint not on file.");
      const h = toIn(pallet.H);
      if (h == null) {
        heightExcludesPallet = true;
        warnings.push(pallet.label + ": pallet height not on file, so heights below exclude the pallet deck.");
      } else { palletH = h; noteEst(pallet.H, pallet.label + " height"); }
      const t = toKg(pallet.tare);
      if (t == null) warnings.push(pallet.label + ": pallet tare not on file; excluded from gross.");
      else { palletTareKg = t; noteEst(pallet.tare, pallet.label + " tare"); }
    }

    let sheetKg = 0, sheetThk = 0;
    if (input.slipSheets) {
      const slip = byId(config.components, "slip_sheet");
      if (!slip) warnings.push("Slip sheets are on but no slip_sheet component is configured; ignored.");
      else {
        const w = toKg(slip.wt), t = toIn(slip.thk);
        if (w == null) warnings.push("Slip sheet weight not on file; excluded from gross.");
        else { sheetKg = w; noteEst(slip.wt, "Slip sheet weight"); }
        if (t == null) warnings.push("Slip sheet thickness not on file; excluded from height.");
        else { sheetThk = t; noteEst(slip.thk, "Slip sheet thickness"); }
      }
    }
    const bulge = Math.max(0, Number(input.bulgeIn) || 0);

    const pattern = pallet ? patternFor(pack, pallet.id) : null;
    if (pattern && pattern.est && perLayer === pattern.perLayer && layers === pattern.layers) {
      estimates.push(pack.label + " pallet pattern " + pattern.perLayer + " × " + pattern.layers + " (not yet confirmed as standard)");
    }
    if (pallet && !pattern && perLayer && layers) {
      warnings.push("No saved pallet pattern for " + pack.label + " on " + pallet.label + "; using the per-layer and layer counts entered.");
    }

    if (errors.length) return { ok: false, errors, warnings, estimates, pack, pallet };

    const unitsPerPallet = perLayer * layers;
    const fullPallets = Math.floor(qty / unitsPerPallet);
    const remainder = qty % unitsPerPallet;
    const partialLayers = remainder ? Math.ceil(remainder / perLayer) : 0;

    function palletCalc(units, nLayers) {
      const sheets = input.slipSheets ? nLayers + 1 : 0;
      const H = nLayers * unitH + palletH + sheets * sheetThk + nLayers * bulge;
      const net = units * netKg;
      const packaging = units * (tareKg + linerKg) + sheets * sheetKg;
      return {
        units, layers: nLayers, sheets,
        L_in: palletL, W_in: palletW, H_in: H,
        netKg: net, packagingKg: packaging, palletKg: palletTareKg,
        grossKg: net + packaging + palletTareKg,
      };
    }

    const groups = [];
    if (fullPallets) groups.push(Object.assign({ kind: "full", count: fullPallets }, palletCalc(unitsPerPallet, layers)));
    if (remainder) groups.push(Object.assign({ kind: "partial", count: 1 }, palletCalc(remainder, partialLayers)));
    const palletCount = fullPallets + (remainder ? 1 : 0);

    // Overhang: a saved pattern can carry a known overhang (drums on a 40 in
    // side); independently, a layer whose footprint area exceeds the deck
    // cannot fit without overhanging.
    let overhangIn = null;
    if (pattern && pattern.perLayer === perLayer) {
      const oh = toIn(pattern.overhang);
      if (oh && oh > 0) {
        overhangIn = oh;
        warnings.push("This pattern overhangs the deck by about " + round(oh, 1) + " in.");
      }
    }
    const layerArea = perLayer * unitL * unitW;
    const deckArea = palletL * palletW;
    if (layerArea > deckArea * 1.0001) {
      warnings.push(perLayer + " " + plural(unitWord) + " per layer need " + Math.round(layerArea) + " sq in but the deck is " + Math.round(deckArea) + " sq in; the load will overhang.");
    }

    // One row per physical pallet, in load order (full pallets, then partial).
    const rows = [];
    let n = 1;
    groups.forEach(g => {
      for (let i = 0; i < g.count; i++) {
        rows.push({ pallet_no: n++, units: g.units, layers: g.layers, netKg: g.netKg, grossKg: g.grossKg, L_in: g.L_in, W_in: g.W_in, H_in: g.H_in });
      }
    });

    let doubleStack = null;
    let separatorsKg = 0;
    if (input.doubleStack) {
      const ds = pack.doubleStack || {};
      if (!ds.allowed) {
        warnings.push(pack.label + " is not set up for double-stacking; ignored.");
      } else {
        const sep = byId(config.components, ds.separator);
        const sepThk = sep ? toIn(sep.thk) : null;
        const sepKg = sep ? toKg(sep.wt) : null;
        if (sep) noteEst(sep.thk, sep.label + " thickness");
        if (ds.separator && sepThk == null) warnings.push("Separator thickness not on file; excluded from stack height.");
        if (ds.separator && sepKg == null) warnings.push("Separator weight not on file; excluded from gross.");
        let maxH = 0, heaviest = 0, stacks = 0;
        for (let i = 0; i < rows.length; i += 2) {
          const a = rows[i], b = rows[i + 1];
          if (b) {
            stacks++;
            maxH = Math.max(maxH, a.H_in + (sepThk || 0) + b.H_in);
            heaviest = Math.max(heaviest, a.grossKg + b.grossKg + (sepKg || 0));
          } else {
            maxH = Math.max(maxH, a.H_in);
            heaviest = Math.max(heaviest, a.grossKg);
          }
        }
        separatorsKg = stacks * (sepKg || 0);
        doubleStack = {
          positions: Math.ceil(rows.length / 2),
          stacks,
          maxHeightIn: maxH,
          heaviestPositionKg: heaviest,
          measuredHeightIn: toIn(ds.measuredHeight),
        };
      }
    }
    const floorPositions = doubleStack ? doubleStack.positions : palletCount;

    const totals = {
      palletCount,
      floorPositions,
      netKg: qty * netKg,
      grossKg: groups.reduce((s, g) => s + g.count * g.grossKg, 0) + separatorsKg,
    };

    let container = null;
    if (input.containerId) {
      const c = byId(config.containers, input.containerId);
      if (c) {
        const cH = toIn(c.H);
        const tallest = doubleStack ? doubleStack.maxHeightIn : Math.max.apply(null, groups.map(g => g.H_in));
        if (cH != null && tallest > cH) {
          warnings.push("Tallest position (" + round(tallest, 1) + " in) exceeds the " + c.label + " interior height (" + round(cH, 1) + " in).");
        }
        const key = footprintKey(pallet);
        const cap = c.floorPositions && key ? posInt(c.floorPositions[key]) : null;
        if (cap) container = { label: c.label, capacity: cap, containersNeeded: Math.ceil(floorPositions / cap), interiorHeightIn: cH };
        else {
          container = { label: c.label, capacity: null, containersNeeded: null, interiorHeightIn: cH };
          warnings.push("No floor-position count on file for " + (key || "this") + " pallets in the " + c.label + ".");
        }
      }
    }

    return {
      ok: true, errors, warnings, estimates,
      pack, pallet, unitWord, quantity: qty, conversion,
      unitsPerPallet, fullPallets,
      partial: remainder ? { units: remainder, layers: partialLayers } : null,
      palletCount, floorPositions, groups, rows, doubleStack, totals, container,
      overhangIn, heightExcludesPallet,
    };
  }

  // Packing-list columns. The order comes from config.packingList.columns so
  // the team can match whatever layout the SO packing list uses.
  const PACKING_COLUMNS = {
    pallet_no: { label: () => "Pallet #", get: r => r.pallet_no },
    units: { label: res => "Qty (" + plural(res.unitWord) + ")", get: r => r.units },
    layers: { label: () => "Layers", get: r => r.layers },
    net_kg: { label: () => "Net (kg)", get: r => round(r.netKg, 1) },
    gross_kg: { label: () => "Gross (kg)", get: r => round(r.grossKg, 1) },
    net_lb: { label: () => "Net (lb)", get: r => Math.round(r.netKg / LB_TO_KG) },
    gross_lb: { label: () => "Gross (lb)", get: r => Math.round(r.grossKg / LB_TO_KG) },
    L_cm: { label: () => "L (cm)", get: r => Math.round(r.L_in * 2.54) },
    W_cm: { label: () => "W (cm)", get: r => Math.round(r.W_in * 2.54) },
    H_cm: { label: () => "H (cm)", get: r => Math.round(r.H_in * 2.54) },
    L_in: { label: () => "L (in)", get: r => round(r.L_in, 1) },
    W_in: { label: () => "W (in)", get: r => round(r.W_in, 1) },
    H_in: { label: () => "H (in)", get: r => round(r.H_in, 1) },
    dims_cm: { label: () => "Dims L × W × H (cm)", get: r => [r.L_in, r.W_in, r.H_in].map(x => Math.round(x * 2.54)).join(" × ") },
    dims_in: { label: () => "Dims L × W × H (in)", get: r => [r.L_in, r.W_in, r.H_in].map(x => round(x, 1)).join(" × ") },
  };

  function packingList(config, result) {
    const cols = ((config.packingList && config.packingList.columns) || []).filter(k => PACKING_COLUMNS[k]);
    const use = cols.length ? cols : ["pallet_no", "units", "net_kg", "gross_kg", "L_cm", "W_cm", "H_cm"];
    const header = use.map(k => PACKING_COLUMNS[k].label(result));
    const rows = (result.rows || []).map(r => use.map(k => PACKING_COLUMNS[k].get(r)));
    return { header, rows, tsv: [header].concat(rows).map(r => r.join("\t")).join("\n") };
  }

  // Structural validation shared by the server (on save) and the UI (before
  // save, and in the raw JSON editor). Returns a list of human-readable
  // problems; empty means valid.
  function validateConfig(config) {
    const problems = [];
    if (!config || typeof config !== "object" || Array.isArray(config)) return ["Config must be a JSON object."];
    const ID_RE = /^[A-Za-z0-9_-]{1,64}$/;
    const lists = ["packs", "pallets", "components", "containers"];
    lists.forEach(k => { if (!Array.isArray(config[k])) problems.push("\"" + k + "\" must be a list."); });
    if (problems.length) return problems;

    const measure = (m, where, units) => {
      if (m == null) return;
      if (typeof m !== "object" || Array.isArray(m)) { problems.push(where + " must be {\"v\": number, \"u\": unit}."); return; }
      if (m.v != null && (typeof m.v !== "number" || !isFinite(m.v) || m.v < 0)) problems.push(where + " value must be a non-negative number or empty.");
      if (m.v != null && units.indexOf(m.u) === -1) problems.push(where + " unit must be one of " + units.join(", ") + ".");
    };
    const ids = {};
    lists.forEach(k => {
      ids[k] = new Set();
      config[k].forEach((item, i) => {
        const where = k + "[" + i + "]";
        if (!item || typeof item !== "object") { problems.push(where + " must be an object."); return; }
        if (typeof item.id !== "string" || !ID_RE.test(item.id)) problems.push(where + ": id must be 1–64 letters, digits, - or _.");
        else if (ids[k].has(item.id)) problems.push(where + ": duplicate id \"" + item.id + "\".");
        else ids[k].add(item.id);
        if (typeof item.label !== "string" || !item.label.trim()) problems.push(where + " (" + item.id + "): label is required.");
      });
    });

    config.pallets.forEach(p => {
      if (!p || !p.id) return;
      ["L", "W", "H"].forEach(f => measure(p[f], "Pallet " + p.id + " " + f, LENGTH_UNITS));
      measure(p.tare, "Pallet " + p.id + " tare", WEIGHT_UNITS);
      if (p.defaultUnits && ["metric", "imperial"].indexOf(p.defaultUnits) === -1) problems.push("Pallet " + p.id + ": defaultUnits must be metric or imperial.");
    });
    config.components.forEach(c => {
      if (!c || !c.id) return;
      measure(c.wt, "Component " + c.id + " weight", WEIGHT_UNITS);
      measure(c.thk, "Component " + c.id + " thickness", LENGTH_UNITS);
    });
    config.containers.forEach(c => {
      if (!c || !c.id) return;
      ["L", "W", "H"].forEach(f => measure(c[f], "Container " + c.id + " " + f, LENGTH_UNITS));
      measure(c.tare, "Container " + c.id + " tare", WEIGHT_UNITS);
      if (c.floorPositions != null && (typeof c.floorPositions !== "object" || Array.isArray(c.floorPositions))) {
        problems.push("Container " + c.id + ": floorPositions must be an object like {\"48x40\": 18}.");
      }
    });
    config.packs.forEach(p => {
      if (!p || !p.id) return;
      const w = "Pack " + p.id;
      if (p.shape && ["box", "cylinder"].indexOf(p.shape) === -1) problems.push(w + ": shape must be box or cylinder.");
      if (p.qtyEntry && ["kg", "units"].indexOf(p.qtyEntry) === -1) problems.push(w + ": default quantity entry must be kg or units.");
      ["L", "W", "H", "dia", "bulgePerLayer"].forEach(f => measure(p[f], w + " " + f, LENGTH_UNITS));
      ["net", "tare"].forEach(f => measure(p[f], w + " " + f, WEIGHT_UNITS));
      if (p.liner && !ids.components.has(p.liner)) problems.push(w + ": liner \"" + p.liner + "\" is not a configured component.");
      if (p.defaultPallet && !ids.pallets.has(p.defaultPallet)) problems.push(w + ": default pallet \"" + p.defaultPallet + "\" is not a configured pallet.");
      if (p.patterns != null && !Array.isArray(p.patterns)) problems.push(w + ": patterns must be a list.");
      const seen = new Set();
      (Array.isArray(p.patterns) ? p.patterns : []).forEach((pt, i) => {
        const pw = w + " pattern " + (i + 1);
        if (!pt || !ids.pallets.has(pt.pallet)) problems.push(pw + ": pallet \"" + (pt && pt.pallet) + "\" is not a configured pallet.");
        else if (seen.has(pt.pallet)) problems.push(pw + ": more than one pattern for pallet \"" + pt.pallet + "\".");
        else seen.add(pt.pallet);
        if (pt && pt.perLayer != null && !posInt(pt.perLayer)) problems.push(pw + ": per layer must be a whole number above 0.");
        if (pt && pt.layers != null && !posInt(pt.layers)) problems.push(pw + ": layers must be a whole number above 0.");
        if (pt) measure(pt.overhang, pw + " overhang", LENGTH_UNITS);
      });
      const ds = p.doubleStack;
      if (ds && ds.separator && !ids.components.has(ds.separator)) problems.push(w + ": double-stack separator \"" + ds.separator + "\" is not a configured component.");
      if (ds) measure(ds.measuredHeight, w + " measured double-stack height", LENGTH_UNITS);
    });
    const cols = config.packingList && config.packingList.columns;
    if (cols != null) {
      if (!Array.isArray(cols)) problems.push("packingList.columns must be a list.");
      else cols.forEach(k => { if (!PACKING_COLUMNS[k]) problems.push("packingList.columns: unknown column \"" + k + "\". Known: " + Object.keys(PACKING_COLUMNS).join(", ") + "."); });
    }
    return problems;
  }

  return {
    LB_TO_KG, IN_TO_MM, LENGTH_UNITS, WEIGHT_UNITS, PACKING_COLUMNS,
    toIn, toKg, has, round, plural, footprintKey, patternFor, defaultsFor, qtyModeFor, unitsForKg,
    calculate, packingList, validateConfig,
  };
});
