// RSPO mass-balance report, built from the all-time Cin7 movement ledger.
//
// For a date range it answers the three things the RSPO audit asks for:
//   1. Each RSPO oil in and out: opening stock, purchased (PO receipts),
//      used in production, adjustments, closing stock — plus which products
//      the oil was used to make and which purchases (lot, supplier) came in.
//   2. Each RSPO-claimed product: kg produced, and the RSPO oil % in it,
//      both from the Cin7 BOM (formulation) and traced through the actual
//      production lots (what was really consumed).
//   3. Utilization: RSPO oil in claimed product vs oil purchased.
//
// Traced oil content works lot by lot: a lot's oil fraction is the oil kg
// consumed by the runs that produced it (recursively, through WIP lots)
// divided by the kg those runs produced. Pure function of the inputs so it
// can be unit-tested without the server.

const isKg = u => /^kg$/i.test(String(u || "").trim());
const NOT_RECEIVED = "(not received yet)";
const r3 = x => Math.round(x * 1000) / 1000;
const r4 = x => Math.round(x * 10000) / 10000;
const lotKey = (sku, batch) => sku + "|" + (batch || "");
const baseRef = ref => String(ref || "").replace(/\/.*$/, "");

function indexLines(lines) {
  const byRef = new Map();
  const prodByLot = new Map();       // sku|batch -> MO production rows (In)
  const convByLot = new Map();       // sku|batch -> assembly ("FG-") document rows that created it
  for (const l of lines) {
    if (!byRef.has(l.reference)) byRef.set(l.reference, []);
    byRef.get(l.reference).push(l);
    const k = lotKey(l.sku, l.batch);
    if (l.ref_type === "MO" && l.movement_type === "In" && l.qty_in > 0) {
      if (!prodByLot.has(k)) prodByLot.set(k, []);
      prodByLot.get(k).push(l);
    }
    if (l.ref_type === "FG" && l.movement_type === "In" && l.qty_in > 0) {
      if (!convByLot.has(k)) convByLot.set(k, []);
      convByLot.get(k).push(l);
    }
  }
  return { byRef, prodByLot, convByLot };
}

// Inputs to an assembly document: the lots it consumed (kg rows only, so
// packaging counted in Each doesn't dilute the product).
const convInputs = (ix, ref) => (ix.byRef.get(ref) || []).filter(r => r.movement_type === "Out" && r.qty_out > 0 && isKg(r.unit));

// SKUs a lot came from via Cin7 assembly documents (e.g. 867v kg lots
// re-booked as "EU CFC Lot", 505v boxes re-booked as a "-kg" selling SKU).
// Includes the lot's own SKU.
function makeSourcesFn(ix) {
  const memo = new Map();
  function sources(sku, batch, depth = 0) {
    const k = lotKey(sku, batch);
    if (memo.has(k)) return memo.get(k);
    const out = new Set([sku]);
    memo.set(k, out);
    if (depth < 4) {
      for (const conv of ix.convByLot.get(k) || []) {
        for (const r of convInputs(ix, conv.reference)) for (const s of sources(r.sku, r.batch, depth + 1)) out.add(s);
      }
    }
    return out;
  }
  return sources;
}

// kg per unit for a lot created by assembly documents in a non-kg unit
// (input kg ÷ output qty). Null if it can't be worked out.
function kgPerUnitOf(ix, sku, batch) {
  let inKg = 0, outQty = 0;
  for (const conv of ix.convByLot.get(lotKey(sku, batch)) || []) {
    outQty += conv.qty_in;
    inKg += convInputs(ix, conv.reference).reduce((a, r) => a + r.qty_out, 0);
  }
  return outQty > 0 && inKg > 0 ? inKg / outQty : null;
}

// Rows that belong to one production run: the run's own rows plus any rows
// booked on the bare MO number (MO-level consumption with no run suffix).
function runRows(ix, ref) {
  const own = ix.byRef.get(ref) || [];
  const b = baseRef(ref);
  return b !== ref ? own.concat(ix.byRef.get(b) || []) : own;
}

// Oil fraction (kg oil per kg of lot), per oil SKU. Memoised; cycles → {}.
function makeContentFn(ix, oilSet) {
  const memo = new Map();
  const visiting = new Set();
  // Lots whose runs recorded far less output than input weight — usually a
  // packout entered in boxes or drums in a field labelled kg (late 2025).
  const suspect = new Map();
  function content(sku, batch) {
    if (oilSet.has(sku)) return { [sku]: 1 };
    const k = lotKey(sku, batch);
    if (memo.has(k)) return memo.get(k);
    if (visiting.has(k)) return {};
    const prods = ix.prodByLot.get(k);
    if (!prods) {
      // Not made on an MO: if an assembly document re-booked other lots into
      // this one, it's the same product, so it carries their average oil
      // fraction (by input qty). Averaging rather than dividing by output qty
      // keeps this right when the document mixes units (boxes in, kg out).
      const convs = ix.convByLot.get(k);
      if (!convs) { memo.set(k, {}); return {}; }
      visiting.add(k);
      let inQty = 0;
      const oilQty = {};
      for (const conv of convs) {
        for (const r of convInputs(ix, conv.reference)) {
          inQty += r.qty_out;
          const c = content(r.sku, r.batch);
          for (const o in c) oilQty[o] = (oilQty[o] || 0) + r.qty_out * c[o];
        }
      }
      visiting.delete(k);
      const res = {};
      if (inQty > 0) for (const o in oilQty) res[o] = oilQty[o] / inQty;
      memo.set(k, res);
      return res;
    }
    visiting.add(k);
    let outKg = 0, inKg = 0;
    const oilKg = {};
    const refs = [...new Set(prods.map(p => p.reference))];
    for (const ref of refs) {
      const rows = runRows(ix, ref);
      for (const r of rows) {
        if (r.movement_type === "In" && r.qty_in > 0 && isKg(r.unit)) outKg += r.qty_in;
      }
      for (const r of rows) {
        if (r.movement_type !== "Out" || !(r.qty_out > 0)) continue;
        if (isKg(r.unit)) inKg += r.qty_out;
        const c = content(r.sku, r.batch);
        for (const o in c) oilKg[o] = (oilKg[o] || 0) + r.qty_out * c[o];
      }
    }
    visiting.delete(k);
    let denom = outKg;
    if (outKg > 0 && inKg > 0 && outKg / inKg < 0.5) {
      denom = inKg;
      suspect.set(k, { sku, lot: batch || null, refs, outKg, inKg });
    }
    const res = {};
    if (denom > 0) for (const o in oilKg) res[o] = Math.min(1, oilKg[o] / denom);
    memo.set(k, res);
    return res;
  }
  content.suspect = suspect;
  return content;
}

// SKU → claim labels, for every SKU upstream of a claimed product (so a WIP
// liquor that gets packed into 505.EU is tagged "505.EU").
function upstreamTags(lines, products) {
  const inputsOf = new Map();  // output sku -> Set(input skus)
  const byRef = new Map();
  for (const l of lines) {
    if (l.ref_type !== "MO") continue;
    if (!byRef.has(l.reference)) byRef.set(l.reference, { ins: new Set(), outs: new Set() });
    const e = byRef.get(l.reference);
    if (l.movement_type === "In" && l.qty_in > 0) e.outs.add(l.sku);
    if (l.movement_type === "Out" && l.qty_out > 0) e.ins.add(l.sku);
  }
  for (const { ins, outs } of byRef.values()) {
    for (const o of outs) {
      if (!inputsOf.has(o)) inputsOf.set(o, new Set());
      for (const i of ins) if (i !== o) inputsOf.get(o).add(i);
    }
  }
  const tags = new Map();
  for (const p of products) {
    const seen = new Set();
    const stack = [...(p.skus || [])];
    while (stack.length) {
      const s = stack.pop();
      if (seen.has(s)) continue;
      seen.add(s);
      if (!tags.has(s)) tags.set(s, new Set());
      tags.get(s).add(p.label);
      for (const i of inputsOf.get(s) || []) stack.push(i);
    }
  }
  return tags;
}

function buildRspoReport({ lines, config, from, to, bomFor = () => null, supplierFor = () => null, customerFor = () => null }) {
  const oils = (config && config.oils) || [];
  const products = (config && config.products) || [];
  const claimCustomers = ((config && config.claimCustomers) || []).map(s => String(s).trim().toLowerCase()).filter(Boolean);
  const oilSet = new Set(oils.map(o => o.sku));
  const inRange = d => d >= from && d <= to;
  const warnings = [];
  const ix = indexLines(lines);
  const content = makeContentFn(ix, oilSet);
  const sourcesOf = makeSourcesFn(ix);
  const tags = upstreamTags(lines, products);

  let ledgerStart = null, ledgerEnd = null;
  for (const l of lines) {
    if (!ledgerStart || l.movement_date < ledgerStart) ledgerStart = l.movement_date;
    if (!ledgerEnd || l.movement_date > ledgerEnd) ledgerEnd = l.movement_date;
  }
  if (ledgerStart && from < ledgerStart) {
    warnings.push(`The movement ledger starts ${ledgerStart}, after the report start ${from}; opening balances can't include earlier history.`);
  }

  // ── Claimed products ──────────────────────────────────────────────────────
  const productNames = new Map();
  for (const l of lines) if (l.product && !productNames.has(l.sku)) productNames.set(l.sku, l.product);

  const productRows = products.map(p => {
    const skus = p.skus || [];
    const rowOils = (p.oils || []).filter(o => oilSet.has(o));
    const out = {
      label: p.label, skus, oilsFilter: rowOils, note: p.note || "",
      producedKg: 0, lotCount: 0, tracedOilKg: {}, bomOilKg: {}, runs: [], skuDetail: [], warnings: [],
    };
    if (!skus.length) {
      out.warnings.push("No SKU mapped yet. Add it in RSPO settings.");
      return out;
    }
    const lots = new Set();
    let unmatchedKg = 0;
    let bomCoveredKg = 0;   // produced kg of SKUs that have a BOM
    for (const sku of skus) {
      const bom = bomFor(sku);
      let skuKg = 0;
      let skuTraced = 0;
      const unmatchedBefore = unmatchedKg;
      for (const [k, rows] of ix.prodByLot) {
        if (!k.startsWith(sku + "|")) continue;
        // Only trace lots actually produced in the period.
        if (!rows.some(r => inRange(r.movement_date) && isKg(r.unit))) continue;
        const batch = k.slice(sku.length + 1);
        const c = content(sku, batch);
        const matches = !rowOils.length || rowOils.some(o => c[o] > 0);
        for (const r of rows) {
          if (!inRange(r.movement_date) || !isKg(r.unit)) continue;
          if (!matches) { unmatchedKg += r.qty_in; continue; }
          skuKg += r.qty_in;
          lots.add(k);
          const traced = {};
          for (const o in c) {
            const kg = r.qty_in * c[o];
            traced[o] = kg;
            skuTraced += kg;
            out.tracedOilKg[o] = (out.tracedOilKg[o] || 0) + kg;
          }
          out.runs.push({ date: r.movement_date, ref: r.reference, sku, lot: batch || null, kg: r3(r.qty_in), tracedOilKg: r3(Object.values(traced).reduce((a, b) => a + b, 0)), tracedByOil: Object.fromEntries(Object.entries(traced).map(([o, v]) => [o, r3(v)])) });
        }
      }
      out.producedKg += skuKg;
      const bomPct = bom && bom.byOil ? Object.fromEntries(Object.entries(bom.byOil).filter(([o]) => !rowOils.length || rowOils.includes(o))) : null;
      if (bomPct) {
        bomCoveredKg += skuKg;
        for (const o in bomPct) out.bomOilKg[o] = (out.bomOilKg[o] || 0) + skuKg * bomPct[o];
      }
      const skuBomPct = bomPct ? Object.values(bomPct).reduce((a, b) => a + b, 0) : null;
      const skuTracedPct = skuKg ? skuTraced / skuKg : null;
      out.skuDetail.push({
        sku, product: productNames.get(sku) || null, producedKg: r3(skuKg),
        bomFound: !!(bom && bom.found), bomVersion: bom ? bom.version || null : null,
        bomPct: skuBomPct != null ? r4(skuBomPct) : null,
        bomPctAllOils: bom && bom.byOil ? r4(Object.values(bom.byOil).reduce((a, b) => a + b, 0)) : null,
        tracedPct: skuTracedPct != null ? r4(skuTracedPct) : null,
      });
      if (skuBomPct && skuTracedPct != null && skuTraced > 0 && Math.abs(skuTracedPct - skuBomPct) > 0.03) {
        out.warnings.push(`${sku}: traced oil (${(skuTracedPct * 100).toFixed(1)}%) differs from the BOM (${(skuBomPct * 100).toFixed(1)}%) by more than 3 points. Check for zero or missing consumption entries on its runs, or a BOM change.`);
      }
      if (!productNames.has(sku)) out.warnings.push(`${sku} has no movements in the ledger. Check the SKU.`);
      else if (!skuKg && unmatchedKg === unmatchedBefore) out.warnings.push(`No ${sku} produced in this period.`);
      if (skuKg && !(bom && bom.found)) out.warnings.push(`No Cin7 BOM synced for ${sku}, so there's no formulation %.`);
      else if (skuKg && bomPct && !Object.values(bomPct).some(v => v > 0)) {
        out.warnings.push(rowOils.length && Object.values(bom.byOil || {}).some(v => v > 0)
          ? `${sku}'s default BOM version uses a different RSPO oil than this line; check the version in Cin7.`
          : `${sku}'s default BOM has no RSPO oil in it.`);
      }
    }
    if (unmatchedKg > 0.001) {
      out.warnings.push(`${Math.round(unmatchedKg).toLocaleString("en-US")} kg of ${skus.join(", ")} was produced in this period without ${rowOils.join(" / ")} traced in it, so it isn't counted on this line.`);
    }
    out.lotCount = lots.size;
    out.runs.sort((a, b) => a.date.localeCompare(b.date) || a.ref.localeCompare(b.ref));
    const tracedTotal = Object.values(out.tracedOilKg).reduce((a, b) => a + b, 0);
    const bomTotal = Object.values(out.bomOilKg).reduce((a, b) => a + b, 0);
    out.tracedPct = out.producedKg ? r4(tracedTotal / out.producedKg) : null;
    // BOM % covers only SKUs that have a BOM (missing ones are warned above).
    out.bomPct = bomCoveredKg > 0 ? r4(bomTotal / bomCoveredKg) : null;
    out.bomCoveredKg = r3(bomCoveredKg);
    if (out.producedKg && tracedTotal === 0) out.warnings.push("No RSPO oil could be traced into this production. Check the oil consumption entries on its runs.");
    out.producedKg = r3(out.producedKg);
    out.tracedOilKgTotal = r3(tracedTotal);
    out.bomOilKgTotal = out.bomPct != null ? r3(bomTotal) : null;
    for (const k of ["tracedOilKg", "bomOilKg"]) out[k] = Object.fromEntries(Object.entries(out[k]).map(([o, v]) => [o, r3(v)]));
    return out;
  });

  // ── Shipments of claimed product ──────────────────────────────────────────
  // Every sales-order line in the period whose lot is a claimed product,
  // directly or via an assembly document that re-booked it under a selling
  // SKU. Oil shipped = shipped kg × the lot's traced oil fraction.
  const shipments = [];
  let missingCustomer = 0;
  for (const p of productRows) Object.assign(p, { shippedKg: 0, shippedClaimKg: 0, shippedOilKg: {}, shipmentCount: 0 });
  for (const l of lines) {
    if (l.ref_type !== "SO" || !inRange(l.movement_date)) continue;
    const qty = (l.qty_out || 0) - (l.qty_in || 0);   // an SO "In" is a return
    if (!qty) continue;
    const srcs = sourcesOf(l.sku, l.batch);
    const c = content(l.sku, l.batch);
    const row = productRows.find(p => p.skus.some(s => srcs.has(s)) && (!p.oilsFilter.length || p.oilsFilter.some(o => c[o] > 0)));
    if (!row) continue;
    const kgPerUnit = isKg(l.unit) ? 1 : kgPerUnitOf(ix, l.sku, l.batch);
    const kg = kgPerUnit != null ? qty * kgPerUnit : null;
    const byOil = {};
    if (kg != null) for (const o in c) byOil[o] = kg * c[o];
    const oilKg = Object.values(byOil).reduce((a, b) => a + b, 0);
    const customer = customerFor(baseRef(l.reference));
    if (!customer) missingCustomer++;
    const toClaimCustomer = claimCustomers.length && customer ? claimCustomers.some(n => customer.toLowerCase().includes(n)) : null;
    shipments.push({
      date: l.movement_date, ref: l.reference, customer: customer || null, toClaimCustomer,
      product: row.label, sku: l.sku, viaAssembly: !row.skus.includes(l.sku), lot: l.batch || null,
      qty: r3(qty), unit: l.unit || null, kg: kg != null ? r3(kg) : null, isReturn: qty < 0,
      oilKg: r3(oilKg), byOil: Object.fromEntries(Object.entries(byOil).map(([o, v]) => [o, r3(v)])),
    });
    row.shipmentCount++;
    if (kg == null) { row.warnings.push(`${l.reference}: ${qty} ${l.unit} of ${l.sku} couldn't be converted to kg.`); continue; }
    row.shippedKg += kg;
    if (toClaimCustomer) row.shippedClaimKg += kg;
    for (const o in byOil) row.shippedOilKg[o] = (row.shippedOilKg[o] || 0) + byOil[o];
    if (!Object.keys(c).length) row.warnings.push(`${l.reference} lot ${l.batch || "?"}: no RSPO oil could be traced into the shipped product.`);
  }
  shipments.sort((a, b) => a.date.localeCompare(b.date) || a.ref.localeCompare(b.ref));
  for (const p of productRows) {
    p.shippedOilKgTotal = r3(Object.values(p.shippedOilKg).reduce((a, b) => a + b, 0));
    p.shippedOilKg = Object.fromEntries(Object.entries(p.shippedOilKg).map(([o, v]) => [o, r3(v)]));
    p.shippedKg = r3(p.shippedKg);
    p.shippedClaimKg = claimCustomers.length ? r3(p.shippedClaimKg) : null;
  }
  if (missingCustomer) {
    warnings.push(`${missingCustomer} shipment line${missingCustomer === 1 ? " has" : "s have"} no customer name yet; customers come from the nightly Cin7 sales-order sync.`);
  }
  for (const s of content.suspect.values()) {
    // Only lots that carry RSPO oil change an oil %; skip the rest.
    const c = content(s.sku, s.lot);
    if (!Object.values(c).some(v => v > 0)) continue;
    const fmt = x => Math.round(x).toLocaleString("en-US");
    warnings.push(`Lot ${s.lot || "?"} (${s.sku}) on ${s.refs.join(", ")}: output entered as ${fmt(s.outKg)} from ${fmt(s.inKg)} kg of inputs, likely boxes or drums rather than kg. Its oil % is worked out from input weight instead.`);
  }

  // ── Oils: in / out ────────────────────────────────────────────────────────
  const oilRows = oils.map(o => {
    const rows = lines.filter(l => l.sku === o.sku);
    const r = {
      sku: o.sku, label: o.label, model: o.model || null,
      opening: 0, purchased: 0, used: 0, adjustments: 0, transfers: 0, other: 0, closing: 0,
      usedBy: [], purchases: [], claimedTracedKg: 0, claimedBomKg: 0,
    };
    const usedBy = new Map();
    for (const l of rows) {
      const net = (l.qty_in || 0) - (l.qty_out || 0);
      if (l.movement_date < from) { r.opening += net; continue; }
      if (l.movement_date > to) continue;
      if (l.ref_type === "PO") {
        r.purchased += net;
        if (l.qty_in > 0) {
          const s = supplierFor(baseRef(l.reference));
          r.purchases.push({ date: l.movement_date, ref: l.reference, lot: l.batch || null, kg: r3(l.qty_in), location: l.location || null, supplier: s || null });
        }
      } else if (l.ref_type === "MO") {
        r.used -= net;
        if (l.qty_out > 0) {
          let outs = (ix.byRef.get(l.reference) || []).filter(x => x.movement_type === "In" && x.qty_in > 0);
          // Run has no output of its own (e.g. booked on the bare MO number):
          // fall back to everything any run of that MO produced.
          if (!outs.length) outs = [...ix.byRef.entries()]
            .filter(([ref]) => baseRef(ref) === baseRef(l.reference))
            .flatMap(([, v]) => v)
            .filter(x => x.movement_type === "In" && x.qty_in > 0);
          const outKg = {};
          for (const x of outs) outKg[x.sku] = (outKg[x.sku] || 0) + x.qty_in;
          const totalOut = Object.values(outKg).reduce((a, b) => a + b, 0);
          const shares = totalOut > 0 ? Object.entries(outKg).map(([s, q]) => [s, q / totalOut]) : [[NOT_RECEIVED, 1]];
          for (const [s, f] of shares) {
            if (!usedBy.has(s)) usedBy.set(s, {
              sku: s, kg: 0, runs: new Set(),
              product: s === NOT_RECEIVED ? "Oil consumed on runs whose output isn't received in Cin7 yet (work in progress)" : (productNames.get(s) || null),
            });
            const u = usedBy.get(s);
            u.kg += l.qty_out * f;
            u.runs.add(l.reference);
          }
        }
      } else if (l.ref_type === "ST") r.adjustments += net;
      else if (l.ref_type === "TR") r.transfers += net;
      else r.other += net;
    }
    r.closing = r.opening + r.purchased - r.used + r.adjustments + r.transfers + r.other;
    // A product line restricted to other oils (506.EU SG vs MB share a SKU)
    // isn't fed by this oil, so don't tag it.
    const fedByThisOil = label => {
      const p = products.find(x => x.label === label);
      return !p || !(p.oils && p.oils.length) || p.oils.includes(o.sku);
    };
    r.usedBy = [...usedBy.values()].map(u => ({
      sku: u.sku, product: u.product, kg: r3(u.kg), runCount: u.runs.size,
      runs: [...u.runs].sort(),
      feeds: tags.has(u.sku) ? [...tags.get(u.sku)].filter(fedByThisOil) : [],
    })).sort((a, b) => b.kg - a.kg);
    r.purchases.sort((a, b) => a.date.localeCompare(b.date));
    r.shippedTracedKg = 0;
    for (const p of productRows) {
      r.claimedTracedKg += p.tracedOilKg[o.sku] || 0;
      r.claimedBomKg += p.bomOilKg[o.sku] || 0;
      r.shippedTracedKg += p.shippedOilKg[o.sku] || 0;
    }
    r.utilizationTraced = r.purchased > 0 ? r4(r.claimedTracedKg / r.purchased) : null;
    r.utilizationBom = r.purchased > 0 && r.claimedBomKg > 0 ? r4(r.claimedBomKg / r.purchased) : null;
    r.shippedOverPurchased = r.purchased > 0 ? r4(r.shippedTracedKg / r.purchased) : null;
    for (const k of ["opening", "purchased", "used", "adjustments", "transfers", "other", "closing", "claimedTracedKg", "claimedBomKg", "shippedTracedKg"]) r[k] = r3(r[k]);
    if (!rows.length) warnings.push(`${o.sku} (${o.label}) has no movements in the ledger. Check the SKU.`);
    if (r.opening < -0.5) warnings.push(`${o.sku} opening stock is negative (${Math.round(r.opening)} kg); some receipts may predate the ledger or were entered as stock adjustments.`);
    return r;
  });

  return {
    from, to, ledgerStart, ledgerEnd,
    oils: oilRows,
    products: productRows,
    shipments,
    claimCustomers: (config && config.claimCustomers) || [],
    totals: {
      shippedKg: r3(productRows.reduce((a, p) => a + p.shippedKg, 0)),
      shippedClaimKg: claimCustomers.length ? r3(productRows.reduce((a, p) => a + (p.shippedClaimKg || 0), 0)) : null,
      shippedOilKg: r3(productRows.reduce((a, p) => a + p.shippedOilKgTotal, 0)),
      producedKg: r3(productRows.reduce((a, p) => a + p.producedKg, 0)),
      claimedTracedKg: r3(productRows.reduce((a, p) => a + (p.tracedOilKgTotal || 0), 0)),
      claimedBomKg: r3(productRows.reduce((a, p) => a + (p.bomOilKgTotal || 0), 0)),
      purchasedKg: r3(oilRows.reduce((a, o) => a + o.purchased, 0)),
    },
    warnings,
  };
}

// Config validation, shared by the PUT endpoint and the editor.
function validateRspoConfig(c) {
  const problems = [];
  if (!c || typeof c !== "object") return ["Settings must be an object."];
  if (!Array.isArray(c.oils) || !Array.isArray(c.products)) return ["Settings need an oils list and a products list."];
  const SKU = /^[A-Za-z0-9][A-Za-z0-9 ._\-\/]{1,63}$/;
  const oilSkus = new Set();
  c.oils.forEach((o, i) => {
    if (!o || !SKU.test(String(o.sku || ""))) problems.push(`Oil ${i + 1}: SKU is missing or invalid.`);
    else if (oilSkus.has(o.sku)) problems.push(`Oil ${o.sku} is listed twice.`);
    else oilSkus.add(o.sku);
    if (!o || !String(o.label || "").trim()) problems.push(`Oil ${i + 1}: name is required.`);
    if (o && o.model && !["SG", "MB", "IP", "B&C"].includes(o.model)) problems.push(`Oil ${o.sku}: supply model must be SG, MB, IP or B&C.`);
  });
  c.products.forEach((p, i) => {
    if (!p || !String(p.label || "").trim()) problems.push(`Product ${i + 1}: name is required.`);
    if (p && p.skus != null && !Array.isArray(p.skus)) problems.push(`Product ${i + 1}: SKUs must be a list.`);
    (p && Array.isArray(p.skus) ? p.skus : []).forEach(s => { if (!SKU.test(String(s))) problems.push(`Product ${p.label || i + 1}: "${s}" isn't a valid SKU.`); });
    (p && Array.isArray(p.oils) ? p.oils : []).forEach(s => { if (!oilSkus.has(s)) problems.push(`Product ${p.label || i + 1}: oil "${s}" isn't in the oils list.`); });
  });
  if (c.claimCustomers != null && (!Array.isArray(c.claimCustomers) || c.claimCustomers.some(s => typeof s !== "string"))) {
    problems.push("Claim customers must be a list of names.");
  }
  return problems;
}

module.exports = { buildRspoReport, validateRspoConfig, _internals: { indexLines, makeContentFn, upstreamTags } };
