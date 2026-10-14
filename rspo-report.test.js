const { buildRspoReport, validateRspoConfig } = require("./lib/rspo-report");
const seed = require("./lib/rspo-seed.json");

const OIL = "RM-OIL-SG", OIL_MB = "RM-OIL-MB";
let n = 0;
const mv = (date, reference, sku, type, qty, batch, extra = {}) => Object.assign({
  movement_date: date, reference, ref_type: reference.slice(0, 2),
  ref_number: reference.replace(/\/.*$/, ""), sku, product: sku + " name", unit: "Kg",
  batch, movement_type: type, qty_in: type === "In" ? qty : 0, qty_out: type === "Out" ? qty : 0,
  location: "Mason", _n: n++,
}, extra);

const lines = [
  mv("2025-12-01", "PO-00001", OIL, "In", 200, "S0"),
  mv("2025-12-20", "MO-00009/1", "FG-X", "In", 400, "F0"),            // before range
  mv("2026-01-05", "PO-00002", OIL, "In", 1000, "S1"),
  // WIP step: 300 oil + 700 sugar → 1,000 kg WIP-A (30% oil)
  mv("2026-02-01", "MO-00001/1", OIL, "Out", 300, "S1"),
  mv("2026-02-01", "MO-00001/1", "RM-SUGAR", "Out", 700, "G1"),
  mv("2026-02-01", "MO-00001/1", "WIP-A", "In", 1000, "W1"),
  // Packout: 500 kg WIP-A → 500 kg FG-X
  mv("2026-02-10", "MO-00002/1", "WIP-A", "Out", 500, "W1"),
  mv("2026-02-10", "MO-00002/1", "FG-X", "In", 500, "F1"),
  // Oil consumed on a run that hasn't recorded its output yet
  mv("2026-03-01", "MO-00003/1", OIL, "Out", 100, "S1"),
  mv("2026-03-05", "ST-00001", OIL, "Out", 20, "S1"),
  // MB oil into a separate lot of the same FG SKU
  mv("2026-01-06", "PO-00003", OIL_MB, "In", 500, "M1"),
  mv("2026-04-01", "MO-00004/1", OIL_MB, "Out", 100, "M1"),
  mv("2026-04-01", "MO-00004/1", "RM-SUGAR", "Out", 300, "G1"),
  mv("2026-04-01", "MO-00004/1", "FG-X", "In", 400, "F2"),
  // Rework loop: a lot consumed into itself must not hang the trace
  mv("2026-05-01", "MO-00005/1", "WIP-B", "Out", 10, "R1"),
  mv("2026-05-01", "MO-00005/1", OIL, "Out", 50, "S1"),
  mv("2026-05-01", "MO-00005/1", "WIP-B", "In", 60, "R1"),
];

const config = {
  oils: [{ sku: OIL, label: "SG oil", model: "SG" }, { sku: OIL_MB, label: "MB oil", model: "MB" }],
  products: [
    { label: "X (SG)", skus: ["FG-X"], oils: [OIL] },
    { label: "X (MB)", skus: ["FG-X"], oils: [OIL_MB] },
    { label: "Unmapped", skus: [] },
  ],
};
const bomFor = sku => sku === "FG-X" ? { found: true, version: "1", byOil: { [OIL]: 0.3, [OIL_MB]: 0 } } : null;
const report = () => buildRspoReport({
  lines, config, from: "2026-01-01", to: "2026-12-31", bomFor, supplierFor: ref => ref === "PO-00002" ? "Cargill BV" : null,
});

describe("RSPO report", () => {
  it("traces oil through a WIP step into the finished product", () => {
    const sg = report().products[0];
    expect(sg.producedKg).toBe(500);          // F1 only; F0 is before the range, F2 has MB oil
    expect(sg.tracedOilKgTotal).toBeCloseTo(150, 6);
    expect(sg.tracedPct).toBeCloseTo(0.3, 6);
    expect(sg.bomPct).toBeCloseTo(0.3, 6);
    expect(sg.bomOilKgTotal).toBeCloseTo(150, 6);
    expect(sg.runs).toEqual([expect.objectContaining({ ref: "MO-00002/1", lot: "F1", kg: 500, tracedOilKg: 150 })]);
  });

  it("splits one SKU across product lines by the oil actually traced", () => {
    const [sg, mb] = report().products;
    expect(mb.producedKg).toBe(400);
    expect(mb.tracedPct).toBeCloseTo(0.25, 6);
    expect(sg.warnings.join(" ")).toMatch(/400 kg of FG-X was produced in this period without RM-OIL-SG/);
    expect(mb.warnings.join(" ")).toMatch(/default BOM version uses a different RSPO oil/);
  });

  it("balances each oil: opening + purchased − used ± adjustments = closing", () => {
    const o = report().oils[0];
    expect(o.opening).toBe(200);
    expect(o.purchased).toBe(1000);
    expect(o.used).toBe(450);                  // 300 + 100 + 50
    expect(o.adjustments).toBe(-20);
    expect(o.closing).toBe(730);
    expect(o.purchases).toEqual([expect.objectContaining({ ref: "PO-00002", lot: "S1", kg: 1000, supplier: "Cargill BV" })]);
    expect(o.claimedTracedKg).toBeCloseTo(150, 6);
    expect(o.utilizationTraced).toBeCloseTo(0.15, 6);
  });

  it("shows what the oil was used to make, tagging claimed products and unreceived runs", () => {
    const used = report().oils[0].usedBy;
    const wip = used.find(u => u.sku === "WIP-A");
    expect(wip.kg).toBe(300);
    expect(wip.feeds).toEqual(["X (SG)"]);     // not the MB line, which uses a different oil
    expect(used.find(u => u.sku === "(not received yet)").kg).toBe(100);
  });

  it("averages BOM % only over SKUs that have a BOM", () => {
    const r = buildRspoReport({
      lines, from: "2026-01-01", to: "2026-12-31", bomFor,
      config: { oils: config.oils, products: [{ label: "Both", skus: ["FG-X", "WIP-A"] }] },
    }).products[0];
    expect(r.producedKg).toBe(1900);           // FG-X 900 + WIP-A 1,000
    expect(r.bomCoveredKg).toBe(900);          // WIP-A has no BOM
    expect(r.bomPct).toBeCloseTo(0.3, 6);      // not 270 / 1,900
    expect(r.warnings.join(" ")).toMatch(/No Cin7 BOM synced for WIP-A/);
  });

  it("flags unmapped product lines and survives a rework loop", () => {
    const r = report();
    expect(r.products[2].warnings[0]).toMatch(/No SKU mapped/);
    expect(r.oils[0].usedBy.find(u => u.sku === "WIP-B").kg).toBe(50);
  });
});

describe("RSPO settings", () => {
  it("starter settings are valid", () => {
    expect(validateRspoConfig(seed)).toEqual([]);
  });
  it("rejects duplicate oils, bad models and unknown oil filters", () => {
    const bad = { oils: [{ sku: "RM-1", label: "a", model: "XX" }, { sku: "RM-1", label: "b" }], products: [{ label: "P", skus: ["FG-1"], oils: ["RM-9"] }] };
    const p = validateRspoConfig(bad).join(" ");
    expect(p).toMatch(/listed twice/);
    expect(p).toMatch(/supply model/);
    expect(p).toMatch(/oil "RM-9" isn't in the oils list/);
  });
});
