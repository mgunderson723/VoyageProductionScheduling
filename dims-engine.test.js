// Regression tests for the DIMs calculator, anchored on loads that were
// weighed and measured on the Mason dock (see the DIMs Calculator Spec,
// "Validation points"). If a config edit or engine change moves these
// numbers, the test tells you before a freight quote does.
const Dims = require("./public/dims-engine");
const seed = require("./lib/dims-seed.json");

// Quantities in these tests are unit counts (cases / drums) unless a test
// passes qtyMode: "kg".
const run = (packId, quantity, overrides = {}) => {
  const d = Dims.defaultsFor(seed, packId, overrides.palletId);
  return Dims.calculate(seed, Object.assign({ packId, quantity }, d, { qtyMode: "units" }, overrides));
};

describe("DIMs seed config", () => {
  it("passes validation", () => {
    expect(Dims.validateConfig(seed)).toEqual([]);
  });

  it("rejects broken references", () => {
    const bad = JSON.parse(JSON.stringify(seed));
    bad.packs[0].defaultPallet = "nope";
    bad.packs[0].patterns[0].pallet = "nope";
    const problems = Dims.validateConfig(bad);
    expect(problems.some(p => p.includes("default pallet \"nope\""))).toBe(true);
    expect(problems.some(p => p.includes("pallet \"nope\" is not a configured pallet"))).toBe(true);
  });
});

describe("Validation loads", () => {
  it("CFC 10 kg BIB, 80 cases, EU 40×48, 9 slip sheets → 853.05 kg, ~58 in", () => {
    const r = run("cfc_bib_10kg", 80);
    expect(r.ok).toBe(true);
    expect(r.palletCount).toBe(1);
    expect(r.groups[0].sheets).toBe(9);
    expect(r.totals.grossKg).toBeCloseTo(853.05, 2);
    // 8 × 6.625 + 4.625 + 9 × 0.0625; spec range 58.1–58.4 in, measured 58 in
    expect(r.groups[0].H_in).toBeCloseTo(58.1875, 4);
    expect(r.groups[0].H_in).toBeGreaterThanOrEqual(58.1);
    expect(r.groups[0].H_in).toBeLessThanOrEqual(58.4);
  });

  it("PFS 54 cases, GMA pallet → ~40 in and ~1,105 lb (measured 40 in / 501 kg)", () => {
    const r = run("pfs_case", 54);
    expect(r.ok).toBe(true);
    expect(r.palletCount).toBe(1);
    expect(r.groups[0].H_in).toBeCloseTo(39.95, 2); // 6 × 5.625 + 5.0 + 6 × 0.2 bulge
    expect(Math.abs(r.groups[0].H_in - 40)).toBeLessThan(0.5);
    expect(r.totals.grossKg / Dims.LB_TO_KG).toBeCloseTo(1106.5, 1);
    expect(Math.abs(r.totals.grossKg - 501)).toBeLessThan(2);
  });

  it("PFS without the bulge allowance matches the 38.75 in OD-spec figure", () => {
    const r = run("pfs_case", 54, { bulgeIn: 0 });
    expect(r.groups[0].H_in).toBeCloseTo(38.75, 2);
  });

  it("Drum double-stack on EU 40×48 lands within 1 in of the measured 79 in", () => {
    const r = run("drum_55gal", 6, { doubleStack: true });
    expect(r.ok).toBe(true);
    expect(r.palletCount).toBe(2);
    expect(r.floorPositions).toBe(1);
    expect(Math.abs(r.doubleStack.maxHeightIn - 79)).toBeLessThan(1);
    expect(r.doubleStack.measuredHeightIn).toBe(79);
    expect(r.overhangIn).toBeCloseTo(1.3, 5);
  });
});

describe("Quantity entered in kg", () => {
  it("defaults to kg for every pack except PFS, which is entered in cases", () => {
    expect(Dims.defaultsFor(seed, "cfc_bib_10kg").qtyMode).toBe("kg");
    expect(Dims.defaultsFor(seed, "cfc_box_25kg").qtyMode).toBe("kg");
    expect(Dims.defaultsFor(seed, "drum_55gal").qtyMode).toBe("kg");
    expect(Dims.defaultsFor(seed, "pfs_case").qtyMode).toBe("units");
  });

  it("800 kg of CFC is exactly 80 cases and the 853.05 kg validation pallet", () => {
    const r = run("cfc_bib_10kg", 800, { qtyMode: "kg" });
    expect(r.quantity).toBe(80);
    expect(r.conversion.roundedUp).toBe(false);
    expect(r.totals.grossKg).toBeCloseTo(853.05, 2);
  });

  it("rounds a part-filled unit up", () => {
    const r = run("cfc_bib_10kg", 805, { qtyMode: "kg" });
    expect(r.quantity).toBe(81);
    expect(r.conversion.unitsExact).toBeCloseTo(80.5, 6);
    expect(r.conversion.roundedUp).toBe(true);
  });

  it("doesn't add a unit through floating-point noise on lb-denominated nets", () => {
    const kg = 54 * 18.75 * Dims.LB_TO_KG; // 54 PFS cases expressed in kg
    const r = run("pfs_case", kg, { qtyMode: "kg" });
    expect(r.quantity).toBe(54);
    expect(r.conversion.roundedUp).toBe(false);
  });

  it("1,575 kg of liquor is 7 drums", () => {
    expect(run("drum_55gal", 1575, { qtyMode: "kg" }).quantity).toBe(7);
  });

  it("rejects a zero or blank kg quantity", () => {
    const r = run("cfc_bib_10kg", 0, { qtyMode: "kg" });
    expect(r.ok).toBe(false);
    expect(r.errors.join(" ")).toMatch(/quantity in kg/);
  });
});

describe("Pallet math", () => {
  it("splits into full pallets plus a partial with the right layer count", () => {
    const r = run("cfc_bib_10kg", 175); // 80 per pallet → 2 full + 15 cases (2 layers)
    expect(r.fullPallets).toBe(2);
    expect(r.partial).toEqual({ units: 15, layers: 2 });
    expect(r.palletCount).toBe(3);
    expect(r.rows.map(x => x.units)).toEqual([80, 80, 15]);
    expect(r.rows[2].layers).toBe(2);
    expect(r.totals.netKg).toBeCloseTo(1750, 6);
  });

  it("flags estimates that feed the result", () => {
    const r = run("pfs_case", 54);
    expect(r.estimates.some(e => e.includes("tare"))).toBe(true);
    expect(r.estimates.some(e => e.includes("pattern 9 × 6"))).toBe(true);
  });

  it("refuses to calculate a pack with no dims on file and says what's missing", () => {
    const r = run("cfc_box_25kg", 40, { palletId: "eu_40x48" });
    expect(r.ok).toBe(false);
    expect(r.errors.join(" ")).toMatch(/outside height not on file/);
    expect(r.errors.join(" ")).toMatch(/per layer not set/);
  });

  it("warns when the tallest position exceeds the container height", () => {
    const r = run("drum_55gal", 6, { doubleStack: true, containerId: "hapag_45r1_reefer" });
    expect(r.container.capacity).toBe(18);
    expect(r.warnings.some(w => w.includes("exceeds"))).toBe(false); // ~79.8 in < 95.5 in
    const tall = run("cfc_bib_10kg", 160, { layers: 16, containerId: "hapag_45r1_reefer" }); // ~111.7 in
    expect(tall.warnings.some(w => w.includes("exceeds the Hapag 45R1 reefer"))).toBe(true);
  });

  it("builds a packing list in the configured column order", () => {
    const r = run("cfc_bib_10kg", 90);
    const pl = Dims.packingList(seed, r);
    expect(pl.header[0]).toBe("Pallet #");
    expect(pl.rows).toHaveLength(2);
    expect(pl.rows[0][1]).toBe(80);
    expect(pl.tsv.split("\n")).toHaveLength(3);
  });
});
