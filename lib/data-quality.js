// Data-quality overlay for traceability.
//
// Four known windows where Voyage's inventory ledger has noise the auditor
// needs context for. Each window defines a date range and a predicate — a
// movement gets annotated only if BOTH its date is in-range AND the row's
// characteristics match the window's rule.
//
// Dates are env-tunable so we can refine without redeploy. Defaults reflect
// m.gunderson's operational read as of 2026-07-15.
//
// Windows overlap intentionally: a stock adjustment on a VC- SKU in
// October 2025 gets both the 'lot-migration' and 'vc-via-st' notes because
// both explanations plausibly apply and the auditor should see both angles.

function envDate(name, fallback) {
  const v = process.env[name];
  if (v && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  return fallback;
}

function buildWindows() {
  return [
    {
      id: "pre-standup",
      from: "2025-01-01",
      to: envDate("TRACEABILITY_DQ_PRESTANDUP_UNTIL", "2025-04-30"),
      applies: () => true,
      note: "Pre-facility-standup: ERP was in bring-up; historical entries before this window may be incomplete. Absence of a receipt or movement for this lot before this date is not necessarily evidence it didn't happen — it may not have been recorded.",
    },
    {
      id: "lot-migration",
      from: envDate("TRACEABILITY_DQ_LOTMIGRATION_FROM", "2025-10-01"),
      to: envDate("TRACEABILITY_DQ_LOTMIGRATION_UNTIL", "2025-12-31"),
      applies: (mv) => mv.ref_type === "ST",
      note: "Lot-code migration window: many stock adjustments during this period were paper reclassifications (moving items from lot-only tracking to lot+pallet tracking), not physical movements. If two ST entries offset around this time, they are likely a relabel, not shrinkage.",
    },
    {
      id: "vc-via-st",
      from: envDate("TRACEABILITY_DQ_LOTMIGRATION_FROM", "2025-10-01"),
      to: envDate("TRACEABILITY_DQ_LOTMIGRATION_UNTIL", "2025-12-31"),
      applies: (mv) => mv.ref_type === "ST" && String(mv.sku || "").startsWith("VC-"),
      note: "VC receipt via stock adjustment: during this window a bad process caused several volatile-compound receipts to be entered as ST (stock adjustment) rather than PO. A stock-adjustment-in on a VC- SKU here is likely a legitimate supplier receipt without a linked PO.",
    },
    {
      id: "facility-count",
      from: envDate("TRACEABILITY_DQ_FACILITY_COUNT_FROM", "2026-04-01"),
      to: envDate("TRACEABILITY_DQ_FACILITY_COUNT_UNTIL", "2026-06-30"),
      applies: (mv) => mv.ref_type === "ST",
      note: "Full facility inventory count reconciliation: large stock adjustments during this window are inventory-count corrections following the departure of the accountant responsible for the Q4 2025 discrepancies. Treat as reconciliation, not operational shrinkage/gain.",
    },
  ];
}

function annotateMovement(mv, windows = buildWindows()) {
  const iso = String(mv.movement_date || "").slice(0, 10);
  const notes = [];
  for (const w of windows) {
    if (iso < w.from || iso > w.to) continue;
    if (!w.applies(mv)) continue;
    notes.push({ id: w.id, from: w.from, to: w.to, note: w.note });
  }
  return notes.length ? { ...mv, dq_notes: notes } : mv;
}

function annotateMovements(rows) {
  const windows = buildWindows();
  return rows.map((r) => annotateMovement(r, windows));
}

// For narrative summaries: how many DQ-flagged movements a result set has,
// grouped by window id. Bot uses this to caveat findings.
function summarizeDataQuality(rows) {
  const windows = buildWindows();
  const counts = {};
  for (const w of windows) counts[w.id] = 0;
  for (const r of rows) {
    for (const w of windows) {
      const iso = String(r.movement_date || "").slice(0, 10);
      if (iso >= w.from && iso <= w.to && w.applies(r)) counts[w.id]++;
    }
  }
  const flagged = Object.entries(counts)
    .filter(([, n]) => n > 0)
    .map(([id, n]) => {
      const w = windows.find((x) => x.id === id);
      return { id, count: n, window: `${w.from}..${w.to}`, note: w.note };
    });
  return { total_movements: rows.length, flagged_windows: flagged };
}

module.exports = {
  buildWindows,
  annotateMovement,
  annotateMovements,
  summarizeDataQuality,
};
