// Line-level parser for Cin7 "Inventory Movement Details" report.
//
// The existing parseInventoryRows in server.js aggregates to SKU/month
// rollups and throws individual movement lines away — fine for the dashboard,
// but useless for SQF lot traceability. This module preserves every row so
// downstream traceability endpoints can look up "every movement of lot X",
// "everything that flowed through MO-Y", etc.
//
// Expected columns (Cin7 default; case-insensitive header matching):
//   Date, Location, Reference, SKU, Document reference, Product, Unit,
//   Expiry date, Batch #, Type, Quantity in, Quantity out, Cost in, Cost out
//
// Ref-type taxonomy (extracted from the Reference cell prefix):
//   PO = Purchase Order receipt
//   MO = Manufacturing Order (production activity)
//   ST = Stock adjustment
//   TR = Transfer between locations
//   SO = Sales Order shipment
//   FG = Non-SO dispatch (samples, manual out, R&D)  -- rename if Cin7 clarifies
//   OTHER = anything unrecognized

const XLSX = require("xlsx");

const REF_TYPES = new Set(["PO", "MO", "ST", "TR", "SO", "FG"]);

function classifyRef(ref) {
  const s = String(ref || "").trim();
  if (!s) return { ref_type: null, ref_number: null };
  const m = s.match(/^([A-Z]{2,3})-([0-9A-Za-z]+)(?:\/\d+)?/);
  if (!m) return { ref_type: "OTHER", ref_number: s };
  const prefix = m[1].toUpperCase();
  return {
    ref_type: REF_TYPES.has(prefix) ? prefix : "OTHER",
    // Strip the /N batch suffix so all rows for one MO share a ref_number
    // ('MO-00293/1', 'MO-00293/2' → 'MO-00293'). PO/ST/TR/SO/FG have no
    // slash suffix in practice but we handle uniformly.
    ref_number: `${prefix}-${m[2]}`,
  };
}

// Cin7 dates come in as either JS Date objects (when cellDates: true) or
// strings like '28-Mar-2025'. Return YYYY-MM-DD or null.
const MONTH_MAP = { jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06", jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12" };
function parseDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) {
    // Guard the sentinel '9999-12-31' Cin7 uses for non-expiring items —
    // still a valid date, just treated as "no expiry" by convention.
    if (v.getUTCFullYear() > 9000) return null;
    return v.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (m) {
    const mo = MONTH_MAP[m[2].toLowerCase()];
    if (!mo) return null;
    return `${m[3]}-${mo}-${m[1].padStart(2, "0")}`;
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return null;
}

function parseNum(v) {
  if (v == null || v === "") return null;
  const n = parseFloat(String(v).replace(/,/g, ""));
  return isFinite(n) ? n : null;
}

// Given the raw rows (array-of-arrays) from XLSX.utils.sheet_to_json with
// header:1, locate the header row by looking for the distinctive column
// combo. The report has 4-5 rows of metadata before the real header.
function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 25); i++) {
    const fields = (rows[i] || []).map(f => String(f || "").toLowerCase().trim());
    if (fields.includes("sku") && fields.includes("batch #") && fields.includes("type")) {
      return i;
    }
  }
  return -1;
}

function mapColumns(headerRow) {
  const hdr = (headerRow || []).map(h => String(h || "").toLowerCase().trim());
  const col = {};
  hdr.forEach((h, i) => {
    if (h === "date") col.date = i;
    else if (h === "location") col.location = i;
    else if (h === "reference") col.reference = i;
    else if (h === "sku") col.sku = i;
    else if (h === "document reference") col.doc_ref = i;
    else if (h === "product") col.product = i;
    else if (h === "unit") col.unit = i;
    else if (h === "expiry date") col.expiry = i;
    else if (h === "batch #") col.batch = i;
    else if (h === "type") col.movement_type = i;
    else if (h === "quantity in") col.qty_in = i;
    else if (h === "quantity out") col.qty_out = i;
    else if (h === "cost in") col.cost_in = i;
    else if (h === "cost out") col.cost_out = i;
  });
  const required = ["date", "sku", "qty_in", "qty_out"];
  for (const k of required) {
    if (col[k] === undefined) throw new Error(`Movement report missing required column: ${k}`);
  }
  return col;
}

function rowToLine(row, col) {
  const dateIso = parseDate(row[col.date]);
  if (!dateIso) return null;
  const sku = String(row[col.sku] || "").trim();
  if (!sku) return null;
  const rawRef = col.reference !== undefined ? String(row[col.reference] || "").trim() : "";
  const { ref_type, ref_number } = classifyRef(rawRef);
  return {
    movement_date: dateIso,
    location: col.location !== undefined ? String(row[col.location] || "").trim() || null : null,
    reference: rawRef || null,
    ref_type,
    ref_number,
    sku,
    document_reference: col.doc_ref !== undefined ? String(row[col.doc_ref] || "").trim() || null : null,
    product: col.product !== undefined ? String(row[col.product] || "").trim() || null : null,
    unit: col.unit !== undefined ? String(row[col.unit] || "").trim() || null : null,
    expiry_date: col.expiry !== undefined ? parseDate(row[col.expiry]) : null,
    batch: col.batch !== undefined ? (String(row[col.batch] || "").trim() || null) : null,
    movement_type: col.movement_type !== undefined ? String(row[col.movement_type] || "").trim() || null : null,
    qty_in: parseNum(row[col.qty_in]) || 0,
    qty_out: parseNum(row[col.qty_out]) || 0,
    cost_in: col.cost_in !== undefined ? parseNum(row[col.cost_in]) : null,
    cost_out: col.cost_out !== undefined ? parseNum(row[col.cost_out]) : null,
  };
}

// Public: parse a buffer (XLSX or CSV) into an array of line objects.
// Second return is { min_date, max_date, row_count } for the ingest to
// know which date range to delete before insert.
function parseMovementFile(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  if (!wb.SheetNames.length) throw new Error("Workbook has no sheets");
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const hdrIdx = findHeaderRow(rows);
  if (hdrIdx === -1) throw new Error("Could not find header row (expected SKU + Batch # + Type on the same row)");
  const col = mapColumns(rows[hdrIdx]);
  const lines = [];
  let minDate = null;
  let maxDate = null;
  for (let i = hdrIdx + 1; i < rows.length; i++) {
    const line = rowToLine(rows[i], col);
    if (!line) continue;
    lines.push(line);
    if (!minDate || line.movement_date < minDate) minDate = line.movement_date;
    if (!maxDate || line.movement_date > maxDate) maxDate = line.movement_date;
  }
  return { lines, min_date: minDate, max_date: maxDate, row_count: lines.length };
}

module.exports = {
  parseMovementFile,
  classifyRef,
  parseDate,
  // Exposed for tests
  _rowToLine: rowToLine,
  _findHeaderRow: findHeaderRow,
  _mapColumns: mapColumns,
};
