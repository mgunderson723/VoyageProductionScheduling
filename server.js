const express = require("express");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const multer = require("multer");
const XLSX = require("xlsx");
const Anthropic = require("@anthropic-ai/sdk");
const cron      = require("node-cron");

// ── Excel import helpers ──────────────────────────────────────────────────────

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

const MACHINE_MAP = {
  "seed cleaning":                    "seed_clean",
  "alk/roaster":                      "roaster",
  "west mac (cbe)":                   "west_mac",
  "east mac (cbs)":                   "east_mac",
  "1250 mac":                         "mac_1250",
  "5k mac packout":                   "mac_packout",
  "mac packout":                      "mac_packout",
  "pouching":                         "pouching",
  "mass line / conch / depositing":   "MULTI",
  "mass line (changeover)":           "conching",
  "fat melter":                       "fat_melter",
  "refining":                         "refining",
  "conching":                         "conching",
  "depositing":                       "depositing",
};
const MASS_MACHINES = ["refining", "conching", "depositing"];

function mapMachine(raw) {
  if (!raw) return null;
  const key = raw.toLowerCase().replace(/\s+/g, " ").trim();
  if (MACHINE_MAP[key]) return MACHINE_MAP[key];
  // Prefix match — handles variants like "Alk/Roaster (ECOM)", "Alk/Roaster (APAC)"
  for (const [k, v] of Object.entries(MACHINE_MAP)) {
    if (key.startsWith(k)) return v;
  }
  return null;
}

function parseMOs(raw) {
  return (String(raw || "").match(/MO-\d+/g) || []);
}

function parseDate(raw) {
  if (!raw) return null;
  if (raw instanceof Date) {
    const y = raw.getFullYear();
    const m = String(raw.getMonth() + 1).padStart(2, "0");
    const d = String(raw.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  const m = String(raw).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`;
  return null;
}

// Words that, when they immediately follow a "<number> kg" token, indicate the
// number is a PACKAGE SIZE (bag-in-box, drum, pouch, etc.) — not the
// production quantity. Critical: the previous parser greedily matched any
// "X kg" and any leading number, so notes like "Packout · 25kg BIB" set the
// order qty to 25 (the BIB size), and notes like "Grinding · 7-day cycle"
// set it to 7 (the day count). The fix: only accept kg/MT matches that are
// NOT followed by these package words, then take the largest.
const PACKAGE_WORD_RE = /^\s*(BIB|drum|pouch|tote|sack|bag|box|pail|jar|case|bottle)/i;

function parseQty(raw) {
  if (!raw) return 0;
  const s = String(raw).replace(/,/g, "");
  const candidates = [];
  // Match "X MT" — convert to kg
  const mtRe = /(\d+(?:\.\d+)?)\s*MT\b/gi;
  let m;
  while ((m = mtRe.exec(s)) !== null) {
    const trailing = s.slice(mtRe.lastIndex);
    if (!PACKAGE_WORD_RE.test(trailing)) candidates.push(parseFloat(m[1]) * 1000);
  }
  // Match "X kg" — exclude when followed by a package word
  const kgRe = /(\d+(?:\.\d+)?)\s*kg\b/gi;
  while ((m = kgRe.exec(s)) !== null) {
    const trailing = s.slice(kgRe.lastIndex);
    if (!PACKAGE_WORD_RE.test(trailing)) candidates.push(parseFloat(m[1]));
  }
  if (!candidates.length) return 0;
  // Production qty is the largest kg-tagged number on the line (e.g. 4,300 kg
  // wins over a stray 50 kg test reference, and certainly over package sizes).
  return Math.round(Math.max(...candidates));
}

function detectAttribs(sku, machine) {
  const u = (sku + " " + machine).toUpperCase();
  let cat = "liquor", sub = "liquor", temper = null, region = null;
  if (/COATING|INCLUSION|CHIP|860|859|865|815|810/.test(u)) {
    cat = "chocolate"; sub = "chocolate";
  } else if (u.includes("COFFEE")) {
    cat = u.includes("GROUND") ? "coffee_ground" : "coffee_beans";
    sub = cat;
  }
  if (u.includes("CBE")) temper = "cbe";
  else if (u.includes("CBS")) temper = "cbs";
  if (/-EU\b|\.EU\b/.test(u)) region = "eu";
  else if (/-US\b|\.US\b/.test(u)) region = "us";
  return { cat, sub, temper, region };
}

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, "data");

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

app.use(express.json({ limit: "50mb" }));
app.use(express.static(path.join(__dirname, "public")));

// Helper to read/write JSON files
function readData(key) {
  const file = path.join(DATA_DIR, `${key}.json`);
  if (!fs.existsSync(file)) return null;
  try {
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (e) {
    return null;
  }
}

function writeData(key, data) {
  const file = path.join(DATA_DIR, `${key}.json`);
  fs.writeFileSync(file, JSON.stringify(data, null, 2), "utf8");
}

// ── Auth ──────────────────────────────────────────────────────────────────────
//
// All /api/* endpoints require a valid session cookie except /api/login and
// /api/logout. The cookie is an HMAC-signed payload "<userId>.<expiresMs>.<sig>".
// The hash function below MUST match public/index.html#hashPassword so existing
// stored user records remain valid.

const SESSION_SECRET   = process.env.SESSION_SECRET || crypto.randomBytes(32).toString("hex");
const SESSION_TTL_MS   = 7 * 24 * 60 * 60 * 1000; // 7 days
const SESSION_COOKIE   = "vfsession";
const IS_PROD          = !!process.env.RAILWAY_ENVIRONMENT || process.env.NODE_ENV === "production";

if (!process.env.SESSION_SECRET) {
  console.warn("[auth] SESSION_SECRET not set — generated an ephemeral secret. Sessions will be invalidated on every restart. Set SESSION_SECRET in env to fix.");
}

function hashPassword(pw) {
  let h = 0;
  for (let i = 0; i < pw.length; i++) { h = ((h << 5) - h) + pw.charCodeAt(i); h |= 0; }
  return "h_" + Math.abs(h).toString(36) + "_" + pw.length;
}

function signSession(userId) {
  const expires = Date.now() + SESSION_TTL_MS;
  const payload = `${userId}.${expires}`;
  const sig = crypto.createHmac("sha256", SESSION_SECRET).update(payload).digest("base64url");
  return `${payload}.${sig}`;
}

function verifySession(value) {
  if (!value) return null;
  const parts = value.split(".");
  if (parts.length !== 3) return null;
  const [userId, expiresStr, sig] = parts;
  const payload = `${userId}.${expiresStr}`;
  const expected = crypto.createHmac("sha256", SESSION_SECRET).update(payload).digest("base64url");
  if (sig.length !== expected.length) return null;
  let sigBuf, expBuf;
  try { sigBuf = Buffer.from(sig); expBuf = Buffer.from(expected); } catch (e) { return null; }
  if (!crypto.timingSafeEqual(sigBuf, expBuf)) return null;
  if (Number(expiresStr) < Date.now()) return null;
  return userId;
}

function parseCookies(header) {
  const out = {};
  if (!header) return out;
  for (const part of header.split(";")) {
    const idx = part.indexOf("=");
    if (idx === -1) continue;
    const k = part.slice(0, idx).trim();
    const v = part.slice(idx + 1).trim();
    out[k] = decodeURIComponent(v);
  }
  return out;
}

function getSessionUserId(req) {
  const cookies = parseCookies(req.headers.cookie);
  return verifySession(cookies[SESSION_COOKIE]);
}

function buildCookie(value, maxAgeSec) {
  const parts = [
    `${SESSION_COOKIE}=${encodeURIComponent(value)}`,
    "Path=/",
    "HttpOnly",
    "SameSite=Lax",
    `Max-Age=${maxAgeSec}`,
  ];
  if (IS_PROD) parts.push("Secure");
  return parts.join("; ");
}

// Seed a default admin if no users exist yet, so a fresh deploy is loginnable.
function seedDefaultAdminIfMissing() {
  const users = readData("vf_users");
  if (Array.isArray(users) && users.length > 0) return;
  const defaultUser = {
    id: "master",
    username: "productionadmin",
    password: hashPassword("productionadmin1800"),
    role: "admin",
    created: new Date().toISOString().slice(0, 10),
  };
  writeData("vf_users", [defaultUser]);
  console.log("[auth] Seeded default admin user 'productionadmin' (no prior users found).");
}
seedDefaultAdminIfMissing();

// Public auth endpoints (registered BEFORE the requireAuth middleware below).
app.post("/api/login", (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) {
    return res.status(400).json({ ok: false, error: "Missing username or password" });
  }
  const users = readData("vf_users") || [];
  const u = users.find(x => x && x.username && x.username.toLowerCase() === String(username).toLowerCase());
  if (!u || u.password !== hashPassword(password)) {
    return res.status(401).json({ ok: false, error: "Invalid credentials" });
  }
  res.setHeader("Set-Cookie", buildCookie(signSession(u.id), Math.floor(SESSION_TTL_MS / 1000)));
  res.json({ ok: true, user: { id: u.id, username: u.username, role: u.role } });
});

app.post("/api/logout", (req, res) => {
  res.setHeader("Set-Cookie", buildCookie("", 0));
  res.json({ ok: true });
});

// Everything else under /api/* requires authentication.
app.use("/api", (req, res, next) => {
  if (req.path === "/login" || req.path === "/logout") return next();
  const userId = getSessionUserId(req);
  if (!userId) return res.status(401).json({ ok: false, error: "Not authenticated" });
  req.userId = userId;
  next();
});

// GET /api/me — used by the front-end to resume a session on page load.
app.get("/api/me", (req, res) => {
  const users = readData("vf_users") || [];
  const u = users.find(x => x && x.id === req.userId);
  if (!u) return res.status(401).json({ ok: false, error: "User no longer exists" });
  res.json({ ok: true, user: { id: u.id, username: u.username, role: u.role } });
});

// GET data by key
app.get("/api/data/:key", (req, res) => {
  const key = req.params.key.replace(/[^a-z0-9_-]/gi, "");
  const data = readData(key);
  if (data === null) {
    return res.json({ exists: false, value: null });
  }
  res.json({ exists: true, value: data });
});

// PUT data by key
app.put("/api/data/:key", (req, res) => {
  const key = req.params.key.replace(/[^a-z0-9_-]/gi, "");
  try {
    // Audit: when vf_orders is updated, diff old vs new and append per-change
    // entries to vf_audit_log. Other keys are written through unchanged.
    if (key === "vf_orders") {
      try { auditOrdersChange(readData("vf_orders") || [], req.body.value || [], req); }
      catch (e) { console.error("[audit] failed to record changes:", e.message); }
    }
    writeData(key, req.body.value);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── Order audit log ──────────────────────────────────────────────────────────
//
// Every PUT to vf_orders runs through auditOrdersChange() which diffs the old
// and new arrays and appends one or more entries to vf_audit_log. The log is
// append-only and capped at AUDIT_MAX entries (oldest evicted) so the file
// can't grow unbounded.

const AUDIT_MAX = 5000;
const AUDITED_FIELDS = [
  "orderId", "sku", "due", "start", "end", "cat", "sub", "region", "temper",
  "machine", "qty", "batches", "total", "status", "priority", "notes",
  "confirmed",
];

function actorFromRequest(req) {
  const users = readData("vf_users") || [];
  const u = users.find(x => x && x.id === req.userId);
  return {
    userId: req.userId || null,
    userName: u ? u.username : "(unknown)",
  };
}

function diffOrder(oldO, newO) {
  const changes = {};
  for (const f of AUDITED_FIELDS) {
    const a = oldO ? oldO[f] : undefined;
    const b = newO ? newO[f] : undefined;
    if (a !== b) changes[f] = { from: a == null ? null : a, to: b == null ? null : b };
  }
  return changes;
}

function appendAuditEntries(entries) {
  if (!entries.length) return;
  const log = readData("vf_audit_log") || [];
  log.push(...entries);
  // Keep only the latest AUDIT_MAX
  const trimmed = log.length > AUDIT_MAX ? log.slice(log.length - AUDIT_MAX) : log;
  writeData("vf_audit_log", trimmed);
}

function auditOrdersChange(oldOrders, newOrders, req) {
  const actor = actorFromRequest(req);
  const source = (req.body && typeof req.body.source === "string") ? req.body.source : "manual";
  const ts = new Date().toISOString();
  const oldById = new Map(oldOrders.map(o => [o.id, o]));
  const newById = new Map(newOrders.map(o => [o.id, o]));
  const entries = [];

  // Created orders
  for (const [id, o] of newById) {
    if (!oldById.has(id)) {
      entries.push({
        ts, ...actor, source,
        action: "create",
        orderId: o.orderId || null,
        entityId: id,
        snapshot: pickAuditedFields(o),
      });
    }
  }
  // Deleted orders
  for (const [id, o] of oldById) {
    if (!newById.has(id)) {
      entries.push({
        ts, ...actor, source,
        action: "delete",
        orderId: o.orderId || null,
        entityId: id,
        snapshot: pickAuditedFields(o),
      });
    }
  }
  // Updated orders
  for (const [id, newO] of newById) {
    const oldO = oldById.get(id);
    if (!oldO) continue;
    const changes = diffOrder(oldO, newO);
    if (Object.keys(changes).length === 0) continue;
    entries.push({
      ts, ...actor, source,
      action: "update",
      orderId: newO.orderId || oldO.orderId || null,
      entityId: id,
      changes,
    });
  }
  appendAuditEntries(entries);
}

function pickAuditedFields(o) {
  const out = {};
  for (const f of AUDITED_FIELDS) if (o[f] !== undefined) out[f] = o[f];
  return out;
}

// GET /api/audit-log — recent entries (admin only)
// requireAdmin is hoisted (function declaration) — defined further below.
app.get("/api/audit-log", (req, res, next) => requireAdmin(req, res, next), (req, res) => {
  const log = readData("vf_audit_log") || [];
  // Newest first, limit 500 by default
  const limit = Math.min(parseInt(req.query.limit, 10) || 500, AUDIT_MAX);
  res.json({ ok: true, total: log.length, entries: log.slice(-limit).reverse() });
});

// POST /api/import-excel/parse — parse an uploaded .xlsx production schedule
app.post("/api/import-excel/parse", upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, error: "No file uploaded" });
  try {
    const wb = XLSX.read(req.file.buffer, { type: "buffer", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    // header:1 → array-of-arrays; defval:'' → empty cells become ""
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    const results = [];

    // Row 0 = title, row 1 = header — data starts at row 2
    // Column layout: A=Section, B=Line/Machine, C=MO#, D=SKU, E=Start, F=End, G=Qty/Notes
    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      const rawMachine = String(row[1] || "").trim();
      const rawMO      = String(row[2] || "").trim();
      const rawSKU     = String(row[3] || "").trim();
      const rawStart   = row[4];
      const rawEnd     = row[5];
      const rawNotes   = String(row[6] || "").trim();

      // Skip section headers and blank rows (no SKU = no real data)
      if (!rawSKU) continue;

      const machKey = mapMachine(rawMachine);
      if (!machKey) continue;   // unrecognised machine — skip

      const mos   = parseMOs(rawMO);
      const start = parseDate(rawStart);
      const end   = parseDate(rawEnd);
      const qty   = parseQty(rawNotes);
      const attribs = detectAttribs(rawSKU, rawMachine);

      if (machKey === "MULTI" && mos.length >= 2) {
        // Mass line rows carry one MO per machine (refining → conching → depositing)
        mos.forEach((mo, idx) => {
          results.push({
            orderId: mo, sku: rawSKU,
            machine: MASS_MACHINES[idx] || "conching",
            start, end, qty, batches: 1, total: qty,
            ...attribs, status: "queued", priority: "med", due: end, notes: rawNotes, confirmed: false,
          });
        });
      } else {
        // If no MO number, generate a stable placeholder from WIP code + machine + start date
        const wipMatch = rawSKU.match(/WIP[-\s]?([\w-]+)/i);
        const wipCode = wipMatch ? wipMatch[1].replace(/\s+/g, "-") : "UNK";
        const orderId = mos[0] || `TBD-${machKey}-${(start || "").replace(/-/g, "")}`;
        const extra = mos.slice(1);
        const notes = [rawNotes, extra.length ? `Also: ${extra.join(", ")}` : ""]
          .filter(Boolean).join(" · ");
        results.push({
          orderId, sku: rawSKU,
          machine: machKey === "MULTI" ? "conching" : machKey,
          start, end, qty, batches: 1, total: qty,
          ...attribs, status: "queued", priority: "med", due: end, notes, confirmed: false,
        });
      }
    }

    const filtered = results.filter(r => r.start);   // drop rows with no parseable start date
    res.json({ ok: true, orders: filtered });
  } catch (e) {
    console.error("Excel parse error:", e);
    res.status(400).json({ ok: false, error: e.message });
  }
});

// ── Claude AI Chat ────────────────────────────────────────────────────────────

const AI_SYSTEM = `You are a production scheduling assistant for Voyage Foods. You help manage the production schedule across multiple machines and zones.

Machines (use these exact keys):
- Zone 1 Seed prep: seed_clean (Seed Cleaning), roaster (Alk/Roaster)
- Zone 1 Mcintyres: east_mac (East Mac CBS), west_mac (West Mac CBE), mac_1250 (1250 Mac), mac_packout (Mac Packout), pouching (Pouching)
- Zone 2 Chocolate: fat_melter (Fat Melter), refining (Refining), conching (Conching), depositing (Depositing)

Capacity & runtime constraints (apply these when recommending slots):
- east_mac / west_mac (Mcintyres): Runtime 12 days. Min 2,250 kg / Max 4,300 kg per batch. EU and US products cannot be mixed on the same run.
- refining: Runtime 1 day. Max capacity 6,000 kg/day (4 batches × 1,500 kg). Min 500 kg per batch.
- conching: Runtime 1 day. Min 3,000 kg / Max 6,000 kg per run.
- roaster: ~325 kg per batch, up to ~4 batches/shift (~1,300 kg/day).
- fat_melter: 24 hr melt cycle. Runs simultaneously with liquor melting.
- Liquor melting (pre-conching): 72 hr.
- Finished chocolate pipeline: roasting → fat melting (24 hr, parallel with liquor melting 72 hr) → refining (1 day) → conching (1 day) → depositing.
- Liquor pipeline: roasting → fat melting (24 hr) → Mcintyre (12 days) → packout.

Order statuses: queued, in-progress, complete, on-hold
Order priorities: high, med, low

When recommending a production slot:
1. Call get_orders to see what's currently scheduled.
2. Call find_available_slots for the relevant machine(s) to identify open windows.
3. Check that the requested quantity fits within machine capacity constraints above.
4. If the product requires multiple machines (e.g. finished chocolate), find slots across the full pipeline.
5. Present 2–3 concrete options (with start/end dates) and explain the trade-offs.
6. Do NOT book anything — only recommend. The user must ask you to create or update an order to make it happen.

IMPORTANT: Before calling any write tools (add_order, shift_machine_orders, update_order_dates, update_order_status, update_order_quantity, delete_order), you MUST:
1. Use get_orders to see the current state
2. Clearly describe to the user exactly what changes you plan to make
3. Wait for them to explicitly confirm (e.g. "yes", "go ahead", "confirm") before executing writes

For delete_order specifically: always name the order ID and SKU in your confirmation request, as deletion is permanent and cannot be undone.

Dates are always in YYYY-MM-DD format.`;

// Tools that mutate vf_orders. Used by /api/chat to set a dataChanged flag
// in the response so the front-end refreshes its local state regardless of
// what wording the model used in its reply text.
const MUTATING_AI_TOOLS = new Set([
  "shift_machine_orders",
  "update_order_dates",
  "update_order_status",
  "add_order",
  "update_order_quantity",
  "delete_order",
]);

const AI_TOOLS = [
  {
    name: "get_orders",
    description: "Get all current work orders from the production schedule. Always call this first to understand the current state before proposing changes.",
    input_schema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "shift_machine_orders",
    description: "Shift all production orders for a specific machine by a number of days. Updates start, end, and due dates. Only affects non-completed orders.",
    input_schema: {
      type: "object",
      properties: {
        machine: { type: "string", description: "Machine key (e.g. east_mac, west_mac, roaster, refining, conching, depositing, mac_1250, mac_packout, seed_clean, fat_melter)" },
        days: { type: "number", description: "Number of days to shift (positive = forward in time, negative = backward)" },
      },
      required: ["machine", "days"],
    },
  },
  {
    name: "update_order_dates",
    description: "Update the start and/or end date of a specific work order by its order ID.",
    input_schema: {
      type: "object",
      properties: {
        order_id: { type: "string", description: "The orderId field of the order (e.g. MO-12345 or TBD-roaster-20260401)" },
        start: { type: "string", description: "New start date in YYYY-MM-DD format (omit to leave unchanged)" },
        end: { type: "string", description: "New end date in YYYY-MM-DD format (omit to leave unchanged)" },
      },
      required: ["order_id"],
    },
  },
  {
    name: "update_order_status",
    description: "Update the status of a specific work order.",
    input_schema: {
      type: "object",
      properties: {
        order_id: { type: "string", description: "The orderId field of the order" },
        status: { type: "string", enum: ["queued", "in-progress", "complete", "on-hold"] },
      },
      required: ["order_id", "status"],
    },
  },
  {
    name: "add_order",
    description: "Create a new work order on the production schedule. Only call this after the user has confirmed they want to proceed with a specific slot.",
    input_schema: {
      type: "object",
      properties: {
        sku:      { type: "string", description: "SKU / product description" },
        machine:  { type: "string", description: "Machine key (e.g. east_mac, conching, roaster)" },
        start:    { type: "string", description: "Start date YYYY-MM-DD" },
        end:      { type: "string", description: "End date YYYY-MM-DD" },
        qty:      { type: "number", description: "Batch quantity in kg" },
        orderId:  { type: "string", description: "MO number if known (e.g. MO-00999), otherwise omit and one will be generated" },
        priority: { type: "string", enum: ["high", "med", "low"], description: "Priority (default: med)" },
        notes:    { type: "string", description: "Any notes or special instructions" },
      },
      required: ["sku", "machine", "start", "end"],
    },
  },
  {
    name: "find_available_slots",
    description: "Scan the current schedule for available gaps on a machine that could fit a new production order. Returns concrete date windows. Always call this before recommending a slot to the user.",
    input_schema: {
      type: "object",
      properties: {
        machine:        { type: "string",  description: "Machine key to check" },
        duration_days:  { type: "number",  description: "How many calendar days the order needs" },
        qty:            { type: "number",  description: "Quantity in kg (used to validate against capacity)" },
        earliest_start: { type: "string",  description: "Don't suggest slots before this date (YYYY-MM-DD). Defaults to today." },
        count:          { type: "number",  description: "How many slot options to return (default 3)" },
      },
      required: ["machine", "duration_days"],
    },
  },
  {
    name: "update_order_quantity",
    description: "Update the batch quantity (kg) of a specific work order. Also recalculates the total.",
    input_schema: {
      type: "object",
      properties: {
        order_id: { type: "string", description: "The orderId field of the order" },
        qty: { type: "number", description: "New batch quantity in kg" },
        batches: { type: "number", description: "Number of batches (optional, defaults to existing value)" },
      },
      required: ["order_id", "qty"],
    },
  },
  {
    name: "delete_order",
    description: "Permanently delete a work order from the schedule. Use only after the user has explicitly confirmed the deletion.",
    input_schema: {
      type: "object",
      properties: {
        order_id: { type: "string", description: "The orderId field of the order to delete" },
      },
      required: ["order_id"],
    },
  },
];

function addDays(dateStr, days) {
  if (!dateStr) return dateStr;
  const d = new Date(dateStr + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function daysBetween(a, b) {
  const msPerDay = 86400000;
  return Math.floor((new Date(b + "T00:00:00Z") - new Date(a + "T00:00:00Z")) / msPerDay);
}

function shiftDate(dateStr, days) {
  if (!dateStr) return dateStr;
  const d = new Date(dateStr + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

async function executeAITool(name, input) {
  const orders = readData("vf_orders") || [];
  switch (name) {
    case "get_orders": {
      const summary = orders.map(o => ({
        orderId: o.orderId || o.id,
        sku: o.sku,
        machine: o.machine,
        start: o.start,
        end: o.end,
        status: o.status,
        priority: o.priority,
        qty: o.qty,
      }));
      return { count: summary.length, orders: summary };
    }

    case "shift_machine_orders": {
      const { machine, days } = input;
      const affected = orders.filter(o => o.machine === machine && o.status !== "complete");
      if (!affected.length) {
        return { ok: true, affected: 0, message: `No active orders found for machine '${machine}'` };
      }
      affected.forEach(order => {
        const idx = orders.findIndex(o => o.id === order.id);
        if (idx !== -1) {
          orders[idx].start = shiftDate(orders[idx].start, days);
          orders[idx].end   = shiftDate(orders[idx].end,   days);
          orders[idx].due   = shiftDate(orders[idx].due,   days);
        }
      });
      writeData("vf_orders", orders);
      return {
        ok: true,
        affected: affected.length,
        message: `Shifted ${affected.length} order(s) on '${machine}' by ${days > 0 ? "+" : ""}${days} days`,
        orders: affected.map(o => o.orderId || o.id),
      };
    }

    case "update_order_dates": {
      const { order_id, start, end } = input;
      const idx = orders.findIndex(o => o.orderId === order_id || o.id === order_id);
      if (idx === -1) return { ok: false, error: `Order '${order_id}' not found` };
      if (start) orders[idx].start = start;
      if (end)   { orders[idx].end = end; orders[idx].due = end; }
      writeData("vf_orders", orders);
      return { ok: true, message: `Updated order '${order_id}'`, start: orders[idx].start, end: orders[idx].end };
    }

    case "update_order_status": {
      const { order_id, status } = input;
      const idx = orders.findIndex(o => o.orderId === order_id || o.id === order_id);
      if (idx === -1) return { ok: false, error: `Order '${order_id}' not found` };
      orders[idx].status = status;
      writeData("vf_orders", orders);
      return { ok: true, message: `Order '${order_id}' status set to '${status}'` };
    }

    case "add_order": {
      const { sku, machine, start, end, qty = 0, orderId, priority = "med", notes = "" } = input;
      const id = `ai-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
      const generatedOrderId = orderId || `TBD-${machine}-${(start || "").replace(/-/g, "")}`;
      const newOrder = {
        id, orderId: generatedOrderId, sku, machine,
        start, end, due: end,
        qty, batches: 1, total: qty,
        status: "queued", priority, notes,
        cat: "liquor", sub: "liquor",
        confirmed: false,
      };
      orders.push(newOrder);
      writeData("vf_orders", orders);
      return { ok: true, message: `Created order '${generatedOrderId}' for ${sku} on ${machine} (${start} → ${end})`, order: newOrder };
    }

    case "find_available_slots": {
      const { machine, duration_days, qty, earliest_start, count = 3 } = input;
      const today = new Date().toISOString().slice(0, 10);
      const from  = earliest_start || today;

      // Active orders on this machine that overlap or start after `from`, sorted by start
      const active = orders
        .filter(o => o.machine === machine && o.status !== "complete" && o.start && (o.end || o.start) >= from)
        .sort((a, b) => a.start.localeCompare(b.start));

      const slots = [];

      // Walk through the timeline looking for gaps
      let cursor = from;
      for (const order of active) {
        const orderStart = order.start;
        const orderEnd   = addDays(order.end || order.start, 1); // day after order ends

        if (orderStart > cursor) {
          const gap = daysBetween(cursor, orderStart);
          if (gap >= duration_days) {
            const slotEnd = addDays(cursor, duration_days);
            slots.push({
              start:      cursor,
              end:        slotEnd,
              gap_days:   gap,
              fits:       true,
              note:       slots.length === 0 ? "Earliest available gap" : "Gap in schedule",
            });
            if (slots.length >= count) break;
          }
        }
        // Advance cursor past this order
        if (orderEnd > cursor) cursor = orderEnd;
      }

      // Fill remaining options after the last scheduled order
      while (slots.length < count) {
        slots.push({
          start:    cursor,
          end:      addDays(cursor, duration_days),
          gap_days: null,
          fits:     true,
          note:     "After end of current schedule",
        });
        cursor = addDays(cursor, duration_days + 1);
      }

      // Capacity sanity check
      const capacityWarnings = [];
      if ((machine === "east_mac" || machine === "west_mac") && qty) {
        if (qty < 2250) capacityWarnings.push(`Qty ${qty} kg is below minimum 2,250 kg for Mcintyres`);
        if (qty > 4300) capacityWarnings.push(`Qty ${qty} kg exceeds maximum 4,300 kg for Mcintyres`);
      }
      if (machine === "refining" && qty) {
        if (qty > 6000) capacityWarnings.push(`Qty ${qty} kg exceeds daily max 6,000 kg for Refining`);
      }
      if (machine === "conching" && qty) {
        if (qty < 3000) capacityWarnings.push(`Qty ${qty} kg is below minimum 3,000 kg for Conching`);
        if (qty > 6000) capacityWarnings.push(`Qty ${qty} kg exceeds maximum 6,000 kg for Conching`);
      }

      return { machine, duration_days, qty: qty || null, slots, capacityWarnings };
    }

    case "update_order_quantity": {
      const { order_id, qty, batches } = input;
      const idx = orders.findIndex(o => o.orderId === order_id || o.id === order_id);
      if (idx === -1) return { ok: false, error: `Order '${order_id}' not found` };
      orders[idx].qty = qty;
      if (batches !== undefined) orders[idx].batches = batches;
      orders[idx].total = qty * (orders[idx].batches || 1);
      writeData("vf_orders", orders);
      return { ok: true, message: `Updated '${order_id}' quantity to ${qty} kg (total: ${orders[idx].total} kg)` };
    }

    case "delete_order": {
      const { order_id } = input;
      const idx = orders.findIndex(o => o.orderId === order_id || o.id === order_id);
      if (idx === -1) return { ok: false, error: `Order '${order_id}' not found` };
      const deleted = orders.splice(idx, 1)[0];
      writeData("vf_orders", orders);
      return { ok: true, message: `Deleted order '${order_id}' (${deleted.sku || ""})` };
    }

    default:
      return { ok: false, error: `Unknown tool: ${name}` };
  }
}

app.post("/api/chat", async (req, res) => {
  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(503).json({ ok: false, error: "AI assistant is not configured (missing ANTHROPIC_API_KEY)" });
  }
  const { messages } = req.body;
  if (!messages || !Array.isArray(messages)) {
    return res.status(400).json({ ok: false, error: "messages array required" });
  }

  try {
    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
    let currentMessages = messages.map(m => ({
      role: m.role,
      content: m.content,
    }));
    let response;
    const mutatedTools = []; // names of mutating tools invoked this turn

    // Agentic loop — max 8 tool-use rounds
    for (let i = 0; i < 8; i++) {
      response = await client.messages.create({
        model: "claude-opus-4-6",
        max_tokens: 4096,
        system: AI_SYSTEM,
        tools: AI_TOOLS,
        messages: currentMessages,
      });

      if (response.stop_reason !== "tool_use") break;

      // Execute all tool calls
      const toolUseBlocks = response.content.filter(b => b.type === "tool_use");
      currentMessages.push({ role: "assistant", content: response.content });

      const toolResults = await Promise.all(
        toolUseBlocks.map(async block => {
          // Track mutating tool invocations so the front-end knows to refresh
          // its local data — replaces the fragile reply-text regex matching.
          if (MUTATING_AI_TOOLS.has(block.name)) {
            mutatedTools.push(block.name);
          }
          return {
            type: "tool_result",
            tool_use_id: block.id,
            content: JSON.stringify(await executeAITool(block.name, block.input)),
          };
        })
      );
      currentMessages.push({ role: "user", content: toolResults });
    }

    // Extract the final text reply
    const reply = (response.content || [])
      .filter(b => b.type === "text")
      .map(b => b.text)
      .join("\n")
      .trim();

    // Append final assistant message to the conversation
    currentMessages.push({ role: "assistant", content: response.content });

    res.json({
      ok: true,
      reply,
      messages: currentMessages,
      dataChanged: mutatedTools.length > 0,
      mutatedTools,
    });
  } catch (e) {
    console.error("AI chat error:", e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── Cin7 Core sync ────────────────────────────────────────────────────────────
// Field-name map — if the first run shows different keys in /api/sync-cin7/test,
// update the values here and redeploy.
const C7 = {
  base:       "https://inventory.dearsystems.com/ExternalApi/v2",
  path:       "/inventorymovements",
  arrKey:     "InventoryMovements",   // top-level array key in response
  sku:        "SKU",
  name:       "Name",
  category:   "Category",
  unit:       "Unit",
  date:       "Date",                 // ISO string or YYYY-MM-DD
  reason:     "Reason",               // reference type (e.g. "Manufacturing Order")
  reference:  "Reference",            // MO number (e.g. "MO-00728")
  qtyIn:      "QtyIN",
  qtyOut:     "QtyOUT",
};

const CIN7_SYNC_DAYS = parseInt(process.env.CIN7_SYNC_DAYS || "90", 10);

function cin7Headers() {
  return {
    "api-auth-accountid":      process.env.CIN7_ACCOUNT_ID,
    "api-auth-applicationkey": process.env.CIN7_APPLICATION_KEY,
    "Content-Type":            "application/json",
  };
}

async function fetchCin7Movements(days) {
  if (!process.env.CIN7_ACCOUNT_ID || !process.env.CIN7_APPLICATION_KEY) {
    throw new Error("CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY environment variable not set");
  }
  const windowDays = days ?? CIN7_SYNC_DAYS;
  const end = new Date();
  let start = new Date();
  start.setDate(start.getDate() - windowDays);
  // Snap start to the 1st of its month when the caller requested a short window,
  // so the per-month merge in performCin7Sync doesn't replace a month bucket with
  // a partial fetch and drop earlier days.
  if (days != null) {
    start = new Date(start.getFullYear(), start.getMonth(), 1);
  }
  const startDate = start.toISOString().slice(0, 10);
  const endDate   = end.toISOString().slice(0, 10);

  const all = [];
  let page  = 1;
  const limit = 1000;

  while (true) {
    const url = `${C7.base}${C7.path}?Page=${page}&Limit=${limit}&StartDate=${startDate}&EndDate=${endDate}`;
    const resp = await fetch(url, { headers: cin7Headers() });
    const ct   = resp.headers.get("content-type") || "";
    const text = await resp.text().catch(() => "");
    if (!resp.ok) {
      throw new Error(`Cin7 API ${resp.status} ${resp.statusText} [${ct}]: ${text.slice(0, 300)}`);
    }
    let data;
    try {
      data = JSON.parse(text);
    } catch (_) {
      throw new Error(`Cin7 API ${resp.status} returned non-JSON [${ct}]: ${text.slice(0, 300)}`);
    }
    const batch = data[C7.arrKey] || [];
    all.push(...batch);
    // Stop when we get fewer records than the page size
    if (batch.length < limit) break;
    page++;
  }
  return { movements: all, startDate, endDate };
}

function buildInventoryFromCin7(movements) {
  const skuMap   = new Map();
  const catAgg   = {};
  const rtAgg    = {};
  const moMap    = new Map();
  const monthSet = new Set();

  for (const m of movements) {
    const sku  = m[C7.sku]  || m.Sku  || m.sku  || "";  if (!sku)  continue;
    const date = String(m[C7.date] || "").slice(0, 10);
    const month = date.slice(0, 7);                        if (!month) continue;

    const prod = m[C7.name]      || "";
    const cat  = m[C7.category]  || "";
    const unit = m[C7.unit]      || "kg";
    const rt   = m[C7.reason]    || "";
    const ref  = m[C7.reference] || "";
    const inb  = parseFloat(m[C7.qtyIn]  || 0) || 0;
    const outb = parseFloat(m[C7.qtyOut] || 0) || 0;

    monthSet.add(month);

    if (!skuMap.has(sku)) {
      skuMap.set(sku, { s: sku, p: prod, c: cat, u: unit, m: {}, rt: {}, ti: 0, to: 0, net: 0, bc: 0 });
    }
    const entry = skuMap.get(sku);
    if (!entry.m[month]) entry.m[month] = { i: 0, o: 0 };
    entry.m[month].i += inb;
    entry.m[month].o += outb;
    entry.ti  += inb;
    entry.to  += outb;
    entry.net += inb - outb;

    if (rt) {
      if (!entry.rt[rt])       entry.rt[rt] = { i: 0, o: 0 };
      entry.rt[rt].i += inb;  entry.rt[rt].o += outb;
    }
    if (!catAgg[cat])        catAgg[cat] = {};
    if (!catAgg[cat][month]) catAgg[cat][month] = { i: 0, o: 0 };
    catAgg[cat][month].i += inb; catAgg[cat][month].o += outb;

    if (!rtAgg[rt])        rtAgg[rt] = {};
    if (!rtAgg[rt][month]) rtAgg[rt][month] = { i: 0, o: 0 };
    rtAgg[rt][month].i += inb; rtAgg[rt][month].o += outb;

    // Track per-MO movements
    const moMatch = ref.match(/MO-\d+/);
    if (moMatch) {
      const moId = moMatch[0];
      if (!moMap.has(moId)) moMap.set(moId, { mo: moId, sku, prod, totalIn: 0, totalOut: 0 });
      moMap.get(moId).totalIn  += inb;
      moMap.get(moId).totalOut += outb;
    }
  }

  return {
    invSku:      [...skuMap.values()],
    invCat:      catAgg,
    invRt:       rtAgg,
    months:      [...monthSet].sort(),
    moMovements: Object.fromEntries(moMap),
  };
}

async function performCin7Sync(days) {
  const { movements, startDate, endDate } = await fetchCin7Movements(days);
  const fresh    = buildInventoryFromCin7(movements);
  const existing = readData("inventory");

  let merged;
  if (!existing || !existing.invSku || !existing.invSku.length) {
    // No prior data — just write fresh
    merged = fresh;
  } else {
    // Merge by calendar month:
    //   months inside the sync window  → replaced by fresh Cin7 data
    //   months outside the sync window → kept from existing storage
    const syncMonths = new Set(fresh.months);

    // ── invSku ──────────────────────────────────────────────────────────────
    const skuMap = new Map();

    // Seed with existing data, stripping months that the fresh pull covers
    for (const e of existing.invSku) {
      const entry = { ...e, m: {}, rt: { ...e.rt }, ti: 0, to: 0, net: 0, bc: e.bc || 0 };
      for (const mo in e.m) {
        if (!syncMonths.has(mo)) {
          entry.m[mo]  = e.m[mo];
          entry.ti    += e.m[mo].i;
          entry.to    += e.m[mo].o;
          entry.net   += e.m[mo].i - e.m[mo].o;
        }
      }
      skuMap.set(e.s, entry);
    }

    // Layer fresh months on top
    for (const e of fresh.invSku) {
      if (!skuMap.has(e.s)) {
        skuMap.set(e.s, { ...e });
      } else {
        const entry = skuMap.get(e.s);
        for (const mo in e.m) {
          entry.m[mo]  = e.m[mo];
          entry.ti    += e.m[mo].i;
          entry.to    += e.m[mo].o;
          entry.net   += e.m[mo].i - e.m[mo].o;
        }
        // Fresh rt data wins for sync-window ref types
        for (const rt in e.rt) entry.rt[rt] = e.rt[rt];
        if (!entry.p && e.p) entry.p = e.p;
      }
    }

    // ── catAgg / rtAgg — replace sync-window months, keep the rest ──────────
    const mergeSrc = (old, next) => {
      const out = {};
      for (const key in old) {
        out[key] = {};
        for (const mo in old[key]) {
          if (!syncMonths.has(mo)) out[key][mo] = old[key][mo];
        }
      }
      for (const key in next) {
        if (!out[key]) out[key] = {};
        for (const mo in next[key]) out[key][mo] = next[key][mo];
      }
      return out;
    };

    // ── moMovements — keep historical, fresh wins for any MO it contains ────
    const mergedMo = { ...(existing.moMovements || {}), ...fresh.moMovements };

    merged = {
      invSku:      [...skuMap.values()],
      invCat:      mergeSrc(existing.invCat || {}, fresh.invCat),
      invRt:       mergeSrc(existing.invRt  || {}, fresh.invRt),
      months:      [...new Set([...(existing.months || []), ...fresh.months])].sort(),
      moMovements: mergedMo,
    };
  }

  writeData("inventory", merged);

  const status = {
    ok:            true,
    lastSync:      new Date().toISOString(),
    source:        "cin7",
    movementCount: movements.length,
    skuCount:      merged.invSku.length,
    moCount:       Object.keys(merged.moMovements).length,
    startDate,
    endDate,
  };
  writeData("vf_sync_status", status);
  return status;
}

// POST /api/sync-cin7 — DISABLED. Cin7 deprecated /inventorymovements (2026)
// without a v2 replacement that exposes production-receipt data, so the API
// path can't reconstruct what we need. Inventory data is now loaded by
// uploading the daily Cin7 "Inventory Movement Details" report (CSV/XLSX)
// via the Traceability tab. Re-enable this if Cin7 ships a working endpoint.
app.post("/api/sync-cin7", async (req, res) => {
  res.status(503).json({
    ok: false,
    disabled: true,
    error: "Cin7 API sync is disabled — Cin7 removed the /inventorymovements endpoint. Upload the daily Inventory Movement Details report on the Traceability tab instead.",
  });
});

// GET /api/sync-cin7/status — last sync metadata
app.get("/api/sync-cin7/status", (req, res) => {
  const status = readData("vf_sync_status") || { ok: false, lastSync: null };
  res.json(status);
});

// GET /api/sync-cin7/test — fetch one page raw, for verifying field names
app.get("/api/sync-cin7/test", async (req, res) => {
  try {
    if (!process.env.CIN7_ACCOUNT_ID || !process.env.CIN7_APPLICATION_KEY) {
      return res.status(503).json({ ok: false, error: "CIN7 credentials not configured" });
    }
    const url  = `${C7.base}${C7.path}?Page=1&Limit=3`;
    const resp = await fetch(url, { headers: cin7Headers() });
    const ct   = resp.headers.get("content-type") || "";
    const text = await resp.text().catch(() => "");
    let sample;
    let parseError = null;
    try {
      sample = JSON.parse(text);
    } catch (e) {
      parseError = e.message;
      sample = text.slice(0, 500);
    }
    res.json({ ok: resp.ok, status: resp.status, contentType: ct, fieldMap: C7, sample, parseError });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// Daily Cin7 movement-sync cron is DISABLED — see the /api/sync-cin7 comment.
// Inventory MOVEMENTS still come from manual report uploads on the Traceability
// tab. The on-hand snapshot below is a separate, working API path.
console.log("[Cin7] Movement sync disabled — using manual report uploads instead.");

// ── Cin7 on-hand inventory (Phase 1 of the MRP feature) ──────────────────────
//
// Cin7 Core's V1 endpoint /ExternalApi/ProductAvailability returns a snapshot
// of stock per SKU × Location × Batch — the only Cin7 endpoint we've found
// that gives us live on-hand quantities. It does NOT include movement
// history or MO linkage; for that we still rely on the manual upload path.

const C7_ONHAND_URL = "https://inventory.dearsystems.com/ExternalApi/ProductAvailability";

async function fetchCin7OnHand() {
  if (!process.env.CIN7_ACCOUNT_ID || !process.env.CIN7_APPLICATION_KEY) {
    throw new Error("CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY environment variable not set");
  }
  const all = [];
  let page = 1;
  const limit = 1000;
  while (true) {
    const url = `${C7_ONHAND_URL}?Page=${page}&Limit=${limit}`;
    const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
    const ct = resp.headers.get("content-type") || "";
    const text = await resp.text().catch(() => "");
    if (!resp.ok) {
      throw new Error(`Cin7 ProductAvailability ${resp.status} ${resp.statusText} [${ct}]: ${text.slice(0, 300)}`);
    }
    let data;
    try { data = JSON.parse(text); }
    catch (_) { throw new Error(`Cin7 ProductAvailability ${resp.status} returned non-JSON [${ct}]: ${text.slice(0, 300)}`); }
    const batch = data.ProductAvailability || [];
    all.push(...batch);
    if (batch.length < limit) break;
    page++;
    if (page > 50) throw new Error("Cin7 ProductAvailability pagination exceeded 50 pages — aborting");
  }
  return all;
}

async function performCin7OnHandSync() {
  const rows = await fetchCin7OnHand();
  const now = new Date().toISOString();
  // Per-SKU rollup: sum across batches/locations for the MRP engine
  const bySku = {};
  for (const r of rows) {
    const sku = r.SKU;
    if (!sku) continue;
    if (!bySku[sku]) {
      bySku[sku] = {
        sku,
        name: r.Name || "",
        onHand: 0,
        allocated: 0,
        available: 0,
        onOrder: 0,
        inTransit: 0,
        stockOnHand: 0,
        locations: new Set(),
        batches: 0,
        nextDelivery: null,
      };
    }
    const a = bySku[sku];
    a.onHand += Number(r.OnHand) || 0;
    a.allocated += Number(r.Allocated) || 0;
    a.available += Number(r.Available) || 0;
    a.onOrder += Number(r.OnOrder) || 0;
    a.inTransit += Number(r.InTransit) || 0;
    a.stockOnHand += Number(r.StockOnHand) || 0;
    a.batches += 1;
    if (r.Location) a.locations.add(r.Location);
    if (r.NextDeliveryDate && (!a.nextDelivery || r.NextDeliveryDate < a.nextDelivery)) {
      a.nextDelivery = r.NextDeliveryDate;
    }
  }
  // Convert location Sets to sorted arrays for serialization
  const skuRollup = Object.values(bySku).map(a => ({ ...a, locations: [...a.locations].sort() }));
  const blob = {
    lastSync: now,
    rowCount: rows.length,
    skuCount: skuRollup.length,
    rows,           // raw per SKU/location/batch records
    bySku: skuRollup, // rolled up per-SKU totals (what MRP uses)
  };
  writeData("inventory_onhand", blob);
  return { ok: true, lastSync: now, rowCount: rows.length, skuCount: skuRollup.length };
}

// Admin-only middleware
function requireAdmin(req, res, next) {
  const users = readData("vf_users") || [];
  const u = users.find(x => x && x.id === req.userId);
  if (!u || u.role !== "admin") {
    return res.status(403).json({ ok: false, error: "Admin role required" });
  }
  next();
}

// POST /api/cin7/onhand/sync — admin-triggered live pull
app.post("/api/cin7/onhand/sync", requireAdmin, async (req, res) => {
  try {
    const status = await performCin7OnHandSync();
    res.json(status);
  } catch (e) {
    console.error("[Cin7 OnHand] Sync error:", e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// GET /api/cin7/onhand — return the cached on-hand snapshot (any authed user)
app.get("/api/cin7/onhand", (req, res) => {
  const blob = readData("inventory_onhand");
  if (!blob) return res.json({ ok: true, lastSync: null, bySku: [], rows: [] });
  res.json({ ok: true, ...blob });
});

// GET /api/cin7/onhand/status — last sync metadata only (cheap)
app.get("/api/cin7/onhand/status", (req, res) => {
  const blob = readData("inventory_onhand");
  if (!blob) return res.json({ ok: true, lastSync: null });
  res.json({ ok: true, lastSync: blob.lastSync, rowCount: blob.rowCount, skuCount: blob.skuCount });
});

// ── MRP Phase 2: BOM + supply settings (lead times, safety stock) ────────────
//
// Storage shape:
//   vf_boms              — { lastImport, parents: { <sku>: [{ version, ... }] } }
//   vf_supply_settings   — { lastImport, defaults: {...}, perSku: { <sku>: {...} } }
//
// Default version (VersionDefault=Yes) is used for MRP requirements; other
// versions are kept in the parents[sku] array so a UI can show them all.

// Map BOM WorkCentreName → app machine key (or null for "priority 2 — map later")
const WC_TO_MACHINE = {
  "Refine":               "refining",
  "Conch":                "conching",
  "Drops and Pack":       "depositing",
  "Pack in Pouch":        "pouching",
  "Pack from MAC":        "mac_packout",
  "Clean Seeds":          "seed_clean",
  "Grape Seeds":          "roaster",
  "Ground BFC":           "grinder",
  "Liquor":               "__liquor_split__",  // resolved from fat-type RM
  "Final Blend":          null,
  "Hazelnut Free Spread": null,
  "Concentrate":          null,
  "Paste":                null,
  "Testing work center":  null,
};

// Production lead time per machine, in days (sourced from Capacity & Ops tab)
const MACHINE_LEAD_DAYS = {
  seed_clean:  1,
  roaster:     1,
  east_mac:    12,
  west_mac:    12,
  mac_1250:    1,
  mac_packout: 10,
  pouching:    1,
  fat_melter:  3,
  refining:    1,
  conching:    1,
  depositing:  1,
  grinder:     1,   // 1000 kg / 2-hr shift; defaults to 1 day for typical runs
};

// CSV line parser that handles quoted fields and commas-inside-quotes
function parseCsvLine(line) {
  const out = []; let cur = ""; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQ) {
      if (ch === '"' && line[i+1] === '"') { cur += '"'; i++; }
      else if (ch === '"') inQ = false;
      else cur += ch;
    } else {
      if (ch === '"') inQ = true;
      else if (ch === ',') { out.push(cur); cur = ""; }
      else cur += ch;
    }
  }
  out.push(cur);
  return out.map(s => s.trim());
}

// Parse Cin7 BOM CSV (Action,ProductSKU,...). Returns the canonical vf_boms blob.
function parseBomCsv(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 2) throw new Error("BOM CSV has no data rows");
  const hdr = parseCsvLine(lines[0]);
  const idx = {};
  hdr.forEach((h, i) => idx[h] = i);
  const required = ["ProductSKU", "ItemType", "ComponentSKU_ResourceCode", "Quantity", "Version"];
  for (const r of required) if (idx[r] === undefined) throw new Error(`BOM CSV missing column: ${r}`);

  // First pass: bucket rows by parent + version
  const bucket = {}; // key = parent + '|' + version
  for (let i = 1; i < lines.length; i++) {
    const f = parseCsvLine(lines[i]);
    const parent = f[idx.ProductSKU];
    if (!parent) continue;
    const itemType = f[idx.ItemType];
    if (itemType === "Output" || itemType === "Resource") continue;
    if (itemType !== "Component") continue;
    const componentSku = f[idx.ComponentSKU_ResourceCode];
    if (!componentSku) continue;
    const version = String(f[idx.Version] || "1");
    const key = parent + "|" + version;
    if (!bucket[key]) {
      bucket[key] = {
        parent,
        parentName: f[idx.ProductName] || "",
        version,
        versionName: f[idx.VersionName] || "",
        isDefault: (f[idx.VersionDefault] || "").toLowerCase() === "yes",
        // QuantityToProduce — the batch size the BOM is defined for. Component
        // qtys must be divided by this to get the per-unit-of-parent ratio.
        // Defaults to 1 if missing/blank/zero (treats blanks as a no-op).
        qtyToProduce: parseFloat(f[idx.QuantityToProduce]) || 1,
        runSize: parseFloat(f[idx.RunSize]) || 0,
        minQty: parseFloat(f[idx.MinQuantity]) || 0,
        maxQty: parseFloat(f[idx.MaxQuantity]) || 0,
        productionLeadTimeRaw: parseInt(f[idx.ProductionLeadTime], 10) || 0,
        components: [],
      };
    }
    bucket[key].components.push({
      sku: componentSku,
      name: f[idx.ComponentName_ResourceName] || "",
      qty: parseFloat(f[idx.Quantity]) || 0,
      wastagePct: parseFloat(f[idx.WastagePercent_ForStockComponentOnly]) || 0,
      op: parseInt(f[idx.OperationSequence], 10) || 1,
      opName: f[idx.OperationName] || "",
      workCentre: f[idx.WorkCentreName] || "",
    });
  }

  // Second pass: derive machine + production lead time per BOM
  const parents = {};
  for (const key of Object.keys(bucket)) {
    const b = bucket[key];
    b.machine = deriveMachineFromBom(b);
    b.productionLeadTime = b.machine ? (MACHINE_LEAD_DAYS[b.machine] || null) : null;
    if (!parents[b.parent]) parents[b.parent] = [];
    parents[b.parent].push(b);
  }
  // Sort each parent's versions: default first, then by version number
  for (const sku of Object.keys(parents)) {
    parents[sku].sort((a, b) => {
      if (a.isDefault !== b.isDefault) return a.isDefault ? -1 : 1;
      return Number(a.version) - Number(b.version);
    });
  }
  return {
    lastImport: new Date().toISOString(),
    parents,
    parentCount: Object.keys(parents).length,
    rowCount: Object.values(parents).reduce((s, vs) => s + vs.reduce((s2, v) => s2 + v.components.length, 0), 0),
  };
}

// Resolve the machine key for a BOM. Most work centres map directly; Liquor
// splits to east_mac / west_mac based on whether the components include a CBE
// (RM-120002-00 Coberine) or CBS (RM-120004-00 PK-100) fat.
function deriveMachineFromBom(bom) {
  if (!bom.components.length) return null;
  // Use the first operation's work centre as the primary
  const wc = bom.components[0].workCentre;
  const m = WC_TO_MACHINE[wc];
  if (m === undefined) return null; // unknown WC
  if (m === null) return null;       // priority-2, deliberately unmapped
  if (m !== "__liquor_split__") return m;
  // Liquor split — scan all components for a fat SKU
  const skus = new Set(bom.components.map(c => c.sku));
  if (skus.has("RM-120002-00")) return "west_mac"; // Coberine = CBE
  if (skus.has("RM-120004-00")) return "east_mac"; // PK-100 = CBS
  return null;
}

// Parse the lead-time CSV (SKU, Product Name, "WOS & Lead").
// Per-user: when the cell reads "Xwks & Y Days", use ONLY the days portion.
// "contract" → flag as contract-managed (no PO suggestions).
// "more to -00" / similar → alias resolved later in normalizeSupplySettings.
function parseLeadTimeCsv(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return [];
  const hdr = parseCsvLine(lines[0]).map(h => h.toLowerCase());
  const skuIdx = hdr.findIndex(h => h === "sku");
  const ltIdx  = hdr.findIndex(h => h.includes("lead") || h === "wos & lead");
  if (skuIdx === -1 || ltIdx === -1) throw new Error("Lead-time CSV needs 'SKU' and 'WOS & Lead' (or 'Lead') columns");
  const out = [];
  for (let i = 1; i < lines.length; i++) {
    const f = parseCsvLine(lines[i]);
    const sku = f[skuIdx]; if (!sku) continue;
    const raw = (f[ltIdx] || "").trim();
    const parsed = parseLeadTimeCell(raw);
    out.push({ sku, raw, ...parsed });
  }
  return out;
}

// "X Days" / "Xwks & Y Days" / "contract" / "more to -00" / "" → structured
function parseLeadTimeCell(raw) {
  if (!raw) return { leadTimeDays: null, isContract: false, alias: null };
  const lower = raw.toLowerCase();
  if (lower.includes("contract")) return { leadTimeDays: null, isContract: true, alias: null };
  // "more to -00" or "more to RM-XXXX" → alias
  const aliasMatch = raw.match(/more to\s+(\S+)/i);
  if (aliasMatch) return { leadTimeDays: null, isContract: false, alias: aliasMatch[1] };
  // "Xwks & Y Days" — per user, take ONLY the days portion
  const both = raw.match(/(\d+)\s*wk[s]?\s*&\s*(\d+)\s*day/i);
  if (both) return { leadTimeDays: parseInt(both[2], 10), isContract: false, alias: null };
  // "X Days"
  const days = raw.match(/(\d+)\s*day/i);
  if (days) return { leadTimeDays: parseInt(days[1], 10), isContract: false, alias: null };
  return { leadTimeDays: null, isContract: false, alias: null };
}

// Normalize raw lead-time entries into the canonical perSku map. Resolves
// aliases ("more to -00" → copy target's settings) and drops malformed rows.
function normalizeSupplySettings(rawEntries, defaults) {
  const perSku = {};
  // First pass: direct entries
  for (const e of rawEntries) {
    if (!e.sku || e.alias) continue;
    perSku[e.sku] = {
      leadTimeDays: e.leadTimeDays,
      isContract: e.isContract,
      isAlias: false,
      raw: e.raw,
    };
  }
  // Second pass: alias entries point to direct entries
  for (const e of rawEntries) {
    if (!e.alias) continue;
    // alias target like "-00" means "same as <prefix>-00" — resolve by prefix match
    let targetSku = e.alias;
    if (targetSku.startsWith("-")) {
      // user shorthand — find a sibling with the same prefix
      const prefix = e.sku.replace(/-\d+$/, "");
      targetSku = prefix + targetSku;
    }
    const target = perSku[targetSku];
    if (target) {
      perSku[e.sku] = { ...target, isAlias: true, aliasOf: targetSku, raw: e.raw };
    } else {
      perSku[e.sku] = { leadTimeDays: null, isContract: false, isAlias: true, aliasOf: targetSku, aliasUnresolved: true, raw: e.raw };
    }
  }
  return { lastImport: new Date().toISOString(), defaults, perSku };
}

// Recursively expand a BOM to leaf-RM requirements for a given parent + qty.
// Returns { leaves: { sku: {qty, name, leafSku} }, intermediates: [{sku, qty, version, depth}] }
// Cycles are detected via the visited set; a leaf is anything not in vf_boms.parents.
function expandBom(parents, parentSku, qty, opts) {
  opts = opts || {};
  const applyWastage = opts.applyWastage !== false; // default true
  const visited = new Set();
  const leaves = {};
  const trail = [];

  function recurse(sku, needed, depth) {
    if (depth > 12) throw new Error(`BOM recursion too deep at ${sku}`);
    if (visited.has(sku)) {
      // cycle — log and treat as leaf to bail out gracefully
      if (!leaves[sku]) leaves[sku] = { sku, qty: 0, name: "(cycle detected)", isCycle: true };
      leaves[sku].qty += needed;
      return;
    }
    const versions = parents[sku];
    if (!versions || !versions.length) {
      // leaf RM
      if (!leaves[sku]) leaves[sku] = { sku, qty: 0, name: "" };
      leaves[sku].qty += needed;
      return;
    }
    // Pick default version (first after sort)
    const bom = versions[0];
    visited.add(sku);
    trail.push({ sku, qty: needed, version: bom.version, depth });
    // Normalize each component qty to "per 1 unit of parent" by dividing by
    // the BOM's QuantityToProduce. The Cin7 export defines BOMs at a batch
    // size (e.g. 25 kg of FG-860 needs 25 kg of WIP-5100043, not 25 kg per
    // 1 kg of FG). Without this division, requirements get inflated by the
    // batch size.
    const batchSize = bom.qtyToProduce || 1;
    for (const c of bom.components) {
      const perUnit = c.qty / batchSize;
      const eff = applyWastage ? perUnit * (1 + (c.wastagePct || 0) / 100) : perUnit;
      recurse(c.sku, needed * eff, depth + 1);
    }
    visited.delete(sku);
  }

  recurse(parentSku, qty, 0);
  return { leaves, intermediates: trail };
}

// POST /api/boms/import — admin-only, accepts raw CSV text in body { csv: "..." }
app.post("/api/boms/import", requireAdmin, (req, res) => {
  try {
    const csv = req.body && req.body.csv;
    if (!csv) return res.status(400).json({ ok: false, error: "Missing 'csv' field in body" });
    const blob = parseBomCsv(csv);
    writeData("vf_boms", blob);
    res.json({ ok: true, ...blob, parents: undefined });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// POST /api/supply-settings/import — admin-only. Body: { leadTimeCsv?, packagingCsv?, defaults? }
app.post("/api/supply-settings/import", requireAdmin, (req, res) => {
  try {
    const body = req.body || {};
    const defaults = Object.assign(
      { leadTimeDays: 30, safetyStockDays: 14, packagingDefaultDays: 14 },
      body.defaults || {},
    );
    const all = [];
    if (body.leadTimeCsv) all.push(...parseLeadTimeCsv(body.leadTimeCsv));
    if (body.packagingCsv) all.push(...parseLeadTimeCsv(body.packagingCsv));
    // Apply per-prefix defaults: PK-* without explicit value → packagingDefaultDays
    for (const e of all) {
      if (e.leadTimeDays == null && !e.isContract && !e.alias && e.sku && e.sku.startsWith("PK-")) {
        e.leadTimeDays = defaults.packagingDefaultDays;
        e.appliedPackagingDefault = true;
      }
    }
    const blob = normalizeSupplySettings(all, defaults);
    writeData("vf_supply_settings", blob);
    res.json({ ok: true, lastImport: blob.lastImport, defaults: blob.defaults, skuCount: Object.keys(blob.perSku).length });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// GET /api/boms — full BOM blob (any authed user)
app.get("/api/boms", (req, res) => {
  const blob = readData("vf_boms");
  if (!blob) return res.json({ ok: true, lastImport: null, parents: {}, parentCount: 0 });
  res.json({ ok: true, ...blob });
});

// GET /api/boms/expand?sku=X&qty=Y[&wastage=0] — recursive expansion for testing/MRP
app.get("/api/boms/expand", (req, res) => {
  try {
    const sku = req.query.sku;
    const qty = parseFloat(req.query.qty);
    if (!sku || !isFinite(qty)) return res.status(400).json({ ok: false, error: "Need sku and numeric qty" });
    const blob = readData("vf_boms");
    if (!blob) return res.status(404).json({ ok: false, error: "No BOMs imported" });
    const result = expandBom(blob.parents, sku, qty, { applyWastage: req.query.wastage !== "0" });
    res.json({ ok: true, sku, qty, ...result });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// GET /api/supply-settings — current settings
app.get("/api/supply-settings", (req, res) => {
  const blob = readData("vf_supply_settings");
  if (!blob) return res.json({ ok: true, lastImport: null, defaults: { leadTimeDays: 30, safetyStockDays: 14, packagingDefaultDays: 14 }, perSku: {} });
  res.json({ ok: true, ...blob });
});

// PUT /api/supply-settings — admin-only full-blob replace. The UI edits one
// SKU at a time but each save sends the entire perSku map + defaults so the
// server doesn't need diff logic. Body: { defaults: {...}, perSku: {...} }.
// Validates shapes and rejects malformed input rather than letting bad data
// poison the MRP calc downstream.
app.put("/api/supply-settings", requireAdmin, (req, res) => {
  try {
    const body = req.body || {};
    const defaults = Object.assign(
      { leadTimeDays: 30, safetyStockDays: 14, packagingDefaultDays: 14 },
      body.defaults || {},
    );
    // Coerce default values to non-negative integers
    for (const k of ["leadTimeDays", "safetyStockDays", "packagingDefaultDays"]) {
      const n = parseInt(defaults[k], 10);
      if (!isFinite(n) || n < 0) return res.status(400).json({ ok: false, error: `defaults.${k} must be a non-negative integer` });
      defaults[k] = n;
    }
    const inSku = body.perSku || {};
    if (typeof inSku !== "object" || Array.isArray(inSku)) return res.status(400).json({ ok: false, error: "perSku must be an object" });
    const cleanSku = {};
    for (const sku of Object.keys(inSku)) {
      if (!sku || typeof sku !== "string") continue;
      const v = inSku[sku] || {};
      const entry = {
        leadTimeDays: v.leadTimeDays == null ? null : (parseInt(v.leadTimeDays, 10) || 0),
        isContract: !!v.isContract,
        isAlias: !!v.isAlias,
      };
      if (entry.isAlias && v.aliasOf) entry.aliasOf = String(v.aliasOf);
      if (v.raw != null) entry.raw = String(v.raw);
      // If isContract or isAlias, leadTimeDays may be null. If neither and
      // the value is null, the SKU effectively falls back to the default.
      cleanSku[sku] = entry;
    }
    const blob = {
      lastImport: new Date().toISOString(),
      defaults,
      perSku: cleanSku,
    };
    writeData("vf_supply_settings", blob);
    res.json({ ok: true, lastImport: blob.lastImport, skuCount: Object.keys(cleanSku).length });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// Nightly on-hand sync at 06:30 UTC (just after the manual movement-upload window)
if (process.env.CIN7_ACCOUNT_ID && process.env.CIN7_APPLICATION_KEY) {
  cron.schedule("30 6 * * *", async () => {
    console.log("[Cin7 OnHand] Nightly sync starting…");
    try {
      const s = await performCin7OnHandSync();
      console.log(`[Cin7 OnHand] Sync done — ${s.rowCount} rows, ${s.skuCount} SKUs`);
    } catch (e) {
      console.error("[Cin7 OnHand] Nightly sync failed:", e.message);
    }
  });
  console.log("[Cin7 OnHand] Nightly sync scheduled at 06:30 UTC");
}

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

if (require.main === module) {
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`VF Production Scheduling running on port ${PORT}`);
  });
}

module.exports = app;
