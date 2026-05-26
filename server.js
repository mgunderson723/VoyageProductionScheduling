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

// Everything else under /api/* requires authentication, except a small set
// of webhook endpoints that have their own auth (shared secret in header).
const SESSION_BYPASS_PATHS = new Set([
  "/login",
  "/logout",
  "/cin7/inventory-movements", // Apps Script auto-sync, gated by X-VF-Sync-Secret
]);
app.use("/api", (req, res, next) => {
  if (SESSION_BYPASS_PATHS.has(req.path)) return next();
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
  "confirmed", "actualQty",
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
- Zone 2 Other: grinder (Ground BFC line — 1,000 kg per 2-hr shift)

Capacity & runtime constraints (apply these when recommending slots):
- east_mac / west_mac (Mcintyres): Runtime 12 days. Min 2,250 kg / Max 4,300 kg per batch. EU and US products cannot be mixed on the same run.
- refining: Runtime 1 day. Max capacity 6,000 kg/day (4 batches × 1,500 kg). Min 500 kg per batch.
- conching: Runtime 1 day. Min 3,000 kg / Max 6,000 kg per run.
- roaster: ~325 kg per batch, up to ~4 batches/shift (~1,300 kg/day).
- fat_melter: CBE-based fat 24 hr melt cycle; CBS-based fat 72 hr. Runs simultaneously with liquor melting.
- Liquor melting (pre-conching): 72 hr.

PRODUCTION RECIPE PIPELINE (read carefully — quantities scale DOWN through the chain, not 1:1):

Finished Chocolate (CFC) requires (in this backward order):
  4. CFC FG ← refining → conching → depositing (Zone 2 chocolate line; ~3 days end-to-end)
  3. Liquor (10–20% of the FG recipe by weight, recipe-dependent)
  2. Liquor itself ← Mcintyre run (12 day cycle, east_mac CBS / west_mac CBE)
  1. Liquor BOM = ~70% roasted seeds + ~30% melted fat
       ↳ Roasted seeds ← roaster
       ↳ Melted fat ← fat_melter (24 hr CBE / 72 hr CBS)

So 7,000 kg of finished CFC ≠ 7,000 kg of every input. Rough sanity check:
  - 7,000 kg CFC needs ~700–1,400 kg of liquor (10–20%)
  - That liquor needs ~490–980 kg roasted seeds (70% of liquor)
  - Plus ~210–420 kg melted fat (30% of liquor)
  - The remaining ~5,600–6,300 kg of CFC weight comes from sugar, packaging, flavorings, etc. as defined in the FG's BOM.

Liquor (sold-as-is) production: roasting → fat melting (24/72 hr) → Mcintyre (12 days) → packout. Liquor IS made on the Mcintyre, not on the chocolate line.

CRITICAL: DO NOT recommend production schedules using straight-line 1:1 scaling of the requested FG qty through every upstream stage. ALWAYS call bom_expand FIRST for any multi-stage production planning so you have BOM-driven quantities at each step. If the user names a product loosely (e.g. "CFC 506 EU"), use find_bom to resolve to the actual SKU before bom_expand.

Order statuses: queued, in-progress, complete, on-hold
Order priorities: high, med, low
Order confirmation: 'confirmed' boolean (default false on new). Tentative orders get visual cue on calendar but are still real schedule entries.

PRODUCT CATEGORIES (the front-end's edit modal validates these — picking the wrong combo will silently strip fields when a user opens the order):
- cat='coffee', sub ∈ {coffee_beans, coffee_ground} — coffee line, runs on the grinder
- cat='liquor',    sub='liquor'    — chocolate liquor + Mcintyre line products (incl. Final Blends like PFS Final Blend)
- cat='chocolate', sub='chocolate' — finished chocolate, fat melter, refining/conching/depositing, AND pouching (per the design system: "purple = finished chocolate, including final blends, bars, pouched product")
- region: 'eu' or 'us', only meaningful for liquor
- temper: 'cbe' or 'cbs' — east_mac is always cbs, west_mac is always cbe; chocolate-line machines run both, supply explicitly when known

When you call add_order: if the machine maps unambiguously, omitting cat/sub is fine — defaults are inferred. But if you know the product (e.g. PFS pouches → cat='chocolate', sub='chocolate'), pass cat/sub explicitly. Use update_order_metadata to fix cat/sub/region/temper/machine on existing orders.

When recommending a production slot for a finished good or WIP:
1. Call get_orders to see what's currently scheduled.
2. If the user's product reference is loose, call find_bom to get the right SKU.
3. Call bom_expand to get per-stage quantities and the upstream WIPs that must be produced. The intermediateStages array tells you each stage's machine + its production lead time.
4. For multi-stage products, schedule BACKWARD from the desired completion date:
   - Last step (e.g. depositing) ends at the target completion date
   - Each upstream stage must finish before its downstream stage starts (i.e. work the lead times backward)
   - Roasting + fat melting can run in parallel (both feed Mcintyre in the liquor case, or feed downstream stages directly for CFC)
5. Check that each stage's required qty fits within that machine's capacity constraints above.
6. Use find_available_slots (or get_orders + reason about gaps) to pick concrete dates that don't collide with existing scheduled work.
7. Present 2–3 concrete options spanning the full pipeline (each option = a complete set of stage dates) and explain the trade-offs (lead time risk, fat type, batch size fit, etc.).
8. Do NOT book anything — only recommend. The user must ask you to create or update orders to make it happen.

For metadata edits — renaming an MO (orderId), changing the SKU on a row, editing notes, flipping priority, marking confirmed/tentative, or recording actual produced qty — use update_order_metadata. It accepts any subset of those fields in one call. Do NOT just describe the change in text; if the user asked you to change a field, you MUST call this tool to persist it, otherwise the user will see no effect when they click the order on the calendar.

IMPORTANT: Before calling any write tools (add_order, shift_machine_orders, update_order_dates, update_order_status, update_order_quantity, update_order_metadata, delete_order), you MUST:
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
  "update_order_metadata",
]);

// Category model — must match SUBTYPES in public/index.html. The front-end's
// edit modal validates (cat, sub) on save; if either is empty or sub doesn't
// belong to cat, save throws and the row never persists. Bot tools enforce
// the same constraints so AI-created orders can be opened + edited cleanly.
const CATEGORY_SUBS = {
  coffee:    ["coffee_beans", "coffee_ground"],
  liquor:    ["liquor"],
  chocolate: ["chocolate"],
};
const ALL_SUBS = Object.values(CATEGORY_SUBS).flat();
const VALID_REGIONS = ["eu", "us"];
const VALID_TEMPERS = ["cbe", "cbs"];

// Default cat/sub for a given machine when the bot doesn't supply them.
// Pouching is intentionally chocolate (matches the design-system note that
// "purple = finished chocolate (final blends, bars, pouched product)") and
// puts pouched FGs on the chocolate-coloured chip rather than the generic
// teal "grind" fallback. grinder is the only coffee-line machine.
function inferCatSubFromMachine(machine) {
  switch (machine) {
    case "grinder":
      return { cat: "coffee", sub: "coffee_ground" };
    case "fat_melter":
    case "refining":
    case "conching":
    case "depositing":
    case "pouching":
      return { cat: "chocolate", sub: "chocolate" };
    case "roaster":
    case "seed_clean":
    case "east_mac":
    case "west_mac":
    case "mac_1250":
    case "mac_packout":
    default:
      return { cat: "liquor", sub: "liquor" };
  }
}

// Default temper inference from machine, used only when the bot omits it.
// Only the Mcintyres are unambiguously typed (east=CBS, west=CBE); everything
// else is left null because chocolate-line machines run both tempers.
function inferTemperFromMachine(machine) {
  if (machine === "east_mac") return "cbs";
  if (machine === "west_mac") return "cbe";
  return "";
}

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
    description: "Create a new work order on the production schedule. Only call this after the user has confirmed they want to proceed with a specific slot. PICK THE RIGHT cat AND sub for the product — these drive the calendar chip color and are validated by the edit modal. Defaults are inferred from machine if you omit them, but those defaults are only correct when the machine maps unambiguously (Mcintyres → liquor, chocolate-line → chocolate, grinder → coffee, pouching → chocolate). Override when you know better — e.g. PFS Final Blend on mac_1250 is technically a peanut-free spread and still uses cat='liquor' sub='liquor', but if you're booking coffee on the grinder, set cat='coffee'.",
    input_schema: {
      type: "object",
      properties: {
        sku:      { type: "string", description: "SKU / product description" },
        machine:  { type: "string", description: "Machine key (e.g. east_mac, conching, roaster, pouching)" },
        start:    { type: "string", description: "Start date YYYY-MM-DD" },
        end:      { type: "string", description: "End date YYYY-MM-DD" },
        qty:      { type: "number", description: "Batch quantity in kg" },
        batches:  { type: "number", description: "Number of batches (default 1). Total = qty × batches." },
        orderId:  { type: "string", description: "MO number if known (e.g. MO-00999), otherwise omit and one will be generated" },
        priority: { type: "string", enum: ["high", "med", "low"], description: "Priority (default: med)" },
        notes:    { type: "string", description: "Any notes or special instructions" },
        cat:      { type: "string", enum: ["coffee", "liquor", "chocolate"], description: "Product category. Drives chip color and which sub-types are valid. Defaults from machine if omitted." },
        sub:      { type: "string", enum: ["coffee_beans", "coffee_ground", "liquor", "chocolate"], description: "Product sub-type. Must belong to its category: coffee→{coffee_beans, coffee_ground}, liquor→{liquor}, chocolate→{chocolate}. Defaults from machine if omitted." },
        region:   { type: "string", enum: ["eu", "us"], description: "Region — only meaningful for liquor (EU vs US recipes). Omit for non-liquor." },
        temper:   { type: "string", enum: ["cbe", "cbs"], description: "Temper type — CBE or CBS. Auto-set for east_mac (cbs) / west_mac (cbe); supply explicitly for other machines if known." },
        confirmed:{ type: "boolean", description: "Confirmed-for-production flag. New orders default to false (tentative) so the user can review before committing." },
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
  {
    name: "update_order_metadata",
    description: "Update one or more metadata fields on a work order — the catch-all for fields not covered by the dedicated update tools (dates, status, qty). Use this for renames (orderId), SKU corrections, notes edits, priority changes, confirming/un-confirming, recording actual produced qty, fixing the machine assignment, and re-categorizing (cat / sub / region / temper). Only the fields you provide are updated; omit a field to leave it unchanged. Always confirm with the user before calling this on an existing order.",
    input_schema: {
      type: "object",
      properties: {
        order_id: { type: "string", description: "The CURRENT orderId of the order to update (used to find the row). To rename, also pass new_order_id." },
        new_order_id: { type: "string", description: "Replace the order's orderId / name with this value. Use for renaming MOs (e.g. 'MO-00783' → 'MO-00783-A')." },
        sku: { type: "string", description: "Replace the SKU. Pass the full SKU string." },
        notes: { type: "string", description: "Replace the notes/free-text field. Pass an empty string to clear." },
        priority: { type: "string", enum: ["low", "med", "high"], description: "Set priority." },
        confirmed: { type: "boolean", description: "Set the confirmed-for-production flag. true = confirmed, false = tentative." },
        actual_qty: { type: "number", description: "Set the actual produced qty (kg). Use when actual differs from planned (e.g. abandoned mid-run). Pass 0 to clear or null to remove the override." },
        machine: { type: "string", description: "Re-assign to a different machine line (e.g. 'pouching', 'east_mac'). Use the machine keys from the system prompt — passing an unknown key is an error." },
        cat: { type: "string", enum: ["coffee", "liquor", "chocolate"], description: "Re-categorize. If you change cat you usually need to update sub too (and possibly clear region/temper)." },
        sub: { type: "string", enum: ["coffee_beans", "coffee_ground", "liquor", "chocolate"], description: "Sub-type. Must belong to its category." },
        region: { type: "string", enum: ["eu", "us"], description: "Region (liquor only). Pass empty string to clear." },
        temper: { type: "string", enum: ["cbe", "cbs"], description: "Temper (cbe / cbs). Pass empty string to clear." },
        batches: { type: "number", description: "Number of batches. Total auto-recalcs as qty × batches." },
      },
      required: ["order_id"],
    },
  },
  {
    name: "find_bom",
    description: "Search the BOM library for parent SKUs matching a query. Use this when the user names a product (e.g. 'CFC 506 EU' or 'PFS pouch') and you need to resolve it to a real SKU before calling bom_expand. Returns up to 20 matches. If the user gives an exact SKU, you can skip this and call bom_expand directly.",
    input_schema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Substring to match against parent SKU or product name (case-insensitive)" },
      },
      required: ["query"],
    },
  },
  {
    name: "bom_expand",
    description: "CRITICAL for any multi-stage production planning. Recursively expand a parent SKU's BOM down to leaf-level raw-material requirements for a given quantity. Returns a list of every leaf RM and the kg required, plus the WIP intermediates traversed (with their own qtys, so you can plan upstream production stages). Honors the BOM's wastage% per edge. Use this BEFORE recommending production schedules for finished goods or WIPs — it tells you how much you actually need at each stage of the pipeline (e.g. 7,000 kg of finished chocolate does NOT require 7,000 kg of roasted seeds; the BOM ratios are typically far smaller). The 'machine' field on each intermediate identifies which production line that stage runs on, so you can sequence the schedule.",
    input_schema: {
      type: "object",
      properties: {
        sku: { type: "string", description: "Parent SKU to expand (FG-… or WIP-…)" },
        qty: { type: "number", description: "Production quantity in kg (or units, for packaged FGs)" },
      },
      required: ["sku", "qty"],
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
      const { sku, machine, start, end, qty = 0, batches = 1, orderId, priority = "med", notes = "" } = input;
      // Resolve cat/sub: respect the bot's explicit values; otherwise infer
      // from machine. Validate that sub belongs to cat — if either is bad,
      // fail loudly rather than persisting an order the modal can't save.
      let cat = input.cat;
      let sub = input.sub;
      if (!cat || !sub) {
        const inferred = inferCatSubFromMachine(machine);
        cat = cat || inferred.cat;
        sub = sub || inferred.sub;
      }
      if (!CATEGORY_SUBS[cat]) {
        return { ok: false, error: `Invalid cat '${cat}'. Valid: ${Object.keys(CATEGORY_SUBS).join(", ")}` };
      }
      if (!CATEGORY_SUBS[cat].includes(sub)) {
        return { ok: false, error: `sub '${sub}' is not valid for cat '${cat}'. Valid subs for ${cat}: ${CATEGORY_SUBS[cat].join(", ")}` };
      }
      // Region only applies to liquor; silently drop for other cats so we
      // don't pollute downstream filters.
      let region = input.region != null ? String(input.region) : "";
      if (region && !VALID_REGIONS.includes(region)) {
        return { ok: false, error: `Invalid region '${region}'. Valid: ${VALID_REGIONS.join(", ")}` };
      }
      if (cat !== "liquor") region = "";
      // Temper: prefer explicit input, otherwise infer from machine.
      let temper = input.temper != null ? String(input.temper) : inferTemperFromMachine(machine);
      if (temper && !VALID_TEMPERS.includes(temper)) {
        return { ok: false, error: `Invalid temper '${temper}'. Valid: ${VALID_TEMPERS.join(", ")}` };
      }
      const confirmed = input.confirmed === undefined ? false : !!input.confirmed;
      const id = `ai-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
      const generatedOrderId = orderId || `TBD-${machine}-${(start || "").replace(/-/g, "")}`;
      const newOrder = {
        id, orderId: generatedOrderId, sku, machine,
        start, end, due: end,
        qty, batches, total: qty * batches,
        status: "queued", priority, notes,
        cat, sub, region, temper,
        confirmed,
      };
      orders.push(newOrder);
      writeData("vf_orders", orders);
      return { ok: true, message: `Created order '${generatedOrderId}' for ${sku} on ${machine} (${start} → ${end}, ${cat}/${sub}${temper ? `/${temper}` : ""}${region ? `/${region}` : ""})`, order: newOrder };
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

    case "update_order_metadata": {
      const { order_id } = input;
      const idx = orders.findIndex(o => o.orderId === order_id || o.id === order_id);
      if (idx === -1) return { ok: false, error: `Order '${order_id}' not found` };
      const order = orders[idx];
      const before = { ...order };
      const changes = [];
      if (input.new_order_id !== undefined && input.new_order_id !== order.orderId) {
        const newId = String(input.new_order_id).trim();
        if (!newId) return { ok: false, error: "new_order_id cannot be empty" };
        // Detect a duplicate orderId (otherwise renaming creates an ambiguous lookup)
        const conflict = orders.find((o, i) => i !== idx && o.orderId === newId);
        if (conflict) return { ok: false, error: `Cannot rename — '${newId}' already exists on another order (id ${conflict.id})` };
        order.orderId = newId;
        changes.push(`orderId: '${before.orderId}' → '${newId}'`);
      }
      if (input.sku !== undefined && input.sku !== order.sku) {
        order.sku = String(input.sku);
        changes.push(`sku: '${before.sku || ""}' → '${order.sku}'`);
      }
      if (input.notes !== undefined && input.notes !== order.notes) {
        order.notes = String(input.notes);
        changes.push("notes updated");
      }
      if (input.priority !== undefined && input.priority !== order.priority) {
        const allowed = ["low", "med", "high"];
        if (!allowed.includes(input.priority)) return { ok: false, error: `priority must be one of ${allowed.join(", ")}` };
        order.priority = input.priority;
        changes.push(`priority: '${before.priority || ""}' → '${input.priority}'`);
      }
      if (input.confirmed !== undefined && !!input.confirmed !== !!order.confirmed) {
        order.confirmed = !!input.confirmed;
        changes.push(`confirmed: ${!!before.confirmed} → ${!!input.confirmed}`);
      }
      if (input.actual_qty !== undefined) {
        const newActual = (input.actual_qty === null || input.actual_qty === "") ? null : Number(input.actual_qty);
        if (newActual !== null && !isFinite(newActual)) return { ok: false, error: "actual_qty must be a number or null" };
        if (newActual !== (order.actualQty == null ? null : Number(order.actualQty))) {
          order.actualQty = newActual;
          changes.push(`actualQty: ${before.actualQty == null ? "null" : before.actualQty} → ${newActual == null ? "null" : newActual}`);
        }
      }
      // Machine reassignment — list of valid keys mirrors the modal dropdown
      // in public/index.html. An invalid key would silently disappear in the
      // UI on next save (the original symptom we're guarding against), so we
      // reject it here.
      const VALID_MACHINES = [
        "roaster", "seed_clean", "east_mac", "west_mac", "mac_1250",
        "mac_packout", "pouching", "grinder", "fat_melter", "refining",
        "conching", "depositing",
      ];
      if (input.machine !== undefined && input.machine !== order.machine) {
        const m = String(input.machine);
        if (m && !VALID_MACHINES.includes(m)) {
          return { ok: false, error: `Invalid machine '${m}'. Valid: ${VALID_MACHINES.join(", ")}` };
        }
        order.machine = m;
        changes.push(`machine: '${before.machine || ""}' → '${m}'`);
      }
      // Re-categorize: cat / sub change in lockstep. If only one is supplied,
      // validate the resulting (cat, sub) combo against CATEGORY_SUBS.
      if (input.cat !== undefined || input.sub !== undefined) {
        const newCat = input.cat !== undefined ? String(input.cat) : order.cat;
        const newSub = input.sub !== undefined ? String(input.sub) : order.sub;
        if (!CATEGORY_SUBS[newCat]) {
          return { ok: false, error: `Invalid cat '${newCat}'. Valid: ${Object.keys(CATEGORY_SUBS).join(", ")}` };
        }
        if (!CATEGORY_SUBS[newCat].includes(newSub)) {
          return { ok: false, error: `sub '${newSub}' is not valid for cat '${newCat}'. Valid subs: ${CATEGORY_SUBS[newCat].join(", ")}` };
        }
        if (newCat !== order.cat) { order.cat = newCat; changes.push(`cat: '${before.cat || ""}' → '${newCat}'`); }
        if (newSub !== order.sub) { order.sub = newSub; changes.push(`sub: '${before.sub || ""}' → '${newSub}'`); }
      }
      if (input.region !== undefined && String(input.region) !== (order.region || "")) {
        const r = String(input.region);
        if (r && !VALID_REGIONS.includes(r)) return { ok: false, error: `Invalid region '${r}'. Valid: ${VALID_REGIONS.join(", ")}` };
        order.region = r;
        changes.push(`region: '${before.region || ""}' → '${r}'`);
      }
      if (input.temper !== undefined && String(input.temper) !== (order.temper || "")) {
        const t = String(input.temper);
        if (t && !VALID_TEMPERS.includes(t)) return { ok: false, error: `Invalid temper '${t}'. Valid: ${VALID_TEMPERS.join(", ")}` };
        order.temper = t;
        changes.push(`temper: '${before.temper || ""}' → '${t}'`);
      }
      if (input.batches !== undefined) {
        const b = Number(input.batches);
        if (!isFinite(b) || b < 1) return { ok: false, error: "batches must be a positive number" };
        if (b !== (order.batches || 1)) {
          order.batches = b;
          order.total = (order.qty || 0) * b;
          changes.push(`batches: ${before.batches || 1} → ${b} (total now ${order.total})`);
        }
      }
      if (changes.length === 0) {
        return { ok: true, message: `No changes made to '${order_id}' — fields already match.` };
      }
      writeData("vf_orders", orders);
      return { ok: true, message: `Updated '${order_id}': ${changes.join("; ")}.` };
    }

    case "find_bom": {
      const q = String(input.query || "").trim().toLowerCase();
      if (!q) return { ok: false, error: "query is required" };
      const blob = readData("vf_boms");
      if (!blob || !blob.parents) return { ok: false, error: "No BOMs imported yet" };
      const matches = [];
      for (const sku of Object.keys(blob.parents)) {
        const versions = blob.parents[sku];
        const def = versions[0] || {};
        const hay = (sku + " " + (def.parentName || "")).toLowerCase();
        if (hay.includes(q)) {
          matches.push({
            sku,
            name: def.parentName || "",
            machine: def.machine || null,
            productionLeadTimeDays: def.productionLeadTime,
            qtyToProduce: def.qtyToProduce,
            componentCount: (def.components || []).length,
            versionCount: versions.length,
          });
          if (matches.length >= 20) break;
        }
      }
      return { ok: true, query: q, matchCount: matches.length, matches };
    }

    case "bom_expand": {
      const { sku, qty } = input;
      if (!sku || !isFinite(qty)) return { ok: false, error: "Need sku and numeric qty" };
      const blob = readData("vf_boms");
      if (!blob || !blob.parents) return { ok: false, error: "No BOMs imported yet" };
      if (!blob.parents[sku]) return { ok: false, error: `No BOM defined for '${sku}'. Try find_bom to discover the right SKU.` };
      let result;
      try { result = expandBom(blob.parents, sku, qty, { applyWastage: true }); }
      catch (e) { return { ok: false, error: e.message }; }
      // Enrich the intermediates with each WIP's machine + production lead
      // time so the AI can sequence the upstream stages correctly.
      const enriched = result.intermediates.map(step => {
        const versions = blob.parents[step.sku] || [];
        const bom = versions[0] || {};
        return {
          sku: step.sku,
          qtyKg: Math.round(step.qty * 100) / 100,
          version: step.version,
          depth: step.depth,
          machine: bom.machine || null,
          productionLeadTimeDays: bom.productionLeadTime,
          parentName: bom.parentName || "",
        };
      });
      const leaves = Object.values(result.leaves)
        .map(l => ({ sku: l.sku, qtyKg: Math.round(l.qty * 100) / 100, isCycle: !!l.isCycle }))
        .sort((a, b) => b.qtyKg - a.qtyKg);
      return {
        ok: true,
        parent: sku,
        qty,
        leafRequirements: leaves,
        intermediateStages: enriched,
        note: "leafRequirements = leaf-level raw materials to procure. intermediateStages = each WIP that must be produced (with its machine and lead time) — use these for backward production scheduling.",
      };
    }

    default:
      return { ok: false, error: `Unknown tool: ${name}` };
  }
}

app.post("/api/chat", requireOrderEdit, async (req, res) => {
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

// Order-edit middleware — admin or operator only. Used to gate /api/chat
// (the AI bot has mutating tools like add_order/delete_order/etc.) so the
// new 'planner' and existing 'viewer' roles can't bypass canEditOrders()
// by chatting at the bot. Front-end also hides the chat widget for these
// roles; this is the defence-in-depth layer.
function requireOrderEdit(req, res, next) {
  const users = readData("vf_users") || [];
  const u = users.find(x => x && x.id === req.userId);
  if (!u || (u.role !== "admin" && u.role !== "operator")) {
    return res.status(403).json({ ok: false, error: "Order-edit role required (admin or operator). The chat assistant is disabled for read-only roles." });
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

// ── Cin7 product cost cache ─────────────────────────────────────────────────
// Pulls AverageCost per SKU from /product (paginated). Used by the MRP $$
// summary view to translate suggested PO qtys into capital commitments.
async function fetchCin7ProductCosts() {
  if (!process.env.CIN7_ACCOUNT_ID || !process.env.CIN7_APPLICATION_KEY) {
    throw new Error("CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY environment variable not set");
  }
  const all = [];
  let page = 1;
  const limit = 1000;
  while (true) {
    const url = `https://inventory.dearsystems.com/ExternalApi/v2/product?Page=${page}&Limit=${limit}`;
    const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
    const ct = resp.headers.get("content-type") || "";
    const text = await resp.text().catch(() => "");
    if (!resp.ok) throw new Error(`Cin7 product list ${resp.status} [${ct}]: ${text.slice(0, 300)}`);
    let data;
    try { data = JSON.parse(text); }
    catch (_) { throw new Error(`Cin7 product list ${resp.status} non-JSON [${ct}]: ${text.slice(0, 300)}`); }
    const batch = data.Products || [];
    all.push(...batch);
    if (batch.length < limit) break;
    page++;
    if (page > 20) throw new Error("Cin7 product pagination exceeded 20 pages — aborting");
  }
  return all;
}

async function performCin7ProductCostsSync() {
  const products = await fetchCin7ProductCosts();
  const now = new Date().toISOString();
  const bySku = {};
  let withCost = 0;
  for (const p of products) {
    if (!p.SKU) continue;
    const cost = Number(p.AverageCost) || 0;
    bySku[p.SKU] = {
      sku: p.SKU,
      name: p.Name || "",
      averageCost: cost,
      costingMethod: p.CostingMethod || "",
      category: p.Category || "",
    };
    if (cost > 0) withCost++;
  }
  const blob = {
    lastSync: now,
    productCount: products.length,
    withCostCount: withCost,
    currency: "USD",
    bySku,
  };
  writeData("product_costs", blob);
  return { ok: true, lastSync: now, productCount: products.length, withCostCount: withCost };
}

// POST /api/cin7/product-costs/sync — admin-triggered live pull
app.post("/api/cin7/product-costs/sync", requireAdmin, async (req, res) => {
  try {
    const status = await performCin7ProductCostsSync();
    res.json(status);
  } catch (e) {
    console.error("[Cin7 Costs] Sync error:", e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// GET /api/cin7/product-costs — return the cached cost map (any authed user)
app.get("/api/cin7/product-costs", (req, res) => {
  const blob = readData("product_costs");
  if (!blob) return res.json({ ok: true, lastSync: null, bySku: {} });
  res.json({ ok: true, ...blob });
});

// ── Cin7 production-run sync (powers the Error Reporting tab) ───────────────
//
// Why: the daily Inventory Movement Details report doesn't emit a row when an
// operator records 0 consumption on a BOM line — no movement, no record. So
// the original movement-diff detector couldn't catch the case it was meant
// to. Production Runs, on the other hand, ALWAYS store both ExpectedQuantity
// (per BOM) and Quantity (operator-entered actual) on every component line.
// Comparing those two directly is the right primitive.
//
// API used (Cin7 Core v2):
//   GET /production/orderList?CompletionDateFrom=<today-7d>
//     → mixed list of Production Orders + Runs (Type field is "O"|"R")
//   GET /production/order/run?ProductionOrderID=<guid>
//     → { Runs: [ { Status, Number, Operations: [ { Components: [...] } ] } ] }
//
// Component field mapping (counterintuitive — be careful):
//   Quantity         = ACTUAL consumed (operator entry)
//   ExpectedQuantity = REQUIRED per BOM
//   WastageQty       = wastage (treated as legit consumption, not flagged)
//
// Flag rule: ExpectedQuantity > 0 AND Quantity === 0. Strict zero on actuals,
// only when something was expected. Catches both operator-entered zeros and
// the case where the line never got an actual recorded at all.
//
// Rate limit: Cin7 caps at 60 calls/min. We sleep 1100ms between detail
// calls — a 7-day window with ~50 unique parent orders takes ~1 minute.

const C7_PROD_BASE = "https://inventory.dearsystems.com/ExternalApi/v2";
const PRODUCTION_RUN_WINDOW_DAYS = 7;
const PRODUCTION_RUN_RATE_LIMIT_MS = 1100;

function sleepMs(ms) { return new Promise(r => setTimeout(r, ms)); }

async function fetchProductionOrderList(completionDateFrom) {
  if (!process.env.CIN7_ACCOUNT_ID || !process.env.CIN7_APPLICATION_KEY) {
    throw new Error("CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY environment variable not set");
  }
  const all = [];
  let page = 1;
  const limit = 200;
  while (true) {
    const url = `${C7_PROD_BASE}/production/orderList?Page=${page}&Limit=${limit}&CompletionDateFrom=${completionDateFrom}`;
    const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
    const ct = resp.headers.get("content-type") || "";
    const text = await resp.text().catch(() => "");
    if (!resp.ok) {
      throw new Error(`Cin7 productionOrderList ${resp.status} ${resp.statusText} [${ct}]: ${text.slice(0, 300)}`);
    }
    let data;
    try { data = JSON.parse(text); }
    catch (_) { throw new Error(`Cin7 productionOrderList ${resp.status} non-JSON [${ct}]: ${text.slice(0, 300)}`); }
    const batch = data.ProductionOrderListItems || [];
    all.push(...batch);
    if (batch.length < limit) break;
    page++;
    if (page > 50) throw new Error("Cin7 productionOrderList pagination exceeded 50 pages — aborting");
  }
  return all;
}

async function fetchProductionRunDetail(productionOrderID) {
  const url = `${C7_PROD_BASE}/production/order/run?ProductionOrderID=${encodeURIComponent(productionOrderID)}`;
  const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
  const ct = resp.headers.get("content-type") || "";
  const text = await resp.text().catch(() => "");
  if (!resp.ok) {
    throw new Error(`Cin7 productionRun ${resp.status} ${resp.statusText} [${ct}]: ${text.slice(0, 300)}`);
  }
  try { return JSON.parse(text); }
  catch (_) { throw new Error(`Cin7 productionRun ${resp.status} non-JSON [${ct}]: ${text.slice(0, 300)}`); }
}

async function performProductionRunSync() {
  const now = new Date();
  const end = now.toISOString().slice(0, 10);
  const startD = new Date(now);
  startD.setUTCDate(startD.getUTCDate() - (PRODUCTION_RUN_WINDOW_DAYS - 1));
  const start = startD.toISOString().slice(0, 10);

  // 1. List recent activity. The endpoint returns both Production Orders and
  //    Production Runs (Type=O|R). We only care about COMPLETED runs.
  //
  // Status-field gotcha: the spec's example shows both `Status` (DRAFT,
  // AUTHORIZED, RELEASED, VOIDED) and `OrderStatus` (DRAFT, PLANNED,
  // RELEASED, IN PROGRESS, COMPLETED, VOIDED) on each list row. For Runs,
  // the field semantics aren't explicitly documented — checking BOTH for
  // "COMPLETED" is the safe move.
  const list = await fetchProductionOrderList(start);

  // Diagnostic counts — show up in Railway logs so we can see what the
  // tenant's data actually looks like if filtering misses things.
  const typeCounts = {};
  const statusCounts = {};
  const orderStatusCounts = {};
  for (const it of list) {
    if (!it) continue;
    const t = it.Type || "?";
    typeCounts[t] = (typeCounts[t] || 0) + 1;
    const s = (it.Status || "(blank)");
    statusCounts[s] = (statusCounts[s] || 0) + 1;
    const os = (it.OrderStatus || "(blank)");
    orderStatusCounts[os] = (orderStatusCounts[os] || 0) + 1;
  }
  console.log(`[ProdRunSync] List response: ${list.length} rows · Types ${JSON.stringify(typeCounts)} · Status ${JSON.stringify(statusCounts)} · OrderStatus ${JSON.stringify(orderStatusCounts)}`);

  const isCompletedRow = it => {
    const s = String(it.Status || "").toUpperCase();
    const os = String(it.OrderStatus || "").toUpperCase();
    return s === "COMPLETED" || os === "COMPLETED";
  };
  const completedRunRows = list.filter(it => it && it.Type === "R" && isCompletedRow(it));

  // 2. Group by parent ProductionOrderID — one detail call per unique order.
  //    Track which RunIDs are the completed ones we care about (an order may
  //    have other runs we should skip).
  const completedRunsByOrder = new Map();
  for (const row of completedRunRows) {
    if (!row.ProductionOrderID || !row.TaskID) continue;
    if (!completedRunsByOrder.has(row.ProductionOrderID)) {
      completedRunsByOrder.set(row.ProductionOrderID, {
        orderNumber: row.OrderNumber || "",
        runs: new Map(),
      });
    }
    completedRunsByOrder.get(row.ProductionOrderID).runs.set(row.TaskID, {
      runID: row.TaskID,
      completionDate: row.CompletionDate ? String(row.CompletionDate).slice(0, 10) : null,
      productSKU: row.ProductSKU || "",
      productName: row.ProductName || "",
      locationName: row.LocationName || "",
    });
  }

  // Product-category lookup so the Production Output tab can group by
  // Voyage's product families (Chocolate / Coffee / Spreads / ...). The
  // product_costs cache mirrors Cin7's per-product Category string. Runs
  // whose FG SKU isn't in the cache fall through to "Other".
  const costsBlob = readData("product_costs") || { bySku: {} };
  const categoryOf = sku => {
    const c = costsBlob.bySku && costsBlob.bySku[sku];
    return (c && c.category) ? c.category : "Other";
  };

  // 3. For each unique parent order, fetch run detail. Walk operations →
  //    components (for the Error Reporting flagged-line detector) AND the
  //    Run.Output[] array (for the Production Output tab's actuals view).
  const flagged = [];
  const allCompletedRuns = [];
  let detailCallsMade = 0;
  let detailFailures = 0;
  const orderIDs = Array.from(completedRunsByOrder.keys());

  for (let i = 0; i < orderIDs.length; i++) {
    const orderID = orderIDs[i];
    const orderInfo = completedRunsByOrder.get(orderID);
    let detail;
    try {
      detail = await fetchProductionRunDetail(orderID);
      detailCallsMade++;
    } catch (e) {
      detailFailures++;
      console.error(`[ProdRunSync] Detail fetch failed for ${orderID} (${orderInfo.orderNumber}):`, e.message);
      // Throttle even on failure so we don't hammer the rate limit
      if (i < orderIDs.length - 1) await sleepMs(PRODUCTION_RUN_RATE_LIMIT_MS);
      continue;
    }

    const orderNumber = detail.OrderNumber || orderInfo.orderNumber || "";
    const runs = Array.isArray(detail.Runs) ? detail.Runs : [];
    for (const run of runs) {
      const runMeta = orderInfo.runs.get(run.RunID);
      if (!runMeta) continue; // not one of the completed runs from the list
      const runNumber = run.Number != null ? String(run.Number) : "";
      const moRef = orderNumber + (runNumber ? `/${runNumber}` : "");

      // Walk components across all operations on this run (Error Reporting path)
      const allComponentsOnRun = [];
      const flaggedOnRun = [];
      const workCentersSeen = [];
      for (const op of (run.Operations || [])) {
        const workCenter = op.WorkCenterName || op.Name || "";
        if (workCenter && !workCentersSeen.includes(workCenter)) workCentersSeen.push(workCenter);
        for (const c of (op.Components || [])) {
          const expected = Number(c.ExpectedQuantity) || 0;
          const actual = Number(c.Quantity) || 0;
          const lineRow = {
            sku: c.ProductCode || "",
            product: c.ProductName || "",
            batch: c.BatchSN || "",
            location: c.LocationName || "",
            unit: c.Unit || "",
            expected,
            actual,
            wastage: Number(c.WastageQty) || 0,
            workCenter,
          };
          allComponentsOnRun.push(lineRow);
          if (expected > 0 && actual === 0) flaggedOnRun.push(lineRow);
        }
      }

      if (flaggedOnRun.length > 0) {
        // Sibling lines = all components on the run that aren't themselves flagged
        const flaggedSet = new Set(flaggedOnRun);
        const siblings = allComponentsOnRun.filter(l => !flaggedSet.has(l));
        flagged.push({
          moRef,
          orderNumber,
          runNumber,
          completionDate: runMeta.completionDate,
          fgSKU: runMeta.productSKU,
          fgProduct: runMeta.productName,
          location: runMeta.locationName,
          flagged: flaggedOnRun,
          siblings,
        });
      }

      // Production Output path — collect what was actually finished on this run.
      // Run.Output[] is the canonical "finished products that landed in stock"
      // list (vs Operations[].OutputProducts which can double-count intermediate
      // products). Multiple outputs per run are possible for multi-output BOMs;
      // we keep them as separate rows. Work-center attribution is the LAST
      // operation in the sequence, since that's where the FG is realized.
      const outputs = Array.isArray(run.Output) ? run.Output : [];
      const lastWorkCenter = workCentersSeen.length ? workCentersSeen[workCentersSeen.length - 1] : "";
      const outputRows = outputs.map(o => {
        const sku = o.ProductCode || "";
        return {
          sku,
          product: o.ProductName || "",
          category: categoryOf(sku),
          qty: Number(o.Quantity) || 0,
          wastage: Number(o.WastageQuantity) || 0,
          unit: o.Unit || "",
          batch: o.BatchSN || "",
          location: o.LocationName || runMeta.locationName || "",
        };
      });
      const runOutputQty = outputRows.reduce((s, r) => s + r.qty, 0);

      // Include the run even if Output[] is empty so a row exists in the feed;
      // the FG-from-list metadata gives us at least the planned SKU/name.
      allCompletedRuns.push({
        moRef,
        orderNumber,
        runNumber,
        completionDate: runMeta.completionDate,
        fgSKU: runMeta.productSKU,
        fgProduct: runMeta.productName,
        fgCategory: categoryOf(runMeta.productSKU),
        location: runMeta.locationName,
        workCenter: lastWorkCenter,
        workCenters: workCentersSeen,
        outputs: outputRows,
        outputQty: runOutputQty,
      });
    }

    if (i < orderIDs.length - 1) await sleepMs(PRODUCTION_RUN_RATE_LIMIT_MS);
  }

  // Sort both feeds newest first
  flagged.sort((a, b) => (b.completionDate || "").localeCompare(a.completionDate || ""));
  allCompletedRuns.sort((a, b) => (b.completionDate || "").localeCompare(a.completionDate || ""));

  const blob = {
    lastSync: new Date().toISOString(),
    windowStart: start,
    windowEnd: end,
    listRowCount: list.length,
    listTypeCounts: typeCounts,
    listStatusCounts: statusCounts,
    listOrderStatusCounts: orderStatusCounts,
    completedRunsScanned: completedRunRows.length,
    parentOrdersScanned: orderIDs.length,
    detailCallsMade,
    detailFailures,
    flaggedRunCount: flagged.length,
    flaggedLineCount: flagged.reduce((s, g) => s + g.flagged.length, 0),
    flagged,
    allCompletedRuns,
  };
  writeData("production_run_errors_t7d", blob);
  return blob;
}

// Back-compat alias — old name still used by the cron + manual-sync endpoint
// callers. The function body now also populates the Production Output feed.
const performProductionRunErrorSync = performProductionRunSync;

// POST /api/cin7/production-runs/sync — manual trigger (any authed user).
// Returns the same summary the cron logs. Used by the "Sync now" button on
// the Error Reporting tab so the user can battle-test without waiting for
// the 07:00 UTC cron.
app.post("/api/cin7/production-runs/sync", async (req, res) => {
  try {
    const blob = await performProductionRunErrorSync();
    res.json({
      ok: true,
      lastSync: blob.lastSync,
      windowStart: blob.windowStart,
      windowEnd: blob.windowEnd,
      listRowCount: blob.listRowCount,
      listTypeCounts: blob.listTypeCounts,
      listStatusCounts: blob.listStatusCounts,
      listOrderStatusCounts: blob.listOrderStatusCounts,
      completedRunsScanned: blob.completedRunsScanned,
      parentOrdersScanned: blob.parentOrdersScanned,
      detailCallsMade: blob.detailCallsMade,
      detailFailures: blob.detailFailures,
      flaggedRunCount: blob.flaggedRunCount,
      flaggedLineCount: blob.flaggedLineCount,
    });
  } catch (e) {
    console.error("[ProdRunSync] Manual sync error:", e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── Cin7 Inventory Movement Details ingest (Apps Script → app webhook) ───────
//
// The user's Cin7 daily report drops an XLSX file into a Google Drive folder.
// A small Apps Script bound to that folder reads the file each morning and
// POSTs it to this endpoint with a shared secret. We parse it server-side
// (same shape as the Traceability tab's manual CSV upload) and replace the
// inventory data — driving auto-promote and the calendar ✓ marks without any
// human action.
//
// Auth: header X-VF-Sync-Secret must match env INVENTORY_SYNC_SECRET.
// Body: { filename, contentType, contentBase64, fileLastModified? }

function parseInventoryXlsx(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  return parseInventoryRows(rows);
}

// Same aggregation logic as public/traceability_explorer.jsx#parseInventoryCSV
// but takes a 2-D array of rows (from sheet_to_json) instead of CSV text.
// Output shape MUST match so downstream consumers (Live Inventory KPIs, MO
// Status auto-promote, calendar ✓ marks) work identically.
function parseInventoryRows(rows) {
  if (rows.length < 2) throw new Error("File has no data rows");

  // Find the header row — first row containing both 'sku' and a 'quantity in' column
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const fields = rows[i].map(f => String(f || "").toLowerCase().trim());
    if (fields.includes("sku") && (fields.includes("quantity in") || fields.includes("inbound"))) {
      hdrIdx = i; break;
    }
  }
  if (hdrIdx === -1) throw new Error("Could not find header row with SKU + Quantity columns");

  const hdr = rows[hdrIdx].map(h => String(h || "").toLowerCase().trim());
  const colMap = {};
  hdr.forEach((h, i) => {
    if (h === "sku") colMap.sku = i;
    else if (h === "product") colMap.product = i;
    else if (h === "category") colMap.category = i;
    else if (h === "unit") colMap.unit = i;
    else if (h === "date") colMap.date = i;
    else if (h === "month") colMap.month = i;
    else if (h === "reference type" || h === "ref_type") colMap.refType = i;
    else if (h === "quantity in" || h === "inbound") colMap.qtyIn = i;
    else if (h === "quantity out" || h === "outbound") colMap.qtyOut = i;
    else if (h === "batch #" || h === "batch_count") colMap.batch = i;
    else if (h === "reference") colMap.reference = i;
  });
  if (colMap.sku === undefined) throw new Error("Missing required column: SKU");
  if (colMap.qtyIn === undefined) throw new Error("Missing required column: Quantity in");
  if (colMap.date === undefined && colMap.month === undefined) throw new Error("Missing required column: Date or Month");

  const skuMap = new Map();
  const batchSets = new Map();
  const catAgg = {};
  const rtAgg = {};
  const moMap = new Map();
  const monthSet = new Set();
  const parseNum = s => { if (s == null || s === "") return 0; return parseFloat(String(s).replace(/,/g, "")) || 0; };
  const dateToMonth = d => {
    if (!d) return null;
    const s = (d instanceof Date) ? d.toISOString().slice(0, 10) : String(d);
    const mMap = { jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12" };
    let m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (m) { const mo = mMap[m[2].toLowerCase()]; return mo ? m[3]+"-"+mo : null; }
    m = s.match(/^(\d{4})-(\d{2})/);
    if (m) return m[1]+"-"+m[2];
    return null;
  };

  let rowCount = 0;
  for (let i = hdrIdx + 1; i < rows.length; i++) {
    const f = rows[i];
    if (!f || f.length < 3) continue;
    const sku = String(f[colMap.sku] || "").trim();
    if (!sku) continue;
    const month = colMap.month !== undefined ? String(f[colMap.month] || "") : dateToMonth(f[colMap.date]);
    if (!month) continue;
    const prod = colMap.product !== undefined ? String(f[colMap.product] || "") : "";
    const cat = colMap.category !== undefined ? String(f[colMap.category] || "") : "";
    const unit = colMap.unit !== undefined ? String(f[colMap.unit] || "") : "";
    const rt = colMap.refType !== undefined ? String(f[colMap.refType] || "") : "";
    const inb = parseNum(f[colMap.qtyIn]);
    const outb = colMap.qtyOut !== undefined ? parseNum(f[colMap.qtyOut]) : 0;
    const batch = colMap.batch !== undefined ? String(f[colMap.batch] || "") : "";
    monthSet.add(month);
    rowCount++;

    if (!skuMap.has(sku)) skuMap.set(sku, { s: sku, p: prod, c: cat, u: unit, m: {}, rt: {}, ti: 0, to: 0, net: 0, bc: 0 });
    if (!batchSets.has(sku)) batchSets.set(sku, new Set());
    const entry = skuMap.get(sku);
    if (!entry.m[month]) entry.m[month] = { i: 0, o: 0 };
    entry.m[month].i += inb; entry.m[month].o += outb;
    entry.ti += inb; entry.to += outb; entry.net += (inb - outb);
    if (batch) batchSets.get(sku).add(batch);
    if (rt) {
      if (!entry.rt[rt]) entry.rt[rt] = { i: 0, o: 0 };
      entry.rt[rt].i += inb; entry.rt[rt].o += outb;
    }
    if (!catAgg[cat]) catAgg[cat] = {};
    if (!catAgg[cat][month]) catAgg[cat][month] = { i: 0, o: 0 };
    catAgg[cat][month].i += inb; catAgg[cat][month].o += outb;
    if (!rtAgg[rt]) rtAgg[rt] = {};
    if (!rtAgg[rt][month]) rtAgg[rt][month] = { i: 0, o: 0 };
    rtAgg[rt][month].i += inb; rtAgg[rt][month].o += outb;

    const ref = colMap.reference !== undefined ? String(f[colMap.reference] || "") : "";
    const moMatch = ref.match(/MO-\d+/);
    if (moMatch) {
      const moId = moMatch[0];
      if (!moMap.has(moId)) moMap.set(moId, { mo: moId, sku, prod, totalIn: 0, totalOut: 0 });
      const me = moMap.get(moId);
      me.totalIn += inb;
      me.totalOut += outb;
    }
  }
  for (const [sku, entry] of skuMap) entry.bc = batchSets.get(sku).size;
  const months = [...monthSet].sort();
  const invSku = [...skuMap.values()];
  if (!invSku.length) throw new Error("No valid data rows found");
  return {
    invSku,
    invCat: catAgg,
    invRt: rtAgg,
    months,
    rowCount,
    moMovements: Object.fromEntries(moMap),
  };
}

// Max-merge for the daily auto-sync against a rolling 15-day Cin7 window.
//
// Why MAX and not REPLACE: each daily file is a partial slice (only
// transactions whose stock-movement-date falls in the last 15 days). For an
// MO that ran 12 days starting outside that window, the file would report a
// partial total. Per-month-REPLACE would clobber the full bulk-backfilled
// month with the partial 15-day slice — eroding history one day at a time.
//
// Strategy: for each (SKU, month) bucket, take MAX of i and o independently
// across existing vs delta. Same for invCat[cat][month], invRt[rt][month],
// and moMovements per MO. The "fullest" snapshot ever seen wins.
//
// Assumes monotonic-growth semantics: production-receipt qtys only go up over
// time. Voiding a posted transaction would reduce the true total — MAX would
// incorrectly preserve the pre-void value. Rare enough that re-running a bulk
// backfill is an acceptable fix when needed.
function mergeInventoryByMonth(existing, delta) {
  // Helper: pick larger { i, o } pair, treating missing as zero
  const maxIO = (a, b) => ({
    i: Math.max((a && a.i) || 0, (b && b.i) || 0),
    o: Math.max((a && a.o) || 0, (b && b.o) || 0),
  });
  // Helper: per-key per-month max-merge for catAgg / rtAgg
  const mergeMonthlyMax = (oldAgg, newAgg) => {
    const out = {};
    const keys = new Set([...Object.keys(oldAgg || {}), ...Object.keys(newAgg || {})]);
    for (const k of keys) {
      out[k] = {};
      const months = new Set([...Object.keys((oldAgg || {})[k] || {}), ...Object.keys((newAgg || {})[k] || {})]);
      for (const m of months) {
        out[k][m] = maxIO((oldAgg || {})[k] && oldAgg[k][m], (newAgg || {})[k] && newAgg[k][m]);
      }
    }
    return out;
  };

  // Index existing + delta SKUs by SKU id
  const existingBySku = new Map((existing.invSku || []).map(e => [e.s, e]));
  const deltaBySku    = new Map((delta.invSku || []).map(e => [e.s, e]));
  const allSkus = new Set([...existingBySku.keys(), ...deltaBySku.keys()]);

  const merged = [];
  for (const sku of allSkus) {
    const ex = existingBySku.get(sku);
    const dx = deltaBySku.get(sku);
    // Union of months from both sides
    const months = new Set([
      ...Object.keys((ex && ex.m) || {}),
      ...Object.keys((dx && dx.m) || {}),
    ]);
    const m = {};
    let ti = 0, to = 0, net = 0;
    for (const month of months) {
      const v = maxIO((ex && ex.m && ex.m[month]) || null, (dx && dx.m && dx.m[month]) || null);
      m[month] = v;
      ti += v.i; to += v.o; net += v.i - v.o;
    }
    // Per-SKU rt totals: union of ref types, max per type
    const rt = {};
    const rtKeys = new Set([
      ...Object.keys((ex && ex.rt) || {}),
      ...Object.keys((dx && dx.rt) || {}),
    ]);
    for (const r of rtKeys) {
      rt[r] = maxIO((ex && ex.rt && ex.rt[r]) || null, (dx && dx.rt && dx.rt[r]) || null);
    }
    // Batch count: take the larger
    const bc = Math.max((ex && ex.bc) || 0, (dx && dx.bc) || 0);
    merged.push({
      s: sku,
      p: (dx && dx.p) || (ex && ex.p) || "",
      c: (dx && dx.c) || (ex && ex.c) || "",
      u: (dx && dx.u) || (ex && ex.u) || "",
      m, rt, ti, to, net, bc,
    });
  }

  // moMovements: per-MO max of totalIn and totalOut independently. Carry
  // over the SKU/prod metadata from whichever side has it (delta wins on
  // ties because it's likely fresher).
  const mergedMo = {};
  const moKeys = new Set([
    ...Object.keys(existing.moMovements || {}),
    ...Object.keys(delta.moMovements || {}),
  ]);
  for (const mo of moKeys) {
    const ex = (existing.moMovements || {})[mo];
    const dx = (delta.moMovements || {})[mo];
    mergedMo[mo] = {
      mo,
      sku: (dx && dx.sku) || (ex && ex.sku) || "",
      prod: (dx && dx.prod) || (ex && ex.prod) || "",
      totalIn: Math.max((ex && ex.totalIn) || 0, (dx && dx.totalIn) || 0),
      totalOut: Math.max((ex && ex.totalOut) || 0, (dx && dx.totalOut) || 0),
    };
  }

  return {
    invSku: merged,
    invCat: mergeMonthlyMax(existing.invCat, delta.invCat),
    invRt:  mergeMonthlyMax(existing.invRt,  delta.invRt),
    months: [...new Set([...(existing.months || []), ...(delta.months || [])])].sort(),
    moMovements: mergedMo,
  };
}

app.post("/api/cin7/inventory-movements", async (req, res) => {
  // Auth: shared secret. Returns 401 instead of standard requireAuth so this
  // endpoint can be hit by Apps Script without a session cookie.
  const expected = process.env.INVENTORY_SYNC_SECRET;
  if (!expected) return res.status(503).json({ ok: false, error: "INVENTORY_SYNC_SECRET not configured on server" });
  const provided = req.headers["x-vf-sync-secret"] || req.headers["X-VF-Sync-Secret"];
  if (!provided || provided !== expected) return res.status(401).json({ ok: false, error: "Invalid sync secret" });

  try {
    const body = req.body || {};
    const { filename, contentBase64, fileLastModified } = body;
    if (!contentBase64) return res.status(400).json({ ok: false, error: "Missing contentBase64 in body" });
    const buffer = Buffer.from(contentBase64, "base64");
    let parsed;
    try {
      parsed = parseInventoryXlsx(buffer);
    } catch (e) {
      return res.status(400).json({ ok: false, error: "Parse failed: " + e.message });
    }

    // Per-month-replace merge — preserves history outside the rolling window
    const existing = readData("inventory") || {};
    const merged = mergeInventoryByMonth(existing, parsed);

    const finalData = {
      ...merged,
      lastSync: new Date().toISOString(),
      lastSyncSource: "gdrive-auto-sync",
      lastSyncFile: filename || null,
      lastSyncFileModified: fileLastModified || null,
    };
    writeData("inventory", finalData);

    console.log(`[gdrive-sync] Ingested ${filename || "(unnamed)"} — ${parsed.rowCount} rows · ${parsed.invSku.length} new-window SKUs · ${Object.keys(parsed.moMovements).length} new-window MOs · merged total: ${finalData.invSku.length} SKUs · ${Object.keys(finalData.moMovements).length} MOs`);
    res.json({
      ok: true,
      filename: filename || null,
      windowRowCount: parsed.rowCount,
      windowSkuCount: parsed.invSku.length,
      windowMoCount: Object.keys(parsed.moMovements).length,
      mergedSkuCount: finalData.invSku.length,
      mergedMoCount: Object.keys(finalData.moMovements).length,
      windowMonths: parsed.months,
      allMonths: finalData.months,
      lastSync: finalData.lastSync,
    });
  } catch (e) {
    console.error("[gdrive-sync] Ingest error:", e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── Error Reporting: trailing-7-day production-run zero-actual detector ─────
//
// Reads the production_run_errors_t7d blob (populated by the 07:00 UTC cron
// or the manual Sync now button). The detection is done in the sync function
// — this endpoint just serves the precomputed result.
//
// Grouped by full MO reference (e.g. "MO-00774/3") to keep per-batch
// granularity. Sibling lines (other components on the same run) are returned
// for context.
app.get("/api/error-reporting/zero-out-mo-bom", (req, res) => {
  const blob = readData("production_run_errors_t7d");
  if (!blob || !Array.isArray(blob.flagged)) {
    return res.json({
      ok: true,
      lastSync: null,
      windowStart: null,
      windowEnd: null,
      completedRunsScanned: 0,
      parentOrdersScanned: 0,
      detailFailures: 0,
      flaggedRunCount: 0,
      flaggedLineCount: 0,
      flagged: [],
    });
  }

  res.json({
    ok: true,
    lastSync: blob.lastSync || null,
    windowStart: blob.windowStart || null,
    windowEnd: blob.windowEnd || null,
    completedRunsScanned: blob.completedRunsScanned || 0,
    parentOrdersScanned: blob.parentOrdersScanned || 0,
    detailFailures: blob.detailFailures || 0,
    flaggedRunCount: blob.flaggedRunCount || 0,
    flaggedLineCount: blob.flaggedLineCount || 0,
    flagged: blob.flagged,
  });
});

// ── Production Output: trailing-7-day "what did we produce" feed ────────────
//
// Same source blob as the Error Reporting endpoint above — populated by the
// shared 07:00 UTC cron / Sync now button. This endpoint serves the
// allCompletedRuns[] slice, rolled up by product category for the
// Production output tab.
//
// Response shape:
//   {
//     ok, lastSync, windowStart, windowEnd,
//     totalQty:        sum across all outputs (treat units as 'kg' for headline KPI),
//     totalsByUnit:    { kg: 5210, Each: 14000, ... } — for the truthful breakdown
//     runCount:        number of completed runs in the window
//     byCategory:      [ { category, totalQty, unit, runs: [{moRef, fgSKU, fgProduct, qty, completionDate, workCenter, outputs:[...]}] } ]
//     byWorkCenter:    [ { workCenter, totalQty, unit } ]
//   }
app.get("/api/production-output/last-7d", (req, res) => {
  const blob = readData("production_run_errors_t7d");
  if (!blob || !Array.isArray(blob.allCompletedRuns)) {
    return res.json({
      ok: true,
      lastSync: null,
      windowStart: null,
      windowEnd: null,
      totalQty: 0,
      totalsByUnit: {},
      runCount: 0,
      byCategory: [],
      byWorkCenter: [],
    });
  }

  // Category rollup. Each run contributes its outputs to a category bucket.
  // Multiple FG SKUs in one run can land in different categories — we honor
  // each output's individual category (vs the run-level fgCategory which is
  // just the headline product).
  const catMap = new Map();   // category → { category, totalQty, totalsByUnit, runs: [] }
  const wcMap = new Map();    // workCenter → { workCenter, totalQty, totalsByUnit }
  const totalsByUnit = {};
  let totalQty = 0;

  const ensureCat = name => {
    if (!catMap.has(name)) catMap.set(name, { category: name, totalQty: 0, totalsByUnit: {}, runs: [] });
    return catMap.get(name);
  };
  const ensureWc = name => {
    if (!wcMap.has(name)) wcMap.set(name, { workCenter: name, totalQty: 0, totalsByUnit: {} });
    return wcMap.get(name);
  };

  for (const run of blob.allCompletedRuns) {
    const outputs = Array.isArray(run.outputs) ? run.outputs : [];
    // If a run has multiple outputs, they may be in different categories.
    // Group them by category for the per-run rendering.
    const outputsByCat = new Map();
    for (const o of outputs) {
      const cat = o.category || run.fgCategory || "Other";
      if (!outputsByCat.has(cat)) outputsByCat.set(cat, []);
      outputsByCat.get(cat).push(o);
      // Accumulate cross-totals
      totalQty += o.qty;
      totalsByUnit[o.unit || "(no unit)"] = (totalsByUnit[o.unit || "(no unit)"] || 0) + o.qty;
      const cBucket = ensureCat(cat);
      cBucket.totalQty += o.qty;
      cBucket.totalsByUnit[o.unit || "(no unit)"] = (cBucket.totalsByUnit[o.unit || "(no unit)"] || 0) + o.qty;
    }
    // Work-center attribution: run-level (last operation), as built in sync.
    const wcName = run.workCenter || "(unattributed)";
    const wcBucket = ensureWc(wcName);
    const runQty = outputs.reduce((s, o) => s + o.qty, 0);
    wcBucket.totalQty += runQty;
    // Attribute units proportionally to the dominant unit on the run
    for (const o of outputs) {
      wcBucket.totalsByUnit[o.unit || "(no unit)"] = (wcBucket.totalsByUnit[o.unit || "(no unit)"] || 0) + o.qty;
    }

    // Emit per-category run rows. If a run had no outputs at all, still emit
    // a single row under its FG's category so the feed isn't lossy.
    if (outputs.length === 0) {
      const cat = run.fgCategory || "Other";
      ensureCat(cat).runs.push({
        moRef: run.moRef,
        completionDate: run.completionDate,
        fgSKU: run.fgSKU,
        fgProduct: run.fgProduct,
        workCenter: run.workCenter,
        location: run.location,
        qty: 0,
        unit: "",
        outputs: [],
      });
    } else {
      for (const [cat, catOutputs] of outputsByCat) {
        const catQty = catOutputs.reduce((s, o) => s + o.qty, 0);
        const dominantUnit = catOutputs.length === 1 ? catOutputs[0].unit
          : (catOutputs.reduce((a, b) => a.qty >= b.qty ? a : b)).unit || "";
        ensureCat(cat).runs.push({
          moRef: run.moRef,
          completionDate: run.completionDate,
          fgSKU: run.fgSKU,
          fgProduct: run.fgProduct,
          workCenter: run.workCenter,
          location: run.location,
          qty: catQty,
          unit: dominantUnit,
          outputs: catOutputs,
        });
      }
    }
  }

  // Sort categories by total qty desc, runs within each category by date desc
  const byCategory = Array.from(catMap.values()).sort((a, b) => b.totalQty - a.totalQty);
  for (const c of byCategory) c.runs.sort((a, b) => (b.completionDate || "").localeCompare(a.completionDate || ""));
  const byWorkCenter = Array.from(wcMap.values()).sort((a, b) => b.totalQty - a.totalQty);

  res.json({
    ok: true,
    lastSync: blob.lastSync || null,
    windowStart: blob.windowStart || null,
    windowEnd: blob.windowEnd || null,
    totalQty,
    totalsByUnit,
    runCount: blob.allCompletedRuns.length,
    byCategory,
    byWorkCenter,
    // Diagnostic info so the tab can surface "the sync ran but the filter
    // dropped everything" cases without needing Railway log access.
    listRowCount: blob.listRowCount || 0,
    listTypeCounts: blob.listTypeCounts || {},
    listStatusCounts: blob.listStatusCounts || {},
    listOrderStatusCounts: blob.listOrderStatusCounts || {},
    completedRunsScanned: blob.completedRunsScanned || 0,
    parentOrdersScanned: blob.parentOrdersScanned || 0,
  });
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

  // Product costs run a bit later so the on-hand sync doesn't fight for the
  // Cin7 rate limit. Costs change slowly so daily cadence is overkill but
  // matches the on-hand pattern and means the MRP $$ summary is always fresh.
  cron.schedule("45 6 * * *", async () => {
    console.log("[Cin7 Costs] Nightly sync starting…");
    try {
      const s = await performCin7ProductCostsSync();
      console.log(`[Cin7 Costs] Sync done — ${s.productCount} products, ${s.withCostCount} with non-zero cost`);
    } catch (e) {
      console.error("[Cin7 Costs] Nightly sync failed:", e.message);
    }
  });
  console.log("[Cin7 Costs] Nightly sync scheduled at 06:45 UTC");

  // Production-run error sync — flags BOM input lines with actual=0 across
  // completed runs in the trailing 7 days. Runs after on-hand + costs so we
  // don't compete with them for the 60/min Cin7 rate budget.
  cron.schedule("0 7 * * *", async () => {
    console.log("[ProdRunSync] Nightly sync starting…");
    try {
      const s = await performProductionRunErrorSync();
      console.log(`[ProdRunSync] Sync done — ${s.completedRunsScanned} completed runs in window, ${s.parentOrdersScanned} parent orders fetched (${s.detailFailures} failures), ${s.flaggedRunCount} runs flagged with ${s.flaggedLineCount} zero-actual lines`);
    } catch (e) {
      console.error("[ProdRunSync] Nightly sync failed:", e.message);
    }
  });
  console.log("[ProdRunSync] Nightly sync scheduled at 07:00 UTC");
}

// ── MRP Phase 3: requirements engine ─────────────────────────────────────────
//
// Forward-walk allocation MRP. For each production order in the horizon:
//   1. Recursively expand the FG's BOM to leaf-RM requirements
//   2. Date-stamp each requirement at the order's start date (when the
//      material is needed on the production floor)
//   3. Sort all requirements globally by date
//   4. Walk forward, allocating from (on-hand + on-order + in-transit) per SKU
//   5. When we hit a shortfall, that's an at-risk MO + a suggested PO
//      (qty = the shortfall, must-order-by = need-date − leadTimeDays)
//
// Toggles:
//   - includeUnconfirmed: when false, orders with confirmed===false are
//     skipped entirely (conservative — only buys for committed plan)
//   - applyWastage: when true, expandBom multiplies each component qty by
//     (1 + wastagePct/100)
//   - horizonDays: how far forward to look (default 120 ≈ 4 months)

function getMrpInputs() {
  const orders = readData("vf_orders") || [];
  const bomBlob = readData("vf_boms") || { parents: {} };
  const supplyBlob = readData("vf_supply_settings") || { defaults: {}, perSku: {} };
  const onHandBlob = readData("inventory_onhand") || { bySku: [] };
  const onHandBySku = {};
  for (const row of onHandBlob.bySku || []) onHandBySku[row.sku] = row;
  // Product costs (optional — MRP runs without them, dollar fields are null)
  const costsBlob = readData("product_costs") || { bySku: {} };
  return {
    orders,
    bomParents: bomBlob.parents || {},
    supply: supplyBlob,
    onHandBySku,
    costsBySku: costsBlob.bySku || {},
    costsLastSync: costsBlob.lastSync || null,
  };
}

// Resolve effective lead time for a SKU. Aliases follow once. Contract SKUs
// are flagged so the PO logic can exclude them.
function resolveLeadTime(sku, supply) {
  const defaults = supply.defaults || {};
  const perSku = supply.perSku || {};
  const seen = new Set();
  let cur = sku;
  while (cur && perSku[cur] && perSku[cur].isAlias && !seen.has(cur)) {
    seen.add(cur);
    cur = perSku[cur].aliasOf;
  }
  const entry = perSku[cur];
  if (entry && entry.isContract) {
    return { leadTimeDays: null, isContract: true, source: "contract" };
  }
  if (entry && entry.leadTimeDays != null) {
    return { leadTimeDays: entry.leadTimeDays, isContract: false, source: cur === sku ? "explicit" : `alias:${cur}` };
  }
  // Per-prefix default for packaging
  if (sku && sku.startsWith("PK-") && defaults.packagingDefaultDays != null) {
    return { leadTimeDays: defaults.packagingDefaultDays, isContract: false, source: "packaging-default" };
  }
  return { leadTimeDays: defaults.leadTimeDays != null ? defaults.leadTimeDays : 30, isContract: false, source: "default" };
}

function mrpAddDays(dateStr, days) {
  if (!dateStr) return null;
  const d = new Date(dateStr + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

// Order.sku field formats vary across the import history — sometimes it's a
// bare SKU ("FG-888-810-00-US"), sometimes "SKU · Description"
// ("WIP-5100810-US · 810 CBE_US..."), sometimes "Description  SKU"
// ("505.EU CBE Liquor  WIP-5100042-EU"). To find the right BOM, try a direct
// match first, then look for any BOM parent SKU as a word-boundary substring.
function extractBomSku(orderSku, bomParents) {
  if (!orderSku) return null;
  if (bomParents[orderSku]) return orderSku;
  // Try cleaned forms first (split on common separators, strip whitespace)
  const candidates = orderSku.split(/[·,]/).map(s => s.trim());
  for (const c of candidates) {
    if (bomParents[c]) return c;
  }
  // Substring search across all BOM parents — match longest first so
  // "WIP-5100810" prefers the more specific over a "WIP-5100" substring
  const allSkus = Object.keys(bomParents).sort((a, b) => b.length - a.length);
  for (const sku of allSkus) {
    // Word-boundary regex with escaped special chars
    const re = new RegExp("(^|[^A-Za-z0-9-])" + sku.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "([^A-Za-z0-9-]|$)");
    if (re.test(orderSku)) return sku;
  }
  return null;
}

// Build a flat list of dated leaf-RM requirements from the production plan.
// Each entry: { sku, qtyKg, neededByDate, sourceOrderId, sourceFgSku, sourceQty }
function buildRequirements(orders, bomParents, opts) {
  const horizonEnd = mrpAddDays(opts.today, opts.horizonDays);
  const requirements = [];
  const skipped = { unconfirmed: 0, complete: 0, noStart: 0, outsideHorizon: 0, noBom: 0 };
  const noBomExamples = [];

  for (const o of orders) {
    if (!opts.includeUnconfirmed && o.confirmed === false) { skipped.unconfirmed++; continue; }
    if (o.status === "complete") { skipped.complete++; continue; }
    if (!o.start) { skipped.noStart++; continue; }
    if (o.start > horizonEnd) { skipped.outsideHorizon++; continue; }

    const plannedQty = o.total || (o.qty || 0) * (o.batches || 1);
    if (!o.sku || plannedQty <= 0) { skipped.noBom++; continue; }

    const fgSku = extractBomSku(o.sku, bomParents);
    if (!fgSku) {
      skipped.noBom++;
      if (noBomExamples.length < 5) noBomExamples.push({ orderId: o.orderId, sku: o.sku });
      continue;
    }

    let expansion;
    try { expansion = expandBom(bomParents, fgSku, plannedQty, { applyWastage: opts.applyWastage }); }
    catch (e) { skipped.noBom++; continue; }

    const neededBy = o.start < opts.today ? opts.today : o.start;
    for (const leaf of Object.values(expansion.leaves || {})) {
      if (leaf.qty <= 0) continue;
      // Filter out non-procurable BOM leaves (labor, scrap, output products,
      // anything that isn't a real raw material or packaging SKU). MRP only
      // suggests POs for things that get bought.
      if (!isProcurable(leaf.sku)) continue;
      requirements.push({
        sku: leaf.sku,
        qtyKg: leaf.qty,
        neededByDate: neededBy,
        sourceOrderId: o.orderId || o.id,
        sourceFgSku: fgSku,
        sourceFgQty: plannedQty,
      });
    }
  }
  return { requirements, skipped, noBomExamples };
}

// A BOM leaf is "procurable" if it looks like a real RM, packaging, or
// supplier-purchased component (FG-FL-* flavors). Excludes labor lines,
// SCRAP outputs, and anything else that doesn't get bought.
function isProcurable(sku) {
  if (!sku) return false;
  const s = String(sku).toUpperCase();
  // Known non-procurable patterns
  if (s.startsWith("LABOR")) return false;
  if (s === "SCRAP") return false;
  if (s.startsWith("WIDGET")) return false;   // test data
  if (s.startsWith("TEST")) return false;     // test data
  if (s.startsWith("L ") || s === "L") return false; // bare labor codes
  // Procurable prefixes (matches what the supply settings cover)
  return /^(RM-|PK-|FG-FL|FLV-|VC-|ING-)/.test(s);
}

// Forward-walk allocation. For each SKU, walk requirements in date order,
// drawing from running supply (on-hand + on-order). When supply hits zero,
// every subsequent need becomes a shortfall → suggested PO.
function allocateAndPlan(requirements, onHandBySku, supply, today) {
  // Group requirements by SKU
  const bySku = {};
  for (const r of requirements) {
    if (!bySku[r.sku]) bySku[r.sku] = [];
    bySku[r.sku].push(r);
  }

  const skuResults = [];
  const allAtRiskOrders = new Map(); // orderId → { orderId, shortages: [{sku, qtyShort, neededByDate}] }
  const suggestedPOs = [];

  for (const sku of Object.keys(bySku)) {
    const reqs = bySku[sku].sort((a, b) => a.neededByDate.localeCompare(b.neededByDate));
    const lead = resolveLeadTime(sku, supply);
    const onHand = onHandBySku[sku] || { onHand: 0, allocated: 0, available: 0, onOrder: 0, inTransit: 0, name: "" };

    // Starting supply pool: available (on-hand minus already-allocated to other things in Cin7)
    // + onOrder + inTransit. We model this as a single pool for v1; per-receipt-date
    // bucketing would be a Phase 4 refinement.
    let supplyPool = (onHand.available || 0) + (onHand.onOrder || 0) + (onHand.inTransit || 0);
    const startingSupply = supplyPool;

    let totalDemand = 0;
    let totalShort = 0;
    let earliestShortDate = null;
    const allocations = [];

    for (const r of reqs) {
      totalDemand += r.qtyKg;
      const allocFromPool = Math.min(supplyPool, r.qtyKg);
      const shortage = r.qtyKg - allocFromPool;
      supplyPool -= allocFromPool;
      allocations.push({
        ...r,
        allocatedFromSupply: allocFromPool,
        shortage,
        runningSupplyAfter: supplyPool,
      });
      if (shortage > 0) {
        totalShort += shortage;
        if (!earliestShortDate || r.neededByDate < earliestShortDate) earliestShortDate = r.neededByDate;
        // Record at-risk for this MO
        if (!allAtRiskOrders.has(r.sourceOrderId)) {
          allAtRiskOrders.set(r.sourceOrderId, {
            orderId: r.sourceOrderId,
            sourceFgSku: r.sourceFgSku,
            sourceFgQty: r.sourceFgQty,
            shortages: [],
          });
        }
        allAtRiskOrders.get(r.sourceOrderId).shortages.push({
          sku: r.sku,
          qtyShort: shortage,
          neededByDate: r.neededByDate,
        });
      }
    }

    // Suggest a PO if there's a shortfall AND the SKU isn't contract-managed
    if (totalShort > 0 && !lead.isContract) {
      const mustOrderBy = lead.leadTimeDays != null ? mrpAddDays(earliestShortDate, -lead.leadTimeDays) : null;
      const isOverdue = mustOrderBy != null && mustOrderBy < today;
      suggestedPOs.push({
        sku,
        name: onHand.name || "",
        qtyToOrder: Math.ceil(totalShort),
        earliestNeedDate: earliestShortDate,
        leadTimeDays: lead.leadTimeDays,
        leadTimeSource: lead.source,
        mustOrderByDate: mustOrderBy,
        isOverdue,
        projectedReceiptDate: lead.leadTimeDays != null ? mrpAddDays(today, lead.leadTimeDays) : null,
      });
    }

    skuResults.push({
      sku,
      name: onHand.name || "",
      isContract: lead.isContract,
      leadTimeDays: lead.leadTimeDays,
      leadTimeSource: lead.source,
      startingSupply,
      onHand: onHand.onHand || 0,
      available: onHand.available || 0,
      onOrder: onHand.onOrder || 0,
      inTransit: onHand.inTransit || 0,
      totalDemand,
      totalShort,
      allocations, // detailed timeline for drill-down
    });
  }

  // Sort outputs for stable presentation
  suggestedPOs.sort((a, b) => {
    // Overdue first, then by must-order-by date
    if (a.isOverdue !== b.isOverdue) return a.isOverdue ? -1 : 1;
    return (a.mustOrderByDate || "").localeCompare(b.mustOrderByDate || "");
  });

  const atRiskOrders = [...allAtRiskOrders.values()].sort((a, b) => {
    const aMin = a.shortages.reduce((m, s) => m && m < s.neededByDate ? m : s.neededByDate, null);
    const bMin = b.shortages.reduce((m, s) => m && m < s.neededByDate ? m : s.neededByDate, null);
    return (aMin || "").localeCompare(bMin || "");
  });

  return { skuResults, suggestedPOs, atRiskOrders };
}

// GET /api/mrp/run?includeUnconfirmed=0&horizonDays=120&applyWastage=1
app.get("/api/mrp/run", (req, res) => {
  try {
    const includeUnconfirmed = req.query.includeUnconfirmed === "1" || req.query.includeUnconfirmed === "true";
    const applyWastage = req.query.applyWastage !== "0" && req.query.applyWastage !== "false";
    const horizonDays = Math.max(1, Math.min(365, parseInt(req.query.horizonDays, 10) || 120));
    const today = new Date().toISOString().slice(0, 10);

    const { orders, bomParents, supply, onHandBySku, costsBySku, costsLastSync } = getMrpInputs();
    const { requirements, skipped, noBomExamples } = buildRequirements(orders, bomParents, {
      today, horizonDays, includeUnconfirmed, applyWastage,
    });
    const { skuResults, suggestedPOs, atRiskOrders } = allocateAndPlan(requirements, onHandBySku, supply, today);

    // Enrich each suggested PO with $$ — unit cost from Cin7 product cache,
    // line cost = unitCost × qtyToOrder. Mark missing-cost SKUs explicitly so
    // the UI can flag them rather than silently treating them as $0.
    let totalDollars = 0;
    let overdueDollars = 0;
    let missingCostCount = 0;
    const dollarsByMonth = {}; // YYYY-MM -> { total, overdue, count }
    for (const po of suggestedPOs) {
      const c = costsBySku[po.sku];
      const unitCost = c && c.averageCost > 0 ? c.averageCost : null;
      po.unitCost = unitCost;
      po.lineCost = unitCost != null ? unitCost * po.qtyToOrder : null;
      po.costMissing = unitCost == null;
      if (po.costMissing) missingCostCount++;
      else {
        totalDollars += po.lineCost;
        if (po.isOverdue) overdueDollars += po.lineCost;
        // Bucket by must-order-by month (or earliestNeedDate as fallback)
        const bucketDate = po.mustOrderByDate || po.earliestNeedDate;
        if (bucketDate) {
          const month = bucketDate.slice(0, 7);
          if (!dollarsByMonth[month]) dollarsByMonth[month] = { month, total: 0, overdue: 0, count: 0 };
          dollarsByMonth[month].total += po.lineCost;
          if (po.isOverdue) dollarsByMonth[month].overdue += po.lineCost;
          dollarsByMonth[month].count += 1;
        }
      }
    }
    const dollarsByMonthArr = Object.values(dollarsByMonth).sort((a, b) => a.month.localeCompare(b.month));

    res.json({
      ok: true,
      runAt: new Date().toISOString(),
      today,
      settings: { includeUnconfirmed, applyWastage, horizonDays },
      summary: {
        ordersConsidered: orders.length - (skipped.unconfirmed + skipped.complete + skipped.noStart + skipped.outsideHorizon + skipped.noBom),
        ordersSkipped: skipped,
        noBomExamples,
        requirementCount: requirements.length,
        skuCount: skuResults.length,
        atRiskOrderCount: atRiskOrders.length,
        suggestedPoCount: suggestedPOs.length,
        overduePoCount: suggestedPOs.filter(p => p.isOverdue).length,
        // $$ planning
        currency: "USD",
        totalDollars: Math.round(totalDollars * 100) / 100,
        overdueDollars: Math.round(overdueDollars * 100) / 100,
        missingCostCount,
        costsLastSync,
        dollarsByMonth: dollarsByMonthArr,
      },
      suggestedPOs,
      atRiskOrders,
      skuResults,
    });
  } catch (e) {
    console.error("[mrp] Run failed:", e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

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
