const express = require("express");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const multer = require("multer");
const XLSX = require("xlsx");
const Anthropic = require("@anthropic-ai/sdk");
const cron      = require("node-cron");
const { parseMovementFile } = require("./lib/movement-parser");
const { annotateMovements, summarizeDataQuality } = require("./lib/data-quality");

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

// ── Bug reports ─────────────────────────────────────────────────────────────
//
// User-submitted bug reports captured from the in-app 🐛 button. Auto-
// attached context (tab, URL, user-agent, recent console errors) lets us
// debug without needing to ping the user for repro steps.
//
// Storage: vf_bug_reports = { reports: [...] }, capped at BUG_REPORT_MAX
// entries (oldest evicted) so the file stays bounded even if a flood
// happens.
const BUG_REPORT_MAX = 500;
const BUG_REPORT_STATUSES = new Set(["new", "acknowledged", "resolved"]);

function readBugReports() {
  const blob = readData("vf_bug_reports");
  if (!blob || !Array.isArray(blob.reports)) return { reports: [] };
  return blob;
}

// POST /api/bug-reports — any authed user. Body:
//   { summary, description, expected?, tab?, url?, userAgent?, contextErrors? }
app.post("/api/bug-reports", (req, res) => {
  try {
    const body = req.body || {};
    const summary = String(body.summary || "").trim();
    const description = String(body.description || "").trim();
    if (!summary) return res.status(400).json({ ok: false, error: "summary is required" });
    if (!description) return res.status(400).json({ ok: false, error: "description is required" });
    // Cap text lengths to keep the blob bounded
    const clip = (s, n) => String(s || "").slice(0, n);
    const users = readData("vf_users") || [];
    const u = users.find(x => x && x.id === req.userId);
    const entry = {
      id: "bug_" + crypto.randomBytes(8).toString("hex"),
      ts: new Date().toISOString(),
      userId: req.userId || null,
      userName: u ? u.username : "(unknown)",
      role: u ? u.role : null,
      summary: clip(summary, 200),
      description: clip(description, 5000),
      expected: clip(body.expected, 5000),
      tab: clip(body.tab, 80),
      url: clip(body.url, 500),
      userAgent: clip(body.userAgent, 500),
      contextErrors: Array.isArray(body.contextErrors)
        ? body.contextErrors.slice(0, 10).map(e => ({
            ts: clip(e && e.ts, 40),
            kind: clip(e && e.kind, 40),
            message: clip(e && e.message, 500),
            source: clip(e && e.source, 200),
            lineno: typeof (e && e.lineno) === "number" ? e.lineno : null,
            url: clip(e && e.url, 500),
            status: typeof (e && e.status) === "number" ? e.status : null,
          }))
        : [],
      status: "new",
    };
    const blob = readBugReports();
    blob.reports.push(entry);
    if (blob.reports.length > BUG_REPORT_MAX) {
      blob.reports = blob.reports.slice(blob.reports.length - BUG_REPORT_MAX);
    }
    writeData("vf_bug_reports", blob);
    res.json({ ok: true, id: entry.id });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// GET /api/bug-reports — admin only. Returns newest first.
app.get("/api/bug-reports", (req, res, next) => requireAdmin(req, res, next), (req, res) => {
  const blob = readBugReports();
  const reports = blob.reports.slice().reverse();
  const counts = { new: 0, acknowledged: 0, resolved: 0 };
  for (const r of blob.reports) counts[r.status] = (counts[r.status] || 0) + 1;
  res.json({ ok: true, total: blob.reports.length, counts, reports });
});

// PATCH /api/bug-reports/:id — admin only. Body: { status }.
app.patch("/api/bug-reports/:id", (req, res, next) => requireAdmin(req, res, next), (req, res) => {
  try {
    const status = String((req.body || {}).status || "").trim().toLowerCase();
    if (!BUG_REPORT_STATUSES.has(status)) {
      return res.status(400).json({ ok: false, error: "status must be one of: new, acknowledged, resolved" });
    }
    const blob = readBugReports();
    const r = blob.reports.find(x => x && x.id === req.params.id);
    if (!r) return res.status(404).json({ ok: false, error: "not found" });
    const users = readData("vf_users") || [];
    const actor = users.find(x => x && x.id === req.userId);
    r.status = status;
    if (status === "resolved") {
      r.resolvedTs = new Date().toISOString();
      r.resolvedBy = actor ? actor.username : "(unknown)";
    } else {
      delete r.resolvedTs;
      delete r.resolvedBy;
    }
    writeData("vf_bug_reports", blob);
    res.json({ ok: true, report: r });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
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
- east_mac / west_mac (Mcintyres): Runtime 10 days. Min 2,250 kg / Max 4,300 kg per batch. EU and US products cannot be mixed on the same run. Shift count does NOT change Mcintyre throughput — runtime is fixed.
- refining: Runtime 1 day. Max capacity 6,000 kg/day (4 batches × 1,500 kg). Min 500 kg per batch.
- conching: Runtime 1 day. Min 3,000 kg / Max 6,000 kg per run.
- seed_clean: Throughput 1,302 kg/shift (linear: 2,604 kg at 2 shifts, 3,906 kg at 3 shifts). Same machine for chocolate (grapeseeds) and coffee (chickpeas).
- roaster (Alk/Roaster): Per-product throughput, multi-shift linear.
  - Grapeseeds (chocolate): 847 kg/shift (1,694 at 2 shifts, 2,541 at 3 shifts).
  - Chickpeas (coffee): 1,163 kg/shift (2,326 at 2 shifts, 3,489 at 3 shifts).
  - When sizing roaster capacity for a plan, use the throughput for the SPECIFIC product going through (chickpeas roast faster than grapeseeds).
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

Liquor (sold-as-is) production: roasting → fat melting (24/72 hr) → Mcintyre (10 days) → packout. Liquor IS made on the Mcintyre, not on the chocolate line.

CRITICAL: DO NOT recommend production schedules using straight-line 1:1 scaling of the requested FG qty through every upstream stage. ALWAYS call bom_expand FIRST for any multi-stage production planning so you have BOM-driven quantities at each step. If the user names a product loosely (e.g. "CFC 506 EU"), use find_bom to resolve to the actual SKU before bom_expand.

Order statuses: queued, in-progress, complete, on-hold
Order priorities: high, med, low
Order confirmation: 'confirmed' boolean (default false on new). Tentative orders get visual cue on calendar but are still real schedule entries.

PRODUCT CATEGORIES (the front-end's edit modal validates these — picking the wrong combo will silently strip fields when a user opens the order):
- cat='coffee', sub ∈ {coffee_beans, coffee_ground} — coffee line, runs on the grinder
- cat='liquor',    sub='liquor'    — chocolate liquor + Mcintyre line products (incl. Final Blends like PFS Final Blend)
- cat='chocolate', sub='chocolate' — finished chocolate, fat melter, refining/conching/depositing, AND pouching (per the design system: "purple = finished chocolate, including final blends, bars, pouched product")
- region: 'eu' or 'us', only meaningful for liquor
- temper: 'cbe' or 'cbs' — east_mac DEFAULTS to cbs, west_mac DEFAULTS to cbe; chocolate-line machines run both, supply explicitly when known
- The Mcintyre fat-type default is a soft constraint, not a hard rule. It exists because a fat-type changeover on a Mcintyre is expensive (long cleanout, lost time, wastage on first batch back). If the user EXPLICITLY asks you to "ignore the fat type constraint", "run cross-fat-type", "use west_mac for CBS" (or similar), you can recommend scheduling the off-default temper — but you MUST surface the trade-off in your reply ("This requires a fat-type changeover on the Mcintyre — expect ~X extra hours plus first-batch wastage; the team should confirm they're OK paying that cost"). Default behavior with no explicit override: stick to the fat-type assignment.

When you call add_order: if the machine maps unambiguously, omitting cat/sub is fine — defaults are inferred. But if you know the product (e.g. PFS pouches → cat='chocolate', sub='chocolate'), pass cat/sub explicitly. Use update_order_metadata to fix cat/sub/region/temper/machine on existing orders.

When recommending a production slot for a finished good or WIP:
1. Call get_orders to see what's currently scheduled.
2. If the user's product reference is loose, call find_bom to get the right SKU.
3. Call bom_expand to get per-stage quantities and the upstream WIPs that must be produced. The intermediateStages array tells you each stage's machine + its production lead time.
4. **Call get_on_hand on every intermediate WIP SKU returned by bom_expand** (plus the FG SKU itself if relevant). For each stage where Available ≥ required qty: SKIP the upstream production for that stage and tell the user explicitly ("100 kg of WIP-X already on hand — no need to roast it"). For stages where Available is positive but less than required: reduce the qty for that stage by the available amount and note the reduction.
5. For multi-stage products that still need upstream production, schedule BACKWARD from the desired completion date:
   - Last step (e.g. depositing) ends at the target completion date
   - Each upstream stage must finish before its downstream stage starts (i.e. work the lead times backward)
   - Roasting + fat melting can run in parallel (both feed Mcintyre in the liquor case, or feed downstream stages directly for CFC)
6. Check that each stage's required qty fits within that machine's capacity constraints above.
7. Use find_available_slots (or get_orders + reason about gaps) to pick concrete dates that don't collide with existing scheduled work.
8. Present 2–3 concrete options spanning the full pipeline (each option = a complete set of stage dates) and explain the trade-offs (lead time risk, fat type, batch size fit, etc.). Surface any on-hand substitution in the recap so the user sees the savings.
9. Do NOT book anything — only recommend. The user must ask you to create or update orders to make it happen.

For metadata edits — renaming an MO (orderId), changing the SKU on a row, editing notes, flipping priority, marking confirmed/tentative, or recording actual produced qty — use update_order_metadata. It accepts any subset of those fields in one call. Do NOT just describe the change in text; if the user asked you to change a field, you MUST call this tool to persist it, otherwise the user will see no effect when they click the order on the calendar.

IMPORTANT: Before calling any write tools (add_order, shift_machine_orders, update_order_dates, update_order_status, update_order_quantity, update_order_metadata, delete_order), you MUST:
1. Use get_orders to see the current state
2. Clearly describe to the user exactly what changes you plan to make
3. Wait for them to explicitly confirm (e.g. "yes", "go ahead", "confirm") before executing writes

For delete_order specifically: always name the order ID and SKU in your confirmation request, as deletion is permanent and cannot be undone.

DUPLICATE PREVENTION (very important — recent operator complaints stem from this):
- Before calling add_order, ALWAYS scan get_orders results for any existing order with the SAME SKU and SAME machine whose date range overlaps your proposed start..end window.
- If a match exists, DO NOT add a new order silently. Surface it: "I found an existing order for {SKU} on {machine} from {start} to {end} with {qty} kg. Did you mean to UPDATE this one (use update_order_quantity / update_order_dates) instead of creating a new one?"
- Only proceed with add_order after the user explicitly confirms they want a separate additional order. In that case, pass allow_duplicate=true so the server-side guard doesn't reject it.
- The server-side add_order tool ALSO enforces this — if you skip the check, it will refuse and return a duplicate error. Treat that error as a hint to re-read get_orders and reconcile.

AMBIGUITY HANDLING:
- If a user request could mean two different things (e.g. "add 100kg of X" could mean "create a new order" or "update the existing one to 100kg"), ASK before acting. Do not guess.
- If a quantity, date, machine, or SKU isn't clearly specified, ASK rather than infer.
- Never create multiple orders for the same SKU on the same date unless the user explicitly says they want multiple batches.

MRP & MATERIAL QUESTIONS:
- For ANY question involving materials, purchase orders, ordering, capital commitment, raw-material shortages, supplier orders, or "why is this PO/order looking off", call run_mrp FIRST to see the engine's current output.
- run_mrp returns: summary (total $, PO counts, at-risk count), top suggested POs (in-window), deferred POs (beyond PO horizon), and at-risk MOs. Suggested POs are sorted by line cost desc — biggest commitments first.
- Default to poHorizonDays=30 (what we'd actually order this month). If the user asks about longer-range commitments, increase or set to 0.
- If the user mentions the pipeline tab or a forward look, set includeDrafts=true so pipeline opportunities count as demand.
- includeCompanions=true adds synthetic companion-demand orders per the rules in MRP Setup → Companion demand (e.g. "when liquor is scheduled, generate a flavor-pack order at the same date"). These flow through normal BOM expansion and PO suggestions. Default OFF. Turn ON when the user asks about flavor procurement, air-freight companion products, or "what do I need to order that isn't in the liquor BOM?". Requirements from companion demand carry isCompanionDemand=true + companionDriverOrderId/companionDriverSku so you can attribute them.
- run_mrp is READ-ONLY — it doesn't queue anything for approval. You can call it freely.

DEMAND ATTRIBUTION (read this carefully — this is the #1 source of bot errors on this app):
- When you need to explain WHERE a PO's demand comes from, which orders DROVE a raw-material requirement, or WHY a specific SKU shows up in MRP — you MUST call trace_po_demand on that SKU. NEVER guess based on order size, product names, BOM intuition, or pattern matching ("this is the biggest order so it must be the cause").
- BOMs are recursive and your model of any specific recipe is OFTEN WRONG. A product name like "Nut Free Spread" might contain sunflower paste, not chickpeas — you can't know what's in a BOM without expanding it. trace_po_demand gives you the ground truth.
- If trace_po_demand returns an empty sources list for a SKU you expected to find demand for, the demand for that SKU isn't coming from where you thought. Tell the user clearly: "I was wrong — this RM isn't actually driven by [order you mentioned]. The real sources are [whatever the tool returned]."

FILTER MIRRORING (mandatory, no exceptions):
- When the user states MRP settings in their prompt (e.g. "PO horizon 90", "exclude orders before 6/11", "include pipeline drafts", "120-day planning horizon") — you MUST pass those EXACT values to every tool call (run_mrp, trace_po_demand, get_on_hand). Failure to do this is the most common cause of your numbers diverging from what the user sees in the UI.
- If the user references "MRP" without naming explicit filters, ask them: "Which MRP settings should I use — same as what's currently on your UI? (planning horizon, PO horizon, includeDrafts, excludeBefore)" Don't guess.
- Specifically: excludeBefore is the most common one to forget. If the user said "ignore orders before X" anywhere in the thread, pass excludeBefore=X to EVERY trace_po_demand and run_mrp call until the user changes it. This is non-negotiable.

PO-TOTAL ARITHMETIC (also mandatory):
- A "PO" total = the NET kg you'd commit to a supplier. It equals run_mrp.suggestedPOs[].qtyToOrder. It is NOT the same as gross BOM-expanded demand.
- Before quoting any PO total in your reply, find that exact SKU in run_mrp's suggestedPOs and use qtyToOrder VERBATIM. Do NOT sum bom_expand or trace_po_demand contributions to produce a PO total — those are gross demand and have NOT been netted against on-hand / on-order / in-transit.
- If your computed gross demand is more than 10% different from run_mrp's qtyToOrder for the same SKU, your math is wrong. Re-call the tools with the correct filters; do NOT report the wrong number anyway.
- When explaining a PO, show the gross→net story explicitly: "Gross BOM demand from these orders is X kg. On-hand + on-order covers Y kg. Net PO need (= MRP's suggested qtyToOrder) is Z kg." That makes the chain auditable.

PACKOUT-SCHEDULING GUARD (respect it):
- MRP silently skips MOs scheduled on a packout machine (depositing, pouching, mac_packout) unless their SKU starts with FG-. Only FG-* SKUs carry the packaging BOM (box + bag + intermediate). A packout MO coded against a WIP-, WIP2-, WIP3-, or RM- SKU is a scheduling error and would expand the wrong BOM.
- run_mrp response includes packoutSkipped.count and packoutSkipped.examples. When count > 0, surface it in your reply and list the affected MOs so ops can re-code them. Do NOT silently ignore — this is a signal that some real demand is missing from the current MRP output because those MOs got skipped.

UOM AWARENESS (critical for VC-* and other non-kg SKUs):
- Every suggestedPO includes a "uom" field with Cin7's native unit of measure (Kg, g, Each, case, etc.). ALWAYS include the UOM when quoting qtyToOrder. "Order 500 g of VC-XYZ" not "order 500 of VC-XYZ" — the difference between grams and kg is a 1000× ordering hazard.
- Flavor concentrates (VC-*) are stored in Cin7 as GRAMS, not kg. Don't mentally convert unless the user asks — the qtyToOrder value IS the number they'll enter in Cin7's PO screen, so trust the qty + its uom label.
- If uom is null/missing, note that in your response ("UOM not synced yet — verify against Cin7 before ordering") rather than assuming kg.

FG-LEVEL NETTING (default ON in run_mrp — behavior you must narrate accurately):
- run_mrp now nets FG on-hand against planned production BEFORE expanding BOMs. If a scheduled MO or pipeline draft would produce FG-XXX and there is already FG-XXX.available in stock, the RM demand for that draft is expanded on the NET production qty (planned minus what stock covers), not the gross planned qty. This is real MRP behavior and matches how a scheduler would think.
- run_mrp responses include an fgNettingSummary array — FG SKUs where on-hand stock offset planned production. When it is non-empty, SURFACE THIS in your reply. Each row has these distinct fields you MUST report correctly:
    rawOnHandKg              — total physical inventory in Cin7
    allocatedToSalesOrdersKg — qty already committed to open SOs (unavailable for netting)
    startingAvailableKg      — rawOnHand minus allocated; THE POOL FOR FUTURE NETTING
    totalConsumedKg          — how much of that pool FIFO consumed against planned production
    availableRemainingKg     — startingAvailable minus totalConsumed
  Sample narration: "FG-888-860 has 4,938 kg on hand, but 4,500 kg is already allocated to open SOs (SO-XXXXX). Only 438 kg is available for future netting — FIFO applied that to the Sep pipeline draft first, leaving 0 kg for Dec." NEVER report the raw on-hand as if it were available — that misleads the user into thinking netting is broken when it's actually working correctly against a smaller-than-expected pool. Ordering is FIFO by need-by-date across all planned orders (real MOs AND pipeline drafts) — earliest need consumes first.
- trace_po_demand rows now include sourceFgGrossQty (pre-netting), sourceFgQty (net, drives RM), and sourceFgOffsetKg (FG on-hand absorbed). If gross does not equal net for a row, mention it in your attribution table: "PIPELINE-Cargill US-2026-12 · FG-888-858 · 15,000 kg planned → 5,136 kg offset by FG on-hand → 9,864 kg net → 3,156 kg RM contribution."
- Only set netFgOnHand=false if the user explicitly asks for a "gross production plan" or "what would we need to buy if the FG shelf were empty" scenario. Never disable it silently.

WIP-LEVEL NETTING (multi-level MRP, always on, no toggle):
- MRP also cuts off recursion at WIP SKUs that are already being produced by another scheduled MO. Example: MO-00933 produces WIP1-XXX and MO-00934 consumes WIP1-XXX in its BOM. Without this cutoff, expanding MO-00934 would recurse through WIP1's recipe and double-count all the RMs (sugar, fat, etc.) that MO-00933's expansion already accounted for. With the cutoff, MO-00934's expansion stops at WIP1 (treats it as a leaf) because MO-00933 is already producing it.
- run_mrp responses include wipNettingSummary — an array of WIP SKUs where cutoff fired, with fields: consumedKg (total demand from downstream MOs), plannedProductionKg (what other MOs will make), onHandKg (available WIP stock), coverageGapKg (max(0, consumed − planned − onHand)), downstreamOrderCount.
- If coverageGapKg > 0 for a WIP, that means the scheduled production is UNDER-planned relative to downstream consumption. RM demand for that gap is NOT reflected in this MRP run. SURFACE THIS to the user as a warning: "WIP1-XXX has a 500 kg planning gap — scheduled production of X kg doesn't cover downstream consumption of Y kg. Ops should schedule an additional MO to close the gap; the RM contribution for the gap is not in the current PO suggestions."
- If coverageGapKg = 0 (planned production + on-hand fully covers consumption), no user-visible narration needed unless they ask. Just don't be surprised that expanding downstream MOs shows sugar/fat demand as zero — that's expected because the producing MO's expansion has it.

PER-ORDER ATTRIBUTION TABLE (always include when listing demand drivers):
- When trace_po_demand returns multiple sources, show them as a table with these columns: source orderId, start date (for synthetic pipeline drafts this is the *computed production start* — ship month minus a channel-specific production lead time; label as "pipeline draft, prod start X (ships Y)" so the user can see both), FG SKU + qty, kg of the leaf RM contributed, % of total.
- Stale orders that should have been filtered out are MOST visible in this table — their start dates will be before the user's excludeBefore cutoff and the user can immediately see the violation. Including this table is your built-in audit trail.

Combine the output with bom_expand when the user wants the per-FG breakdown: trace_po_demand tells you which orders drive the leaf demand; bom_expand on those orders' FG SKUs shows the recursive chain.

QUEUED FOR APPROVAL (how write tools work now):
- Every call to a mutating tool (add_order, shift_machine_orders, update_order_dates, update_order_status, update_order_quantity, update_order_metadata, delete_order) is QUEUED, not executed immediately. The user sees a preview card in the UI and must click "Apply" before the action takes effect.
- The tool result you receive will say "status: queued_for_approval" — that means the change is staged, not committed. The schedule will NOT update for your next get_orders call until the user applies.
- In your reply, narrate what you queued in plain language ("I've queued an add_order for {SKU} on {machine}, qty {qty}, {start}..{end}. Click Apply to commit.") so the preview card has clear context.
- Multiple queued actions accumulate. If you queue several in one turn, list them all in your reply.

TRACEABILITY (SQF audit prep — Q3/Q4 2026):
- Three tools cover lot-level questions: trace_lot (single lot history), trace_fg_lineage (FG → components upstream + shipments downstream), trace_rm_history (RM receipt + consumption ledger).
- Every response includes data_quality flags for any movement falling in one of the four known-noisy windows. You MUST surface these to the user verbatim in your narration — do not silently paper over them. The auditor will ask about the same anomalies the flags call out.
- The four windows (as of 2026-07-15, tunable via env):
    1. Pre-standup (2025-01-01 → 2025-04-30): historical entries incomplete; absence of a receipt before this window is not proof it didn't happen.
    2. Lot-code migration (2025-10-01 → 2025-12-31): many stock adjustments were paper relabels (lot-only → lot+pallet tracking), not physical movements. Two offsetting ST rows around this time are likely a relabel.
    3. VC-via-stock-adjustment (2025-10-01 → 2025-12-31): flavor-component (VC-*) receipts entered as ST rather than PO due to bad process. A ST-in on a VC- SKU here is likely a legitimate supplier receipt without linked PO.
    4. Full facility count (2026-04-01 → 2026-06-30): large ST corrections are inventory-count reconciliation, not shrinkage.
- When the user asks "what happened to lot X", "how was FG lot Y built", "trace this batch back to the raw materials" — call the trace_* tool. Do NOT reason from BOM structure or MRP data — the lot-level movement history is ground truth; BOM expansion is a plan of intent, movement history is what actually happened.
- Reference-type prefixes in the response: PO (purchase order receipt), MO (production activity — 'MO-NNNNN/N' means batch N of the MO), ST (stock adjustment), TR (transfer between locations), SO (sales order shipment), FG (Cin7 'Assembly' — UOM conversion or lot consolidation, legacy at Voyage, paper-only movement).
- Supplier attribution: trace_fg_lineage terminal RM/VC/PK leaves whose origin is a PO include a "supplier" field with {name, code, orderDate} pulled from the ACTUAL Cin7 PO (via the nightly purchase_orders sync). When surfacing a lot's chain, mention the supplier next to the origin PO — e.g. "RM-110000-00 lot 45131 came in via PO-00040 · supplier: Cargill Inc · PO ordered 2025-03-15". This is the actual supplier on the PO, not a SKU default, so no caveat needed. Lots with an ST-* origin (stock adjustment) have no supplier field — that's honest, ST lots didn't come in through a supplier receipt at all.
- Customer attribution: trace_fg_lineage / trace_lot / trace_fg_lineage's downstream sections and the forward-trace endpoint now include per-shipment {customer, customer_reference, so_order_date, so_ship_date} on any Out row with ref_type=SO, pulled from the actual Cin7 sale via the nightly sales_orders sync. When surfacing where a FG lot ended up, name the customer — e.g. "FG-604-102-00 lot 0903174175 shipped 90 cases on 2026-06-27 via SO-00119 to Revolution Foods". Ship-date caveat: two dates can differ on a shipment — the movement_date is when the inventory decrement was posted in Cin7 (may lag the physical event), while so_ship_date is what the SO record shows as the promised or actual ship date. If they differ meaningfully, surface both. Lots that were transferred (TR) or adjusted-out (ST) rather than shipped will not have a customer.
- Date fidelity caveat (applies broadly): every date in the trace reflects what ops entered into Cin7 at the time of the record. If a clerk logged a PO receipt three days after the physical goods arrived without back-dating the movement, the trace shows the entry date, not the actual receipt date. Our code faithfully reports what's in Cin7 — the fix for date-accuracy audit issues is upstream data hygiene (train ops to record actual event dates in Cin7, not just entry timestamps), not something we can correct downstream. If a user asks "why does this date look wrong" or auditor questions arise about date fidelity, surface this framing.

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
        allow_duplicate: { type: "boolean", description: "Bypass the same-SKU/same-machine/overlapping-date-range duplicate guard. Only set to true after the USER has explicitly confirmed they want a separate parallel order on top of an existing one. Defaults to false." },
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
  {
    name: "trace_po_demand",
    description: "GROUND-TRUTH source attribution for a raw material's demand. Returns the list of orders/pipeline drafts whose BOM expansion produced demand for this RM SKU, with per-source kg contribution. ALWAYS call this BEFORE explaining where a PO's demand comes from — do NOT guess based on order size, product names, or your own intuition about what a BOM might contain. BOMs are recursive and your model of any specific recipe is often wrong (e.g., a 'Nut Free Spread' might contain sunflower paste, not chickpeas). If you tell the user 'this PO is driven by order X' and X doesn't appear in this tool's output, you're hallucinating — the truth is whoever IS in the output. Mirrors run_mrp's settings (poHorizonDays / horizonDays / includeUnconfirmed / includeDrafts / excludeBefore) for consistency.",
    input_schema: {
      type: "object",
      properties: {
        sku: { type: "string", description: "The raw material / leaf SKU whose demand sources you want to trace (e.g. RM-110007)" },
        horizonDays:        { type: "number",  description: "Planning horizon (default 120) — mirror run_mrp" },
        includeUnconfirmed: { type: "boolean", description: "Default false" },
        includeDrafts:      { type: "boolean", description: "Default false — set true if you want pipeline drafts considered" },
        includeCompanions:  { type: "boolean", description: "Default false — set true to include synthetic companion-demand orders (see run_mrp docs on companion rules)" },
        excludeBefore:      { type: "string",  description: "YYYY-MM-DD optional" },
      },
      required: ["sku"],
    },
  },
  {
    name: "get_on_hand",
    description: "Check current on-hand inventory for one or more SKUs. CRITICAL for multi-stage planning — when sequencing upstream production (after bom_expand), call this on the intermediate WIP SKUs to see if any are already in stock. If Available covers the requirement, recommend SCALING DOWN or SKIPPING the upstream production stage and explain it to the user. Returns OnHand, Allocated, Available, OnOrder, InTransit, and locations per SKU. Available = OnHand − Allocated is the right number to compare against new requirements (anything Allocated is already promised to other orders). The on-hand snapshot is daily (06:30 UTC) so it's reasonably fresh but not real-time — note that to the user if the freshness matters.",
    input_schema: {
      type: "object",
      properties: {
        skus: { type: "array", items: { type: "string" }, description: "List of SKUs to check. Cap of 50 per call." },
      },
      required: ["skus"],
    },
  },
  {
    name: "run_mrp",
    description: "Read-only. Runs the MRP engine against the current schedule + BOMs + on-hand inventory and returns a COMPACT summary: total $ committed, suggested POs (top by line cost), deferred POs, at-risk MOs. Use this when the user asks ANY question involving materials, ordering, capital commitment, supplier orders, raw material shortages, or 'why is this PO/order looking off'. The response also lets you spot anomalies (oversized POs, unexpected demand, at-risk MOs). To dig deeper into a suspicious PO, follow up with get_orders to see what's scheduled and bom_expand on those orders to trace the demand chain. Default settings mirror the UI: poHorizonDays=30 (only suggest POs that must be ordered within 30 days), horizonDays=120 (demand window), includeUnconfirmed=false. Pipeline drafts are NOT included unless includeDrafts=true. Each suggestedPO carries a `uom` field (Cin7 native unit — Kg/g/Each/case); ALWAYS include the UOM when quoting qtyToOrder to the user, because VC-* flavor SKUs are stored in grams and misreading grams as kg is a 1000× ordering hazard.",
    input_schema: {
      type: "object",
      properties: {
        poHorizonDays:      { type: "number",  description: "Only suggest POs whose must-order-by date falls within this many days from today. Default 30. Set to 0 to disable the cap and see every PO across the whole demand window." },
        horizonDays:        { type: "number",  description: "Planning horizon for demand — how far ahead MRP looks at orders. Default 120." },
        includeUnconfirmed: { type: "boolean", description: "Include tentative/unconfirmed orders. Default false." },
        includeDrafts:      { type: "boolean", description: "Include pipeline drafts as additional demand. Default false." },
        includeCompanions:  { type: "boolean", description: "Include companion-demand rules: when a driver SKU (e.g., chocolate liquor) is scheduled, synthetic orders are generated for its configured companion SKUs (e.g., flavor packs) at the same need-date. Flow through the normal MRP pipeline — BOM expansion, FG netting, PO suggestions. Default false. Ask the user before enabling if they haven't mentioned it — this can materially change the PO $$ if rules exist. Companion-derived requirements are flagged isCompanionDemand=true with companionDriverOrderId/companionDriverSku so you can attribute them." },
        excludeBefore:      { type: "string",  description: "Skip orders with start date before this (YYYY-MM-DD). Useful for filtering stale TBD orders." },
        topN:               { type: "number",  description: "How many top-$ suggested POs to return. Default 20." },
        netFgOnHand:        { type: "boolean", description: "Net finished-good on-hand inventory against planned production before computing RM demand. Default true. If a scheduled MO or pipeline draft would produce FG that's already sitting in stock (on-hand minus SO allocations), MRP subtracts that qty from planned production first, then expands the BOM on the reduced qty. This is what real MRP does; only turn OFF (netFgOnHand=false) when you want a 'gross production plan' view that shows what a full run would require ignoring existing FG inventory." },
      },
    },
  },
  {
    name: "trace_lot",
    description: "Return the full chronological movement history for a specific lot code (batch #) — every PO receipt, MO consumption, stock adjustment, transfer, sales-order dispatch that touched this lot. Use when the user asks 'what happened to lot X', 'where did lot Y come from', 'audit trail for lot Z'. Response includes data-quality flags if any movements fall in one of the four known-noisy windows (pre-standup Q2 2025, lot-code migration Q4 2025, VC-via-stock-adjustment Q4 2025, full facility count Q2 2026) — SURFACE those flags verbatim in your reply.",
    input_schema: {
      type: "object",
      properties: { lot: { type: "string", description: "Lot code / batch # exactly as it appears in Cin7 (case-sensitive, may contain spaces or hyphens)." } },
      required: ["lot"],
    },
  },
  {
    name: "trace_fg_lineage",
    description: "For a finished-good (or WIP) lot code, walk the assembly graph upstream (FG lot → producing MO → BOM inputs consumed → recurse until raw material receipts) AND downstream (where the FG shipped — SO, TR, ST, FG-assembly consolidation). Use when the user asks 'how was FG lot X built', 'what went into this pack', 'trace this shipment back to the raw material batch', 'if RM lot Y was recalled, which FG lots would we recall'.",
    input_schema: {
      type: "object",
      properties: {
        lot:       { type: "string", description: "Lot code to trace." },
        max_depth: { type: "number", description: "Upstream recursion depth cap (default 5, max 10)." },
      },
      required: ["lot"],
    },
  },
  {
    name: "trace_rm_history",
    description: "For a raw material (RM/VC/PK) — either a specific lot code OR a SKU — return the full receipt/consumption/adjustment ledger with a running on-hand balance. Groups: receipts (PO in), consumptions (MO out), adjustments (ST), transfers (TR). Use when the user asks 'when did we receive this RM', 'where has this ingredient been consumed', 'why did the RM balance drop on date X'.",
    input_schema: {
      type: "object",
      properties: { query: { type: "string", description: "Either a lot code (batch #) or a SKU (RM-/VC-/PK-/WIP-/FG- prefixed)." } },
      required: ["query"],
    },
  },
];

function addDays(dateStr, days) {
  if (!dateStr) return dateStr;
  const d = new Date(dateStr + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

// Shared duplicate detector — used by both the add_order tool handler (to
// REFUSE) and the chat preview path (to WARN). Returns the conflicting
// order or null. Ignores completed orders since those are historical and
// shouldn't block new scheduling.
function findAddOrderConflict(allOrders, input) {
  const norm = s => String(s || "").trim().toLowerCase();
  const newStart = input.start || "";
  const newEnd   = input.end || newStart;
  if (!newStart || !input.sku || !input.machine) return null;
  const rangesOverlap = (aStart, aEnd, bStart, bEnd) => {
    if (!aStart || !bStart) return false;
    const aE = aEnd || aStart, bE = bEnd || bStart;
    return aStart <= bE && bStart <= aE;
  };
  return (allOrders || []).find(o =>
    o && o.status !== "complete"
    && norm(o.sku) === norm(input.sku)
    && o.machine === input.machine
    && rangesOverlap(o.start, o.end, newStart, newEnd)
  ) || null;
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

// Human-readable summary of the user's current MRP-tab settings, injected
// as a banner at the end of the system prompt for each chat turn. Tells
// the bot what filters will be auto-applied if it omits them on tool
// calls — keeps its narration honest.
function buildMrpSettingsBanner(s) {
  if (!s || typeof s !== "object") return "";
  const parts = [];
  if (isFinite(s.horizonDays))   parts.push(`planning horizon ${s.horizonDays}d`);
  if (isFinite(s.poHorizonDays)) parts.push(`PO horizon ${s.poHorizonDays}d`);
  if (s.includeUnconfirmed)      parts.push("includeUnconfirmed=true");
  if (s.includeDrafts)           parts.push("includeDrafts=true");
  if (s.includeCompanions)       parts.push("includeCompanions=true");
  if (s.excludeBefore && /^\d{4}-\d{2}-\d{2}$/.test(String(s.excludeBefore).trim())) {
    parts.push(`excludeBefore=${String(s.excludeBefore).trim()}`);
  }
  if (!parts.length) return "";
  return `USER'S CURRENT MRP TAB SETTINGS (these auto-apply to any run_mrp / trace_po_demand call where you omit the corresponding parameter): ${parts.join(", ")}. If the user explicitly asks for a different scenario in this turn, pass the parameter explicitly to override — otherwise just call the tool and the server will use these. Your narration MUST match whatever the tool's response.settings field reports as the actually-applied values.`;
}

async function executeAITool(name, input, context) {
  // Tool-layer filter inheritance: when the bot omits a filter param on
  // run_mrp / trace_po_demand, fall back to the user's current MRP-tab
  // settings (passed in by the chat handler). This is the structural
  // backstop for the bot's filter-amnesia problem — even when it forgets
  // to pass excludeBefore or includeDrafts, the user's UI setting wins
  // over a hardcoded default.
  const userSettings = (context && context.mrpSettings) || {};
  const pickNum = (paramVal, userVal, fallback) =>
    isFinite(paramVal) ? paramVal : (isFinite(userVal) ? userVal : fallback);
  const pickBool = (paramVal, userVal, fallback) =>
    typeof paramVal === "boolean" ? paramVal : (typeof userVal === "boolean" ? userVal : fallback);
  const pickStr = (paramVal, userVal) =>
    (typeof paramVal === "string" && paramVal.trim()) ? paramVal :
    (typeof userVal === "string" && userVal.trim()) ? userVal : "";
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
      // Duplicate guard — belt-and-suspenders for the prompt-level rule.
      // Same SKU + same machine + overlapping date range = refusal unless
      // the caller explicitly passed allow_duplicate=true. Drives operators
      // toward update_order_* tools instead of stacking parallel orders.
      if (!input.allow_duplicate) {
        const conflict = findAddOrderConflict(orders, { sku, machine, start, end });
        if (conflict) {
          return {
            ok: false,
            duplicate: true,
            error: `Duplicate detected: order '${conflict.orderId}' for SKU '${conflict.sku}' on ${conflict.machine} already covers ${conflict.start}..${conflict.end || conflict.start} (qty ${conflict.qty}, status ${conflict.status}). Use update_order_quantity or update_order_dates on the existing order, OR retry add_order with allow_duplicate=true if the user has confirmed they want a separate parallel order.`,
            existing: {
              orderId: conflict.orderId,
              entityId: conflict.id,
              sku: conflict.sku,
              machine: conflict.machine,
              start: conflict.start,
              end: conflict.end,
              qty: conflict.qty,
              batches: conflict.batches,
              status: conflict.status,
            },
          };
        }
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
      // Same kgPerUnit override MRP applies — keeps bom_expand's output
      // in lockstep with what buildRequirements sees, so the bot's
      // analysis doesn't disagree with the MRP UI on PFS-style SKUs.
      const supplyBlob = readData("vf_supply_settings") || { perSku: {} };
      const overrideKgPerUnit = (supplyBlob.perSku && supplyBlob.perSku[sku] && supplyBlob.perSku[sku].kgPerUnit) || null;
      const expandQty = overrideKgPerUnit ? (qty / overrideKgPerUnit) : qty;
      let result;
      try { result = expandBom(blob.parents, sku, expandQty, { applyWastage: true }); }
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
        kgPerUnit: overrideKgPerUnit,
        adjustedExpansionQty: overrideKgPerUnit ? Math.round(expandQty * 1000) / 1000 : null,
        leafRequirements: leaves,
        intermediateStages: enriched,
        note: overrideKgPerUnit
          ? `kgPerUnit=${overrideKgPerUnit} applied — your qty of ${qty} kg was divided by ${overrideKgPerUnit} to get ${Math.round(expandQty * 100) / 100} batches before BOM expansion. The leaf requirements below reflect the actual material needed for ${qty} kg of finished product.`
          : "leafRequirements = leaf-level raw materials to procure. intermediateStages = each WIP that must be produced (with its machine and lead time) — use these for backward production scheduling.",
      };
    }

    case "trace_po_demand": {
      const targetSku = String(input.sku || "").trim();
      if (!targetSku) return { ok: false, error: "sku required" };

      const horizonDays        = Math.max(1, Math.min(365, pickNum(input.horizonDays, userSettings.horizonDays, 120)));
      const includeUnconfirmed = pickBool(input.includeUnconfirmed, userSettings.includeUnconfirmed, false);
      const includeDrafts      = pickBool(input.includeDrafts, userSettings.includeDrafts, false);
      const includeCompanions  = pickBool(input.includeCompanions, userSettings.includeCompanions, false);
      const excludeBeforeRaw   = pickStr(input.excludeBefore, userSettings.excludeBefore);
      const excludeBeforeDate  = /^\d{4}-\d{2}-\d{2}$/.test(excludeBeforeRaw) ? excludeBeforeRaw : null;
      const today              = new Date().toISOString().slice(0, 10);

      const { orders: rawOrders, bomParents, supply, onHandBySku } = getMrpInputs();
      let mrpOrders = rawOrders;
      if (includeDrafts) {
        const pipelineBlob = readData("vf_pipeline_drafts");
        const drafts = (pipelineBlob && pipelineBlob.drafts) || [];
        const synth = drafts.map(synthPipelineDraftAsOrder).filter(Boolean);
        mrpOrders = rawOrders.concat(synth);
      }
      if (includeCompanions) {
        const cblob = _companionRulesBlob();
        const synths = generateCompanionOrders(mrpOrders, cblob.rules || []);
        if (synths.length) mrpOrders = mrpOrders.concat(synths);
      }

      const netFgOnHand = pickBool(input.netFgOnHand, userSettings.netFgOnHand, true);
      const { requirements, fgNettingSummary, wipNettingSummary } = buildRequirements(mrpOrders, bomParents, {
        today, horizonDays, includeUnconfirmed, applyWastage: true, excludeBeforeDate, supply,
        onHandBySku, netFgOnHand,
      });

      // Filter to entries that contributed to demand for the target SKU.
      // buildRequirements emits one entry per (leaf SKU, source order) so a
      // single FG order can contribute to multiple RMs but only one entry
      // per RM per order — we just sum qtyKg per source for safety.
      const matching = requirements.filter(r => r.sku === targetSku);
      if (matching.length === 0) {
        return {
          ok: true,
          sku: targetSku,
          totalKgNeeded: 0,
          sourceCount: 0,
          sources: [],
          note: `No order or pipeline draft in the current planning window expanded to demand for '${targetSku}'. If you previously claimed this RM was driven by a specific order, you were wrong — tell the user clearly. Possible reasons: the RM isn't in any active BOM, the relevant orders are outside the horizon, or the SKU spelling differs (call find_bom or get_orders to verify).`,
        };
      }

      // Group by (orderId, sourceFgSku) — NOT just orderId. Pipeline drafts
      // for the same customer+month collapse to one orderId
      // (PIPELINE-<customer>-YYYY-MM) but can span multiple SKUs (e.g.
      // Cargill US Dec 2026 = FG-888-859 Inclusion + FG-888-858 Coating).
      // Real MOs can also share an orderId across sub-batches (MO-00293/1,
      // MO-00293/2 both have orderId "MO-00293"). Grouping on orderId alone
      // sums totalQtyKg correctly across contributions but reports
      // sourceFgQty from ONLY the first one seen, so the ratio displayed
      // looks impossible (e.g. "6,000 kg FG needs 7,046 kg RM" when the
      // reality is 4+ pipeline drafts totaling ~26,000 kg of FG). Keying on
      // (orderId, sourceFgSku) sums both sides consistently.
      const bySource = new Map();
      for (const r of matching) {
        const key = (r.sourceOrderId || "(unknown)") + "|" + (r.sourceFgSku || "");
        if (!bySource.has(key)) {
          bySource.set(key, {
            sourceOrderId: r.sourceOrderId,
            sourceFgSku: r.sourceFgSku,
            sourceFgQty: 0,          // net qty (drives RM demand)
            sourceFgGrossQty: 0,     // pre-netting planned qty
            sourceFgOffsetKg: 0,     // FG on-hand that absorbed demand
            totalQtyKg: 0,
            neededByEarliest: null,
            contributionCount: 0,
          });
        }
        const entry = bySource.get(key);
        entry.totalQtyKg += r.qtyKg;
        entry.sourceFgQty += Number(r.sourceFgQty || 0);
        entry.sourceFgGrossQty += Number(r.sourceFgGrossQty != null ? r.sourceFgGrossQty : r.sourceFgQty || 0);
        entry.sourceFgOffsetKg += Number(r.sourceFgOffsetKg || 0);
        entry.contributionCount += 1;
        if (!entry.neededByEarliest || (r.neededByDate && r.neededByDate < entry.neededByEarliest)) {
          entry.neededByEarliest = r.neededByDate;
        }
      }

      const sources = Array.from(bySource.values())
        .map(s => ({ ...s, totalQtyKg: Math.round(s.totalQtyKg * 1000) / 1000 }))
        .sort((a, b) => b.totalQtyKg - a.totalQtyKg);
      const totalKgNeeded = Math.round(matching.reduce((s, r) => s + r.qtyKg, 0) * 1000) / 1000;

      return {
        ok: true,
        sku: targetSku,
        totalKgNeeded,
        sourceCount: sources.length,
        sources,
        settings: { horizonDays, includeUnconfirmed, includeDrafts, includeCompanions, excludeBeforeDate },
        note: "Each entry is one source order/draft whose BOM expansion produced demand for this RM. If you expected an order to appear here and it isn't, that order does NOT actually use this RM (the BOM tree doesn't expand to it) — do NOT attribute demand to it.",
      };
    }

    case "get_on_hand": {
      const skus = Array.isArray(input.skus) ? input.skus.filter(s => typeof s === "string" && s.trim()).slice(0, 50) : [];
      if (skus.length === 0) return { ok: false, error: "skus array required (1-50 SKUs)" };
      const blob = readData("inventory_onhand");
      if (!blob || !Array.isArray(blob.bySku)) {
        return {
          ok: false,
          error: "On-hand snapshot not available. Admin can sync from the Live Inventory tab, or wait for the 06:30 UTC nightly sync.",
        };
      }
      const map = new Map();
      for (const r of blob.bySku) if (r && r.sku) map.set(r.sku, r);
      const bySku = {};
      for (const sku of skus) {
        const row = map.get(sku);
        if (!row) {
          bySku[sku] = { sku, found: false };
        } else {
          bySku[sku] = {
            sku,
            found: true,
            onHand: row.onHand,
            allocated: row.allocated,
            available: row.available,
            onOrder: row.onOrder,
            inTransit: row.inTransit,
            locations: row.locations || [],
          };
        }
      }
      return {
        ok: true,
        lastSync: blob.lastSync,
        bySku,
        note: "Available = OnHand − Allocated. Use Available for new scheduling decisions. Snapshot is daily; warn the user if they need to know it's not real-time.",
      };
    }

    case "run_mrp": {
      const poHorizonDays      = Math.max(0, Math.min(365, pickNum(input.poHorizonDays, userSettings.poHorizonDays, 30)));
      const horizonDays        = Math.max(1, Math.min(365, pickNum(input.horizonDays, userSettings.horizonDays, 120)));
      const includeUnconfirmed = pickBool(input.includeUnconfirmed, userSettings.includeUnconfirmed, false);
      const includeDrafts      = pickBool(input.includeDrafts, userSettings.includeDrafts, false);
      const includeCompanions  = pickBool(input.includeCompanions, userSettings.includeCompanions, false);
      const excludeBeforeRaw   = pickStr(input.excludeBefore, userSettings.excludeBefore);
      const excludeBeforeDate  = /^\d{4}-\d{2}-\d{2}$/.test(excludeBeforeRaw) ? excludeBeforeRaw : null;
      const topN               = Math.max(1, Math.min(100, isFinite(input.topN) ? input.topN : 20));
      const today              = new Date().toISOString().slice(0, 10);
      const poHorizonEndDate   = poHorizonDays > 0 ? mrpAddDays(today, poHorizonDays) : null;

      const { orders: rawOrders, bomParents, supply, onHandBySku, costsBySku } = getMrpInputs();

      // Mirror the endpoint's pipeline-draft synth so the tool can answer
      // "what if we include drafts?" without divergent logic.
      let mrpOrders = rawOrders;
      let draftsCount = 0;
      if (includeDrafts) {
        const pipelineBlob = readData("vf_pipeline_drafts");
        const drafts = (pipelineBlob && pipelineBlob.drafts) || [];
        const synth = drafts.map(synthPipelineDraftAsOrder).filter(Boolean);
        draftsCount = synth.length;
        mrpOrders = rawOrders.concat(synth);
      }
      let companionsCount = 0;
      if (includeCompanions) {
        const cblob = _companionRulesBlob();
        const synths = generateCompanionOrders(mrpOrders, cblob.rules || []);
        companionsCount = synths.length;
        if (synths.length) mrpOrders = mrpOrders.concat(synths);
      }

      const netFgOnHand = pickBool(input.netFgOnHand, userSettings.netFgOnHand, true);
      const { requirements, skipped, fgNettingSummary, wipNettingSummary } = buildRequirements(mrpOrders, bomParents, {
        today, horizonDays, includeUnconfirmed, applyWastage: true, excludeBeforeDate, supply,
        onHandBySku, netFgOnHand,
      });
      const { suggestedPOs, atRiskOrders } = allocateAndPlan(requirements, onHandBySku, supply, today, poHorizonEndDate);

      // Split + enrich with $$
      let totalDollars = 0, overdueDollars = 0, missingCostCount = 0, deferredDollars = 0;
      const inWindow = [], deferred = [];
      for (const po of suggestedPOs) {
        const c = costsBySku[po.sku];
        const unitCost = c && c.averageCost > 0 ? c.averageCost : null;
        po.unitCost = unitCost;
        po.lineCost = unitCost != null ? unitCost * po.qtyToOrder : null;
        po.costMissing = unitCost == null;
        // Cin7 native UOM (Kg, g, Each, case, etc.) — surfaced so callers
        // don't have to reason about whether qtyToOrder is in kg or grams.
        po.uom = (c && c.uom) || null;
        const isDeferred = poHorizonEndDate && po.mustOrderByDate && po.mustOrderByDate > poHorizonEndDate;
        if (isDeferred) {
          deferred.push(po);
          if (po.lineCost) deferredDollars += po.lineCost;
        } else {
          inWindow.push(po);
          if (po.costMissing) missingCostCount++;
          else {
            totalDollars += po.lineCost;
            if (po.isOverdue) overdueDollars += po.lineCost;
          }
        }
      }

      // Sort by line cost desc (nulls last) and trim to topN
      const byCostDesc = (a, b) => (b.lineCost || -1) - (a.lineCost || -1);
      const trim = po => ({
        sku: po.sku, name: po.name,
        qtyToOrder: po.qtyToOrder, unit: po.unit,
        unitCost: po.unitCost, lineCost: po.lineCost,
        mustOrderByDate: po.mustOrderByDate, earliestNeedDate: po.earliestNeedDate,
        leadTimeDays: po.leadTimeDays, leadTimeSource: po.leadTimeSource,
        isOverdue: po.isOverdue, costMissing: po.costMissing,
        onHand: po.onHand, onOrder: po.onOrder,
      });
      const trimAtRisk = o => ({
        orderId: o.orderId, sku: o.sku, machine: o.machine,
        start: o.start, end: o.end, qty: o.qty, status: o.status,
        shortageSkus: (o.shortages || []).slice(0, 5).map(s => ({ sku: s.sku, shortageKg: s.shortageKg })),
      });

      return {
        ok: true,
        runAt: new Date().toISOString(),
        today,
        settings: { poHorizonDays, horizonDays, includeUnconfirmed, includeDrafts, draftsCount, includeCompanions, companionsCount, excludeBeforeDate, netFgOnHand },
        summary: {
          ordersConsidered: mrpOrders.length - (skipped.unconfirmed + skipped.complete + skipped.noStart + skipped.outsideHorizon + skipped.noBom + skipped.excludedByDate + (skipped.packoutFormulaOnly || 0)),
          requirementCount: requirements.length,
          atRiskOrderCount: atRiskOrders.length,
          suggestedPoCount: inWindow.length,
          deferredPoCount: deferred.length,
          overduePoCount: inWindow.filter(p => p.isOverdue).length,
          totalDollars: Math.round(totalDollars * 100) / 100,
          overdueDollars: Math.round(overdueDollars * 100) / 100,
          deferredDollars: Math.round(deferredDollars * 100) / 100,
          missingCostCount,
        },
        fgNettingSummary: fgNettingSummary || [],
        wipNettingSummary: wipNettingSummary || [],
        packoutSkipped: {
          count: skipped.packoutFormulaOnly || 0,
          examples: packoutSkippedExamples || [],
        },
        suggestedPOs: inWindow.slice().sort(byCostDesc).slice(0, topN).map(trim),
        deferredPOs: deferred.slice().sort(byCostDesc).slice(0, topN).map(trim),
        atRiskOrders: atRiskOrders.slice(0, 30).map(trimAtRisk),
        note: "suggestedPOs = top-N by line cost (highest $ first). fgNettingSummary lists FG SKUs where on-hand stock offset planned production (RM demand is net of that). If a demand number looks lower than expected, check fgNettingSummary — the FG might already be in inventory. packoutSkipped.count is MOs on packout machines (depositing/pouching/mac_packout) that were skipped because their SKU is not FG-* — these are scheduling errors (formula SKU on a packout op) and their BOM demand is intentionally excluded. To trace deeper, call trace_po_demand which shows per-order gross vs net qty.",
      };
    }

    case "trace_lot": {
      const lot = String(input.lot || "").trim();
      if (!lot) return { ok: false, error: "trace_lot requires a lot code (batch #)." };
      const idx = getMovementIndex();
      const rows = (idx.byBatch.get(lot) || []).slice().sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
      if (!rows.length) {
        return { ok: true, lot, count: 0, movements: [], note: "No movements found for this lot code. Double-check the exact case/spacing — Cin7 lot codes sometimes contain spaces or hyphens the user may have mistyped." };
      }
      return {
        ok: true, lot, count: rows.length,
        movements: annotateMovements(rows),
        data_quality: summarizeDataQuality(rows),
        note: "movements sorted oldest first. dq_notes on any row explains a known data-quality caveat for that period — surface those to the user verbatim.",
      };
    }

    case "trace_fg_lineage": {
      const lot = String(input.lot || "").trim();
      if (!lot) return { ok: false, error: "trace_fg_lineage requires a lot code." };
      const maxDepth = Math.max(1, Math.min(10, Number(input.max_depth) || 5));
      const idx = getMovementIndex();
      const posByRef = (readData("purchase_orders") || { byRef: {} }).byRef || {};
      const upstream = buildLineageTree(idx, lot, 0, maxDepth, new Set(), posByRef);
      const downstream = findDownstream(idx, lot);
      const all = collectMovementsFromTree(upstream).concat(downstream);
      return {
        ok: true, lot, max_depth: maxDepth,
        upstream: annotateTreeMovements(upstream),
        downstream: annotateMovements(downstream),
        data_quality: summarizeDataQuality(all),
        note: "upstream.inputs is a recursive tree: each input has its own inputs[] for its BOM components. Terminal leaves have inputs=[] and an origin PO/ST showing where the raw material entered inventory. downstream shows every Out movement of the original lot (shipments, transfers, adjustments, consolidations).",
      };
    }

    case "trace_rm_history": {
      const q = String(input.query || "").trim();
      if (!q) return { ok: false, error: "trace_rm_history requires a lot code or SKU." };
      const isSku = /^(RM|VC|PK|WIP|FG)-/i.test(q);
      const idx = getMovementIndex();
      const rows = ((isSku ? idx.bySku.get(q) : idx.byBatch.get(q)) || []).slice().sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
      let running = 0;
      const timeline = rows.map(mv => {
        running += Number(mv.qty_in || 0) - Number(mv.qty_out || 0);
        return { ...mv, running_balance: running };
      });
      return {
        ok: true, query: q, resolved_via: isSku ? "sku" : "batch",
        count: timeline.length, final_balance: running,
        timeline: annotateMovements(timeline),
        data_quality: summarizeDataQuality(timeline),
        note: "running_balance is cumulative qty_in − qty_out across the timeline. For a single-lot query it should end near zero (all consumed) unless there's residual on hand or a data-quality window inflated it. For a SKU query it spans all lots of that SKU.",
      };
    }

    default:
      return { ok: false, error: `Unknown tool: ${name}` };
  }
}

// Human-readable preview for a mutating tool call. Used by /api/chat to
// describe staged actions to both (a) the AI in its synthetic tool_result
// so it can narrate accurately, and (b) the user via the pending-actions
// UI card. Pulls fresh order state for context-sensitive previews like
// the duplicate detection on add_order.
function previewForTool(name, input) {
  const orders = readData("vf_orders") || [];
  const warnings = [];
  const findOrder = id =>
    orders.find(o => o && (o.orderId === id || o.id === id)) || null;
  const fmt = v => (v == null || v === "") ? "—" : String(v);
  switch (name) {
    case "add_order": {
      const sku = fmt(input.sku);
      const machine = fmt(input.machine);
      const start = fmt(input.start);
      const end = fmt(input.end);
      const qty = input.qty || 0, batches = input.batches || 1;
      const orderId = input.orderId || `TBD-${input.machine || "?"}-${(input.start || "").replace(/-/g, "")}`;
      if (!input.allow_duplicate) {
        const conflict = findAddOrderConflict(orders, input);
        if (conflict) {
          warnings.push(`Potential duplicate: existing order '${conflict.orderId}' (${conflict.sku}, ${conflict.machine}, ${conflict.start}..${conflict.end || conflict.start}, qty ${conflict.qty}) overlaps. The user must confirm a parallel order before this commits.`);
        }
      }
      return {
        label: "Create order",
        summary: `${orderId} · ${sku} · ${machine} · ${start} → ${end} · ${qty} kg × ${batches} batch${batches === 1 ? "" : "es"}`,
        warnings,
      };
    }
    case "update_order_dates": {
      const o = findOrder(input.order_id);
      return {
        label: "Reschedule order",
        summary: o
          ? `${o.orderId} (${o.sku || ""} · ${o.machine || ""}) · ${o.start || "?"} → ${fmt(input.start) || o.start || "?"} … ${o.end || "?"} → ${fmt(input.end) || o.end || "?"}`
          : `Order '${input.order_id}' not found — apply will fail`,
        warnings: o ? [] : [`Order '${input.order_id}' not found in current schedule`],
      };
    }
    case "update_order_quantity": {
      const o = findOrder(input.order_id);
      return {
        label: "Change quantity",
        summary: o
          ? `${o.orderId} (${o.sku || ""}) · qty ${o.qty} → ${fmt(input.qty) || o.qty} · batches ${o.batches} → ${fmt(input.batches) || o.batches}`
          : `Order '${input.order_id}' not found — apply will fail`,
        warnings: o ? [] : [`Order '${input.order_id}' not found`],
      };
    }
    case "update_order_status": {
      const o = findOrder(input.order_id);
      return {
        label: "Change status",
        summary: o
          ? `${o.orderId} · status ${o.status || "?"} → ${fmt(input.status)}`
          : `Order '${input.order_id}' not found — apply will fail`,
        warnings: o ? [] : [`Order '${input.order_id}' not found`],
      };
    }
    case "update_order_metadata": {
      const o = findOrder(input.order_id);
      const changes = Object.entries(input)
        .filter(([k, v]) => k !== "order_id" && v !== undefined && v !== "")
        .map(([k, v]) => `${k}=${fmt(v)}`)
        .join(", ");
      return {
        label: "Edit metadata",
        summary: o
          ? `${o.orderId} (${o.sku || ""}) · ${changes || "(no changes specified)"}`
          : `Order '${input.order_id}' not found — apply will fail`,
        warnings: o ? [] : [`Order '${input.order_id}' not found`],
      };
    }
    case "delete_order": {
      const o = findOrder(input.order_id);
      const w = ["DELETE is permanent — cannot be undone after Apply."];
      if (!o) w.push(`Order '${input.order_id}' not found in current schedule`);
      return {
        label: "Delete order",
        summary: o
          ? `${o.orderId} · ${o.sku || ""} · ${o.machine || ""} · ${o.start || "?"}..${o.end || o.start || "?"}`
          : `Order '${input.order_id}'`,
        warnings: w,
      };
    }
    case "shift_machine_orders": {
      const machine = fmt(input.machine);
      const days = input.days || 0;
      const from = fmt(input.from_date);
      const affected = orders.filter(o =>
        o && o.machine === input.machine && o.status !== "complete"
        && o.start && (!input.from_date || o.start >= input.from_date)
      );
      return {
        label: "Shift machine orders",
        summary: `${machine} · ${affected.length} order${affected.length === 1 ? "" : "s"} from ${from} shifted by ${days >= 0 ? "+" : ""}${days} day${Math.abs(days) === 1 ? "" : "s"}`,
        warnings: affected.length === 0 ? [`No matching orders on ${machine} from ${from} — nothing will move`] : [],
      };
    }
    default:
      return {
        label: name,
        summary: `Input: ${JSON.stringify(input)}`,
        warnings: [],
      };
  }
}

app.post("/api/chat", requireOrderEdit, async (req, res) => {
  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(503).json({ ok: false, error: "AI assistant is not configured (missing ANTHROPIC_API_KEY)" });
  }
  const { messages, mrpSettings } = req.body;
  if (!messages || !Array.isArray(messages)) {
    return res.status(400).json({ ok: false, error: "messages array required" });
  }
  // Tool-layer filter inheritance context (see executeAITool comment).
  // mrpSettings snapshots the user's current MRP-tab inputs so the chat
  // handler can use them as defaults whenever the bot omits a filter.
  const toolContext = {
    mrpSettings: (mrpSettings && typeof mrpSettings === "object") ? mrpSettings : {},
  };

  try {
    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
    let currentMessages = messages.map(m => ({
      role: m.role,
      content: m.content,
    }));
    // Surface the snapshot to the model so it can SEE the user's current
    // settings (and override per-tool if the user explicitly asked for
    // something else). The defaults already kick in at the server layer
    // even if the bot ignores this; this just lets the bot narrate
    // accurately.
    const settingsBanner = buildMrpSettingsBanner(toolContext.mrpSettings);
    const systemWithCtx = settingsBanner ? AI_SYSTEM + "\n\n" + settingsBanner : AI_SYSTEM;
    let response;
    const pendingActions = []; // mutating tool calls queued for user approval

    // Agentic loop — max 8 tool-use rounds
    for (let i = 0; i < 8; i++) {
      response = await client.messages.create({
        model: "claude-opus-4-6",
        max_tokens: 4096,
        system: systemWithCtx,
        tools: AI_TOOLS,
        messages: currentMessages,
      });

      if (response.stop_reason !== "tool_use") break;

      // Execute all tool calls
      const toolUseBlocks = response.content.filter(b => b.type === "tool_use");
      currentMessages.push({ role: "assistant", content: response.content });

      const toolResults = await Promise.all(
        toolUseBlocks.map(async block => {
          // Mutating tools are STAGED, not executed. The user reviews the
          // pending-actions card in the chat UI and clicks Apply to commit.
          // This is the load-bearing change behind the "preview before
          // action" UX — and the line that prevents the bot from silently
          // double-booking production.
          if (MUTATING_AI_TOOLS.has(block.name)) {
            const preview = previewForTool(block.name, block.input);
            const action = {
              id: `pa_${pendingActions.length + 1}_${crypto.randomBytes(3).toString("hex")}`,
              tool: block.name,
              input: block.input,
              label: preview.label,
              summary: preview.summary,
              warnings: preview.warnings,
            };
            pendingActions.push(action);
            return {
              type: "tool_result",
              tool_use_id: block.id,
              content: JSON.stringify({
                status: "queued_for_approval",
                pending_action_id: action.id,
                label: preview.label,
                summary: preview.summary,
                warnings: preview.warnings,
                message: "This action is QUEUED. It has NOT been committed to the schedule. The user sees a preview card and must click Apply to make it take effect. Subsequent get_orders calls will NOT reflect this change until the user applies. Narrate clearly in your reply what you queued and why.",
              }),
            };
          }
          // Read-only tools execute normally — with the user-settings
          // context so MRP filter inheritance kicks in.
          return {
            type: "tool_result",
            tool_use_id: block.id,
            content: JSON.stringify(await executeAITool(block.name, block.input, toolContext)),
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
      // dataChanged stays false here — nothing committed yet. The frontend
      // will refresh orders after the user clicks Apply (via the
      // /api/chat/apply-pending response).
      dataChanged: false,
      mutatedTools: [],
      pendingActions,
    });
  } catch (e) {
    console.error("AI chat error:", e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// POST /api/chat/apply-pending — execute a batch of staged mutating tool
// calls that the user reviewed and approved. Body: { actions: [{tool, input}] }.
// Each action runs in order via the same executeAITool the chat handler used
// to use; per-action results come back so the frontend can show successes /
// failures granularly (e.g. one duplicate-detected refusal among five
// successful adds).
app.post("/api/chat/apply-pending", requireOrderEdit, async (req, res) => {
  const body = req.body || {};
  const actions = Array.isArray(body.actions) ? body.actions : null;
  if (!actions) return res.status(400).json({ ok: false, error: "actions array required" });
  if (!actions.length) return res.json({ ok: true, results: [], appliedCount: 0, failedCount: 0 });

  const results = [];
  let appliedCount = 0;
  let failedCount = 0;
  for (const action of actions) {
    const tool = action && action.tool;
    const input = (action && action.input) || {};
    if (!tool || !MUTATING_AI_TOOLS.has(tool)) {
      results.push({ id: action && action.id, tool, ok: false, error: "tool must be in the mutating-tools whitelist" });
      failedCount++;
      continue;
    }
    try {
      const result = await executeAITool(tool, input);
      const ok = !!(result && result.ok !== false && !result.error);
      results.push({ id: action.id, tool, ok, result });
      if (ok) appliedCount++; else failedCount++;
    } catch (e) {
      results.push({ id: action.id, tool, ok: false, error: e.message });
      failedCount++;
    }
  }
  res.json({ ok: true, results, appliedCount, failedCount });
});

// ── Sales pipeline import + drafts ──────────────────────────────────────────
//
// The team's Voyage Pipeline Review.xlsx tracks customer opportunities; we
// surface a curated subset (Include in MRP = Y) as draft production runs
// the user can selectively promote into vf_orders. Drafts live in their
// own blob so the calendar/MRP aren't polluted until the user explicitly
// commits them.
//
// Quirks worth knowing:
// - The "Quantity in MTs" column is in METRIC TONS for everything except
//   Nut Free Spreads, where the team uses it for CASE COUNT. We preserve
//   the case-count semantics with qtyUnit="cases" instead of force-
//   converting to kg.
// - One channel/SKU mismatch exists in the source data (e.g. Mondelez
//   Q4 row has Channel=Chocolate Liquor but SKU=FG-604-* which is NFS).
//   We surface these as dataIssues[] instead of silently fixing.
// - The BOM expansion gives us informational chain context for chocolate
//   FGs. v1 promote creates the FG packout draft only; users can ask the
//   AI chat bot to add upstream stages with proper scheduling.

const PIPELINE_NFS_CHANNEL = "Nut Free Spreads";
const PIPELINE_NFS_SKU_PREFIX = "FG-604-";

function parsePipelineXlsx(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  if (!wb.SheetNames.length) throw new Error("Workbook has no sheets");
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  // Find the header row by scanning for the distinctive label set
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 60); i++) {
    const fields = rows[i].map(f => String(f || "").toLowerCase().trim());
    if (fields.includes("quarter shipping") && fields.some(f => f.startsWith("include in mrp"))) {
      hdrIdx = i; break;
    }
  }
  if (hdrIdx === -1) throw new Error("Could not find header row (expected 'Quarter Shipping' + 'Include in MRP?' on the same row)");

  // Map column names to indices
  const hdr = rows[hdrIdx].map(h => String(h || "").toLowerCase().trim());
  const col = {};
  hdr.forEach((h, i) => {
    if (h === "quarter shipping") col.quarter = i;
    else if (h === "month") col.month = i;
    else if (h === "customer") col.customer = i;
    else if (h === "channel") col.channel = i;
    else if (h.startsWith("quantity in")) col.qty = i;
    else if (h === "confidence level") col.confidence = i;
    else if (h === "sku") col.sku = i;
    else if (h.startsWith("include in mrp")) col.includeInMrp = i;
    else if (h === "comments") col.comments = i;
  });
  for (const k of ["quarter", "month", "customer", "channel", "qty", "sku", "includeInMrp"]) {
    if (col[k] === undefined) throw new Error(`Pipeline header is missing required column: ${k}`);
  }

  const dateToISO = d => {
    if (!d) return null;
    if (d instanceof Date) return d.toISOString().slice(0, 10);
    const s = String(d);
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    return null;
  };
  const numOf = v => {
    if (v == null || v === "") return null;
    const n = parseFloat(String(v).replace(/,/g, ""));
    return isFinite(n) ? n : null;
  };

  const opportunities = [];
  let mrpYesCount = 0, mrpNoCount = 0, blankRowsSkipped = 0;
  for (let i = hdrIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r.length) { blankRowsSkipped++; continue; }
    const customer = String(r[col.customer] || "").trim();
    const sku = String(r[col.sku] || "").trim();
    if (!customer && !sku) { blankRowsSkipped++; continue; }
    const flag = String(r[col.includeInMrp] || "").trim().toUpperCase();
    if (flag === "Y") mrpYesCount++; else if (flag === "N") { mrpNoCount++; continue; } else { blankRowsSkipped++; continue; }

    const channel = String(r[col.channel] || "").trim();
    const qtyValue = numOf(r[col.qty]);
    const confidence = numOf(r[col.confidence]);
    const shipMonth = dateToISO(r[col.month]);
    const quarter = String(r[col.quarter] || "").trim();
    const comments = String(r[col.comments] || "").trim();

    // Unit handling: Nut Free Spreads is case-count even though the
    // header says MTs. Everything else converts MTs → kg.
    const isNfs = channel === PIPELINE_NFS_CHANNEL;
    const qtyUnit = isNfs ? "cases" : "kg";
    const qtyInUnit = isNfs ? (qtyValue || 0) : ((qtyValue || 0) * 1000);

    // Surface known data-quality issues so the user can decide rather
    // than silently fixing them in import.
    const dataIssues = [];
    if (!shipMonth) dataIssues.push("Missing or unparseable Month — production timing can't be inferred");
    if (!sku) dataIssues.push("SKU column is blank — promote will fail until you set one");
    if (channel !== PIPELINE_NFS_CHANNEL && sku.startsWith(PIPELINE_NFS_SKU_PREFIX)) {
      dataIssues.push(`Channel '${channel}' contradicts SKU prefix '${PIPELINE_NFS_SKU_PREFIX}' (looks like a Nut Free Spread)`);
    }

    opportunities.push({
      pipelineRow: i + 1,   // 1-indexed for human reference
      quarter, shipMonth,
      customer, channel,
      qtyRaw: qtyValue, qtyValue: qtyInUnit, qtyUnit,
      confidence,
      fgSKU: sku,
      comments,
      dataIssues,
    });
  }
  return {
    opportunities,
    mrpYesCount,
    mrpNoCount,
    blankRowsSkipped,
    headerRowIndex: hdrIdx,
  };
}

// Pipeline drafts only have a "Ship Month" — no production start. If we
// treat ship month as production start, MRP dates raw-material must-order-by
// off the ship date and understates urgency (chocolate ordered a month late,
// coffee two weeks late, etc.). Ballpark production lead per channel below
// pushes the synth order's start N days before ship month so RM back-dating
// lands in the right zip code. Tune values here as capacity/routing knowledge
// improves; per-SKU override isn't exposed yet.
const PIPELINE_PRODUCTION_LEAD_DAYS = {
  "coffee": 30,
  "nut free spreads": 14,
  "chocolate liquor": 14,
  "finished chocolate": 30,
  default: 30,
};

function pipelineProductionLeadDays(channel) {
  const c = (channel || "").toLowerCase();
  for (const key of Object.keys(PIPELINE_PRODUCTION_LEAD_DAYS)) {
    if (key !== "default" && c.startsWith(key)) return PIPELINE_PRODUCTION_LEAD_DAYS[key];
  }
  return PIPELINE_PRODUCTION_LEAD_DAYS.default;
}

// Convert a pipeline draft into a synth order the MRP can consume. Returns
// null if the draft can't drive demand (missing fgSKU or shipMonth). Shifts
// start/end back by the channel's production lead so MRP dates RM POs off
// the *production* start rather than ship month; due keeps ship month so any
// customer-facing view still shows the promised date.
function synthPipelineDraftAsOrder(d) {
  if (!d || !d.fgSKU || !d.shipMonth) return null;
  const channel = (d.channel || "").toLowerCase();
  let machine = "mac_packout";
  if (channel.startsWith("coffee")) machine = "grinder";
  else if (channel.startsWith("nut free")) machine = "pouching";
  else if (channel.startsWith("chocolate liquor") || channel.startsWith("finished chocolate")) machine = "depositing";
  const leadDays = pipelineProductionLeadDays(channel);
  const productionStart = mrpAddDays(d.shipMonth, -leadDays) || d.shipMonth;
  return {
    id: "draft_" + d.id,
    orderId: "PIPELINE-" + (d.customer || "?") + "-" + ((d.shipMonth || "").slice(0, 7) || d.quarter || ""),
    sku: d.fgSKU,
    machine,
    start: productionStart,
    end: productionStart,
    due: d.shipMonth,
    qty: d.qtyValue || 0,
    batches: 1,
    total: d.qtyValue || 0,
    status: "queued",
    confirmed: false,
    __fromPipelineDraft: true,
    __shipMonth: d.shipMonth,
    __productionLeadDays: leadDays,
  };
}

// Build pipeline drafts from parsed opportunities, attaching the BOM-
// expanded intermediates as informational chain context (we don't auto-
// create upstream draft orders in v1 — promote creates the FG order
// only; the user uses the AI bot for chain scheduling).
function buildPipelineDrafts(opportunities) {
  const bomBlob = readData("vf_boms");
  const bomParents = (bomBlob && bomBlob.parents) || {};
  return opportunities.map(o => {
    let chain = [];
    let chainStatus = "no_bom";
    let chainError = null;
    if (bomParents[o.fgSKU]) {
      try {
        // Use 1 unit / 1 kg as the basis for expansion — we just want the
        // STRUCTURE of upstream stages, not absolute quantities. The user
        // can scale during promote.
        const result = expandBom(bomParents, o.fgSKU, o.qtyValue || 1, { applyWastage: true });
        // Intermediates = recursion trail. Dedupe by sku, keep deepest qty.
        const bySku = new Map();
        for (const step of (result.intermediates || [])) {
          if (step.sku === o.fgSKU) continue;
          if (!bySku.has(step.sku)) bySku.set(step.sku, { sku: step.sku, qty: 0, depth: step.depth });
          const entry = bySku.get(step.sku);
          entry.qty += step.qty;
          entry.depth = Math.max(entry.depth, step.depth);
        }
        chain = Array.from(bySku.values()).sort((a, b) => b.depth - a.depth || a.sku.localeCompare(b.sku));
        chainStatus = chain.length > 0 ? "expanded" : "no_upstream";
      } catch (e) {
        chainStatus = "error";
        chainError = e.message;
      }
    }
    return {
      id: "pd_" + crypto.randomBytes(6).toString("hex"),
      pipelineRow: o.pipelineRow,
      customer: o.customer,
      channel: o.channel,
      quarter: o.quarter,
      shipMonth: o.shipMonth,
      confidence: o.confidence,
      qtyValue: o.qtyValue,
      qtyRaw: o.qtyRaw,
      qtyUnit: o.qtyUnit,
      fgSKU: o.fgSKU,
      comments: o.comments,
      dataIssues: o.dataIssues,
      chain,
      chainStatus,
      chainError,
    };
  });
}

// POST /api/pipeline/import — upload an XLSX, parse it, write drafts blob.
// Admin only — pipeline data is commercially sensitive (customer names,
// won/lost deal sizes, confidence levels) and shouldn't be exposed to
// operators/planners.
app.post("/api/pipeline/import", requireAdmin, upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, error: "No file uploaded (multipart field name: 'file')" });
  let parsed;
  try { parsed = parsePipelineXlsx(req.file.buffer); }
  catch (e) { return res.status(400).json({ ok: false, error: "Parse failed: " + e.message }); }
  const drafts = buildPipelineDrafts(parsed.opportunities);
  const users = readData("vf_users") || [];
  const u = users.find(x => x && x.id === req.userId);
  const blob = {
    lastImport: new Date().toISOString(),
    source: {
      filename: req.file.originalname || null,
      importedBy: u ? u.username : "(unknown)",
      mrpYesCount: parsed.mrpYesCount,
      mrpNoCount: parsed.mrpNoCount,
      blankRowsSkipped: parsed.blankRowsSkipped,
      headerRowIndex: parsed.headerRowIndex,
    },
    drafts,
  };
  writeData("vf_pipeline_drafts", blob);
  res.json({
    ok: true,
    draftCount: drafts.length,
    mrpYesCount: parsed.mrpYesCount,
    mrpNoCount: parsed.mrpNoCount,
    issuesCount: drafts.filter(d => d.dataIssues && d.dataIssues.length).length,
  });
});

// GET /api/pipeline/drafts — return the current drafts blob. Admin only.
app.get("/api/pipeline/drafts", requireAdmin, (req, res) => {
  const blob = readData("vf_pipeline_drafts");
  if (!blob) return res.json({ ok: true, lastImport: null, source: null, drafts: [] });
  res.json({ ok: true, ...blob });
});

// POST /api/pipeline/drafts/promote — body { draftIds: [..] }
// For each draft, create an FG packout order in vf_orders. Runs the same
// duplicate guard as the AI add_order tool. Returns per-draft results so
// the UI can flag partial failures clearly. Successful promotes are
// removed from the drafts blob.
app.post("/api/pipeline/drafts/promote", requireAdmin, (req, res) => {
  const body = req.body || {};
  const draftIds = Array.isArray(body.draftIds) ? body.draftIds : null;
  if (!draftIds || !draftIds.length) return res.status(400).json({ ok: false, error: "draftIds array required" });
  const blob = readData("vf_pipeline_drafts");
  if (!blob || !Array.isArray(blob.drafts)) return res.status(400).json({ ok: false, error: "No drafts to promote" });
  const allowDup = !!body.allow_duplicate;

  const orders = readData("vf_orders") || [];
  const results = [];
  const keepIds = new Set(blob.drafts.map(d => d.id));
  for (const draftId of draftIds) {
    const d = blob.drafts.find(x => x && x.id === draftId);
    if (!d) { results.push({ id: draftId, ok: false, error: "draft not found" }); continue; }
    if (!d.fgSKU) { results.push({ id: draftId, ok: false, error: "draft has no SKU" }); continue; }
    if (!d.shipMonth) { results.push({ id: draftId, ok: false, error: "draft has no ship month" }); continue; }

    // Default scheduling: place start/end on the ship month's first day.
    // We're not trying to do smart backward-planning here — that's the
    // AI bot's job after the user promotes. Conservative: just create
    // the FG row so it shows up on the calendar.
    const start = d.shipMonth;
    const end = d.shipMonth;
    // Channel → machine guess. Falls back to mac_packout for chocolate
    // packout, grinder for coffee. Nut Free Spreads guess: pouching.
    const channel = (d.channel || "").toLowerCase();
    let machine = "mac_packout";
    if (channel.startsWith("coffee")) machine = "grinder";
    else if (channel.startsWith("nut free")) machine = "pouching";
    else if (channel.startsWith("chocolate liquor") || channel.startsWith("finished chocolate")) machine = "depositing";

    const conflict = allowDup ? null : findAddOrderConflict(orders, {
      sku: d.fgSKU, machine, start, end,
    });
    if (conflict) {
      results.push({
        id: draftId, ok: false, duplicate: true,
        error: `Existing order '${conflict.orderId}' (${conflict.sku}, ${conflict.machine}, ${conflict.start}..${conflict.end || conflict.start}) overlaps. Re-promote with allow_duplicate=true if intended.`,
        existing: {
          orderId: conflict.orderId, sku: conflict.sku, machine: conflict.machine,
          start: conflict.start, end: conflict.end, qty: conflict.qty,
        },
      });
      continue;
    }

    // Derive cat/sub from machine; chocolate liquor stays liquor.
    const inferred = inferCatSubFromMachine(machine);
    let cat = inferred.cat, sub = inferred.sub;
    if (channel.startsWith("chocolate liquor")) { cat = "liquor"; sub = "liquor"; }

    const id = `pl-${Date.now()}-${crypto.randomBytes(2).toString("hex")}`;
    const orderId = `TBD-${machine}-${(start || "").replace(/-/g, "")}`;
    const newOrder = {
      id, orderId, sku: d.fgSKU, machine,
      start, end, due: end,
      qty: d.qtyValue || 0, batches: 1, total: d.qtyValue || 0,
      status: "queued", priority: "med",
      notes: `From pipeline: ${d.customer} (${d.quarter}) · conf ${d.confidence != null ? Math.round(d.confidence * 100) + "%" : "?"}${d.comments ? " · " + d.comments : ""}`,
      cat, sub, region: "", temper: "",
      confirmed: false,
      pipelineSource: { customer: d.customer, channel: d.channel, shipMonth: d.shipMonth, confidence: d.confidence, pipelineRow: d.pipelineRow },
    };
    orders.push(newOrder);
    results.push({ id: draftId, ok: true, orderId, machine, qty: newOrder.qty, qtyUnit: d.qtyUnit });
    keepIds.delete(draftId);
  }
  writeData("vf_orders", orders);
  // Strip promoted drafts from the blob
  blob.drafts = blob.drafts.filter(d => keepIds.has(d.id));
  writeData("vf_pipeline_drafts", blob);
  res.json({ ok: true, results, remainingDrafts: blob.drafts.length });
});

// DELETE /api/pipeline/drafts/:id — discard one draft.
app.delete("/api/pipeline/drafts/:id", requireAdmin, (req, res) => {
  const blob = readData("vf_pipeline_drafts");
  if (!blob || !Array.isArray(blob.drafts)) return res.status(404).json({ ok: false, error: "No drafts" });
  const before = blob.drafts.length;
  blob.drafts = blob.drafts.filter(d => d && d.id !== req.params.id);
  if (blob.drafts.length === before) return res.status(404).json({ ok: false, error: "draft not found" });
  writeData("vf_pipeline_drafts", blob);
  res.json({ ok: true, remainingDrafts: blob.drafts.length });
});

// DELETE /api/pipeline/drafts — clear everything.
app.delete("/api/pipeline/drafts", requireAdmin, (req, res) => {
  writeData("vf_pipeline_drafts", { lastImport: null, source: null, drafts: [] });
  res.json({ ok: true });
});

// ── Companion demand rules ──────────────────────────────────────────────────
//
// Some Voyage products drive downstream demand that isn't physically in
// their BOM. Chocolate liquor is the canonical case: when we ship a liquor
// SKU by ocean, the customer usually needs flavor packs to arrive around
// the same time (airfreight). The flavor is a separate FG with its own BOM;
// it isn't a component of the liquor. Historically MRP had no way to see
// this correlation, so flavor RM procurement was manual.
//
// A companion rule pairs a driver SKU with a companion SKU + a per-unit
// ratio ("0.05 kg flavor per 1 kg liquor"). When MRP runs with the
// includeCompanions toggle on, every in-scope driver order synthesizes a
// companion order at the same need-date. The synthetic order flows through
// the normal MRP pipeline — expands the companion's BOM, nets FG on-hand,
// generates RM PO suggestions — with an __fromCompanionRule flag so the UI
// can label it and buildRequirements can bypass the confirmed filter.
//
// Bounded scope for v1: single-level (a companion rule doesn't chain from
// another companion), same need-date as the driver (no lead offset — the
// companion's own production lead time handles the back-scheduling).

function _companionRulesBlob() {
  return readData("vf_companion_rules") || { lastUpdated: null, rules: [] };
}

function _newCompanionRuleId() {
  return "cr_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// GET /api/companion-rules — any authed user (needed to render MRP UI count)
app.get("/api/companion-rules", (req, res) => {
  const blob = _companionRulesBlob();
  res.json({ ok: true, ...blob });
});

// POST /api/companion-rules — admin-only. Body is a single rule; upserts by id.
app.post("/api/companion-rules", requireAdmin, (req, res) => {
  const body = req.body || {};
  const driverSku = String(body.driverSku || "").trim();
  const companionSku = String(body.companionSku || "").trim();
  const qtyPerDriver = Number(body.qtyPerDriver);
  if (!driverSku || !companionSku) {
    return res.status(400).json({ ok: false, error: "driverSku and companionSku required" });
  }
  if (driverSku === companionSku) {
    return res.status(400).json({ ok: false, error: "driver and companion cannot be the same SKU" });
  }
  if (!isFinite(qtyPerDriver) || qtyPerDriver <= 0) {
    return res.status(400).json({ ok: false, error: "qtyPerDriver must be a positive number" });
  }
  const blob = _companionRulesBlob();
  const now = new Date().toISOString();
  const existing = body.id ? blob.rules.find(r => r.id === body.id) : null;
  if (existing) {
    Object.assign(existing, {
      driverSku,
      driverName: String(body.driverName || existing.driverName || "").trim(),
      companionSku,
      companionName: String(body.companionName || existing.companionName || "").trim(),
      qtyPerDriver,
      unit: String(body.unit || existing.unit || "kg").trim(),
      note: String(body.note || "").trim(),
      active: body.active === false ? false : true,
      updatedAt: now,
    });
  } else {
    blob.rules.push({
      id: _newCompanionRuleId(),
      driverSku,
      driverName: String(body.driverName || "").trim(),
      companionSku,
      companionName: String(body.companionName || "").trim(),
      qtyPerDriver,
      unit: String(body.unit || "kg").trim(),
      note: String(body.note || "").trim(),
      active: body.active === false ? false : true,
      createdAt: now,
      updatedAt: now,
    });
  }
  blob.lastUpdated = now;
  writeData("vf_companion_rules", blob);
  res.json({ ok: true, ...blob });
});

// DELETE /api/companion-rules/:id — admin-only
app.delete("/api/companion-rules/:id", requireAdmin, (req, res) => {
  const blob = _companionRulesBlob();
  const before = blob.rules.length;
  blob.rules = blob.rules.filter(r => r.id !== req.params.id);
  if (blob.rules.length === before) return res.status(404).json({ ok: false, error: "rule not found" });
  blob.lastUpdated = new Date().toISOString();
  writeData("vf_companion_rules", blob);
  res.json({ ok: true, ...blob });
});

// Synthesize one companion order per (in-scope driver order, active matching
// rule). Returns [] when there are no rules or no matches. Bounded to a
// single level — synthetic orders themselves are skipped so a companion
// can't recursively trigger another companion.
function generateCompanionOrders(baseOrders, rules) {
  const active = (rules || []).filter(r => r && r.active !== false);
  if (!active.length) return [];
  const byDriver = {};
  for (const r of active) {
    if (!byDriver[r.driverSku]) byDriver[r.driverSku] = [];
    byDriver[r.driverSku].push(r);
  }
  const out = [];
  for (const o of baseOrders) {
    if (!o || !o.sku) continue;
    if (o.__fromCompanionRule) continue; // no chaining
    const matches = byDriver[o.sku];
    if (!matches) continue;
    for (const rule of matches) {
      const driverQty = Number(o.qty) || Number(o.total) || 0;
      const companionQty = driverQty * Number(rule.qtyPerDriver);
      if (!(companionQty > 0)) continue;
      out.push({
        id: "companion_" + o.id + "_" + rule.id,
        orderId: "COMPANION-" + (o.orderId || o.id) + "-" + rule.companionSku,
        sku: rule.companionSku,
        machine: null, // let BOM expansion determine
        start: o.start,
        end: o.end,
        due: o.due,
        qty: companionQty,
        batches: 1,
        total: companionQty,
        status: "queued",
        confirmed: false,
        __fromCompanionRule: true,
        __driverOrderId: o.id,
        __driverOrderRef: o.orderId,
        __driverSku: o.sku,
        __driverQty: driverQty,
        __ruleId: rule.id,
      });
    }
  }
  return out;
}

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
      // Cin7's native UOM (Kg, g, Each, case, etc.). Used by MRP to render
      // "500 g" vs "50 kg" alongside qtyToOrder so operators don't misread
      // grams-UOM SKUs like VC-* flavor concentrates as kg — a silent 1000×
      // ordering hazard we hit when auto-syncing BOMs from Cin7.
      uom: p.UOM || "",
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

// ── Purchase-order supplier sync ────────────────────────────────────────────
//
// Cin7's v2 /purchase endpoint returns every PO with the ACTUAL supplier.
// We index by the PO reference number (the same "PO-00040" format that shows
// up in every inventory-movement row's Reference field) so the traceability
// lineage view can attribute each terminal RM/VC/PK lot to its real
// supplier — not the SKU's default PreferredSupplier (which is stale at
// Voyage and unsafe for audit narration).
//
// Auditors want "who did this lot come from?" — and the honest answer lives
// on the PO, not on the product master.

async function fetchCin7PurchaseOrders() {
  if (!process.env.CIN7_ACCOUNT_ID || !process.env.CIN7_APPLICATION_KEY) {
    throw new Error("CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY environment variable not set");
  }
  const all = [];
  let page = 1;
  const limit = 1000;
  while (true) {
    const url = `https://inventory.dearsystems.com/ExternalApi/v2/purchaseList?Page=${page}&Limit=${limit}`;
    const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
    const ct = resp.headers.get("content-type") || "";
    const text = await resp.text().catch(() => "");
    if (!resp.ok) throw new Error(`Cin7 purchase list ${resp.status} [${ct}]: ${text.slice(0, 300)}`);
    let data;
    try { data = JSON.parse(text); }
    catch (_) { throw new Error(`Cin7 purchase list ${resp.status} non-JSON [${ct}]: ${text.slice(0, 300)}`); }
    // Cin7's purchaseList response uses PurchaseList[]; guard against schema drift.
    const batch = data.PurchaseList || data.Purchases || data.Purchase || [];
    all.push(...batch);
    if (batch.length < limit) break;
    page++;
    if (page > 30) throw new Error("Cin7 purchase pagination exceeded 30 pages — aborting");
  }
  return all;
}

async function performCin7PurchaseOrdersSync() {
  const purchases = await fetchCin7PurchaseOrders();
  const now = new Date().toISOString();
  const byRef = {};
  let withSupplier = 0;
  for (const p of purchases) {
    // Cin7's purchaseList row is a summary — the fields we need are all
    // present here (no need to hit /purchase/{id} for each one, which
    // would be N+1 hell for a nightly job).
    const ref = String(p.OrderNumber || p.CombinedRef || p.ID || "").trim();
    if (!ref) continue;
    const supplier = String(p.Supplier || p.SupplierName || "").trim();
    byRef[ref] = {
      ref,
      supplier: supplier || null,
      supplierCode: String(p.SupplierReference || p.SupplierCode || "").trim() || null,
      orderDate: p.OrderDate ? String(p.OrderDate).slice(0, 10) : null,
      status: p.Status || null,
      total: p.Total != null ? Number(p.Total) : null,
      currency: p.CurrencyCode || null,
    };
    if (supplier) withSupplier++;
  }
  const blob = {
    lastSync: now,
    purchaseCount: purchases.length,
    withSupplierCount: withSupplier,
    byRef,
  };
  writeData("purchase_orders", blob);
  return { ok: true, lastSync: now, purchaseCount: purchases.length, withSupplierCount: withSupplier };
}

// POST /api/cin7/purchase-orders/sync — admin-triggered live pull
app.post("/api/cin7/purchase-orders/sync", requireAdmin, async (req, res) => {
  try {
    const status = await performCin7PurchaseOrdersSync();
    res.json(status);
  } catch (e) {
    console.error("[Cin7 POs] Sync error:", e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// GET /api/cin7/purchase-orders — return the cached PO map (any authed user)
app.get("/api/cin7/purchase-orders", (req, res) => {
  const blob = readData("purchase_orders");
  if (!blob) return res.json({ ok: true, lastSync: null, byRef: {} });
  res.json({ ok: true, ...blob });
});

// ── Sales-order customer sync ────────────────────────────────────────────────
//
// Mirror of the purchase-order sync but for outbound: pulls every SO from
// Cin7's /saleList and indexes by SO reference number so the forward-trace
// view can attribute FG-lot shipments to the actual customer. Same
// "surface real data or nothing" audit stance — no defaults, no guesses.
//
// Auditor use case: "which customer received lot X, and when did it ship?"
// The trace already knows the SO reference from the movement row; this just
// hydrates the customer name.

async function fetchCin7SalesOrders() {
  if (!process.env.CIN7_ACCOUNT_ID || !process.env.CIN7_APPLICATION_KEY) {
    throw new Error("CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY environment variable not set");
  }
  const all = [];
  let page = 1;
  const limit = 1000;
  while (true) {
    const url = `https://inventory.dearsystems.com/ExternalApi/v2/saleList?Page=${page}&Limit=${limit}`;
    const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
    const ct = resp.headers.get("content-type") || "";
    const text = await resp.text().catch(() => "");
    if (!resp.ok) throw new Error(`Cin7 sale list ${resp.status} [${ct}]: ${text.slice(0, 300)}`);
    let data;
    try { data = JSON.parse(text); }
    catch (_) { throw new Error(`Cin7 sale list ${resp.status} non-JSON [${ct}]: ${text.slice(0, 300)}`); }
    // Cin7's saleList response uses SaleList[]; guard against schema drift.
    const batch = data.SaleList || data.Sales || data.Sale || [];
    all.push(...batch);
    if (batch.length < limit) break;
    page++;
    if (page > 30) throw new Error("Cin7 sale pagination exceeded 30 pages — aborting");
  }
  return all;
}

async function performCin7SalesOrdersSync() {
  const sales = await fetchCin7SalesOrders();
  const now = new Date().toISOString();
  const byRef = {};
  let withCustomer = 0;
  for (const s of sales) {
    // Cin7's saleList row includes OrderNumber (the "SO-XXXXX" that appears
    // in inventory movements) and Customer (name). ShipDate / OrderDate are
    // both present as separate fields.
    const ref = String(s.OrderNumber || s.CombinedRef || s.ID || "").trim();
    if (!ref) continue;
    const customer = String(s.Customer || s.CustomerName || "").trim();
    byRef[ref] = {
      ref,
      customer: customer || null,
      customerReference: String(s.CustomerReference || "").trim() || null,
      orderDate: s.OrderDate ? String(s.OrderDate).slice(0, 10) : null,
      shipDate: s.ShipDate ? String(s.ShipDate).slice(0, 10) : null,
      status: s.Status || null,
      total: s.Total != null ? Number(s.Total) : null,
      currency: s.CurrencyCode || null,
    };
    if (customer) withCustomer++;
  }
  const blob = {
    lastSync: now,
    saleCount: sales.length,
    withCustomerCount: withCustomer,
    byRef,
  };
  writeData("sales_orders", blob);
  return { ok: true, lastSync: now, saleCount: sales.length, withCustomerCount: withCustomer };
}

// POST /api/cin7/sales-orders/sync — admin-triggered live pull
app.post("/api/cin7/sales-orders/sync", requireAdmin, async (req, res) => {
  try {
    const status = await performCin7SalesOrdersSync();
    res.json(status);
  } catch (e) {
    console.error("[Cin7 SOs] Sync error:", e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// GET /api/cin7/sales-orders — return the cached SO map (any authed user)
app.get("/api/cin7/sales-orders", (req, res) => {
  const blob = readData("sales_orders");
  if (!blob) return res.json({ ok: true, lastSync: null, byRef: {} });
  res.json({ ok: true, ...blob });
});

// ── Yield bucket config (powers the Yield + Yield Setup tabs) ───────────────
//
// Each completed Production Run is bucketed by its first Output SKU. The
// yield_buckets blob maps SKUs → buckets (e.g., "Seed cleaning", "FG Packout").
// Defaults are seeded from the SKU list Matt shared on 2026-05-29; admins can
// add/remove SKUs via the Yield Setup tab as new products get introduced.
//
// Shape:
//   {
//     lastUpdated: ISO timestamp,
//     buckets: [
//       { id, line, name, order, skus: [...] },
//       ...
//     ]
//   }
//
// WIP1 + WIP2 SKUs are intentionally combined under one bucket ("Liquor → FG
// ready") per Matt's note — they're systemically separate but conceptually
// one unit op for yield purposes.

const DEFAULT_YIELD_BUCKETS = [
  {
    id: "choc-seed-cleaning",
    line: "Chocolate",
    name: "Seed cleaning",
    order: 1,
    skus: ["WIP-5100008"],
  },
  {
    id: "choc-roasted-seeds",
    line: "Chocolate",
    name: "Roasted seeds",
    order: 2,
    skus: ["WIP-5100007"],
  },
  {
    id: "choc-making-liquor",
    line: "Chocolate",
    name: "Making liquor",
    order: 3,
    skus: [
      "WIP-5100011-US", "WIP-5100012-US", "WIP-5100013",
      "WIP-5100042-EU", "WIP-5100043-EU",
      "WIP-5100046", "WIP-5100047", "WIP-5100048", "WIP-5100049", "WIP-5100050",
    ],
  },
  {
    id: "choc-liquor-to-fg-ready",
    line: "Chocolate",
    name: "Liquor → FG ready",
    order: 4,
    skus: [
      // WIP1 (Filling stage)
      "WIP1-5100062-EU", "WIP1-5100068-EU", "WIP1-5100810-US", "WIP1-5100815-US",
      "WIP1-5100820-EU", "WIP1-5100858-US", "WIP1-5100859-US", "WIP1-5100860-US",
      "WIP1-5100862-EU", "WIP1-5100865-US", "WIP1-5100880-EU", "WIP1-5100885-EU",
      // WIP2 (Conching stage) — combined with WIP1 per Matt's note
      "WIP2-5100062-EU", "WIP2-5100068-EU", "WIP2-5100810-US", "WIP2-5100815-US",
      "WIP2-5100820-EU", "WIP2-5100858-US", "WIP2-5100859-US", "WIP2-5100860-US",
      "WIP2-5100862-EU", "WIP2-5100865-US", "WIP2-5100880-EU", "WIP2-5100885-EU",
      // Plain WIP variants (mid-WIP1/WIP2 outputs that exist as single SKUs)
      "WIP-5100811-EU", "WIP-5100813-US", "WIP-5100814-EU", "WIP-5100820-US",
      "WIP-5100861-EU", "WIP-5100863-US", "WIP-5100864-EU", "WIP-5100866-EU",
      "WIP-5100868-US", "WIP-5100869-EU", "WIP-5100880-US", "WIP-5100885-US",
    ],
  },
  {
    id: "choc-fg-packout",
    line: "Chocolate",
    name: "FG Packout",
    order: 5,
    skus: [
      // 800/850/870 inclusion families
      "FG-800-001-00", "FG-800-002-00",
      "FG-850-051-00", "FG-850-053-00", "FG-850-056-00", "FG-850-057-00",
      "FG-870-070-00", "FG-870-071-00",
      // 860 liquor packout
      "FG-860-005-01-EU", "FG-860-005-01-EU (copy of 25kg)", "FG-860-005-01-EU-kg",
      "FG-860-005-02-EU-kg",
      "FG-860-006-00", "FG-860-006-01-EU", "FG-860-006-01-EU-kg", "FG-860-006-02",
      // 860/880 powder packout
      "FG-860-008-00-EU", "FG-860-008-00-EU-kg", "FG-880-000-00",
      // 888-* coating/inclusion FG
      "FG-888-810-00-US", "FG-888-810-00-US/EU-kg",
      "FG-888-811-00-EU", "FG-888-811-00-EU-kg",
      "FG-888-812-00-EU", "FG-888-812-00-EU-kg",
      "FG-888-813-00-US", "FG-888-813-00-US-kg",
      "FG-888-814-00-EU", "FG-888-814-00-EU-kg",
      "FG-888-815-00-US", "FG-888-815-00-US-kg",
      "FG-888-820-00-EU", "FG-888-820-00-US", "FG-888-820-00-US-kg",
      "FG-888-858-00-US", "FG-888-858-00-US-kg",
      "FG-888-859-00-US", "FG-888-859-00-US-kg",
      "FG-888-860-00-US", "FG-888-860-00-US/EU-kg",
      "FG-888-861-00-EU", "FG-888-861-00-EU-kg",
      "FG-888-862-00-EU", "FG-888-862-00-EU-kg",
      "FG-888-863-00-US", "FG-888-863-00-US-kg",
      "FG-888-864-00-EU", "FG-888-864-00-EU-kg",
      "FG-888-865-00-US", "FG-888-865-00-US/EU-kg",
      "FG-888-866-00-EU", "FG-888-866-00-EU-kg",
      "FG-888-867-00-EU", "FG-888-867-00-kg",
      "FG-888-868-00-US", "FG-888-868-00-US-kg",
      "FG-888-869-00-EU", "FG-888-869-00-EU-kg",
      "FG-888-880-00-EU", "FG-888-880-00-US", "FG-888-880-00-US-kg",
      "FG-888-885-00-EU", "FG-888-885-00-US", "FG-888-885-00-US-kg",
    ],
  },
  // Coffee — split per Matt's pick: WIP-5100002 → Roasting, FG-999-* → Packout
  {
    id: "coffee-roasting",
    line: "Coffee",
    name: "Roasting",
    order: 1,
    skus: ["WIP-5100002", "WIP-5100002-TRIAL"],
  },
  {
    id: "coffee-packout",
    line: "Coffee",
    name: "Packout",
    order: 2,
    skus: [
      "FG-999-002-00", "FG-999-002-00-Kg",
      "FG-999-003-00", "FG-999-003-00-Kg",
      "FG-999-010-00", "FG-999-100-00",
    ],
  },
];

function seedYieldBucketsIfMissing() {
  const existing = readData("vf_yield_buckets");
  if (existing && Array.isArray(existing.buckets)) return existing;
  const blob = {
    lastUpdated: new Date().toISOString(),
    seeded: true,
    buckets: DEFAULT_YIELD_BUCKETS,
  };
  writeData("vf_yield_buckets", blob);
  return blob;
}
seedYieldBucketsIfMissing();

// Build a SKU → bucket lookup. Last-writer-wins on duplicates so admins
// can copy/paste the same SKU between buckets while iterating without
// silent splits in the data.
function buildSkuBucketMap(buckets) {
  const map = new Map();
  for (const b of buckets) {
    for (const sku of (b.skus || [])) {
      if (!sku) continue;
      map.set(String(sku).trim(), b);
    }
  }
  return map;
}

// GET /api/yield/buckets — current bucket config (any authed user)
app.get("/api/yield/buckets", (req, res) => {
  const blob = readData("vf_yield_buckets") || seedYieldBucketsIfMissing();
  res.json({ ok: true, ...blob });
});

// PUT /api/yield/buckets — admin only. Body: { buckets: [...] }
// Validates shape and dedupes SKUs within each bucket before persisting.
app.put("/api/yield/buckets", requireAdmin, (req, res) => {
  try {
    const incoming = (req.body && Array.isArray(req.body.buckets)) ? req.body.buckets : null;
    if (!incoming) return res.status(400).json({ ok: false, error: "Body must include `buckets` array" });
    const cleaned = incoming.map(b => ({
      id: String(b.id || "").trim(),
      line: String(b.line || "").trim(),
      name: String(b.name || "").trim(),
      order: Number.isFinite(b.order) ? b.order : 0,
      skus: Array.isArray(b.skus)
        ? [...new Set(b.skus.map(s => String(s || "").trim()).filter(Boolean))]
        : [],
    })).filter(b => b.id && b.line && b.name);
    const blob = {
      lastUpdated: new Date().toISOString(),
      buckets: cleaned,
    };
    writeData("vf_yield_buckets", blob);
    res.json({ ok: true, ...blob });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
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
// 60-day window — wide enough to support trended-yield charts on the Yield
// tab without making the sync take forever. Daily cron at 07:00 UTC handles
// the long fetch fine; manual "Sync now" buttons take ~3-5 min.
const PRODUCTION_RUN_WINDOW_DAYS = 60;
const PRODUCTION_RUN_RATE_LIMIT_MS = 1100;

function sleepMs(ms) { return new Promise(r => setTimeout(r, ms)); }

// Wrap a Cin7 fetch with 429-aware backoff. Cin7 caps at 60 calls/60s on
// a rolling window; when we trip that, the call returns text/plain "You
// have reached 60 calls per 60 seconds API limit." We back off enough to
// fully refill the budget before retrying, up to 3 attempts. After that
// we return the final response (still 429) and let the caller's existing
// !resp.ok branch throw.
async function cin7FetchWithBackoff(url, fetchOpts, label) {
  const MAX_ATTEMPTS = 3;
  let resp;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    resp = await fetch(url, fetchOpts);
    if (resp.status !== 429) return resp;
    if (attempt === MAX_ATTEMPTS) break;
    // Honor Retry-After if Cin7 sets one; otherwise wait the full rolling
    // window (60s) plus a small jitter so multiple callers don't all retry
    // at the same instant.
    const retryAfterRaw = parseInt(resp.headers.get("retry-after") || "0", 10);
    const baseMs = retryAfterRaw > 0 ? retryAfterRaw * 1000 : 60_000;
    const jitterMs = Math.floor(((attempt * 13) % 7) * 1000); // deterministic 0–6s jitter
    const waitMs = baseMs + jitterMs;
    console.warn(`[Cin7 ${label}] 429 rate limit on attempt ${attempt}/${MAX_ATTEMPTS}; backing off ${Math.round(waitMs/1000)}s before retry`);
    await sleepMs(waitMs);
  }
  return resp;
}

async function fetchProductionOrderList(opts) {
  if (!process.env.CIN7_ACCOUNT_ID || !process.env.CIN7_APPLICATION_KEY) {
    throw new Error("CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY environment variable not set");
  }
  opts = opts || {};
  const params = new URLSearchParams();
  params.set("Limit", "200");
  if (opts.status)              params.set("Status",              opts.status);
  if (opts.requiredByDateFrom)  params.set("RequiredByDateFrom",  opts.requiredByDateFrom);
  if (opts.requiredByDateTo)    params.set("RequiredByDateTo",    opts.requiredByDateTo);
  if (opts.completionDateFrom)  params.set("CompletionDateFrom",  opts.completionDateFrom);
  if (opts.completionDateTo)    params.set("CompletionDateTo",    opts.completionDateTo);
  const all = [];
  let page = 1;
  while (true) {
    // Pace pagination calls — same 1.1s gap we already use between detail
    // calls. Without this, a 5+ page list (typical with the 90-day lookback)
    // burns through the 60/min budget in seconds and starves later detail
    // calls, manifesting as 429s mid-sync.
    if (page > 1) await sleepMs(PRODUCTION_RUN_RATE_LIMIT_MS);
    params.set("Page", String(page));
    const url = `${C7_PROD_BASE}/production/orderList?${params.toString()}`;
    const resp = await cin7FetchWithBackoff(url, { headers: cin7Headers(), redirect: "follow" }, "productionOrderList");
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
    if (batch.length < 200) break;
    page++;
    if (page > 50) throw new Error("Cin7 productionOrderList pagination exceeded 50 pages — aborting");
  }
  return all;
}

async function fetchProductionRunDetail(productionOrderID) {
  const url = `${C7_PROD_BASE}/production/order/run?ProductionOrderID=${encodeURIComponent(productionOrderID)}`;
  const resp = await cin7FetchWithBackoff(url, { headers: cin7Headers(), redirect: "follow" }, "productionRun");
  const ct = resp.headers.get("content-type") || "";
  const text = await resp.text().catch(() => "");
  if (!resp.ok) {
    throw new Error(`Cin7 productionRun ${resp.status} ${resp.statusText} [${ct}]: ${text.slice(0, 300)}`);
  }
  try { return JSON.parse(text); }
  catch (_) { throw new Error(`Cin7 productionRun ${resp.status} non-JSON [${ct}]: ${text.slice(0, 300)}`); }
}

// Module-level concurrency lock. Prevents a second sync from kicking off
// while the first is still running — a double-click on "Sync now" used to
// fire two parallel runs that combined to blow past the 60/min rate limit.
let _prodRunSyncInFlight = false;

async function performProductionRunSync() {
  if (_prodRunSyncInFlight) {
    throw new Error("A production-run sync is already running. Wait for it to finish before triggering another.");
  }
  _prodRunSyncInFlight = true;
  try {
    return await _performProductionRunSyncInner();
  } finally {
    _prodRunSyncInFlight = false;
  }
}

async function _performProductionRunSyncInner() {
  const now = new Date();
  const end = now.toISOString().slice(0, 10);
  const startD = new Date(now);
  startD.setUTCDate(startD.getUTCDate() - (PRODUCTION_RUN_WINDOW_DAYS - 1));
  const start = startD.toISOString().slice(0, 10);

  // 1. List recent ACTIVE orders. Earlier attempt used CompletionDateFrom
  //    which filters by the parent ORDER's completion date — but at Voyage,
  //    individual runs frequently complete while the parent order stays In
  //    Progress (multi-batch MOs). That made the filter return 0 rows.
  //
  //    New strategy: pull all non-voided orders due in the last 30 days,
  //    then in each order's detail walk Runs[] and find ones whose
  //    individual Status === "COMPLETED" with a ReceivedDate (or EndDate
  //    fallback) inside our 7-day window.
  // Wider lookback than the run window — we need to fetch orders whose
  // RequiredByDate could plausibly contain a run that completed inside our
  // 60-day window. Some orders complete weeks past their RequiredBy date.
  const lookbackDays = 90;
  const lookbackD = new Date(now);
  lookbackD.setUTCDate(lookbackD.getUTCDate() - lookbackDays);
  const requiredByDateFrom = lookbackD.toISOString().slice(0, 10);
  const list = await fetchProductionOrderList({
    status: "AllButVoided",
    requiredByDateFrom,
  });

  // Diagnostic counts — show in Railway logs and in the endpoint response
  // so we can see what the tenant's data actually looks like.
  const typeCounts = {};
  const statusCounts = {};
  const orderStatusCounts = {};
  for (const it of list) {
    if (!it) continue;
    const t = it.Type || "?";
    typeCounts[t] = (typeCounts[t] || 0) + 1;
    statusCounts[(it.Status || "(blank)")] = (statusCounts[(it.Status || "(blank)")] || 0) + 1;
    orderStatusCounts[(it.OrderStatus || "(blank)")] = (orderStatusCounts[(it.OrderStatus || "(blank)")] || 0) + 1;
  }
  console.log(`[ProdRunSync] List response: ${list.length} rows · Types ${JSON.stringify(typeCounts)} · Status ${JSON.stringify(statusCounts)} · OrderStatus ${JSON.stringify(orderStatusCounts)}`);

  // 2. Build the per-order map — every Order row gets a detail fetch.
  //    Voided orders are already excluded by the Status=AllButVoided filter.
  //    Run rows (Type=R) are ignored at this stage; we'll find run-level
  //    completion in the detail walk below.
  const ordersToFetch = new Map();   // ProductionOrderID → { orderNumber, productSKU, productName, locationName }
  for (const row of list) {
    if (!row || row.Type !== "O") continue;
    if (!row.ProductionOrderID) continue;
    if (ordersToFetch.has(row.ProductionOrderID)) continue;
    ordersToFetch.set(row.ProductionOrderID, {
      orderNumber: row.OrderNumber || "",
      productSKU: row.ProductSKU || "",
      productName: row.ProductName || "",
      locationName: row.LocationName || "",
    });
  }

  // Helper: did this run complete inside our 7-day window? Prefer
  // ReceivedDate (when output landed in stock), fall back to EndDate.
  const isRunInWindow = run => {
    const raw = run.ReceivedDate || run.EndDate || null;
    if (!raw) return false;
    const iso = String(raw).slice(0, 10);
    return iso >= start && iso <= end;
  };
  const isRunCompleted = run => String(run.Status || "").toUpperCase() === "COMPLETED";

  // Product-category lookup so the Production Output tab can group by
  // Voyage's product families (Chocolate / Coffee / Spreads / ...). The
  // product_costs cache mirrors Cin7's per-product Category string. Runs
  // whose FG SKU isn't in the cache fall through to "Other".
  const costsBlob = readData("product_costs") || { bySku: {} };
  const categoryOf = sku => {
    const c = costsBlob.bySku && costsBlob.bySku[sku];
    return (c && c.category) ? c.category : "Other";
  };

  // Yield-bucket lookup so the Yield tab can group runs by unit-op stage.
  // The vf_yield_buckets blob is seeded on startup and editable via the
  // Yield Setup tab. We match by the run's first Output[].ProductCode.
  const yieldBlob = readData("vf_yield_buckets") || { buckets: [] };
  const skuToBucket = buildSkuBucketMap(yieldBlob.buckets || []);
  const bucketLookup = sku => {
    if (!sku) return null;
    const b = skuToBucket.get(String(sku).trim());
    return b ? { id: b.id, line: b.line, name: b.name } : null;
  };

  // 3. For each active order, fetch run detail. Walk Runs[] and keep only
  //    runs whose Status === "COMPLETED" AND whose ReceivedDate/EndDate
  //    lands in the 7-day window. Then walk operations → components (Error
  //    Reporting) and Run.Output[] (Production Output).
  const flagged = [];
  const allCompletedRuns = [];
  let detailCallsMade = 0;
  let detailFailures = 0;
  let runsConsideredTotal = 0;
  let runsCompletedTotal = 0;
  let runsInWindowTotal = 0;
  const orderIDs = Array.from(ordersToFetch.keys());

  for (let i = 0; i < orderIDs.length; i++) {
    const orderID = orderIDs[i];
    const orderInfo = ordersToFetch.get(orderID);
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
      runsConsideredTotal++;
      if (!isRunCompleted(run)) continue;
      runsCompletedTotal++;
      if (!isRunInWindow(run)) continue;
      runsInWindowTotal++;

      const completionDate = String(run.ReceivedDate || run.EndDate || "").slice(0, 10) || null;
      const runMeta = {
        runID: run.RunID,
        completionDate,
        productSKU: orderInfo.productSKU,
        productName: orderInfo.productName,
        locationName: orderInfo.locationName,
      };
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

      // Yield path — total kg in vs total kg out for the run.
      // Inputs: sum every Components[].Quantity across all operations where
      // Unit==="kg" (per Matt's pick: include processing aids, exclude packaging
      // which is typically counted in "Each").
      // Outputs: sum Output[].Quantity where Unit==="kg".
      // Bucket: looked up from the first kg-output SKU; falls back to first
      // output if no kg outputs exist.
      const inputMassKg = allComponentsOnRun.reduce(
        (s, c) => s + (String(c.unit).toLowerCase() === "kg" ? c.actual : 0),
        0,
      );
      const outputMassKg = outputRows.reduce(
        (s, o) => s + (String(o.unit).toLowerCase() === "kg" ? o.qty : 0),
        0,
      );
      const yieldPct = inputMassKg > 0 ? (outputMassKg / inputMassKg) * 100 : null;
      const bucketSourceSku =
        (outputRows.find(o => String(o.unit).toLowerCase() === "kg") || outputRows[0] || {}).sku || runMeta.productSKU;
      const yieldBucket = bucketLookup(bucketSourceSku);

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
        // Yield-tab fields
        inputMassKg: Math.round(inputMassKg * 1000) / 1000,
        outputMassKg: Math.round(outputMassKg * 1000) / 1000,
        yieldPct: yieldPct != null ? Math.round(yieldPct * 100) / 100 : null,
        bucketSourceSku,
        yieldBucketId: yieldBucket ? yieldBucket.id : null,
        yieldBucketLine: yieldBucket ? yieldBucket.line : null,
        yieldBucketName: yieldBucket ? yieldBucket.name : null,
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
    lookbackDays,
    listRowCount: list.length,
    listTypeCounts: typeCounts,
    listStatusCounts: statusCounts,
    listOrderStatusCounts: orderStatusCounts,
    parentOrdersScanned: orderIDs.length,
    runsConsideredTotal,
    runsCompletedTotal,
    runsInWindowTotal,
    completedRunsScanned: runsInWindowTotal,
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
      lookbackDays: blob.lookbackDays,
      listRowCount: blob.listRowCount,
      listTypeCounts: blob.listTypeCounts,
      listStatusCounts: blob.listStatusCounts,
      listOrderStatusCounts: blob.listOrderStatusCounts,
      parentOrdersScanned: blob.parentOrdersScanned,
      runsConsideredTotal: blob.runsConsideredTotal,
      runsCompletedTotal: blob.runsCompletedTotal,
      runsInWindowTotal: blob.runsInWindowTotal,
      completedRunsScanned: blob.completedRunsScanned,
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

    // Traceability side-channel: same buffer, line-level parse, persist to
    // vf_inventory_movements. Non-fatal if it errors — the primary aggregation
    // above is what the dashboard depends on. Line-level is best-effort.
    let traceCounts = null;
    try {
      const traceParsed = parseMovementFile(buffer);
      if (traceParsed.lines.length) {
        traceCounts = persistMovementLines(traceParsed.lines, traceParsed.min_date, traceParsed.max_date, filename || null);
      }
    } catch (traceErr) {
      console.warn(`[gdrive-sync] Traceability side-channel skipped: ${traceErr.message}`);
    }

    console.log(`[gdrive-sync] Ingested ${filename || "(unnamed)"} — ${parsed.rowCount} rows · ${parsed.invSku.length} new-window SKUs · ${Object.keys(parsed.moMovements).length} new-window MOs · merged total: ${finalData.invSku.length} SKUs · ${Object.keys(finalData.moMovements).length} MOs${traceCounts ? ` · traceability +${traceCounts.inserted}/-${traceCounts.deleted} (total ${traceCounts.total})` : ''}`);
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
      traceability: traceCounts,
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

  // The underlying blob now holds a 60-day window (so the Yield tab has
  // enough data for weekly trending), but this endpoint promises 7 days
  // per its name. Filter here so the Production Output tab stays focused.
  const PRODUCTION_OUTPUT_TAB_DAYS = 7;
  const nowD = new Date();
  const sevenAgoD = new Date(nowD);
  sevenAgoD.setUTCDate(sevenAgoD.getUTCDate() - (PRODUCTION_OUTPUT_TAB_DAYS - 1));
  const tabWindowStart = sevenAgoD.toISOString().slice(0, 10);
  const tabWindowEnd = nowD.toISOString().slice(0, 10);
  const runsInTabWindow = blob.allCompletedRuns.filter(r =>
    r && r.completionDate && r.completionDate >= tabWindowStart && r.completionDate <= tabWindowEnd
  );

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

  for (const run of runsInTabWindow) {
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
    // Window reported = the actual 7-day filter applied here, not the
    // 60-day blob window. Keeps the UI label honest.
    windowStart: tabWindowStart,
    windowEnd: tabWindowEnd,
    totalQty,
    totalsByUnit,
    runCount: runsInTabWindow.length,
    byCategory,
    byWorkCenter,
    // Diagnostic info so the tab can surface "the sync ran but the filter
    // dropped everything" cases without needing Railway log access.
    lookbackDays: blob.lookbackDays || null,
    listRowCount: blob.listRowCount || 0,
    listTypeCounts: blob.listTypeCounts || {},
    listStatusCounts: blob.listStatusCounts || {},
    listOrderStatusCounts: blob.listOrderStatusCounts || {},
    parentOrdersScanned: blob.parentOrdersScanned || 0,
    runsConsideredTotal: blob.runsConsideredTotal || 0,
    runsCompletedTotal: blob.runsCompletedTotal || 0,
    runsInWindowTotal: blob.runsInWindowTotal || 0,
    completedRunsScanned: blob.completedRunsScanned || 0,
  });
});

// ── Yield trended endpoint — powers the Yield tab's sparkline cards ─────────
//
// Reads the production-run blob (same source as Error Reporting + Production
// Output) and rolls per-run yield up to weekly buckets per unit-op stage.
// Runs whose output SKU isn't in any yield bucket are surfaced separately as
// `unmapped` so the user can add them to the right list in Yield Setup.
app.get("/api/yield/trended", (req, res) => {
  const blob = readData("production_run_errors_t7d");
  const bucketsBlob = readData("vf_yield_buckets") || { buckets: [] };
  const allBuckets = bucketsBlob.buckets || [];

  if (!blob || !Array.isArray(blob.allCompletedRuns)) {
    return res.json({
      ok: true,
      lastSync: null,
      windowStart: null,
      windowEnd: null,
      runCount: 0,
      buckets: allBuckets.map(b => ({
        id: b.id, line: b.line, name: b.name, order: b.order,
        totalRuns: 0, totalInputKg: 0, totalOutputKg: 0, overallYieldPct: null, weekly: [],
      })),
      unmapped: { runCount: 0, topSkus: [] },
    });
  }

  // Helper: Monday of the ISO week containing `dateStr` (YYYY-MM-DD).
  const mondayOf = dateStr => {
    if (!dateStr) return null;
    const d = new Date(dateStr + "T00:00:00Z");
    if (isNaN(d.getTime())) return null;
    const dow = d.getUTCDay() || 7; // 1..7 with Monday=1
    d.setUTCDate(d.getUTCDate() - (dow - 1));
    return d.toISOString().slice(0, 10);
  };

  // Group runs by bucket id (null for unmapped)
  const runsByBucket = new Map();
  const unmappedBySku = new Map();
  for (const run of blob.allCompletedRuns) {
    const bucketId = run.yieldBucketId || null;
    if (bucketId === null) {
      const sku = run.bucketSourceSku || run.fgSKU || "(no sku)";
      if (!unmappedBySku.has(sku)) {
        unmappedBySku.set(sku, { sku, count: 0, sampleProduct: run.fgProduct || "", sampleMoRef: run.moRef || "" });
      }
      unmappedBySku.get(sku).count += 1;
      continue;
    }
    if (!runsByBucket.has(bucketId)) runsByBucket.set(bucketId, []);
    runsByBucket.get(bucketId).push(run);
  }

  // Build the response: every configured bucket gets a slot (even with 0
  // runs) so the UI can render the full lineup consistently.
  const bucketsOut = allBuckets.map(b => {
    const runs = runsByBucket.get(b.id) || [];
    const weeklyMap = new Map();   // weekStart → { runs, inputKg, outputKg }
    let totalInput = 0, totalOutput = 0, totalRuns = runs.length;
    for (const r of runs) {
      const week = mondayOf(r.completionDate);
      if (!week) continue;
      if (!weeklyMap.has(week)) weeklyMap.set(week, { weekStart: week, runs: 0, inputKg: 0, outputKg: 0 });
      const w = weeklyMap.get(week);
      w.runs += 1;
      w.inputKg += r.inputMassKg || 0;
      w.outputKg += r.outputMassKg || 0;
      totalInput += r.inputMassKg || 0;
      totalOutput += r.outputMassKg || 0;
    }
    const weekly = Array.from(weeklyMap.values())
      .sort((a, b) => a.weekStart.localeCompare(b.weekStart))
      .map(w => ({
        weekStart: w.weekStart,
        runs: w.runs,
        inputKg: Math.round(w.inputKg * 100) / 100,
        outputKg: Math.round(w.outputKg * 100) / 100,
        yieldPct: w.inputKg > 0 ? Math.round((w.outputKg / w.inputKg) * 10000) / 100 : null,
      }));
    return {
      id: b.id,
      line: b.line,
      name: b.name,
      order: b.order,
      totalRuns,
      totalInputKg: Math.round(totalInput * 100) / 100,
      totalOutputKg: Math.round(totalOutput * 100) / 100,
      overallYieldPct: totalInput > 0 ? Math.round((totalOutput / totalInput) * 10000) / 100 : null,
      weekly,
    };
  });

  const unmappedSkus = Array.from(unmappedBySku.values()).sort((a, b) => b.count - a.count);
  const unmappedRunCount = unmappedSkus.reduce((s, x) => s + x.count, 0);

  res.json({
    ok: true,
    lastSync: blob.lastSync || null,
    windowStart: blob.windowStart || null,
    windowEnd: blob.windowEnd || null,
    runCount: blob.allCompletedRuns.length,
    buckets: bucketsOut,
    unmapped: {
      runCount: unmappedRunCount,
      topSkus: unmappedSkus.slice(0, 20),
    },
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

// Machines where the operation is packaging/finishing — the only BOM shape
// that includes box/bag/packaging inputs is the finished-good (FG-*) BOM.
// WIP tiers (WIP-, WIP2-, WIP3-, ...) hold formula or intermediate recipes
// with raw ingredients, not packaging. If a scheduler drops a packout MO
// against any WIP or RM SKU, expanding that BOM would produce spurious
// demand (cocoa/sugar for WIP1, or nothing usable for WIP2/3). MRP skips
// such orders and surfaces them for scheduling cleanup.
const PACKOUT_MACHINES = new Set(["mac_packout", "depositing", "pouching"]);

// Only FG-* SKUs carry the packaging BOM shape MRP needs on packout machines.
const PACKOUT_VALID_SKU_RE = /^FG-/i;

// Production lead time per machine, in days (sourced from Capacity & Ops tab)
const MACHINE_LEAD_DAYS = {
  seed_clean:  1,
  roaster:     1,
  east_mac:    10,
  west_mac:    10,
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
//
// opts.stopAtSkus (optional Set) — treat these SKUs as leaves even if they
// have a BOM. Used by multi-level MRP netting so that when this MO's
// expansion hits a WIP that is separately produced by another scheduled
// MO in the planning window, we DON'T recurse into that WIP's own RMs
// (they're already counted from the other MO's expansion). Without this,
// two MOs where one produces a WIP the other consumes would double-count
// every leaf RM in the WIP's recipe. If a stopped WIP has partial coverage
// (planned production < gross consumption), the caller is responsible for
// expanding the gap separately — this function just honors the cutoff.
function expandBom(parents, parentSku, qty, opts) {
  opts = opts || {};
  const applyWastage = opts.applyWastage !== false; // default true
  const stopAtSkus = opts.stopAtSkus || null;
  const visited = new Set();
  const leaves = {};
  const trail = [];
  const stoppedAt = {}; // sku -> summed qty that hit a cutoff (for observability)

  function recurse(sku, needed, depth) {
    if (depth > 12) throw new Error(`BOM recursion too deep at ${sku}`);
    if (visited.has(sku)) {
      // cycle — log and treat as leaf to bail out gracefully
      if (!leaves[sku]) leaves[sku] = { sku, qty: 0, name: "(cycle detected)", isCycle: true };
      leaves[sku].qty += needed;
      return;
    }
    // Multi-level MRP cutoff: treat this SKU as a leaf because another MO
    // in the plan already produces it. Do NOT add to leaves (it's not a
    // procurable RM) — just record the cutoff for the caller's observability.
    if (stopAtSkus && stopAtSkus.has(sku) && depth > 0) {
      stoppedAt[sku] = (stoppedAt[sku] || 0) + needed;
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
  return { leaves, intermediates: trail, stoppedAt };
}

// POST /api/boms/import — admin-only, accepts raw CSV text in body { csv: "..." }
// Kept as an escape hatch for manual corrections; the primary path is the
// nightly Cin7 API sync below.
app.post("/api/boms/import", requireAdmin, (req, res) => {
  try {
    const csv = req.body && req.body.csv;
    if (!csv) return res.status(400).json({ ok: false, error: "Missing 'csv' field in body" });
    const blob = parseBomCsv(csv);
    blob.source = "csv-upload";
    writeData("vf_boms", blob);
    res.json({ ok: true, ...blob, parents: undefined });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── Cin7 BOM auto-sync (per-product detail fetch) ───────────────────────────
//
// Cin7's list endpoint (/product) returns only summary fields — components
// are exposed exclusively on the detail endpoint (/product?ID=<uuid>). So
// this sync is N+1: fetch the list once, filter to SKUs that plausibly
// have a BOM (Voyage's FG-* and WIP-* prefixes plus anything Cin7 tagged
// Type=Assembled), then hit detail per SKU with rate limiting.
//
// Runs 2–5 minutes wall-clock depending on the BOM count, so the endpoint
// is fire-and-forget: it starts the job and returns immediately; the UI
// polls /sync-status. Failures below 10% overwrite vf_boms; above 10% we
// keep the prior blob to avoid a partial-write corruption.
//
// The CSV upload endpoint above is retained as a manual override — QA can
// still hand-tune BOMs by uploading a modified Cin7 export.

function _bomSleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function fetchCin7BomCandidates() {
  // Reuse fetchCin7ProductCosts — it already paginates /product cleanly.
  // We just filter server-side.
  const products = await fetchCin7ProductCosts();
  return products.filter(p => {
    if (!p.SKU || !p.ID) return false;
    const t = String(p.Type || "").toLowerCase();
    const sku = String(p.SKU).toUpperCase();
    // Match FG-* and any WIP variant (WIP-, WIP2-, WIP3-, ...). Original
    // filter of startsWith("WIP-") missed the whole WIP2-* series because
    // "WIP2-".startsWith("WIP-") is false — a subtle string-prefix bug that
    // silently dropped the entire conching stage from the sync.
    return t.includes("assembl") || sku.startsWith("FG-") || /^WIP\d*-/.test(sku);
  });
}

async function fetchCin7ProductDetail(productID) {
  const url = `https://inventory.dearsystems.com/ExternalApi/v2/product?ID=${encodeURIComponent(productID)}`;
  const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
  const ct = resp.headers.get("content-type") || "";
  const text = await resp.text().catch(() => "");
  if (resp.status === 429) {
    const err = new Error("Cin7 rate limit hit (429)");
    err.retryable = true;
    throw err;
  }
  if (!resp.ok) throw new Error(`Cin7 product detail ${resp.status} [${ct}]: ${text.slice(0, 300)}`);
  let data;
  try { data = JSON.parse(text); }
  catch (_) { throw new Error(`Cin7 product detail ${resp.status} non-JSON [${ct}]: ${text.slice(0, 300)}`); }
  // Some Cin7 tenants wrap in { Products: [...] }, others return the bare object.
  if (data && Array.isArray(data.Products) && data.Products.length) return data.Products[0];
  return data;
}

// Hits the correct Cin7 Production BOM endpoint — confirmed via debug probe.
// Response shape: { ProductID, ProductionBoms: [{ BomID, OutputQuantity,
// Version, Name, IsDefault, ComponentProductionLeadTime, Operations: [{
// Order, Name, WorkCenterName, Components: [{ ProductSku, ProductName,
// Quantity, WastagePercent, ... }] }] }] }
async function fetchCin7ProductionBom(productID) {
  const url = `https://inventory.dearsystems.com/ExternalApi/v2/production/productionBOM?ProductID=${encodeURIComponent(productID)}`;
  const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
  const ct = resp.headers.get("content-type") || "";
  const text = await resp.text().catch(() => "");
  if (resp.status === 429) {
    const err = new Error("Cin7 rate limit hit (429)");
    err.retryable = true;
    throw err;
  }
  // 404 = no BOM for this product — return empty, not an error.
  if (resp.status === 404) return { ProductID: productID, ProductionBoms: [] };
  if (!resp.ok) throw new Error(`Cin7 productionBOM ${resp.status} [${ct}]: ${text.slice(0, 300)}`);
  let data;
  try { data = JSON.parse(text); }
  catch (_) { throw new Error(`Cin7 productionBOM ${resp.status} non-JSON [${ct}]: ${text.slice(0, 300)}`); }
  return data || { ProductID: productID, ProductionBoms: [] };
}

// Translate one Cin7 ProductionBom entry into the vf_boms bucket shape.
// Cin7's response nests components inside Operations[], so we flatten
// operations into a single component list while carrying operation-level
// metadata (Order → op, Name → opName, WorkCenterName → workCentre).
// Mirrors parseBomCsv's projection so downstream MRP/expansion code is
// agnostic to whether the data came from CSV or the API.
function translateCin7ProductionBom(product, bomEntry) {
  const version = String(bomEntry.Version != null ? bomEntry.Version : 1);
  const components = [];
  const operations = Array.isArray(bomEntry.Operations) ? bomEntry.Operations : [];
  for (const op of operations) {
    const opRows = Array.isArray(op.Components) ? op.Components : [];
    const opSeq = parseInt(op.Order, 10) || 1;
    const opName = op.Name || "";
    // Cin7 spells this WorkCenterName (US); parseBomCsv reads WorkCentreName
    // (UK). Both feed deriveMachineFromBom via the same string, so we just
    // pass whatever Cin7 hands us.
    const workCentre = op.WorkCenterName || op.WorkCentreName || "";
    for (const c of opRows) {
      if (!c || !c.ProductSku) continue;
      components.push({
        sku: c.ProductSku,
        name: c.ProductName || "",
        qty: Number(c.Quantity) || 0,
        wastagePct: Number(c.WastagePercent) || 0,
        op: opSeq,
        opName,
        workCentre,
      });
    }
  }
  return {
    parent: product.SKU,
    parentName: product.Name || "",
    version,
    // Cin7's ProductionBom "Name" is what the CSV calls VersionName
    // (e.g. "Drums", "25kg boxes").
    versionName: bomEntry.Name || "",
    isDefault: !!bomEntry.IsDefault,
    qtyToProduce: Number(bomEntry.OutputQuantity) || 1,
    runSize: Number(bomEntry.RunSize) || 0,
    minQty: Number(bomEntry.MinQuantity) || 0,
    maxQty: Number(bomEntry.MaxQuantity) || 0,
    productionLeadTimeRaw: parseInt(bomEntry.ComponentProductionLeadTime, 10) || 0,
    components,
  };
}

async function performCin7BomsSync() {
  const candidates = await fetchCin7BomCandidates();
  const bucket = {};
  const failures = [];
  let detailFetched = 0;
  let withBoms = 0;

  for (let i = 0; i < candidates.length; i++) {
    const p = candidates[i];
    let response;
    try {
      response = await fetchCin7ProductionBom(p.ID);
      detailFetched++;
    } catch (e) {
      failures.push({ sku: p.SKU, error: e.message });
      // Back off harder on 429; otherwise standard pause and continue.
      await _bomSleep(e.retryable ? 5000 : 1100);
      continue;
    }
    const boms = response && Array.isArray(response.ProductionBoms) ? response.ProductionBoms : [];
    if (boms.length) {
      withBoms++;
      for (const b of boms) {
        const entry = translateCin7ProductionBom(p, b);
        if (!entry.components.length) continue; // skip BOM versions with no components
        const key = entry.parent + "|" + entry.version;
        bucket[key] = entry;
      }
    }
    await _bomSleep(1100); // ~55/min to stay under Cin7's 60/min cap
  }

  // Post-process identically to parseBomCsv: derive machine + lead time,
  // group by parent, sort so default version is first.
  const parents = {};
  for (const key of Object.keys(bucket)) {
    const b = bucket[key];
    b.machine = deriveMachineFromBom(b);
    b.productionLeadTime = b.machine ? (MACHINE_LEAD_DAYS[b.machine] || null) : null;
    if (!parents[b.parent]) parents[b.parent] = [];
    parents[b.parent].push(b);
  }
  for (const sku of Object.keys(parents)) {
    parents[sku].sort((a, b) => {
      if (a.isDefault !== b.isDefault) return a.isDefault ? -1 : 1;
      return Number(a.version) - Number(b.version);
    });
  }

  const failureRate = candidates.length ? failures.length / candidates.length : 0;
  if (failureRate > 0.10) {
    const err = new Error(
      `Failure rate ${(failureRate*100).toFixed(1)}% exceeds 10% threshold — keeping prior BOMs. ` +
      `First failures: ${failures.slice(0,3).map(f=>f.sku+': '+f.error).join(' | ')}`
    );
    err.failures = failures;
    err.candidatesScanned = candidates.length;
    err.detailFetched = detailFetched;
    throw err;
  }

  // Zero-BOM guardrail: if we successfully fetched a meaningful number of
  // product detail records and NONE returned ProductionBoms, the API shape
  // has changed (or Cin7 renamed the endpoint again). Overwriting the prior
  // BOMs with an empty result would silently kill MRP downstream, so refuse
  // the write and surface a clear error.
  if (detailFetched >= 10 && withBoms === 0) {
    const err = new Error(
      `${detailFetched} products checked, 0 returned a Production BOM — keeping prior BOMs. ` +
      `Verify /production/productionBOM still works via /api/cin7/boms/debug-production-bom?sku=<any BOM'd SKU>.`
    );
    err.detailFetched = detailFetched;
    throw err;
  }

  const now = new Date().toISOString();
  const parentCount = Object.keys(parents).length;
  const rowCount = Object.values(parents).reduce((s, vs) => s + vs.reduce((s2, v) => s2 + v.components.length, 0), 0);
  const blob = {
    lastImport: now,      // keep existing field name for backward compat
    lastSync: now,
    source: "cin7-api",
    productsScanned: candidates.length,
    detailFetched,
    productsWithBoms: withBoms,
    failureCount: failures.length,
    parents,
    parentCount,
    rowCount,
  };
  writeData("vf_boms", blob);
  return {
    ok: true,
    lastSync: now,
    productsScanned: candidates.length,
    detailFetched,
    productsWithBoms: withBoms,
    parentCount,
    rowCount,
    failureCount: failures.length,
    firstFailures: failures.slice(0, 10),
  };
}

// Fire-and-forget: BOM sync runs 2–5 minutes; can't hold an HTTP request that
// long behind the Railway edge proxy. POST kicks off the background job and
// returns immediately; the UI polls the status endpoint.
let _bomSyncState = { state: "idle", startedAt: null, finishedAt: null, error: null, result: null };

app.post("/api/cin7/boms/sync", requireAdmin, (req, res) => {
  if (_bomSyncState.state === "running") {
    return res.status(409).json({ ok: false, error: "Sync already in progress", state: _bomSyncState });
  }
  _bomSyncState = { state: "running", startedAt: new Date().toISOString(), finishedAt: null, error: null, result: null };
  res.json({ ok: true, message: "Sync started in background — poll /api/cin7/boms/sync-status", state: _bomSyncState });
  performCin7BomsSync()
    .then(result => {
      _bomSyncState = {
        state: "complete",
        startedAt: _bomSyncState.startedAt,
        finishedAt: new Date().toISOString(),
        error: null,
        result,
      };
      console.log(`[Cin7 BOMs] Sync done — ${result.detailFetched} products checked, ${result.productsWithBoms} with BOMs, ${result.parentCount} parents, ${result.failureCount} failures`);
    })
    .catch(e => {
      _bomSyncState = {
        state: "failed",
        startedAt: _bomSyncState.startedAt,
        finishedAt: new Date().toISOString(),
        error: e.message,
        result: null,
      };
      console.error("[Cin7 BOMs] Sync failed:", e.message);
    });
});

app.get("/api/cin7/boms/sync-status", (req, res) => {
  const blob = readData("vf_boms") || {};
  res.json({
    ok: true,
    ..._bomSyncState,
    currentBlob: {
      lastImport: blob.lastImport || null,
      lastSync: blob.lastSync || null,
      source: blob.source || null,
      parentCount: blob.parentCount || 0,
      rowCount: blob.rowCount || 0,
      productsScanned: blob.productsScanned || null,
      productsWithBoms: blob.productsWithBoms || null,
    },
  });
});

// GET /api/cin7/boms/debug-probe — admin diagnostic. Probes a set of
// candidate Cin7 endpoints that likely expose Production BOMs in bulk,
// reports HTTP status + first ~400 chars of response body for each. Used
// once to figure out which endpoint Cin7 exposes for this tenant, so the
// sync can be rewritten as a bulk pull instead of per-product detail.
app.get("/api/cin7/boms/debug-probe", requireAdmin, async (req, res) => {
  // Round 2: try the /production/* namespace (mirrors production/orderList
  // which we know works), per-product nested paths with a real BOM'd product
  // ID, and report/export namespaces that sometimes hold what UI exports hit.
  const knownBomProductId = "f222eab2-f731-45ac-9ba0-8153649e866b"; // from earlier debug-detail call
  const candidates = [
    "/production/bom?Page=1&Limit=5",
    "/production/bomList?Page=1&Limit=5",
    "/production/BOM?Page=1&Limit=5",
    "/production/productionBOM?Page=1&Limit=5",
    "/production/productionBom?Page=1&Limit=5",
    "/production/orderBom?Page=1&Limit=5",
    `/product/bom?ID=${knownBomProductId}`,
    `/product/BOM?ID=${knownBomProductId}`,
    `/product/productionBOM?ID=${knownBomProductId}`,
    `/product/bomList?ID=${knownBomProductId}`,
    "/productBOM?Page=1&Limit=5",
    "/productBOMList?Page=1&Limit=5",
    "/report/productionBOM",
    "/reports/productionBOM",
    "/export/productionBOM",
    "/productionOrderBOM?Page=1&Limit=5",
  ];
  const results = [];
  for (const path of candidates) {
    const url = `https://inventory.dearsystems.com/ExternalApi/v2${path}`;
    try {
      const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
      const text = await resp.text().catch(() => "");
      const trimmed = text.slice(0, 400);
      // Try to detect success even if 200 (some Cin7 endpoints return {"Error":"..."} with 200)
      let looksLikeSuccess = resp.ok && !/error|not found|invalid/i.test(trimmed.slice(0, 100));
      results.push({
        path,
        status: resp.status,
        ok: resp.ok,
        looksLikeSuccess,
        contentType: resp.headers.get("content-type") || "",
        bodyPreview: trimmed,
      });
    } catch (e) {
      results.push({ path, error: e.message });
    }
    // Small pause between probes to be polite to Cin7's rate limit
    await new Promise(r => setTimeout(r, 300));
  }
  const winners = results.filter(r => r.looksLikeSuccess);
  res.json({
    ok: true,
    checked: results.length,
    winnerCount: winners.length,
    winners: winners.map(w => w.path),
    results,
  });
});

// GET /api/cin7/boms/debug-production-bom?sku=FG-XXX — admin diagnostic.
// Hits the confirmed Cin7 endpoint /production/productionBOM?ProductID=<id>
// for one SKU and returns the raw response so we can wire the sync against
// the actual shape (field names for components, wastage, etc.).
app.get("/api/cin7/boms/debug-production-bom", requireAdmin, async (req, res) => {
  try {
    const sku = String(req.query.sku || "").trim();
    if (!sku) return res.status(400).json({ ok: false, error: "sku query param required" });
    const products = await fetchCin7ProductCosts();
    const match = products.find(p => String(p.SKU || "").toUpperCase() === sku.toUpperCase());
    if (!match) return res.status(404).json({ ok: false, error: `SKU ${sku} not found` });
    const url = `https://inventory.dearsystems.com/ExternalApi/v2/production/productionBOM?ProductID=${encodeURIComponent(match.ID)}`;
    const resp = await fetch(url, { headers: cin7Headers(), redirect: "follow" });
    const ct = resp.headers.get("content-type") || "";
    const text = await resp.text().catch(() => "");
    let parsed = null;
    try { parsed = JSON.parse(text); } catch (_) {}
    res.json({
      ok: true,
      sku,
      productID: match.ID,
      productName: match.Name,
      status: resp.status,
      contentType: ct,
      topLevelKeys: parsed && typeof parsed === "object" ? Object.keys(parsed) : null,
      // If the response is a wrapper like { ProductionBOMs: [...] } or bare array
      isArray: Array.isArray(parsed),
      firstEntry: Array.isArray(parsed) ? parsed[0] : (parsed && parsed.ProductionBOMs ? parsed.ProductionBOMs[0] : (parsed && parsed.BillOfMaterials ? parsed.BillOfMaterials[0] : parsed)),
      raw: parsed,
      rawString: text.length > 15000 ? text.slice(0, 15000) + "…[truncated]" : text,
    });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// GET /api/cin7/boms/debug-detail?sku=FG-XXX — admin diagnostic. Fetches
// the raw Cin7 product detail for one SKU and returns the top-level keys
// plus the FULL response (truncated body up to ~50 kB). Used to figure out
// what field name Cin7 actually puts BOM data under when the sync returns
// zero BOMs.
app.get("/api/cin7/boms/debug-detail", requireAdmin, async (req, res) => {
  try {
    const sku = String(req.query.sku || "").trim();
    if (!sku) return res.status(400).json({ ok: false, error: "sku query param required" });
    // Look up the product ID from the cached costs blob (populated by the
    // nightly product-costs sync) — avoids a whole-list refetch just for one lookup.
    // If costs blob is missing / stale, fall through to a fresh list fetch.
    let productID = null;
    let productName = null;
    const costsBlob = readData("product_costs") || {};
    // costs blob is keyed by SKU but doesn't store ID; do the fresh list fetch.
    // (Keeps this endpoint honest — never uses stale IDs.)
    const products = await fetchCin7ProductCosts();
    const match = products.find(p => String(p.SKU || "").toUpperCase() === sku.toUpperCase());
    if (!match) {
      return res.status(404).json({ ok: false, error: `SKU ${sku} not found in Cin7 product list`, checked: products.length });
    }
    productID = match.ID;
    productName = match.Name;
    const detail = await fetchCin7ProductDetail(productID);
    const rawJson = JSON.stringify(detail);
    const truncated = rawJson.length > 50000;
    res.json({
      ok: true,
      sku,
      productID,
      productName,
      topLevelKeys: detail && typeof detail === "object" ? Object.keys(detail) : null,
      rawResponseBytes: rawJson.length,
      rawResponseTruncated: truncated,
      raw: detail,
      rawString: truncated ? rawJson.slice(0, 50000) + "…[truncated]" : rawJson,
    });
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

// GET /api/sku-unit-info — per-order-SKU kg-per-unit lookup used by the
// calendar / MO Status frontend to relabel case-unit actuals (PFS pouching).
//
// The supply-settings perSku map is keyed by canonical FG SKU like
// "FG-604-102-00", but `order.sku` is the operator-entered display string
// like "(604 - pouches) PFS_1.5oz_Pouch_20oz_...". Substring matching is
// unreliable, so the resolution has to happen here where extractBomSku can
// walk the BOM catalog. The client just receives a flat map keyed by the
// exact `order.sku` string and does direct lookups.
//
// Read-only; no auth needed (same access tier as supply-settings GET).
app.get("/api/sku-unit-info", (req, res) => {
  try {
    const orders = readData("vf_orders") || [];
    const bomBlob = readData("vf_boms") || { parents: {} };
    const bomParents = bomBlob.parents || {};
    const supply = readData("vf_supply_settings") || { perSku: {} };
    const perSku = supply.perSku || {};
    const skuMap = {};
    const uniqueSkus = new Set();
    for (const o of orders) if (o && o.sku) uniqueSkus.add(o.sku);
    for (const orderSku of uniqueSkus) {
      const fgSku = extractBomSku(orderSku, bomParents);
      if (!fgSku) continue;
      const entry = perSku[fgSku];
      const kg = entry && Number(entry.kgPerUnit);
      if (isFinite(kg) && kg > 0) skuMap[orderSku] = kg;
    }
    res.json({ ok: true, skuMap });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
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
      // kgPerUnit override — used by buildRequirements + bom_expand to
      // convert an order's qty (in kg) into the BOM's qtyToProduce unit
      // (e.g. cases) before recursive expansion. Only relevant for SKUs
      // whose Cin7 BOM was authored per-case rather than per-kg.
      if (v.kgPerUnit != null && v.kgPerUnit !== "") {
        const n = parseFloat(v.kgPerUnit);
        if (isFinite(n) && n > 0) entry.kgPerUnit = n;
      }
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
  // BOM sync runs first at 05:30 UTC because it's the slowest (multi-minute
  // N+1 detail fetch) and we want it to finish before the on-hand/costs/PO/SO
  // burst hits Cin7's 60/min rate budget at 06:30 and after.
  cron.schedule("30 5 * * *", async () => {
    console.log("[Cin7 BOMs] Nightly sync starting…");
    try {
      const s = await performCin7BomsSync();
      console.log(`[Cin7 BOMs] Sync done — ${s.detailFetched} products checked, ${s.productsWithBoms} with BOMs, ${s.parentCount} parents, ${s.failureCount} failures`);
    } catch (e) {
      console.error("[Cin7 BOMs] Nightly sync failed:", e.message);
    }
  });
  console.log("[Cin7 BOMs] Nightly sync scheduled at 05:30 UTC");

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

  // Purchase-order supplier sync — populates PO ref → actual supplier map
  // for the traceability lineage view. Small payload (a few thousand POs
  // for Voyage), single paginated API call. Scheduled at 06:50 UTC to sit
  // after on-hand (06:30) and costs (06:45) so we don't compete with them
  // for Cin7's 60/min rate budget, and still before production-run at 07:00.
  cron.schedule("50 6 * * *", async () => {
    console.log("[Cin7 POs] Nightly sync starting…");
    try {
      const s = await performCin7PurchaseOrdersSync();
      console.log(`[Cin7 POs] Sync done — ${s.purchaseCount} POs, ${s.withSupplierCount} with supplier`);
    } catch (e) {
      console.error("[Cin7 POs] Nightly sync failed:", e.message);
    }
  });
  console.log("[Cin7 POs] Nightly sync scheduled at 06:50 UTC");

  // Sales-order customer sync — populates SO ref → customer name map for
  // the forward-trace shipment attribution. Scheduled at 06:55 UTC, after
  // POs (06:50) and before production-run (07:00).
  cron.schedule("55 6 * * *", async () => {
    console.log("[Cin7 SOs] Nightly sync starting…");
    try {
      const s = await performCin7SalesOrdersSync();
      console.log(`[Cin7 SOs] Sync done — ${s.saleCount} SOs, ${s.withCustomerCount} with customer`);
    } catch (e) {
      console.error("[Cin7 SOs] Nightly sync failed:", e.message);
    }
  });
  console.log("[Cin7 SOs] Nightly sync scheduled at 06:55 UTC");

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
  const skipped = { unconfirmed: 0, complete: 0, noStart: 0, outsideHorizon: 0, noBom: 0, excludedByDate: 0, packoutFormulaOnly: 0 };
  const packoutSkippedExamples = [];
  const noBomExamples = [];
  // "Exclude before" filter — solves the lingering-TBD problem where
  // uncleared old orders inflate MRP demand for material that's actually
  // already been produced. opts.excludeBeforeDate is a YYYY-MM-DD string.
  const excludeBeforeDate = opts.excludeBeforeDate || null;

  // Compute the in-scope order set ONCE up front — reused by both the FG
  // netting pre-pass and the WIP cutoff pre-pass below. Filters MUST match
  // the main loop's per-order gate so pre-passes only see orders that will
  // actually generate requirements. Critically: completed MOs are excluded
  // (a completed MO's OUTPUT is what on-hand already reflects; letting a
  // completed MO "consume" on-hand double-counts it and drains the netting
  // pool before real planned production gets a chance).
  const inScopeOrders = orders.filter(o => {
    if (!opts.includeUnconfirmed && o.confirmed === false && !o.__fromPipelineDraft && !o.__fromCompanionRule) return false;
    if (o.status === "complete") return false;
    if (!o.start) return false;
    if (o.start > horizonEnd) return false;
    if (excludeBeforeDate && o.start < excludeBeforeDate) return false;
    return true;
  });

  // ── FG-level netting (default ON for planning use) ────────────────────────
  // Real MRP nets on-hand at every BOM level, not just leaf RM. If we already
  // have 5,000 kg of a finished good on the shelf, a pipeline draft for
  // 15,000 kg of that same FG only needs 10,000 kg of NEW production — which
  // means only 10,000 kg × recipe worth of raw materials. Without this step
  // the RM demand is inflated by whatever FG stock we're carrying.
  //
  // Uses onHandBySku[fgSku].available (on-hand minus SO allocations) so we
  // don't double-count FG that's already been committed to a customer
  // shipment. Consumes FIFO by need-by-date so the earliest-scheduled order
  // gets first crack at the free stock (matches how a scheduler would think
  // about it — if the FG is on the shelf right now, cover the nearest need
  // first). Records the offset on each order so trace_po_demand can surface
  // the "netted by FG on-hand" note in attributions.
  const netFgOnHand = opts.netFgOnHand !== false && opts.onHandBySku && Object.keys(opts.onHandBySku).length > 0;
  const fgNetted = new Map(); // fgSku -> {consumedByOrder: Map<orderId, kg>, totalConsumed, offsetBreakdown}
  if (netFgOnHand) {
    const byFg = new Map();
    for (const o of inScopeOrders) {
      if (!o.sku) continue;
      const fg = extractBomSku(o.sku, bomParents);
      if (!fg) continue;
      const plannedQty = o.total || (o.qty || 0) * (o.batches || 1);
      if (plannedQty <= 0) continue;
      const needBy = o.start && o.start < opts.today ? opts.today : o.start;
      if (!byFg.has(fg)) byFg.set(fg, []);
      byFg.get(fg).push({ order: o, fgSku: fg, plannedQty, needBy });
    }
    for (const [fg, entries] of byFg) {
      const oh = opts.onHandBySku[fg];
      if (!oh) continue;
      let available = Math.max(0, Number(oh.available || 0));
      if (available <= 0) continue;
      entries.sort((a, b) => String(a.needBy || "").localeCompare(String(b.needBy || "")));
      const consumedByOrder = new Map();
      const offsetBreakdown = []; // audit trail: which order got which slice
      let totalConsumed = 0;
      for (const e of entries) {
        if (available <= 0) break;
        const take = Math.min(available, e.plannedQty);
        if (take <= 0) continue;
        consumedByOrder.set(e.order.id || e.order.orderId || String(Math.random()), take);
        offsetBreakdown.push({
          orderId: e.order.orderId || e.order.id,
          orderInternalId: e.order.id,
          needBy: e.needBy,
          grossPlannedKg: e.plannedQty,
          offsetKg: take,
          netPlannedKg: e.plannedQty - take,
          isPipelineDraft: !!e.order.__fromPipelineDraft,
          isCompanionDemand: !!e.order.__fromCompanionRule,
        });
        available -= take;
        totalConsumed += take;
        e.netPlannedQty = e.plannedQty - take;
        e.fgOffsetKg = take;
      }
      fgNetted.set(fg, { totalConsumed, availableRemaining: available, consumedByOrder, offsetBreakdown });
    }
    // Build a fast lookup: order.id → { netPlannedQty, fgOffsetKg }
    opts._perOrderNet = new Map();
    for (const [, entries] of byFg) {
      for (const e of entries) {
        if (e.netPlannedQty != null) {
          opts._perOrderNet.set(e.order.id || e.order.orderId, {
            netPlannedQty: e.netPlannedQty,
            fgOffsetKg: e.fgOffsetKg,
            fgSku: e.fgSku,
          });
        }
      }
    }
  }

  // ── Multi-level MRP netting: prevent WIP double-count ────────────────────
  // If MO-A produces WIP-X and MO-B consumes WIP-X in its BOM, gross expansion
  // counts WIP-X's raw materials twice (once from MO-A's expansion, once from
  // MO-B recursing through WIP-X). Fix: when expanding MO-B, treat WIP-X as a
  // leaf so recursion stops there — MO-A's expansion already covers those RMs.
  // Same principle as FG netting above, but for intermediate WIPs.
  //
  // We only cut off SKUs where planned production >= expected consumption
  // (fully covered). For the "partial coverage" case (production < demand)
  // we still cut off — but log the gap so the caller can surface it. The
  // gap represents production that ops SHOULD schedule but hasn't, so the
  // right behavior is to flag it, not to silently pad RM demand.
  const plannedProducersBySku = new Map(); // sku -> total planned qty in scope
  for (const o of inScopeOrders) {
    if (!o.sku) continue;
    const fg = extractBomSku(o.sku, bomParents);
    if (!fg) continue;
    const q = o.total || (o.qty || 0) * (o.batches || 1);
    if (q <= 0) continue;
    plannedProducersBySku.set(fg, (plannedProducersBySku.get(fg) || 0) + q);
  }
  const stopAtSkusGlobal = new Set(plannedProducersBySku.keys());
  const wipCutoffSummary = new Map(); // wipSku -> {totalStoppedKg, sourceOrderIds:[]}

  for (const o of orders) {
    // Pipeline-draft synthetic orders bypass the "unconfirmed" filter —
    // they're inherently unconfirmed (no operator has booked them) and
    // including them in MRP is the whole point of the toggle.
    if (!opts.includeUnconfirmed && o.confirmed === false && !o.__fromPipelineDraft && !o.__fromCompanionRule) { skipped.unconfirmed++; continue; }
    if (o.status === "complete") { skipped.complete++; continue; }
    if (!o.start) { skipped.noStart++; continue; }
    if (o.start > horizonEnd) { skipped.outsideHorizon++; continue; }
    if (excludeBeforeDate && o.start < excludeBeforeDate) { skipped.excludedByDate++; continue; }

    const grossPlannedQty = o.total || (o.qty || 0) * (o.batches || 1);
    if (!o.sku || grossPlannedQty <= 0) { skipped.noBom++; continue; }

    // Packout-scheduling guard: MOs scheduled on a packout machine but coded
    // against a WIP1 formula SKU (or an RM/PK/VC) would expand the wrong BOM
    // and generate spurious raw-ingredient demand. The correct code is a
    // FG-* or a numbered-WIP tier (WIP2-, WIP3-, ...). Skip these entirely
    // so bad scheduling doesn't pollute PO suggestions, and surface a few
    // examples so ops can fix the affected MOs.
    if (PACKOUT_MACHINES.has(o.machine) && !PACKOUT_VALID_SKU_RE.test(o.sku)) {
      skipped.packoutFormulaOnly++;
      if (packoutSkippedExamples.length < 15) {
        packoutSkippedExamples.push({
          orderId: o.orderId || o.id,
          sku: o.sku,
          machine: o.machine,
          start: o.start,
          qty: grossPlannedQty,
        });
      }
      continue;
    }

    const fgSku = extractBomSku(o.sku, bomParents);
    if (!fgSku) {
      skipped.noBom++;
      if (noBomExamples.length < 5) noBomExamples.push({ orderId: o.orderId, sku: o.sku });
      continue;
    }

    // Apply FG-level netting: if this order's FG has on-hand available and
    // we consumed some of it earlier in the FIFO pass, expand at the net
    // production qty (gross minus what FG on-hand covered). Requirements
    // carry both grossFgQty and netFgQty so trace_po_demand can show the
    // offset explicitly.
    const netEntry = opts._perOrderNet && opts._perOrderNet.get(o.id || o.orderId);
    const netPlannedQty = netEntry ? netEntry.netPlannedQty : grossPlannedQty;
    const fgOffsetKg = netEntry ? netEntry.fgOffsetKg : 0;
    if (netPlannedQty <= 0) {
      // FG on-hand fully covers this order — no new production needed, no
      // RM demand. Still record a zero-demand "requirement" so callers can
      // see the order was considered and fully offset.
      continue;
    }

    // kgPerUnit override: if Supply Settings flags this SKU as having a
    // BOM authored per-case (or per-unit-of-some-fixed-mass), the order's
    // qty in kg has to be divided down to the BOM's batch unit before
    // expansion. Without this, a 1,050-kg pouching order against a
    // qtyToProduce=1 BOM is misread as 1,050 cases and overstates
    // material need by ~case-weight (~8x for PFS).
    const overrideKgPerUnit = (opts.supply && opts.supply.perSku && opts.supply.perSku[fgSku] && opts.supply.perSku[fgSku].kgPerUnit) || null;
    const expandQty = overrideKgPerUnit ? (netPlannedQty / overrideKgPerUnit) : netPlannedQty;

    // Cutoff set for THIS order: every other planned producer. Excluding
    // fgSku itself is critical — otherwise the top-level parent would be
    // treated as its own leaf and produce zero requirements.
    const stopSet = new Set(stopAtSkusGlobal);
    stopSet.delete(fgSku);

    let expansion;
    try { expansion = expandBom(bomParents, fgSku, expandQty, { applyWastage: opts.applyWastage, stopAtSkus: stopSet }); }
    catch (e) { skipped.noBom++; continue; }

    // Record cutoffs for observability. If any WIP was stopped, tally the
    // consumption qty so we can compare against the WIP's planned production
    // and warn about partial-coverage gaps.
    for (const [wipSku, stoppedQty] of Object.entries(expansion.stoppedAt || {})) {
      const summary = wipCutoffSummary.get(wipSku) || { totalStoppedKg: 0, sourceOrderIds: [] };
      summary.totalStoppedKg += stoppedQty;
      summary.sourceOrderIds.push(o.orderId || o.id);
      wipCutoffSummary.set(wipSku, summary);
    }

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
        sourceFgQty: netPlannedQty,           // qty this contribution is based on
        sourceFgGrossQty: grossPlannedQty,    // original planned qty pre-netting
        sourceFgOffsetKg: fgOffsetKg,         // kg absorbed by FG on-hand
        // Companion-demand attribution: when this requirement traces to a
        // synthetic order generated from a companion rule, carry the driver
        // order's ID + SKU so trace_po_demand and the UI can render
        // "Companion of Order XYZ (5000 kg FG-LIQUOR)" instead of the
        // synthetic COMPANION-* id which reads as noise.
        isCompanionDemand: !!o.__fromCompanionRule,
        companionDriverOrderId: o.__driverOrderRef || null,
        companionDriverSku: o.__driverSku || null,
        companionDriverQty: o.__driverQty || null,
        companionRuleId: o.__ruleId || null,
      });
    }
  }
  // Summary of FG offsets applied — attached to the return so run_mrp can
  // surface "we skipped X kg of production because FG stock covered it" in
  // its response, which explains the delta vs the pre-netting demand.
  const fgNettingSummary = [];
  if (netFgOnHand) {
    for (const [fgSku, info] of fgNetted) {
      if (info.totalConsumed > 0) {
        const oh = (opts.onHandBySku || {})[fgSku] || {};
        fgNettingSummary.push({
          fgSku,
          // Raw on-hand breakdown so the caller can distinguish "we have
          // 5,000 kg physically" from "only 438 kg is free for netting
          // because 4,562 kg is already promised to open SOs". Otherwise
          // it looks like FIFO ordering is broken when it's actually
          // working correctly against a smaller-than-expected pool.
          rawOnHandKg: Math.round(Number(oh.onHand || 0) * 1000) / 1000,
          allocatedToSalesOrdersKg: Math.round(Number(oh.allocated || 0) * 1000) / 1000,
          startingAvailableKg: Math.round(Number(oh.available || 0) * 1000) / 1000,
          totalConsumedKg: Math.round(info.totalConsumed * 1000) / 1000,
          availableRemainingKg: Math.round(info.availableRemaining * 1000) / 1000,
          ordersOffset: info.consumedByOrder.size,
          // Per-order audit trail — for each order that absorbed some
          // offset, in FIFO order. If ordersOffset shows a count higher
          // than the caller expected, this array tells them WHICH orders
          // grabbed the slack. Includes real MOs (isPipelineDraft=false)
          // and pipeline synths (isPipelineDraft=true) so it's obvious
          // when a real scheduled MO is competing for the pool.
          offsetBreakdown: info.offsetBreakdown || [],
        });
      }
    }
  }

  // Summary of WIP cutoffs applied — for each WIP that got treated as a leaf
  // during downstream MO expansions, report:
  //   plannedProductionKg — how much of this WIP is being made by other in-scope MOs
  //   consumedKg          — how much downstream MOs' BOMs asked for
  //   coverageGapKg       — max(0, consumed - planned - WIP on-hand); if > 0, ops
  //                         has under-scheduled production of this WIP and the RM
  //                         cost of the gap is NOT reflected in this MRP run
  const wipNettingSummary = [];
  const wipOnHandBySku = opts.onHandBySku || {};
  for (const [wipSku, info] of wipCutoffSummary) {
    const planned = plannedProducersBySku.get(wipSku) || 0;
    const oh = wipOnHandBySku[wipSku];
    const availableOnHand = oh ? Math.max(0, Number(oh.available || 0)) : 0;
    const gap = Math.max(0, info.totalStoppedKg - planned - availableOnHand);
    wipNettingSummary.push({
      wipSku,
      consumedKg: Math.round(info.totalStoppedKg * 1000) / 1000,
      plannedProductionKg: Math.round(planned * 1000) / 1000,
      onHandKg: Math.round(availableOnHand * 1000) / 1000,
      coverageGapKg: Math.round(gap * 1000) / 1000,
      downstreamOrderCount: new Set(info.sourceOrderIds).size,
    });
  }
  return { requirements, skipped, noBomExamples, packoutSkippedExamples, fgNettingSummary, wipNettingSummary };
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
function allocateAndPlan(requirements, onHandBySku, supply, today, poHorizonEndDate) {
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
    const allocations = [];
    // Per-shortfall list — each item is one demand event the supply pool
    // couldn't cover. Keeping these granular lets us bucket POs by each
    // shortfall's own must-order-by date instead of lumping the whole
    // forward window into one inflated PO dated by the earliest shortage.
    const shortfalls = [];

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
        const mustOrderBy = lead.leadTimeDays != null ? mrpAddDays(r.neededByDate, -lead.leadTimeDays) : null;
        shortfalls.push({
          qty: shortage,
          neededByDate: r.neededByDate,
          mustOrderByDate: mustOrderBy,
          sourceOrderId: r.sourceOrderId,
        });
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

    // Suggest PO(s) if there are shortfalls AND the SKU isn't contract-managed.
    // Bucket shortfalls by whether their own mustOrderByDate falls within the
    // user's PO horizon — each bucket becomes its own PO so far-future demand
    // doesn't inflate the in-window commitment.
    if (totalShort > 0 && !lead.isContract) {
      const buckets = new Map(); // "in-window" | "deferred" → array of shortfalls
      for (const sf of shortfalls) {
        const bucket = (poHorizonEndDate && sf.mustOrderByDate && sf.mustOrderByDate > poHorizonEndDate)
          ? "deferred"
          : "in-window";
        if (!buckets.has(bucket)) buckets.set(bucket, []);
        buckets.get(bucket).push(sf);
      }
      // Emit PO per bucket — order: in-window first, then deferred
      for (const bucketName of ["in-window", "deferred"]) {
        const list = buckets.get(bucketName);
        if (!list || !list.length) continue;
        const bucketQty = list.reduce((s, x) => s + x.qty, 0);
        const earliestNeed = list.reduce((min, x) =>
          !min || x.neededByDate < min ? x.neededByDate : min, null);
        const earliestMustOrderBy = list.reduce((min, x) =>
          !min || (x.mustOrderByDate && x.mustOrderByDate < min) ? x.mustOrderByDate : min, null);
        suggestedPOs.push({
          sku,
          name: onHand.name || "",
          qtyToOrder: Math.ceil(bucketQty),
          earliestNeedDate: earliestNeed,
          leadTimeDays: lead.leadTimeDays,
          leadTimeSource: lead.source,
          mustOrderByDate: earliestMustOrderBy,
          isOverdue: earliestMustOrderBy != null && earliestMustOrderBy < today,
          projectedReceiptDate: lead.leadTimeDays != null ? mrpAddDays(today, lead.leadTimeDays) : null,
          bucket: bucketName,
        });
      }
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
    // Optional "exclude orders with start < this date" filter — accepts
    // YYYY-MM-DD. Empty/invalid → filter disabled.
    const excludeBeforeRaw = String(req.query.excludeBefore || "").trim();
    const excludeBeforeDate = /^\d{4}-\d{2}-\d{2}$/.test(excludeBeforeRaw) ? excludeBeforeRaw : null;
    const includeDrafts = req.query.includeDrafts === "1" || req.query.includeDrafts === "true";
    const includeCompanions = req.query.includeCompanions === "1" || req.query.includeCompanions === "true";
    // PO horizon — how far ahead the user is willing to commit purchase
    // orders. Suggested POs whose mustOrderByDate falls beyond this
    // window are deferred to a separate bucket (visible in the response
    // but excluded from headline KPIs). Default 30d; 0 or negative
    // disables the cap (legacy behavior).
    const poHorizonDays = Math.max(0, Math.min(365, parseInt(req.query.poHorizonDays, 10) || 30));
    const today = new Date().toISOString().slice(0, 10);
    const poHorizonEndDate = poHorizonDays > 0 ? mrpAddDays(today, poHorizonDays) : null;

    const { orders, bomParents, supply, onHandBySku, costsBySku, costsLastSync } = getMrpInputs();

    // When the toggle is on, synthesize order-shaped objects from each
    // pipeline draft and concat them with the real orders. Synthetic
    // entries are flagged with __fromPipelineDraft so buildRequirements
    // can bypass the unconfirmed filter (drafts are always unconfirmed
    // by definition).
    let mrpOrders = orders;
    let draftsCount = 0;
    if (includeDrafts) {
      const pipelineBlob = readData("vf_pipeline_drafts");
      const drafts = (pipelineBlob && pipelineBlob.drafts) || [];
      const synth = drafts.map(synthPipelineDraftAsOrder).filter(Boolean);
      draftsCount = synth.length;
      mrpOrders = orders.concat(synth);
    }

    // Companion-demand synthesizer runs AFTER draft synthesis so a pipeline
    // draft for a liquor SKU can also trigger companion flavor demand. Skips
    // if disabled or no rules configured.
    let companionsCount = 0;
    if (includeCompanions) {
      const cblob = _companionRulesBlob();
      const synths = generateCompanionOrders(mrpOrders, cblob.rules || []);
      companionsCount = synths.length;
      if (synths.length) mrpOrders = mrpOrders.concat(synths);
    }

    const netFgOnHandFlag = req.query.netFgOnHand !== "false"; // default true
    const { requirements, skipped, noBomExamples, packoutSkippedExamples, fgNettingSummary, wipNettingSummary } = buildRequirements(mrpOrders, bomParents, {
      today, horizonDays, includeUnconfirmed, applyWastage, excludeBeforeDate, supply,
      onHandBySku, netFgOnHand: netFgOnHandFlag,
    });
    const { skuResults, suggestedPOs, atRiskOrders } = allocateAndPlan(requirements, onHandBySku, supply, today, poHorizonEndDate);

    // Enrich each suggested PO with $$ — unit cost from Cin7 product cache,
    // line cost = unitCost × qtyToOrder. Mark missing-cost SKUs explicitly so
    // the UI can flag them rather than silently treating them as $0.
    //
    // The PO horizon split happens here too: POs whose mustOrderByDate is
    // beyond the cap move into deferredPOs and don't count toward the
    // headline KPIs. dollarsByMonth still aggregates across both so the
    // monthly cash-flow projection stays useful as a forward look.
    let totalDollars = 0;          // in-window only
    let overdueDollars = 0;        // in-window only
    let missingCostCount = 0;      // in-window only (deferred tracked separately)
    let deferredDollars = 0;
    let deferredMissingCostCount = 0;
    const dollarsByMonth = {}; // YYYY-MM -> { total, overdue, count }
    const inWindowPOs = [];
    const deferredPOs = [];
    for (const po of suggestedPOs) {
      const c = costsBySku[po.sku];
      const unitCost = c && c.averageCost > 0 ? c.averageCost : null;
      po.unitCost = unitCost;
      po.lineCost = unitCost != null ? unitCost * po.qtyToOrder : null;
      po.costMissing = unitCost == null;
      po.uom = (c && c.uom) || null;
      const deferred = poHorizonEndDate && po.mustOrderByDate && po.mustOrderByDate > poHorizonEndDate;
      po.deferredByHorizon = !!deferred;
      if (deferred) {
        deferredPOs.push(po);
        if (po.costMissing) deferredMissingCostCount++;
        else deferredDollars += po.lineCost;
      } else {
        inWindowPOs.push(po);
        if (po.costMissing) missingCostCount++;
        else {
          totalDollars += po.lineCost;
          if (po.isOverdue) overdueDollars += po.lineCost;
        }
      }
      // dollarsByMonth spans both — gives the user a forward cash flow view
      if (po.lineCost != null) {
        const bucketDate = po.mustOrderByDate || po.earliestNeedDate;
        if (bucketDate) {
          const month = bucketDate.slice(0, 7);
          if (!dollarsByMonth[month]) dollarsByMonth[month] = { month, total: 0, overdue: 0, count: 0 };
          dollarsByMonth[month].total += po.lineCost;
          if (po.isOverdue && !deferred) dollarsByMonth[month].overdue += po.lineCost;
          dollarsByMonth[month].count += 1;
        }
      }
    }
    const dollarsByMonthArr = Object.values(dollarsByMonth).sort((a, b) => a.month.localeCompare(b.month));

    res.json({
      ok: true,
      runAt: new Date().toISOString(),
      today,
      settings: { includeUnconfirmed, applyWastage, horizonDays, excludeBeforeDate, includeDrafts, draftsCount, includeCompanions, companionsCount, poHorizonDays, poHorizonEndDate, netFgOnHand: netFgOnHandFlag },
      fgNettingSummary: fgNettingSummary || [],
      wipNettingSummary: wipNettingSummary || [],
      summary: {
        ordersConsidered: mrpOrders.length - (skipped.unconfirmed + skipped.complete + skipped.noStart + skipped.outsideHorizon + skipped.noBom + skipped.excludedByDate + (skipped.packoutFormulaOnly || 0)),
        ordersSkipped: skipped,
        noBomExamples,
        packoutSkippedExamples,
        requirementCount: requirements.length,
        skuCount: skuResults.length,
        atRiskOrderCount: atRiskOrders.length,
        // PO counts — headline KPIs reflect in-window only
        suggestedPoCount: inWindowPOs.length,
        overduePoCount: inWindowPOs.filter(p => p.isOverdue).length,
        deferredPoCount: deferredPOs.length,
        // $$ planning
        currency: "USD",
        totalDollars: Math.round(totalDollars * 100) / 100,
        overdueDollars: Math.round(overdueDollars * 100) / 100,
        deferredDollars: Math.round(deferredDollars * 100) / 100,
        missingCostCount,
        deferredMissingCostCount,
        costsLastSync,
        dollarsByMonth: dollarsByMonthArr,
      },
      suggestedPOs: inWindowPOs,
      deferredPOs,
      atRiskOrders,
      skuResults,
    });
  } catch (e) {
    console.error("[mrp] Run failed:", e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── Traceability (SQF audit): line-level lot / SKU / MO / PO lookups ────────
//
// Ships as a matt-planner-side capability while FMDS/Masterdata ownership is
// worked out. The daily "Inventory Movement Details" report is exploded into
// per-line rows in the vf_inventory_movements blob and indexed in-memory for
// O(1) lot/SKU/ref lookups. Same shape and endpoints as the FMDS-side design
// so we can retire this cleanly once Masterdata's UI meets audit needs.

// vf_inventory_movements blob shape:
//   { lines: [{...movement...}], last_import: { source_file, at, min_date, max_date, count } }

let _movementIndex = null;

function rebuildMovementIndex() {
  const blob = readData("vf_inventory_movements") || { lines: [] };
  const idx = {
    byBatch: new Map(),
    bySku: new Map(),
    byRef: new Map(),
    byRefNumber: new Map(),
    all: blob.lines,
  };
  for (const m of blob.lines) {
    if (m.batch)      pushMap(idx.byBatch, m.batch, m);
    if (m.sku)        pushMap(idx.bySku, m.sku, m);
    if (m.reference)  pushMap(idx.byRef, m.reference, m);
    if (m.ref_number) pushMap(idx.byRefNumber, m.ref_number, m);
  }
  _movementIndex = idx;
  return idx;
}
function pushMap(map, key, val) {
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(val);
}
function getMovementIndex() {
  if (!_movementIndex) rebuildMovementIndex();
  return _movementIndex;
}

// Replace-by-date-range ingest. Removes any existing rows in [minDate, maxDate]
// then appends the fresh ones. Idempotent for a given file. Called by both the
// manual upload endpoint and the daily gdrive-sync path.
function persistMovementLines(newLines, minDate, maxDate, sourceFile) {
  const blob = readData("vf_inventory_movements") || { lines: [] };
  const nowIso = new Date().toISOString();
  const kept = blob.lines.filter(m => m.movement_date < minDate || m.movement_date > maxDate);
  const stamped = newLines.map(l => ({ ...l, source_file: sourceFile || null, source_upload_at: nowIso }));
  const merged = kept.concat(stamped).sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
  writeData("vf_inventory_movements", {
    lines: merged,
    last_import: { source_file: sourceFile || null, at: nowIso, min_date: minDate, max_date: maxDate, count: newLines.length },
  });
  rebuildMovementIndex();
  return { deleted: blob.lines.length - kept.length, inserted: newLines.length, total: merged.length };
}

// Rebuild index on load so first request doesn't pay startup cost.
try { rebuildMovementIndex(); } catch (_) { /* no blob yet */ }

// POST /api/traceability/movements/import — admin XLSX upload, backfill path.
app.post("/api/traceability/movements/import", requireAdmin, upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, error: "No file uploaded (field name must be 'file')" });
  let parsed;
  try { parsed = parseMovementFile(req.file.buffer); }
  catch (e) { return res.status(400).json({ ok: false, error: "Parse failed: " + e.message }); }
  if (!parsed.lines.length) return res.status(400).json({ ok: false, error: "File contained no valid movement rows" });
  try {
    const counts = persistMovementLines(parsed.lines, parsed.min_date, parsed.max_date, req.file.originalname || null);
    res.json({
      ok: true,
      filename: req.file.originalname,
      row_count: parsed.row_count,
      date_range: { from: parsed.min_date, to: parsed.max_date },
      counts,
    });
  } catch (e) {
    console.error("[traceability] persist failed:", e);
    res.status(500).json({ ok: false, error: "Persist failed: " + e.message });
  }
});

// GET /api/traceability/lot/:code — every movement of a lot code, oldest first.
app.get("/api/traceability/lot/:code", (req, res) => {
  const code = String(req.params.code || "").trim();
  if (!code) return res.status(400).json({ ok: false, error: "Empty lot code" });
  const idx = getMovementIndex();
  const rows = (idx.byBatch.get(code) || []).slice().sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
  res.json({ ok: true, lot: code, count: rows.length, movements: annotateMovements(rows), data_quality: summarizeDataQuality(rows) });
});

// GET /api/traceability/sku/:sku — full history for a SKU (newest first, capped).
app.get("/api/traceability/sku/:sku", (req, res) => {
  const sku = String(req.params.sku || "").trim();
  if (!sku) return res.status(400).json({ ok: false, error: "Empty SKU" });
  const limit = Math.min(parseInt(req.query.limit, 10) || 500, 5000);
  const from = req.query.from || null;
  const to = req.query.to || null;
  const refTypes = (req.query.ref_type || "").split(",").map(s => s.trim().toUpperCase()).filter(Boolean);
  const idx = getMovementIndex();
  let rows = (idx.bySku.get(sku) || []).slice();
  if (from) rows = rows.filter(m => m.movement_date >= from);
  if (to) rows = rows.filter(m => m.movement_date <= to);
  if (refTypes.length) rows = rows.filter(m => refTypes.includes(m.ref_type));
  rows.sort((a, b) => (b.movement_date || "").localeCompare(a.movement_date || ""));
  const clipped = rows.slice(0, limit);
  res.json({ ok: true, sku, filters: { from, to, ref_types: refTypes }, count: clipped.length, total_matches: rows.length, limit, movements: annotateMovements(clipped), data_quality: summarizeDataQuality(rows) });
});

// GET /api/traceability/reference/:ref — full-ref or normalized-number lookup.
app.get("/api/traceability/reference/:ref", (req, res) => {
  const ref = String(req.params.ref || "").trim();
  if (!ref) return res.status(400).json({ ok: false, error: "Empty reference" });
  const isSpecific = ref.includes("/");
  const idx = getMovementIndex();
  const rows = ((isSpecific ? idx.byRef.get(ref) : idx.byRefNumber.get(ref)) || []).slice().sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
  res.json({ ok: true, reference: ref, resolved_via: isSpecific ? "reference" : "ref_number", count: rows.length, movements: annotateMovements(rows), data_quality: summarizeDataQuality(rows) });
});

// GET /api/traceability/summary — dashboard stats.
app.get("/api/traceability/summary", (_req, res) => {
  const idx = getMovementIndex();
  const all = idx.all;
  const byType = {};
  const skus = new Set();
  const lots = new Set();
  let fromDate = null, toDate = null;
  for (const m of all) {
    if (m.sku) skus.add(m.sku);
    if (m.batch) lots.add(m.batch);
    byType[m.ref_type || "null"] = (byType[m.ref_type || "null"] || 0) + 1;
    if (!fromDate || m.movement_date < fromDate) fromDate = m.movement_date;
    if (!toDate || m.movement_date > toDate) toDate = m.movement_date;
  }
  const blob = readData("vf_inventory_movements") || {};
  res.json({
    ok: true,
    totals: { rows: all.length, distinct_lots: lots.size, distinct_skus: skus.size, from_date: fromDate, to_date: toDate },
    by_ref_type: Object.entries(byType).map(([ref_type, n]) => ({ ref_type, n })).sort((a, b) => b.n - a.n),
    last_import: blob.last_import || null,
  });
});

// GET /api/traceability/stale-lots — lot-level triage for floor counts.
//
// Aggregates every movement row keyed on batch to compute per-lot:
//   * last_movement_date — anchor for "days idle"
//   * running_balance     — sum(qty_in) − sum(qty_out); the "supposedly on
//                           hand" figure that QA/ops should physically verify
//   * category            — RM/WIP/VC/PK/FG from SKU prefix
//   * last_location, last_reference — what happened last and where
//   * dq_notes            — flagged if the lot ever sits in a noisy window
//
// Filters (query params, all optional):
//   as_of              YYYY-MM-DD (defaults to today; useful for reproducing
//                      what a report from a specific date would have looked
//                      like). Anchors "days idle" and ignores movements after.
//   min_days_idle      Only include lots idle >= this many days. Default 30.
//   min_balance        Only include lots whose |balance| >= this. Default 0.1
//                      so numeric-noise lots don't dominate the list.
//   include_negative   Include lots with negative running_balance too. Default
//                      true — negatives are the "pre-window receipt missing"
//                      or over-consumed cases that also warrant floor-count.
//   category           Comma-separated filter (RM,VC,PK,WIP,FG,OTHER). Default: all.
//   limit              Cap result size. Default 500, max 5000.
//   sort               "idle_desc" (default), "idle_asc", "balance_desc",
//                      "balance_asc".
function computeLotIdleStats(idx, opts) {
  const byLot = new Map(); // batch -> aggregated stats
  const asOf = opts.asOf || null;
  for (const m of idx.all) {
    if (!m.batch) continue;
    if (asOf && m.movement_date > asOf) continue;
    let stat = byLot.get(m.batch);
    if (!stat) {
      stat = {
        batch: m.batch,
        sku: m.sku || "",
        product: m.product || "",
        unit: m.unit || "",
        category: skuCategory(m.sku),
        first_movement_date: m.movement_date,
        last_movement_date: m.movement_date,
        last_location: m.location || null,
        last_reference: m.reference || null,
        last_ref_type: m.ref_type || null,
        last_movement_type: m.movement_type || null,
        // Earliest non-null expiry across all rows for this lot. All rows
        // for one batch SHOULD carry the same expiry_date, but we take min
        // defensively so a downstream data-quality mismatch never causes us
        // to under-warn (i.e., report a later date than reality).
        expiry_date: m.expiry_date || null,
        total_in: 0,
        total_out: 0,
        movement_count: 0,
      };
      byLot.set(m.batch, stat);
    } else {
      // Update SKU if the current row is more informative (some rows have
      // empty sku but the lot has a real one elsewhere)
      if (!stat.sku && m.sku) stat.sku = m.sku;
      if (!stat.product && m.product) stat.product = m.product;
      if (!stat.unit && m.unit) stat.unit = m.unit;
      if (!stat.category || stat.category === "OTHER") stat.category = skuCategory(m.sku);
      if (m.movement_date < stat.first_movement_date) stat.first_movement_date = m.movement_date;
      if (m.movement_date > stat.last_movement_date) {
        stat.last_movement_date = m.movement_date;
        stat.last_location = m.location || null;
        stat.last_reference = m.reference || null;
        stat.last_ref_type = m.ref_type || null;
        stat.last_movement_type = m.movement_type || null;
      }
      if (m.expiry_date && (!stat.expiry_date || m.expiry_date < stat.expiry_date)) {
        stat.expiry_date = m.expiry_date;
      }
    }
    stat.total_in += Number(m.qty_in || 0);
    stat.total_out += Number(m.qty_out || 0);
    stat.movement_count += 1;
  }
  const today = asOf || new Date().toISOString().slice(0, 10);
  const results = [];
  for (const stat of byLot.values()) {
    stat.running_balance = Math.round((stat.total_in - stat.total_out) * 1000) / 1000;
    stat.days_idle = daysBetween(stat.last_movement_date, today);
    results.push(stat);
  }
  return results;
}

function daysBetween(fromIso, toIso) {
  if (!fromIso || !toIso) return null;
  const a = new Date(fromIso + "T00:00:00Z").getTime();
  const b = new Date(toIso + "T00:00:00Z").getTime();
  return Math.max(0, Math.floor((b - a) / 86400000));
}

app.get("/api/traceability/stale-lots", (req, res) => {
  const asOf = /^\d{4}-\d{2}-\d{2}$/.test(String(req.query.as_of || "")) ? req.query.as_of : null;
  const minDaysIdle = Math.max(0, parseInt(req.query.min_days_idle, 10) || 30);
  const minBalance = Math.max(0, parseFloat(req.query.min_balance || 0.1));
  const includeNegative = req.query.include_negative !== "false";
  const categories = (req.query.category || "").split(",").map(s => s.trim().toUpperCase()).filter(Boolean);
  const limit = Math.min(parseInt(req.query.limit, 10) || 500, 5000);
  const sort = String(req.query.sort || "idle_desc");

  const idx = getMovementIndex();
  const stats = computeLotIdleStats(idx, { asOf });
  const filtered = stats.filter(s => {
    if (s.days_idle == null) return false;
    if (s.days_idle < minDaysIdle) return false;
    const absBal = Math.abs(s.running_balance);
    if (absBal < minBalance) return false;
    if (!includeNegative && s.running_balance < 0) return false;
    if (categories.length && !categories.includes(s.category)) return false;
    return true;
  });

  filtered.sort((a, b) => {
    switch (sort) {
      case "idle_asc":    return a.days_idle - b.days_idle;
      case "balance_desc":return b.running_balance - a.running_balance;
      case "balance_asc": return a.running_balance - b.running_balance;
      case "idle_desc":
      default:            return b.days_idle - a.days_idle;
    }
  });

  res.json({
    ok: true,
    as_of: asOf || new Date().toISOString().slice(0, 10),
    filters: { min_days_idle: minDaysIdle, min_balance: minBalance, include_negative: includeNegative, categories, sort },
    total_matches: filtered.length,
    count: Math.min(filtered.length, limit),
    lots: filtered.slice(0, limit),
  });
});

// GET /api/traceability/expiring-lots — lots approaching (or past) expiry.
//
// Same paper-balance machinery as stale-lots, filtered by expiry_date
// instead of days_idle. Reuses computeLotIdleStats which now carries the
// per-lot expiry_date (earliest across all rows for that batch).
//
// Filters (all optional):
//   days_ahead     Look for lots expiring within this many days. Default 90.
//                  Also includes anything already expired.
//   min_balance    Only lots with running_balance >= this. Default 0.1 to
//                  filter numeric noise / depleted lots.
//   category       Comma-separated (RM,VC,PK,WIP,FG,OTHER). Default: all.
//   include_expired  Include already-expired lots. Default true.
//   limit          Cap result size. Default 500, max 5000.
//   sort           "expiry_asc" (default), "expiry_desc", "balance_desc".
app.get("/api/traceability/expiring-lots", (req, res) => {
  const asOfRaw = String(req.query.as_of || "");
  const asOf = /^\d{4}-\d{2}-\d{2}$/.test(asOfRaw) ? asOfRaw : null;
  const today = asOf || new Date().toISOString().slice(0, 10);
  const daysAhead = Math.max(0, parseInt(req.query.days_ahead, 10) || 90);
  const minBalance = Math.max(0, parseFloat(req.query.min_balance || 0.1));
  const categories = (req.query.category || "").split(",").map(s => s.trim().toUpperCase()).filter(Boolean);
  const includeExpired = req.query.include_expired !== "false";
  const limit = Math.min(parseInt(req.query.limit, 10) || 500, 5000);
  const sort = String(req.query.sort || "expiry_asc");

  const cutoff = mrpAddDays(today, daysAhead);

  const idx = getMovementIndex();
  const stats = computeLotIdleStats(idx, { asOf });

  // Bucket counters — computed pre-filter within the "has expiry + positive
  // balance" universe so the UI can surface "you have N expired lots" even
  // when the user has a specific category filter active.
  let bucketExpired = 0, bucket30 = 0, bucket90 = 0, bucketTotal = 0;

  const filtered = [];
  for (const s of stats) {
    if (!s.expiry_date) continue;
    if (s.running_balance < minBalance) continue;
    // days_until_expiry: negative = already expired
    s.days_until_expiry = daysBetween(today, s.expiry_date) != null
      ? Math.floor((new Date(s.expiry_date + "T00:00:00Z").getTime() - new Date(today + "T00:00:00Z").getTime()) / 86400000)
      : null;
    bucketTotal++;
    if (s.days_until_expiry < 0) bucketExpired++;
    else if (s.days_until_expiry <= 30) bucket30++;
    else if (s.days_until_expiry <= 90) bucket90++;

    if (!includeExpired && s.days_until_expiry < 0) continue;
    if (s.expiry_date > cutoff) continue;
    if (categories.length && !categories.includes(s.category)) continue;
    filtered.push(s);
  }

  filtered.sort((a, b) => {
    switch (sort) {
      case "expiry_desc":  return (b.expiry_date || "").localeCompare(a.expiry_date || "");
      case "balance_desc": return b.running_balance - a.running_balance;
      case "expiry_asc":
      default:             return (a.expiry_date || "").localeCompare(b.expiry_date || "");
    }
  });

  res.json({
    ok: true,
    as_of: today,
    filters: { days_ahead: daysAhead, min_balance: minBalance, categories, include_expired: includeExpired, sort },
    buckets: {
      expired: bucketExpired,
      within_30d: bucket30,
      within_90d: bucket90,
      total_with_expiry: bucketTotal,
    },
    total_matches: filtered.length,
    count: Math.min(filtered.length, limit),
    lots: filtered.slice(0, limit),
  });
});

// ── Lot balance / ledger reconciliation ──────────────────────────────────────
//
// Rolls a lot's movements into six audit-friendly buckets:
//   1. received_qty      Σ qty_in where ref_type ∈ {PO, MO, TR-in} — inbound
//   2. used_in_mfg       Σ qty_out where ref_type = MO — consumed by production
//   3. sold_to_customers Σ qty_out where ref_type ∈ {SO, FG}  — shipped
//   4. net_transferred   Σ TR-in − Σ TR-out — location moves (should ≈ 0)
//   5. net_adjusted_out  Σ ST-out − Σ ST-in — the audit-flagged bucket
//   6. remaining_paper   (1) − (2) − (3) − (4) − (5); the paper-balance answer
//
// Also carries actual_onhand and delta = actual_onhand − remaining_paper so
// callers can see reconciliation discrepancies at a glance.
//
// Note on TR classification: our movement rows carry ref_type = "TR" with the
// direction encoded in qty_in vs qty_out. TR-in adds to received, so we route
// TR-in qty into both received (row 1) AND net_transferred (row 4). Row 4 is
// the intra-lot audit signal; row 1 is the "how much came in" total. This
// avoids double-counting in row 6's reconciliation: received minus consumed
// minus shipped minus transferred (which subtracts the TR-in back out on the
// out side) minus adjusted = paper on hand.
function computeLotBalance(rows) {
  const b = {
    received_qty: 0,
    used_in_mfg: 0,
    sold_to_customers: 0,
    tr_in: 0,
    tr_out: 0,
    st_in: 0,
    st_out: 0,
    other_in: 0,
    other_out: 0,
    total_in: 0,
    total_out: 0,
    first_movement_date: null,
    last_movement_date: null,
    first_reference: null,
    first_ref_type: null,
    sku: null,
    product: null,
    unit: null,
    location: null,
    expiry_date: null,
  };
  const sorted = rows.slice().sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
  for (const m of sorted) {
    const qi = Number(m.qty_in) || 0;
    const qo = Number(m.qty_out) || 0;
    b.total_in += qi;
    b.total_out += qo;
    if (!b.first_movement_date) {
      b.first_movement_date = m.movement_date || null;
      b.first_reference = m.reference || null;
      b.first_ref_type = m.ref_type || null;
    }
    b.last_movement_date = m.movement_date || b.last_movement_date;
    b.sku = b.sku || m.sku || null;
    b.product = b.product || m.product || null;
    b.unit = b.unit || m.unit || null;
    b.location = b.location || m.location || null;
    if (!b.expiry_date && m.expiry_date) b.expiry_date = m.expiry_date;
    const rt = String(m.ref_type || "").toUpperCase();
    if (rt === "PO") {
      b.received_qty += qi;         // POs are always inbound — the "birth" of an RM lot
      b.other_out += qo;            // rare, but preserve
    } else if (rt === "MO") {
      b.received_qty += qi;         // MO outputs are the "birth" of a WIP/FG lot
      b.used_in_mfg += qo;
    } else if (rt === "TR") {
      // Location transfers are NOT new inventory — the same lot moved from
      // one bin/location to another. A well-formed lot has paired TR-out +
      // TR-in for each move, netting to zero. Excluding them from
      // received_qty prevents 6× double-counting when a lot moves through
      // 5–6 locations over its lifetime (spot-caught by matt on lot
      // RRC4290225SPSP: 19,000 kg PO → 128,000 kg reported "received"
      // because each TR-in added another 19k). net_transferred still tracks
      // the imbalance for audit visibility.
      b.tr_in += qi;
      b.tr_out += qo;
    } else if (rt === "SO" || rt === "FG") {
      b.sold_to_customers += qo;
      b.other_in += qi;             // returns, rare
    } else if (rt === "ST") {
      b.st_in += qi;
      b.st_out += qo;
    } else {
      b.other_in += qi;
      b.other_out += qo;
    }
  }
  b.net_transferred = round4(b.tr_in - b.tr_out);
  b.net_adjusted_out = round4(b.st_out - b.st_in);
  // Ratio of |net adjustments| to received. Primary triage signal for the
  // SQF audit: a lot with 21% of its receipt stock-adjusted out (i.e. gone
  // without documented consumption or shipment) needs an explanation. Null
  // when received is 0 so downstream sort/filter can distinguish "no
  // receipt data" from "0% adjusted".
  b.adjustment_pct_of_received = (b.received_qty > 0)
    ? round4(Math.abs(b.net_adjusted_out) / b.received_qty * 100)
    : null;
  // Accounting identity (derives from Σ qty_in − Σ qty_out = balance):
  //   remaining = received − used − sold − net_adjusted_out + net_transferred
  // Adding net_transferred (not subtracting) is correct because TR-in adds
  // to the paper balance and TR-out subtracts; for paired moves the net is
  // zero and the term drops out. For unpaired TR-out (rare, indicates
  // ledger error), net_transferred is negative and remaining decreases —
  // matching the physical intuition that material left the tracked system.
  b.remaining_paper = round4(
    b.received_qty - b.used_in_mfg - b.sold_to_customers - b.net_adjusted_out + b.net_transferred
  );
  // Round the aggregation buckets so ledger arithmetic is exact for audit
  b.received_qty      = round4(b.received_qty);
  b.used_in_mfg       = round4(b.used_in_mfg);
  b.sold_to_customers = round4(b.sold_to_customers);
  b.tr_in             = round4(b.tr_in);
  b.tr_out            = round4(b.tr_out);
  b.st_in             = round4(b.st_in);
  b.st_out            = round4(b.st_out);
  b.other_in          = round4(b.other_in);
  b.other_out         = round4(b.other_out);
  b.total_in          = round4(b.total_in);
  b.total_out         = round4(b.total_out);
  return b;
}
function round4(v) { return Math.round(Number(v || 0) * 10000) / 10000; }

// Build a batch → { qty, locations, sku } index from the raw Cin7 on-hand blob
// (populated by performCin7OnHandSync). We use the raw rows field which
// preserves per-batch/location detail; the top-level bySku rollup discards it.
let _onHandByBatchCache = null;
let _onHandByBatchSyncKey = null;
function getOnHandByBatch() {
  const blob = readData("inventory_onhand");
  if (!blob || !Array.isArray(blob.rows)) return null;
  // Cache invalidation keyed on lastSync — cheap and correct
  if (_onHandByBatchSyncKey === blob.lastSync && _onHandByBatchCache) return _onHandByBatchCache;
  const byBatch = new Map();
  for (const r of blob.rows) {
    const key = String(r.BatchSN || "").trim();
    if (!key) continue;
    let e = byBatch.get(key);
    if (!e) {
      e = { batch: key, sku: r.SKU || "", onHand: 0, allocated: 0, available: 0, locations: new Set() };
      byBatch.set(key, e);
    }
    e.onHand    += Number(r.OnHand)    || 0;
    e.allocated += Number(r.Allocated) || 0;
    e.available += Number(r.Available) || 0;
    if (r.Location) e.locations.add(r.Location);
  }
  _onHandByBatchCache = byBatch;
  _onHandByBatchSyncKey = blob.lastSync;
  return byBatch;
}

function enrichBalanceWithOnHand(b, batch) {
  const onHandIdx = getOnHandByBatch();
  const oh = onHandIdx && onHandIdx.get(batch);
  b.actual_onhand = oh ? round4(oh.onHand) : null;
  b.actual_allocated = oh ? round4(oh.allocated) : null;
  b.actual_available = oh ? round4(oh.available) : null;
  b.actual_locations = oh ? [...oh.locations].sort() : [];
  b.reconciliation_delta = (b.actual_onhand != null) ? round4(b.actual_onhand - b.remaining_paper) : null;
  return b;
}

// Find batch codes that look like sub-lots of a parent — same prefix
// followed by an underscore + arbitrary suffix. Voyage's re-lot process
// (typically via ST split) creates codes like "4529100000_02",
// "4529100000_14_1107", etc. Rolling these into the parent's balance
// gives the true "lifecycle" view; without it, MO consumption tagged to
// a sub-lot code is invisible under the parent.
function findRelatedSubLots(parent, idx) {
  if (!parent) return [];
  const escaped = parent.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp("^" + escaped + "_");
  const out = [];
  for (const [batch] of idx.byBatch) {
    if (batch !== parent && re.test(batch)) out.push(batch);
  }
  return out.sort();
}

// GET /api/traceability/lot-balance/:code — single lot rollup, aggregated
// over parent + auto-detected sub-lots by default.
// Query params:
//   parent_only=1   Skip sub-lot aggregation; show only the exact code.
app.get("/api/traceability/lot-balance/:code", (req, res) => {
  const code = String(req.params.code || "").trim();
  if (!code) return res.status(400).json({ ok: false, error: "Empty lot code" });
  const parentOnly = req.query.parent_only === "1" || req.query.parent_only === "true";
  const idx = getMovementIndex();
  const parentRows = idx.byBatch.get(code) || [];
  if (!parentRows.length) return res.status(404).json({ ok: false, error: `No movements found for lot ${code}` });
  const subLotCodes = parentOnly ? [] : findRelatedSubLots(code, idx);
  const rows = parentRows.slice();
  for (const sc of subLotCodes) {
    const extra = idx.byBatch.get(sc) || [];
    for (const m of extra) rows.push(m);
  }
  const balance = enrichBalanceWithOnHand(computeLotBalance(rows), code);
  balance.batch = code;
  balance.parent_only = parentOnly;
  balance.sub_lot_codes = subLotCodes;
  balance.sub_lot_movement_count = rows.length - parentRows.length;
  balance.movement_count = rows.length;
  balance.category = skuCategory(balance.sku);
  // Per-lot triage: attach the raw ST rows so the UI can render each
  // adjustment with its date + reference + qty + document. This is the
  // 80/20 answer to "what happened to this lot" — one glance tells you
  // whether the adjustment was one big write-off or a lot of small ones,
  // whether there's a document reference for backup, and whether the
  // adjustments cluster in a data-quality window (spot the audit-safe
  // ones vs. the ones actually needing follow-up).
  const stRowsRaw = rows
    .filter(m => String(m.ref_type || "").toUpperCase() === "ST")
    .slice()
    .sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""))
    // Tag which batch code this ST row was booked against so operators can
    // see when a write-off happened on a sub-lot vs the parent.
    .map(m => ({ ...m, batch: m.batch || null }));
  // Route through the noise-window annotator so lot-migration and
  // vc-via-st flags surface on each row — these often explain 80% of the
  // "why" without any documentation being missing.
  const stRows = annotateMovements(stRowsRaw);
  balance.adjustment_rows = stRows.map(m => ({
    movement_date: m.movement_date || null,
    reference: m.reference || null,
    ref_number: m.ref_number || null,
    movement_type: m.movement_type || null,
    batch: m.batch || null,
    qty_in: Number(m.qty_in) || 0,
    qty_out: Number(m.qty_out) || 0,
    net: round4((Number(m.qty_out) || 0) - (Number(m.qty_in) || 0)),
    location: m.location || null,
    document_reference: m.document_reference || null,
    cost_in: Number(m.cost_in) || 0,
    cost_out: Number(m.cost_out) || 0,
    dq_notes: m.dq_notes || null,
  }));
  res.json({ ok: true, balance });
});

// GET /api/traceability/lot-balance-report?days=90&category=...&sku=...&filter=...
//
// Bulk audit report. Default filter matches the SQF pre-assessment scope:
// "any lot with production consumption (MO qty_out) OR customer shipment
// (SO/FG qty_out) in the trailing N days". This is the auditor's zoom-in
// path — pick a shipped FG, work back through the RMs it consumed.
//
// Query params (all optional):
//   days               Trailing window for the activity filter. Default 90.
//   category           Comma-separated (RM,VC,PK,WIP,FG,OTHER). Default: all.
//   sku                Substring match on SKU. Default: none.
//   filter             "activity" (default, matches audit scope),
//                      "adjustments" (only lots with |net_adjusted_out| > 0),
//                      "high_adjustment_pct" (only lots where the adjustment
//                      is a large fraction of the receipt — see
//                      min_adjustment_pct; the "point the team at problem
//                      lots" filter),
//                      "delta" (only lots where reconciliation_delta != 0),
//                      "all" (every lot with any movement).
//   min_adjustment_pct Threshold for high_adjustment_pct filter. Default 20,
//                      matches the "more than 20%" audit call-out.
//   min_activity_qty   Minimum qty_out during window to qualify as "activity".
//                      Default 0.01 so numeric-noise lots don't dominate.
//   limit              Cap result size. Default 500, max 5000.
//   sort               "adjusted_desc" (default), "adjustment_pct_desc",
//                      "delta_desc", "received_desc", "last_movement_desc".
app.get("/api/traceability/lot-balance-report", (req, res) => {
  const days = Math.max(1, Math.min(3650, parseInt(req.query.days, 10) || 90));
  const today = new Date().toISOString().slice(0, 10);
  const activitySince = new Date();
  activitySince.setUTCDate(activitySince.getUTCDate() - days);
  const activitySinceStr = activitySince.toISOString().slice(0, 10);
  const categoryFilter = String(req.query.category || "").split(",").map(s => s.trim().toUpperCase()).filter(Boolean);
  const skuFilter = String(req.query.sku || "").trim().toUpperCase();
  const filterMode = String(req.query.filter || "activity").toLowerCase();
  const minActivity = parseFloat(req.query.min_activity_qty) || 0.01;
  const minAdjPct = parseFloat(req.query.min_adjustment_pct);
  const minAdjPctThreshold = isFinite(minAdjPct) && minAdjPct >= 0 ? minAdjPct : 20;
  const limit = Math.max(1, Math.min(5000, parseInt(req.query.limit, 10) || 500));
  const sort = String(req.query.sort || "adjusted_desc").toLowerCase();

  const idx = getMovementIndex();
  const out = [];
  for (const [batch, rows] of idx.byBatch) {
    // Activity filter: only include lots with qualifying qty_out in window
    if (filterMode === "activity") {
      let qualifies = false;
      for (const m of rows) {
        if (!m.movement_date || m.movement_date < activitySinceStr) continue;
        const qo = Number(m.qty_out) || 0;
        if (qo < minActivity) continue;
        const rt = String(m.ref_type || "").toUpperCase();
        if (rt === "MO" || rt === "SO" || rt === "FG") { qualifies = true; break; }
      }
      if (!qualifies) continue;
    }
    const balance = enrichBalanceWithOnHand(computeLotBalance(rows), batch);
    balance.batch = batch;
    balance.movement_count = rows.length;
    balance.category = skuCategory(balance.sku);

    // Category / SKU filters
    if (categoryFilter.length && !categoryFilter.includes(balance.category)) continue;
    if (skuFilter && !String(balance.sku || "").toUpperCase().includes(skuFilter)) continue;

    // Filter-mode zoom-ins
    if (filterMode === "adjustments" && Math.abs(balance.net_adjusted_out) < minActivity) continue;
    if (filterMode === "high_adjustment_pct") {
      if (balance.adjustment_pct_of_received == null) continue;
      if (balance.adjustment_pct_of_received < minAdjPctThreshold) continue;
    }
    if (filterMode === "delta") {
      if (balance.reconciliation_delta == null) continue;
      if (Math.abs(balance.reconciliation_delta) < minActivity) continue;
    }
    out.push(balance);
  }
  // Sort
  const cmp = {
    adjusted_desc:        (a, b) => Math.abs(b.net_adjusted_out) - Math.abs(a.net_adjusted_out),
    adjustment_pct_desc:  (a, b) => (b.adjustment_pct_of_received || 0) - (a.adjustment_pct_of_received || 0),
    delta_desc:           (a, b) => Math.abs(b.reconciliation_delta || 0) - Math.abs(a.reconciliation_delta || 0),
    received_desc:        (a, b) => (b.received_qty || 0) - (a.received_qty || 0),
    last_movement_desc:   (a, b) => (b.last_movement_date || "").localeCompare(a.last_movement_date || ""),
  }[sort] || ((a, b) => Math.abs(b.net_adjusted_out) - Math.abs(a.net_adjusted_out));
  out.sort(cmp);
  res.json({
    ok: true,
    as_of: today,
    activity_since: activitySinceStr,
    days,
    filter: filterMode,
    total_qualifying: out.length,
    count: Math.min(out.length, limit),
    lots: out.slice(0, limit),
  });
});

// ── Chain traversal ──────────────────────────────────────────────────────────
// FG lot → producing MO → BOM inputs consumed → recurse on each input lot.

function skuCategory(sku) {
  const s = String(sku || "").toUpperCase();
  if (s.startsWith("FG-")) return "FG";
  if (s.startsWith("WIP-")) return "WIP";
  if (s.startsWith("RM-")) return "RM";
  if (s.startsWith("VC-")) return "VC";
  if (s.startsWith("PK-")) return "PK";
  return "OTHER";
}

function findProducingMovement(idx, lot) {
  const rows = idx.byBatch.get(lot) || [];
  return rows.find(m => m.movement_type === "In" && m.ref_type === "MO") || null;
}
function findOriginMovement(idx, lot) {
  const rows = idx.byBatch.get(lot) || [];
  return rows.find(m => m.movement_type === "In" && (m.ref_type === "PO" || m.ref_type === "ST")) || null;
}
function findMoInputs(idx, moRef) {
  return (idx.byRef.get(moRef) || []).filter(m => m.movement_type === "Out");
}
function findDownstream(idx, lot) {
  return (idx.byBatch.get(lot) || []).filter(m => m.movement_type === "Out").slice().sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
}

// posByRef (optional): pass in the purchase_orders.byRef map so terminal
// RM/VC/PK leaves whose origin is a PO get enriched with the ACTUAL
// supplier from that specific PO. Unlike a SKU-default lookup, this
// reflects the real supplier on the specific PO — safe for audit
// narration. Nothing is surfaced for ST-origin lots (stock adjustments
// don't have a supplier by definition).
function buildLineageTree(idx, lot, depth, maxDepth, seen, posByRef) {
  if (depth > maxDepth) return { lot, sku: null, note: "max depth reached", inputs: [] };
  if (seen.has(lot)) return { lot, sku: null, note: "cycle detected", inputs: [] };
  seen.add(lot);
  const producing = findProducingMovement(idx, lot);
  if (!producing) {
    const origin = findOriginMovement(idx, lot);
    let supplier = null;
    if (origin && origin.ref_type === "PO" && posByRef) {
      // Normalize ref: "PO-00040/1" or similar sub-batch suffix → base PO
      const poRef = String(origin.reference || "").replace(/\/.*$/, "");
      const po = posByRef[poRef] || posByRef[origin.reference];
      if (po && po.supplier) {
        supplier = { name: po.supplier, code: po.supplierCode || null, orderDate: po.orderDate || null };
      }
    }
    return {
      lot,
      sku: origin ? origin.sku : null,
      product: origin ? origin.product : null,
      origin: origin || null,
      supplier,
      inputs: [],
    };
  }
  const inputMovements = findMoInputs(idx, producing.reference);
  const inputs = inputMovements.map(im => {
    if (!im.batch) {
      return { lot: null, sku: im.sku, product: im.product, category: skuCategory(im.sku), qty: im.qty_out, unit: im.unit, movement: im, inputs: [] };
    }
    const subtree = buildLineageTree(idx, im.batch, depth + 1, maxDepth, seen, posByRef);
    return { lot: im.batch, sku: im.sku, product: im.product, category: skuCategory(im.sku), qty: im.qty_out, unit: im.unit, movement: im, ...subtree };
  });
  return { lot, sku: producing.sku, product: producing.product, producing_mo: producing.reference, producing_date: producing.movement_date, producing_qty: producing.qty_in, unit: producing.unit, origin: producing, inputs };
}

function collectMovementsFromTree(node, acc = []) {
  if (node.origin) acc.push(node.origin);
  if (node.movement) acc.push(node.movement);
  for (const child of node.inputs || []) collectMovementsFromTree(child, acc);
  return acc;
}
function annotateTreeMovements(node) {
  const out = { ...node };
  if (node.origin) out.origin = annotateMovements([node.origin])[0];
  if (node.movement) out.movement = annotateMovements([node.movement])[0];
  if (node.inputs) out.inputs = node.inputs.map(annotateTreeMovements);
  return out;
}

app.get("/api/traceability/fg-lineage/:lot", (req, res) => {
  const lot = String(req.params.lot || "").trim();
  if (!lot) return res.status(400).json({ ok: false, error: "Empty lot code" });
  const maxDepth = Math.max(1, Math.min(10, parseInt(req.query.max_depth, 10) || 5));
  const idx = getMovementIndex();
  const posByRef = (readData("purchase_orders") || { byRef: {} }).byRef || {};
  const upstream = buildLineageTree(idx, lot, 0, maxDepth, new Set(), posByRef);
  const downstream = findDownstream(idx, lot);
  const all = collectMovementsFromTree(upstream).concat(downstream);
  res.json({
    ok: true, lot, max_depth: maxDepth,
    upstream: annotateTreeMovements(upstream),
    downstream: annotateMovements(downstream),
    data_quality: summarizeDataQuality(all),
  });
});

// ── Forward trace ────────────────────────────────────────────────────────────
// The mirror of the lineage tree: given a lot (or an MO reference), walk
// forward — this lot was consumed by which MOs, those MOs produced which
// lots, and where did those lots end up (SO shipments, transfers, further
// production, adjustments, still-on-hand). Recursive with depth cap.
//
// This is what QA needs to answer "if this raw-material lot is contaminated,
// which finished goods do we recall?" or "where did this specific batch end
// up — did it ship, is it still in stock, was it written off?".
//
// Response shape (recursive):
//   {
//     lot, sku, product, unit, category,
//     still_on_hand,       // ins − outs for THIS lot (best-effort)
//     total_in, total_out, // ins/outs summed
//     consumed_by_mos: [
//       { mo, date, qty_consumed, unit,
//         produced_outputs: [ {lot, sku, ...subtree...} ] }
//     ],
//     shipments: [{ ref: SO-*, date, qty, unit, location, doc_ref }],
//     transfers_out: [{ ref: TR-*, date, qty, unit, from, to }],
//     adjustments_out: [{ ref: ST-*, date, qty, unit }],
//     assemblies: [{ ref: FG-*, date, qty, unit, note: "UOM/consolidation" }],
//     movements: [ ... raw Out rows for reference ... ]
//   }

// sosByRef (optional): pass in the sales_orders.byRef map so SO shipments
// get enriched with the customer name + ship date from the actual SO
// (via nightly Cin7 sales-orders sync). Symmetric with buildLineageTree's
// posByRef enrichment for PO receipts.
function buildForwardTree(idx, lot, depth, maxDepth, seen, sosByRef) {
  if (depth > maxDepth) return { lot, note: "max depth reached", consumed_by_mos: [], shipments: [], transfers_out: [], adjustments_out: [], assemblies: [], still_on_hand: null };
  if (seen.has(lot)) return { lot, note: "cycle detected", consumed_by_mos: [], shipments: [], transfers_out: [], adjustments_out: [], assemblies: [], still_on_hand: null };
  seen.add(lot);

  const all = (idx.byBatch.get(lot) || []).slice().sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
  const ins = all.filter(m => m.movement_type === "In");
  const outs = all.filter(m => m.movement_type === "Out");
  const totalIn = ins.reduce((s, m) => s + Number(m.qty_in || 0), 0);
  const totalOut = outs.reduce((s, m) => s + Number(m.qty_out || 0), 0);
  const sample = ins[0] || outs[0] || {};

  const shipments = [];
  const transfersOut = [];
  const adjustmentsOut = [];
  const assemblies = [];
  const byMo = new Map();

  for (const m of outs) {
    if (m.ref_type === "MO") {
      if (!byMo.has(m.reference)) {
        byMo.set(m.reference, { mo: m.reference, date: m.movement_date, qty: 0, unit: m.unit || sample.unit || null });
      }
      const e = byMo.get(m.reference);
      e.qty += Number(m.qty_out || 0);
      if (!e.date || m.movement_date < e.date) e.date = m.movement_date;
    } else if (m.ref_type === "SO") {
      // Normalize the SO ref (strip any /N sub-line suffix) before lookup.
      const soRef = String(m.reference || "").replace(/\/.*$/, "");
      const so = sosByRef && (sosByRef[soRef] || sosByRef[m.reference]);
      shipments.push({
        ref: m.reference, date: m.movement_date, qty: Number(m.qty_out || 0), unit: m.unit,
        location: m.location, document_reference: m.document_reference,
        customer: so && so.customer ? so.customer : null,
        customer_reference: so && so.customerReference ? so.customerReference : null,
        so_order_date: so && so.orderDate ? so.orderDate : null,
        so_ship_date: so && so.shipDate ? so.shipDate : null,
      });
    } else if (m.ref_type === "TR") {
      transfersOut.push({
        ref: m.reference, date: m.movement_date, qty: Number(m.qty_out || 0), unit: m.unit,
        from: m.location, // ORM movement.location = where it left from on Out rows
      });
    } else if (m.ref_type === "ST") {
      adjustmentsOut.push({
        ref: m.reference, date: m.movement_date, qty: Number(m.qty_out || 0), unit: m.unit,
        document_reference: m.document_reference,
      });
    } else if (m.ref_type === "FG") {
      assemblies.push({
        ref: m.reference, date: m.movement_date, qty: Number(m.qty_out || 0), unit: m.unit,
        note: "Cin7 Assembly (UOM conversion / lot consolidation — paper-only movement)",
      });
    }
  }

  // For each consuming MO, find the outputs of that same MO and recurse.
  const consumedByMos = [];
  for (const entry of byMo.values()) {
    const outputs = (idx.byRef.get(entry.mo) || []).filter(m => m.movement_type === "In");
    const outputTrees = outputs.map(out => {
      if (!out.batch) {
        return {
          lot: null, sku: out.sku, product: out.product, unit: out.unit,
          date: out.movement_date, qty_produced: Number(out.qty_in || 0),
          note: "output row has no batch code — chain terminates here",
          consumed_by_mos: [], shipments: [], transfers_out: [], adjustments_out: [], assemblies: [], still_on_hand: null,
        };
      }
      const subtree = buildForwardTree(idx, out.batch, depth + 1, maxDepth, seen, sosByRef);
      return {
        lot: out.batch, sku: out.sku, product: out.product, unit: out.unit,
        date: out.movement_date, qty_produced: Number(out.qty_in || 0),
        category: skuCategory(out.sku),
        ...subtree,
      };
    });
    consumedByMos.push({
      mo: entry.mo, date: entry.date, qty_consumed: entry.qty, unit: entry.unit,
      produced_outputs: outputTrees,
    });
  }

  return {
    lot,
    sku: sample.sku || null,
    product: sample.product || null,
    unit: sample.unit || null,
    category: skuCategory(sample.sku),
    total_in: Math.round(totalIn * 1000) / 1000,
    total_out: Math.round(totalOut * 1000) / 1000,
    still_on_hand: Math.round((totalIn - totalOut) * 1000) / 1000,
    consumed_by_mos: consumedByMos,
    shipments,
    transfers_out: transfersOut,
    adjustments_out: adjustmentsOut,
    assemblies,
    movements_out: outs,
  };
}

// Collect all movement rows referenced anywhere in the forward tree — for
// data-quality summarization.
function collectMovementsFromForwardTree(node, acc = []) {
  if (!node) return acc;
  if (node.movements_out) for (const m of node.movements_out) acc.push(m);
  for (const c of (node.consumed_by_mos || [])) {
    for (const o of (c.produced_outputs || [])) collectMovementsFromForwardTree(o, acc);
  }
  return acc;
}

// GET /api/traceability/forward/:query — accepts a lot code OR an MO ref
// (with or without the /N batch suffix).
app.get("/api/traceability/forward/:query", (req, res) => {
  const q = String(req.params.query || "").trim();
  if (!q) return res.status(400).json({ ok: false, error: "Empty query" });
  const maxDepth = Math.max(1, Math.min(10, parseInt(req.query.max_depth, 10) || 5));
  const idx = getMovementIndex();
  const sosByRef = (readData("sales_orders") || { byRef: {} }).byRef || {};

  // Ambiguous case: a string like "MO-00XXX/190" could be either an MO
  // reference (MO-00XXX batch 190) or a lot code (where /190 is a Julian
  // date used in newer naming conventions). Older lot codes like
  // "MO-00539/2-10-22-2025" also start with MO- but have extra dashes so
  // they're syntactically distinguishable. Rather than guess by shape,
  // check the data: if the query is present as a batch code, treat it as
  // a lot; otherwise fall back to MO-ref detection. Lot codes are unique
  // identifiers so this preference is safe.
  const strictMoPattern = /^MO-\d+(\/\d+)?$/i.test(q);
  const existsAsBatch = idx.byBatch.has(q);
  const isMoRef = strictMoPattern && !existsAsBatch;
  if (isMoRef) {
    // For an MO ref, forward-trace means: find the outputs of that MO and
    // walk forward from each output lot.
    const specific = q.includes("/");
    const rows = specific ? (idx.byRef.get(q) || []) : (idx.byRefNumber.get(q) || []);
    const outputs = rows.filter(m => m.movement_type === "In");
    const seen = new Set();
    const outputTrees = outputs.map(out => {
      if (!out.batch) {
        return {
          lot: null, sku: out.sku, product: out.product, unit: out.unit,
          date: out.movement_date, qty_produced: Number(out.qty_in || 0),
          note: "output row has no batch code",
          consumed_by_mos: [], shipments: [], transfers_out: [], adjustments_out: [], assemblies: [], still_on_hand: null,
        };
      }
      const subtree = buildForwardTree(idx, out.batch, 0, maxDepth, seen, sosByRef);
      return {
        lot: out.batch, sku: out.sku, product: out.product, unit: out.unit,
        date: out.movement_date, qty_produced: Number(out.qty_in || 0),
        category: skuCategory(out.sku),
        ...subtree,
      };
    });
    const allMovements = [];
    for (const t of outputTrees) collectMovementsFromForwardTree(t, allMovements);
    return res.json({
      ok: true, query: q, kind: "mo", max_depth: maxDepth,
      resolved_via: specific ? "reference" : "ref_number",
      outputs_produced: outputTrees,
      data_quality: summarizeDataQuality(allMovements),
    });
  }

  // Lot code path.
  const tree = buildForwardTree(idx, q, 0, maxDepth, new Set(), sosByRef);
  const allMovements = collectMovementsFromForwardTree(tree);
  res.json({
    ok: true, query: q, kind: "lot", max_depth: maxDepth,
    forward: tree,
    data_quality: summarizeDataQuality(allMovements),
  });
});

app.get("/api/traceability/rm-history/:query", (req, res) => {
  const q = String(req.params.query || "").trim();
  if (!q) return res.status(400).json({ ok: false, error: "Empty query" });
  const isSku = /^(RM|VC|PK|WIP|FG)-/i.test(q);
  const idx = getMovementIndex();
  const rows = ((isSku ? idx.bySku.get(q) : idx.byBatch.get(q)) || []).slice().sort((a, b) => (a.movement_date || "").localeCompare(b.movement_date || ""));
  let running = 0;
  const timeline = rows.map(mv => {
    running += Number(mv.qty_in || 0) - Number(mv.qty_out || 0);
    return { ...mv, running_balance: running };
  });
  const grouped = {
    receipts:     timeline.filter(m => m.movement_type === "In"  && m.ref_type === "PO"),
    consumptions: timeline.filter(m => m.movement_type === "Out" && m.ref_type === "MO"),
    adjustments:  timeline.filter(m => m.ref_type === "ST"),
    transfers:    timeline.filter(m => m.ref_type === "TR"),
    other:        timeline.filter(m => !["PO", "MO", "ST", "TR"].includes(m.ref_type)),
  };
  res.json({
    ok: true, query: q, resolved_via: isSku ? "sku" : "batch",
    count: timeline.length, final_balance: running,
    timeline: annotateMovements(timeline),
    grouped: {
      receipts:     annotateMovements(grouped.receipts),
      consumptions: annotateMovements(grouped.consumptions),
      adjustments:  annotateMovements(grouped.adjustments),
      transfers:    annotateMovements(grouped.transfers),
      other:        annotateMovements(grouped.other),
    },
    data_quality: summarizeDataQuality(timeline),
  });
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
