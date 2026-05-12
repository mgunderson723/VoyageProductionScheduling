// Read-only proxy in front of Cin7 Core.
//
// Threat model: internal users have legitimate read needs but the Cin7 master
// key has read/write. This proxy holds the master key server-side, hands each
// user their own opaque token, and refuses to forward anything that isn't a GET.
//
// Flow per request:
//   1. /healthz                       → 200, no auth
//   2. Authorization: Bearer <token>  → hash, look up in tokens.json. Bad/missing → 401.
//   3. req.method === 'GET'           → otherwise 403 (this is the whole point).
//   4. Forward to https://inventory.dearsystems.com<path> with Cin7 creds injected.
//   5. Append a line to proxy_audit.jsonl no matter what happened.

const express = require("express");
const crypto  = require("crypto");
const fs      = require("fs");
const path    = require("path");

const PORT          = parseInt(process.env.PORT || "3000", 10);
const CIN7_BASE     = "https://inventory.dearsystems.com";
const CIN7_ACCOUNT  = process.env.CIN7_ACCOUNT_ID;
const CIN7_APPKEY   = process.env.CIN7_APPLICATION_KEY;
const DATA_DIR      = process.env.PROXY_DATA_DIR || path.join(__dirname, "data");
const TOKENS_FILE   = path.join(DATA_DIR, "proxy_tokens.json");
const AUDIT_FILE    = path.join(DATA_DIR, "proxy_audit.jsonl");
const FORWARD_TIMEOUT_MS = 30_000;

if (!CIN7_ACCOUNT || !CIN7_APPKEY) {
  console.error("[proxy] CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY env var not set — refusing to start.");
  process.exit(1);
}

fs.mkdirSync(DATA_DIR, { recursive: true });

// ── token store ───────────────────────────────────────────────────────────────
// Tokens are SHA-256 hashed at rest. The raw token is shown to the operator
// exactly once at issuance. Format on disk:
//   [ { token_hash, user_label, created_at, revoked_at | null, note? } ]
function loadTokens() {
  try {
    return JSON.parse(fs.readFileSync(TOKENS_FILE, "utf8"));
  } catch (e) {
    if (e.code === "ENOENT") return [];
    throw e;
  }
}

function hashToken(raw) {
  return crypto.createHash("sha256").update(raw, "utf8").digest("hex");
}

// Cached on startup, reloaded on SIGHUP so token revocations take effect
// without a redeploy. (Issuing tokens still requires the CLI script.)
let TOKENS = loadTokens();
process.on("SIGHUP", () => {
  try {
    TOKENS = loadTokens();
    console.log(`[proxy] reloaded tokens (${TOKENS.length} entries)`);
  } catch (e) {
    console.error("[proxy] token reload failed:", e.message);
  }
});

function lookupToken(raw) {
  if (!raw) return null;
  const h = hashToken(raw);
  const hit = TOKENS.find(t => t.token_hash === h);
  if (!hit || hit.revoked_at) return null;
  return hit;
}

// ── audit log ─────────────────────────────────────────────────────────────────
function audit(entry) {
  const line = JSON.stringify({ ts: new Date().toISOString(), ...entry }) + "\n";
  // Fire-and-forget append. If the disk is full we'd rather drop the audit line
  // than fail the request — alternative would be to refuse requests when audit
  // can't be written, which is stricter but riskier for availability.
  fs.appendFile(AUDIT_FILE, line, (err) => {
    if (err) console.error("[proxy] audit write failed:", err.message);
  });
}

// ── express ───────────────────────────────────────────────────────────────────
const app = express();
app.disable("x-powered-by");
app.set("trust proxy", true); // Railway terminates TLS in front of us

app.get("/healthz", (_req, res) => {
  res.json({ ok: true, tokens: TOKENS.filter(t => !t.revoked_at).length });
});

// Everything else: auth → method check → forward.
app.all("*", async (req, res) => {
  const started = Date.now();
  const ip = req.ip;

  // ── 1. auth ───────────────────────────────────────────────────────────────
  const authHeader = req.get("authorization") || "";
  const m = authHeader.match(/^Bearer\s+(.+)$/i);
  const tokenRecord = m ? lookupToken(m[1].trim()) : null;

  if (!tokenRecord) {
    audit({ user: null, ip, method: req.method, path: req.path, outcome: "auth_fail", status: 401 });
    return res.status(401).json({ error: "invalid or missing bearer token" });
  }

  const user = tokenRecord.user_label;

  // ── 2. method gate (the whole point of this service) ──────────────────────
  if (req.method !== "GET") {
    audit({ user, ip, method: req.method, path: req.path, outcome: "write_blocked", status: 403 });
    return res.status(403).json({
      error: "this proxy is read-only — only GET is permitted",
      attempted_method: req.method,
    });
  }

  // ── 3. forward to Cin7 ────────────────────────────────────────────────────
  const targetUrl = CIN7_BASE + req.originalUrl;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), FORWARD_TIMEOUT_MS);

  try {
    const upstream = await fetch(targetUrl, {
      method: "GET",
      headers: {
        "api-auth-accountid":      CIN7_ACCOUNT,
        "api-auth-applicationkey": CIN7_APPKEY,
        "accept":                  req.get("accept") || "application/json",
      },
      signal: controller.signal,
    });

    const bodyText = await upstream.text();
    const latency  = Date.now() - started;

    audit({
      user, ip, method: "GET", path: req.path, query: req.query,
      outcome: "forwarded", status: upstream.status, cin7_latency_ms: latency,
    });

    // Pass content-type through so JSON stays JSON, CSV stays CSV, etc.
    const ct = upstream.headers.get("content-type");
    if (ct) res.set("content-type", ct);
    res.status(upstream.status).send(bodyText);
  } catch (err) {
    const aborted = err.name === "AbortError";
    audit({
      user, ip, method: "GET", path: req.path, query: req.query,
      outcome: aborted ? "upstream_timeout" : "upstream_error",
      status: 502, error: err.message,
    });
    res.status(502).json({
      error: aborted ? "upstream timeout" : "upstream error",
      detail: err.message,
    });
  } finally {
    clearTimeout(timer);
  }
});

app.listen(PORT, () => {
  console.log(`[proxy] listening on ${PORT} — ${TOKENS.filter(t => !t.revoked_at).length} active tokens`);
});
