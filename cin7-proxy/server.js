// Read-only proxy in front of Cin7 Core.
//
// Threat model: internal users have legitimate read needs but the Cin7 master
// key has read/write. This proxy holds the master key server-side, hands each
// user their own opaque token, and refuses to forward anything that isn't a GET.
//
// Flow per request:
//   1. /healthz                       → 200, no auth
//   2. Authorization: Bearer <token>  → look up in TOKENS map. Bad/missing → 401.
//   3. req.method === 'GET'           → otherwise 403 (this is the whole point).
//   4. Forward to https://inventory.dearsystems.com<path> with Cin7 creds injected.
//   5. Append a line to proxy_audit.jsonl no matter what happened.
//
// Tokens are configured as env vars: PROXY_TOKEN_<USER>=vf_<random>. Each var
// adds one user. Revoking a user = delete the env var and restart. We chose
// this over a file-backed store because Railway doesn't give us a shell to
// run a token-issuance CLI inside the container, and for a handful of users
// env vars are easier to operate anyway.

const express = require("express");
const crypto  = require("crypto");
const fs      = require("fs");
const path    = require("path");

const PORT          = parseInt(process.env.PORT || "3000", 10);
const CIN7_BASE     = "https://inventory.dearsystems.com";
const CIN7_ACCOUNT  = process.env.CIN7_ACCOUNT_ID;
const CIN7_APPKEY   = process.env.CIN7_APPLICATION_KEY;
const DATA_DIR      = process.env.PROXY_DATA_DIR || path.join(__dirname, "data");
const AUDIT_FILE    = path.join(DATA_DIR, "proxy_audit.jsonl");
const FORWARD_TIMEOUT_MS = 30_000;

if (!CIN7_ACCOUNT || !CIN7_APPKEY) {
  console.error("[proxy] CIN7_ACCOUNT_ID or CIN7_APPLICATION_KEY env var not set — refusing to start.");
  process.exit(1);
}

fs.mkdirSync(DATA_DIR, { recursive: true });

// ── token store ───────────────────────────────────────────────────────────────
// Scan process.env on boot for PROXY_TOKEN_<USER>=<value>. We hash the values
// once at startup so the in-memory map doesn't hold raw tokens — only the
// hashes — which slightly limits exposure if anything ever dumps memory.
function hashToken(raw) {
  return crypto.createHash("sha256").update(raw, "utf8").digest("hex");
}

function loadTokensFromEnv() {
  const map = new Map(); // hash → user_label
  for (const [k, v] of Object.entries(process.env)) {
    const m = k.match(/^PROXY_TOKEN_(.+)$/);
    if (!m || !v) continue;
    map.set(hashToken(v), m[1].toLowerCase());
  }
  return map;
}

let TOKENS = loadTokensFromEnv();
console.log(`[proxy] loaded ${TOKENS.size} token(s) from PROXY_TOKEN_* env vars`);

function lookupToken(raw) {
  if (!raw) return null;
  const user = TOKENS.get(hashToken(raw));
  return user || null;
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
  res.json({ ok: true, tokens: TOKENS.size });
});

// Everything else: auth → method check → forward.
app.all("*", async (req, res) => {
  const started = Date.now();
  const ip = req.ip;

  // ── 1. auth ───────────────────────────────────────────────────────────────
  const authHeader = req.get("authorization") || "";
  const m = authHeader.match(/^Bearer\s+(.+)$/i);
  const user = m ? lookupToken(m[1].trim()) : null;

  if (!user) {
    audit({ user: null, ip, method: req.method, path: req.path, outcome: "auth_fail", status: 401 });
    return res.status(401).json({ error: "invalid or missing bearer token" });
  }

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
  console.log(`[proxy] listening on ${PORT} — ${TOKENS.size} active token(s)`);
});
