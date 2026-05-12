#!/usr/bin/env node
// Mint, list, or revoke proxy tokens.
//
//   node scripts/issue-token.js issue <user_label> [--note "..."]
//   node scripts/issue-token.js list
//   node scripts/issue-token.js revoke <user_label>
//
// Tokens are stored hashed in data/proxy_tokens.json. The raw token is printed
// exactly once at issuance time — there is no way to recover it afterwards.
// If a user loses theirs, revoke and reissue.

const crypto = require("crypto");
const fs     = require("fs");
const path   = require("path");

const DATA_DIR    = process.env.PROXY_DATA_DIR || path.join(__dirname, "..", "data");
const TOKENS_FILE = path.join(DATA_DIR, "proxy_tokens.json");

function load() {
  try { return JSON.parse(fs.readFileSync(TOKENS_FILE, "utf8")); }
  catch (e) { if (e.code === "ENOENT") return []; throw e; }
}

function save(tokens) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(TOKENS_FILE, JSON.stringify(tokens, null, 2) + "\n");
}

function hashToken(raw) {
  return crypto.createHash("sha256").update(raw, "utf8").digest("hex");
}

function issue(label, note) {
  if (!label) die("usage: issue <user_label> [--note \"...\"]");
  const tokens = load();
  if (tokens.find(t => t.user_label === label && !t.revoked_at)) {
    die(`an active token for "${label}" already exists — revoke it first`);
  }
  // 32 bytes = 256 bits of entropy. `vf_` prefix makes them recognizable in logs.
  const raw = "vf_" + crypto.randomBytes(32).toString("hex");
  tokens.push({
    token_hash: hashToken(raw),
    user_label: label,
    created_at: new Date().toISOString(),
    revoked_at: null,
    note: note || null,
  });
  save(tokens);
  console.log(`\nIssued token for ${label}:\n`);
  console.log(`  ${raw}\n`);
  console.log("Copy it now — it will never be shown again.");
  console.log("After deploying, the proxy must be restarted (or sent SIGHUP) to pick up the new token.\n");
}

function list() {
  const tokens = load();
  if (!tokens.length) { console.log("no tokens issued yet"); return; }
  for (const t of tokens) {
    const state = t.revoked_at ? `revoked ${t.revoked_at}` : "active";
    console.log(`  ${t.user_label.padEnd(20)} ${state.padEnd(35)} created ${t.created_at}${t.note ? ` — ${t.note}` : ""}`);
  }
}

function revoke(label) {
  if (!label) die("usage: revoke <user_label>");
  const tokens = load();
  const hit = tokens.find(t => t.user_label === label && !t.revoked_at);
  if (!hit) die(`no active token found for "${label}"`);
  hit.revoked_at = new Date().toISOString();
  save(tokens);
  console.log(`Revoked token for ${label}.`);
  console.log("Restart the proxy (or send SIGHUP) for revocation to take effect.");
}

function die(msg) { console.error(msg); process.exit(1); }

// ── arg parsing ───────────────────────────────────────────────────────────────
const [, , cmd, ...rest] = process.argv;
const noteIdx = rest.indexOf("--note");
const note    = noteIdx >= 0 ? rest[noteIdx + 1] : null;
// Strip --note and its value out of positional args; otherwise keep them all.
// (Earlier version filtered on `i !== noteIdx + 1` even when noteIdx === -1,
// which silently dropped the first positional arg.)
const positional = noteIdx >= 0
  ? rest.filter((_, i) => i !== noteIdx && i !== noteIdx + 1)
  : rest;

switch (cmd) {
  case "issue":  issue(positional[0], note); break;
  case "list":   list(); break;
  case "revoke": revoke(positional[0]); break;
  default:       die("usage: issue <user_label> [--note \"...\"] | list | revoke <user_label>");
}
