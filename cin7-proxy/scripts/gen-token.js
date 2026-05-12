#!/usr/bin/env node
// Generate a fresh proxy token value for pasting into a Railway env var.
//
//   node scripts/gen-token.js <user_label>
//
// Prints the env var name + value to set on the cin7-proxy service:
//   PROXY_TOKEN_<USER>=vf_<64-hex>
//
// To revoke a user later: delete that env var in Railway and redeploy.

const crypto = require("crypto");

const label = process.argv[2];
if (!label) {
  console.error("usage: gen-token.js <user_label>");
  process.exit(1);
}

// 32 bytes = 256 bits of entropy. `vf_` prefix makes them recognizable in logs.
const value = "vf_" + crypto.randomBytes(32).toString("hex");
const name  = "PROXY_TOKEN_" + label.toUpperCase().replace(/[^A-Z0-9]/g, "_");

console.log(`\n${name}=${value}\n`);
console.log("→ Set this as an env var on the cin7-proxy Railway service, then redeploy.");
console.log("→ Hand the value (the part after =) to the user over a secure channel.");
console.log("→ Copy it now — it will not be regenerated; if lost, generate a new one.\n");
