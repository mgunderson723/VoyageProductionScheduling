// Daily Slack posts built from the nightly Cin7 production-run sync
// (the production_run_errors_t7d blob):
//
//   1. Error digest — Error reporting flags (zero-actual BOM lines, output
//      greater than input) that haven't been posted before and aren't
//      dismissed. Silent when there's nothing new.
//   2. Production report — completed runs that appeared in Cin7 since the
//      last report, grouped by product category. Posted every day, including
//      "nothing new" days, so a missed entry is noticeable.
//
// Posts go to Slack incoming webhooks (env vars, see webhookFor). What has
// already been posted is tracked in the slack_digest_state blob, and a key
// is only marked as posted after Slack accepts the message, so a failed post
// retries the next day. Builders are pure so they can be tested without Slack.

const STATE_KEY = "slack_digest_state";
const MAX_ITEMS_PER_SECTION = 10;
const SECTION_TEXT_LIMIT = 2900; // Slack caps a section's text at 3,000 chars
const MAX_CATEGORY_SECTIONS = 12;

const esc = s => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const fmtNum = (n, dp = 0) => (n == null || !isFinite(n)) ? "—"
  : Number(n).toLocaleString("en-US", { maximumFractionDigits: dp });
const fmtDate = iso => {
  if (!iso) return "—";
  const d = new Date(String(iso).slice(0, 10) + "T00:00:00Z");
  return isNaN(d) ? String(iso) : d.toLocaleDateString("en-US", { month: "short", day: "numeric", timeZone: "UTC" });
};
const addDays = (iso, n) => {
  const d = new Date(iso + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
};
const plural = (n, word) => n + " " + word + (n === 1 ? "" : "s");
const product = r => [r.fgSKU, r.fgProduct].filter(Boolean).map(esc).join(" ");

function clip(text) {
  if (text.length <= SECTION_TEXT_LIMIT) return text;
  const cut = text.lastIndexOf("\n", SECTION_TEXT_LIMIT - 20);
  return text.slice(0, cut > 0 ? cut : SECTION_TEXT_LIMIT - 20) + "\n_…trimmed_";
}
const section = text => ({ type: "section", text: { type: "mrkdwn", text: clip(text) } });
const context = text => ({ type: "context", elements: [{ type: "mrkdwn", text }] });
const header = text => ({ type: "header", text: { type: "plain_text", text: text.slice(0, 150), emoji: true } });
// Plain mrkdwn links rather than Block Kit buttons: link buttons ping the
// Slack app's interactivity URL, which an incoming-webhook app doesn't have.
const tabLink = (appUrl, tab, label) => appUrl ? `<${appUrl}/?tab=${tab}|${label} →>` : null;

function listSection(title, items, renderItem, moreText) {
  const lines = items.slice(0, MAX_ITEMS_PER_SECTION).map(renderItem);
  if (items.length > MAX_ITEMS_PER_SECTION) {
    lines.push(`_+${items.length - MAX_ITEMS_PER_SECTION} more${moreText ? " — " + moreText : ""}_`);
  }
  return section(`*${title}*\n` + lines.join("\n"));
}

// Sum a run's outputs by unit → "4,300 kg" or "1,200 kg + 40 Each".
function unitTotals(outputs) {
  const byUnit = {};
  (outputs || []).forEach(o => {
    const u = o.unit || "units";
    byUnit[u] = (byUnit[u] || 0) + (Number(o.qty) || 0);
  });
  const parts = Object.entries(byUnit).sort((a, b) => (a[0] === "kg" ? -1 : b[0] === "kg" ? 1 : 0))
    .map(([u, q]) => `${fmtNum(q, 1)} ${esc(u)}`);
  return parts.length ? parts.join(" + ") : "no output recorded";
}

// ── 1. Error digest ─────────────────────────────────────────────────────────
function buildErrorDigest({ blob, dismissed = {}, notified = {}, appUrl = null, firstRun = false }) {
  const zero = (blob && Array.isArray(blob.flagged)) ? blob.flagged : [];
  const mb = (blob && Array.isArray(blob.flaggedMassBalance)) ? blob.flaggedMassBalance : [];
  const open = g => g && g.moRef && !dismissed[g.moRef];
  const newZero = zero.filter(g => open(g) && !notified["zero:" + g.moRef]);
  const newMb = mb.filter(g => open(g) && !notified["mb:" + g.moRef]);
  const keys = newZero.map(g => "zero:" + g.moRef).concat(newMb.map(g => "mb:" + g.moRef));
  if (!keys.length) return { keys, payload: null };

  const link = tabLink(appUrl, "error-reporting", "Open Error reporting");
  const blocks = [header(`Production errors: ${plural(keys.length, "new flag")}`)];
  if (firstRun) {
    blocks.push(context("First report, so this includes every open flag from the sync window. From now on each flag is posted once."));
  }
  if (newZero.length) {
    blocks.push(listSection(`Zero-actual BOM lines (${plural(newZero.length, "run")})`, newZero, g => {
      const lines = g.flagged || [];
      const shown = lines.slice(0, 3).map(l =>
        `${esc(l.sku)}${l.product ? " " + esc(l.product) : ""} (expected ${fmtNum(l.expected, 2)} ${esc(l.unit || "")})`.trim());
      const more = lines.length > 3 ? `, +${lines.length - 3} more` : "";
      return `• *${esc(g.moRef)}* · ${product(g)} · ${fmtDate(g.completionDate)}\n    0 recorded for ${shown.join(", ")}${more}`;
    }, "see Error reporting"));
  }
  if (newMb.length) {
    blocks.push(listSection(`Output greater than input (${plural(newMb.length, "run")})`, newMb, g =>
      `• *${esc(g.moRef)}* · ${product(g)} · ${fmtDate(g.completionDate)}\n` +
      `    output ${fmtNum(g.outputMassKg, 1)} kg vs input ${fmtNum(g.inputMassKg, 1)} kg · ${fmtNum(g.deficitKg, 1)} kg short · yield ${fmtNum(g.yieldPct, 1)}%`,
      "see Error reporting"));
  }
  blocks.push(context(["Each flag is posted once. Dismiss handled flags in the app so they drop off the list.", link].filter(Boolean).join("  ·  ")));
  const text = `Production errors: ${plural(keys.length, "new flag")} (${newZero.length} zero-actual, ${newMb.length} output > input)`;
  return { keys, payload: { text, blocks } };
}

// ── 2. Production report ────────────────────────────────────────────────────
function buildProductionDigest({ blob, seen = {}, appUrl = null, firstRun = false, today, lastPostAt = null }) {
  const runs = (blob && Array.isArray(blob.allCompletedRuns)) ? blob.allCompletedRuns.filter(r => r && r.moRef) : [];
  let newRuns, preSeenKeys = [];
  if (firstRun) {
    // First report only: runs completed yesterday or today. Older runs in the
    // window are marked as seen so they don't flood the channel.
    const cutoff = addDays(today, -1);
    newRuns = runs.filter(r => (r.completionDate || "") >= cutoff);
    preSeenKeys = runs.filter(r => !((r.completionDate || "") >= cutoff)).map(r => r.moRef);
  } else {
    newRuns = runs.filter(r => !seen[r.moRef]);
  }

  // Group outputs by category (a multi-output run can span categories).
  const cats = new Map();
  const allOutputs = [];
  newRuns.forEach(r => {
    const outs = Array.isArray(r.outputs) ? r.outputs : [];
    const byCat = new Map();
    if (!outs.length) byCat.set(r.fgCategory || "Other", []);
    outs.forEach(o => {
      const c = o.category || r.fgCategory || "Other";
      if (!byCat.has(c)) byCat.set(c, []);
      byCat.get(c).push(o);
      allOutputs.push(o);
    });
    byCat.forEach((outputs, c) => {
      if (!cats.has(c)) cats.set(c, { category: c, outputs: [], runs: [] });
      const bucket = cats.get(c);
      bucket.outputs.push(...outputs);
      bucket.runs.push({ run: r, outputs });
    });
  });

  const since = firstRun ? "yesterday" : (lastPostAt ? fmtDate(lastPostAt) : "the last report");
  const blocks = [header("Cin7 production report")];
  if (!newRuns.length) {
    blocks.push(section(firstRun
      ? "No completed production runs in Cin7 for yesterday or today."
      : `No new completed production runs were entered in Cin7 since ${since}.`));
  } else {
    blocks.push(section(`*${plural(newRuns.length, "run")} entered in Cin7 since ${since}* · ${unitTotals(allOutputs)}`));
    const sorted = [...cats.values()].sort((a, b) => b.runs.length - a.runs.length || a.category.localeCompare(b.category));
    sorted.slice(0, MAX_CATEGORY_SECTIONS).forEach(c => {
      const rows = c.runs.sort((a, b) => (b.run.completionDate || "").localeCompare(a.run.completionDate || ""));
      blocks.push(listSection(`${esc(c.category)} — ${unitTotals(c.outputs)} · ${plural(c.runs.length, "run")}`, rows, x =>
        `• *${esc(x.run.moRef)}* · ${product(x.run)} · ${unitTotals(x.outputs)}` +
        `${x.run.workCenter ? " · " + esc(x.run.workCenter) : ""} · completed ${fmtDate(x.run.completionDate)}`,
        "see Production output"));
    });
    if (sorted.length > MAX_CATEGORY_SECTIONS) {
      blocks.push(context(`+${sorted.length - MAX_CATEGORY_SECTIONS} more categories in the app.`));
    }
  }
  const notes = [];
  if (blob && blob.detailFailures > 0) {
    notes.push(`:warning: ${plural(blob.detailFailures, "Cin7 order")} couldn't be read in the last sync; their runs will show up in the next report.`);
  }
  notes.push("Runs are listed once, the first time they appear in Cin7.");
  const link = tabLink(appUrl, "production-output", "Open Production output");
  if (link) notes.push(link);
  blocks.push(context(notes.join("  ·  ")));

  const text = newRuns.length
    ? `Cin7 production: ${plural(newRuns.length, "run")} entered since ${since} (${unitTotals(allOutputs)})`
    : `Cin7 production: no new runs since ${since}`;
  return { runKeys: newRuns.map(r => r.moRef), preSeenKeys, payload: { text, blocks } };
}

function buildSyncFailure(message, appUrl) {
  const link = tabLink(appUrl, "error-reporting", "Open Error reporting");
  return {
    text: "Cin7 production sync failed",
    blocks: [
      header(":warning: Cin7 production sync failed"),
      section(`Last night's sync didn't finish, so today's error and production posts were skipped.\n\`${esc(String(message || "unknown error").slice(0, 400))}\``),
      context(["An admin can retry with Sync now on the Error reporting tab.", link].filter(Boolean).join("  ·  ")),
    ],
  };
}

function buildTestMessage(appUrl) {
  return {
    text: "Voyage production planner can post here",
    blocks: [section(":white_check_mark: The Voyage production planner can post to this channel." + (appUrl ? `\n<${appUrl}|Open the planner →>` : ""))],
  };
}

// Plain-text rendering of a payload, for the admin preview.
function payloadToText(payload) {
  if (!payload) return "";
  return (payload.blocks || []).map(b => {
    if (b.type === "header" || b.type === "section") return b.text.text;
    if (b.type === "context") return b.elements.map(e => e.text).join(" ");
    return "";
  }).filter(Boolean).join("\n\n")
    .replace(/<([^|>]+)\|([^>]+)>/g, "$2 ($1)")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
}

// ── Config ──────────────────────────────────────────────────────────────────
function webhookFor(kind, env) {
  const specific = kind === "errors" ? env.SLACK_ERRORS_WEBHOOK_URL : env.SLACK_PRODUCTION_WEBHOOK_URL;
  const url = String(specific || env.SLACK_WEBHOOK_URL || "").trim();
  return /^https:\/\/hooks\.slack\.com\//.test(url) ? url : null;
}

function appBaseUrl(env) {
  const base = String(env.APP_BASE_URL || (env.RAILWAY_PUBLIC_DOMAIN ? "https://" + env.RAILWAY_PUBLIC_DOMAIN : "")).trim();
  return base ? base.replace(/\/+$/, "") : null;
}

async function postToSlack(url, payload) {
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const body = await resp.text().catch(() => "");
  // Never include the webhook URL in errors or logs — it's a credential.
  if (!resp.ok) throw new Error(`Slack rejected the post (HTTP ${resp.status}${body ? ": " + body.slice(0, 200) : ""})`);
  return body;
}

// Build (and optionally post) both digests, then record what was posted.
// store: { read(key), write(key, value) }. which: "errors" | "production" | "both".
async function runDigests({ store, env, now = new Date(), which = "both", doPost = true, poster = postToSlack }) {
  const blob = store.read("production_run_errors_t7d");
  const dismissed = (store.read("production_run_errors_dismissed") || {}).dismissed || {};
  const state = store.read(STATE_KEY) || {};
  const appUrl = appBaseUrl(env);
  const nowIso = now.toISOString();
  const today = nowIso.slice(0, 10);
  const out = {};
  let changed = false;

  if (which !== "production") {
    const prev = state.errors || null;
    const d = buildErrorDigest({ blob, dismissed, notified: (prev && prev.notified) || {}, appUrl, firstRun: !prev });
    out.errors = { newCount: d.keys.length, posted: false, payload: d.payload, preview: payloadToText(d.payload) };
    const url = webhookFor("errors", env);
    if (doPost && d.payload && !url) out.errors.skipped = "No Slack webhook is set for error posts.";
    if (doPost && d.payload && url) {
      try {
        await poster(url, d.payload);
        const notified = Object.assign({}, (prev && prev.notified) || {});
        d.keys.forEach(k => { notified[k] = nowIso; });
        // Keep keys for flags still open, plus anything posted in the last 90 days.
        const live = new Set([].concat(
          ((blob && blob.flagged) || []).map(g => "zero:" + g.moRef),
          ((blob && blob.flaggedMassBalance) || []).map(g => "mb:" + g.moRef)));
        const floor = addDays(today, -90);
        Object.keys(notified).forEach(k => { if (!live.has(k) && notified[k].slice(0, 10) < floor) delete notified[k]; });
        state.errors = { notified, lastPostAt: nowIso, lastCount: d.keys.length };
        out.errors.posted = true;
        changed = true;
      } catch (e) { out.errors.error = e.message; }
    }
  }

  if (which !== "errors" && !(blob && Array.isArray(blob.allCompletedRuns))) {
    // No sync data yet. Don't start the "seen" list from nothing, or the next
    // report would list every run in the 60-day window as new.
    out.production = { newCount: 0, posted: false, payload: null, preview: "", skipped: "No Cin7 production data yet. Run the production sync first." };
  } else if (which !== "errors") {
    const prev = state.production || null;
    const d = buildProductionDigest({ blob, seen: (prev && prev.seen) || {}, appUrl, firstRun: !prev, today, lastPostAt: prev && prev.lastPostAt });
    out.production = { newCount: d.runKeys.length, posted: false, payload: d.payload, preview: payloadToText(d.payload) };
    const url = webhookFor("production", env);
    if (doPost && !url) out.production.skipped = "No Slack webhook is set for the production report.";
    if (doPost && url) {
      try {
        await poster(url, d.payload);
        const seen = Object.assign({}, (prev && prev.seen) || {});
        d.runKeys.concat(d.preSeenKeys).forEach(k => { seen[k] = nowIso; });
        // Runs that have left the sync window can't come back, so drop them.
        const live = new Set(((blob && blob.allCompletedRuns) || []).map(r => r && r.moRef));
        Object.keys(seen).forEach(k => { if (!live.has(k)) delete seen[k]; });
        state.production = { seen, lastPostAt: nowIso, lastCount: d.runKeys.length };
        out.production.posted = true;
        changed = true;
      } catch (e) { out.production.error = e.message; }
    }
  }

  if (changed) store.write(STATE_KEY, state);
  return out;
}

module.exports = {
  STATE_KEY,
  buildErrorDigest, buildProductionDigest, buildSyncFailure, buildTestMessage,
  payloadToText, webhookFor, appBaseUrl, postToSlack, runDigests,
};
