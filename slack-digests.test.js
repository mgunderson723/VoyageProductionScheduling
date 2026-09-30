const Slack = require("./lib/slack-digests");

const run = (moRef, date, outputs, extra = {}) => Object.assign({
  moRef, completionDate: date, fgSKU: "FG-1", fgProduct: "Thing", fgCategory: "Chocolate",
  workCenter: "West Mac", outputs,
}, extra);
const kg = (qty, category = "Chocolate") => ({ sku: "FG-1", qty, unit: "kg", category });
const zeroFlag = (moRef, date) => ({
  moRef, completionDate: date, fgSKU: "WIP-1", fgProduct: "Liquor",
  flagged: [{ sku: "RM-110000-00", product: "Cane Sugar", expected: 805, actual: 0, unit: "kg" }], siblings: [],
});
const mbFlag = (moRef, date) => ({
  moRef, completionDate: date, fgSKU: "WIP-2", fgProduct: "Liquor",
  inputMassKg: 4100, outputMassKg: 4300, deficitKg: 200, yieldPct: 104.88,
});

function fakeStore(initial) {
  const data = JSON.parse(JSON.stringify(initial));
  return { data, read: k => (data[k] === undefined ? null : JSON.parse(JSON.stringify(data[k]))), write: (k, v) => { data[k] = v; } };
}
const ENV = { SLACK_WEBHOOK_URL: "https://hooks.slack.com/services/T0/B0/xyz", APP_BASE_URL: "https://planner.example.com/" };
const NOW = new Date("2026-09-30T07:05:00Z");

const payloadOk = p => {
  expect(p.blocks.length).toBeLessThanOrEqual(50);
  p.blocks.filter(b => b.type === "section").forEach(b => expect(b.text.text.length).toBeLessThanOrEqual(3000));
};

describe("Error digest", () => {
  const blob = { flagged: [zeroFlag("MO-1/1", "2026-09-29")], flaggedMassBalance: [mbFlag("MO-2/1", "2026-09-29"), mbFlag("MO-3/1", "2026-09-28")] };

  it("posts only flags that are new and not dismissed", () => {
    const d = Slack.buildErrorDigest({ blob, dismissed: { "MO-3/1": {} }, notified: { "zero:MO-1/1": "x" }, appUrl: "https://a" });
    expect(d.keys).toEqual(["mb:MO-2/1"]);
    expect(d.payload.text).toMatch(/1 new flag/);
    const txt = Slack.payloadToText(d.payload);
    expect(txt).toMatch(/MO-2\/1/);
    expect(txt).not.toMatch(/MO-1\/1|MO-3\/1/);
    expect(txt).toMatch(/Open Error reporting →/);
    payloadOk(d.payload);
  });

  it("stays silent when nothing is new", () => {
    const d = Slack.buildErrorDigest({ blob, notified: { "zero:MO-1/1": "x", "mb:MO-2/1": "x", "mb:MO-3/1": "x" } });
    expect(d.payload).toBeNull();
  });

  it("caps long lists and says how many more", () => {
    const many = { flagged: [], flaggedMassBalance: Array.from({ length: 23 }, (_, i) => mbFlag(`MO-${i}/1`, "2026-09-29")) };
    const d = Slack.buildErrorDigest({ blob: many, firstRun: true });
    const txt = Slack.payloadToText(d.payload);
    expect(txt).toMatch(/\+13 more/);
    expect(txt).toMatch(/First report/);
    payloadOk(d.payload);
  });
});

describe("Production report", () => {
  const blob = {
    detailFailures: 0,
    allCompletedRuns: [
      run("MO-10/1", "2026-09-29", [kg(4300, "CBE liquor")]),
      run("MO-11/1", "2026-09-29", [kg(900), { sku: "FG-2", qty: 40, unit: "Each", category: "Chocolate" }]),
      run("MO-12/1", "2026-09-20", [kg(1000)]),
    ],
  };

  it("first report covers only yesterday and today and pre-marks older runs", () => {
    const d = Slack.buildProductionDigest({ blob, firstRun: true, today: "2026-09-30" });
    expect(d.runKeys.sort()).toEqual(["MO-10/1", "MO-11/1"]);
    expect(d.preSeenKeys).toEqual(["MO-12/1"]);
    const txt = Slack.payloadToText(d.payload);
    expect(txt).toMatch(/2 runs entered in Cin7 since yesterday/);
    expect(txt).toMatch(/5,200 kg \+ 40 Each/);
    expect(txt).toMatch(/CBE liquor — 4,300 kg · 1 run/);
    payloadOk(d.payload);
  });

  it("later reports list only runs not seen before, whatever their completion date", () => {
    const d = Slack.buildProductionDigest({ blob, seen: { "MO-10/1": "x", "MO-11/1": "x" }, today: "2026-10-01", lastPostAt: "2026-09-30T07:05:00Z" });
    expect(d.runKeys).toEqual(["MO-12/1"]);
    expect(Slack.payloadToText(d.payload)).toMatch(/since Sep 30/);
  });

  it("says so when nothing new was entered, and flags a partial sync", () => {
    const d = Slack.buildProductionDigest({ blob: Object.assign({}, blob, { detailFailures: 2 }), seen: { "MO-10/1": 1, "MO-11/1": 1, "MO-12/1": 1 }, today: "2026-10-01" });
    const txt = Slack.payloadToText(d.payload);
    expect(txt).toMatch(/No new completed production runs/);
    expect(txt).toMatch(/2 Cin7 orders couldn't be read/);
  });
});

describe("runDigests bookkeeping", () => {
  const base = {
    production_run_errors_t7d: {
      flagged: [zeroFlag("MO-1/1", "2026-09-29")],
      flaggedMassBalance: [mbFlag("MO-2/1", "2026-09-29")],
      allCompletedRuns: [run("MO-10/1", "2026-09-29", [kg(4300)]), run("MO-12/1", "2026-09-01", [kg(10)])],
    },
    production_run_errors_dismissed: { dismissed: {} },
  };

  it("posts both, then posts nothing new the next day", async () => {
    const store = fakeStore(base);
    const sent = [];
    const poster = async (url, p) => { sent.push({ url, p }); };
    const r1 = await Slack.runDigests({ store, env: ENV, now: NOW, poster });
    expect(r1.errors.posted).toBe(true);
    expect(r1.errors.newCount).toBe(2);
    expect(r1.production.posted).toBe(true);
    expect(r1.production.newCount).toBe(1); // first report: MO-10/1 only
    expect(sent).toHaveLength(2);
    expect(Slack.payloadToText(sent[0].p)).toMatch(/https:\/\/planner\.example\.com\/\?tab=error-reporting/);

    const r2 = await Slack.runDigests({ store, env: ENV, now: new Date("2026-10-01T07:05:00Z"), poster });
    expect(r2.errors.payload).toBeNull();      // no new flags → no error post
    expect(r2.production.newCount).toBe(0);    // MO-12/1 was pre-marked seen
    expect(sent).toHaveLength(3);              // production still posts "nothing new"
  });

  it("does not mark anything as posted when Slack rejects the message", async () => {
    const store = fakeStore(base);
    const r = await Slack.runDigests({ store, env: ENV, now: NOW, poster: async () => { throw new Error("Slack rejected the post (HTTP 400)"); } });
    expect(r.errors.error).toMatch(/HTTP 400/);
    expect(store.data.slack_digest_state).toBeUndefined();
  });

  it("previews without posting or saving", async () => {
    const store = fakeStore(base);
    let calls = 0;
    const r = await Slack.runDigests({ store, env: ENV, now: NOW, doPost: false, poster: async () => { calls++; } });
    expect(calls).toBe(0);
    expect(r.errors.preview).toMatch(/Production errors: 2 new flags/);
    expect(store.data.slack_digest_state).toBeUndefined();
  });

  it("skips cleanly without a webhook, and never starts the seen-list before any sync data exists", async () => {
    const noHook = await Slack.runDigests({ store: fakeStore(base), env: {}, now: NOW, poster: async () => {} });
    expect(noHook.errors.skipped).toMatch(/No Slack webhook/);
    const empty = await Slack.runDigests({ store: fakeStore({}), env: ENV, now: NOW, poster: async () => {} });
    expect(empty.production.skipped).toMatch(/No Cin7 production data yet/);
  });

  it("uses per-post webhooks when set and ignores anything that isn't a Slack webhook", () => {
    const env = { SLACK_WEBHOOK_URL: "https://hooks.slack.com/services/a", SLACK_ERRORS_WEBHOOK_URL: "https://hooks.slack.com/services/b" };
    expect(Slack.webhookFor("errors", env)).toMatch(/\/b$/);
    expect(Slack.webhookFor("production", env)).toMatch(/\/a$/);
    expect(Slack.webhookFor("errors", { SLACK_WEBHOOK_URL: "https://evil.example.com/x" })).toBeNull();
  });
});
