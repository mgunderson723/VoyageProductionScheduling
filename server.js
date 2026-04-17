const express = require("express");
const path = require("path");
const fs = require("fs");
const multer = require("multer");
const XLSX = require("xlsx");
const Anthropic = require("@anthropic-ai/sdk");

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

function parseQty(raw) {
  if (!raw) return 0;
  const s = String(raw).replace(/,/g, "");
  const mt = s.match(/(\d+(?:\.\d+)?)\s*MT\b/i);
  if (mt) return Math.round(parseFloat(mt[1]) * 1000);
  const kg = s.match(/(\d+(?:\.\d+)?)\s*kg\b/i);
  if (kg) return Math.round(parseFloat(kg[1]));
  const n = s.match(/(\d+(?:\.\d+)?)/);
  return n ? Math.round(parseFloat(n[1])) : 0;
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
    writeData(key, req.body.value);
    res.json({ ok: true });
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
            ...attribs, status: "queued", priority: "med", due: end, notes: rawNotes,
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
          ...attribs, status: "queued", priority: "med", due: end, notes,
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

Order statuses: queued, in-progress, complete, on-hold
Order priorities: high, med, low

IMPORTANT: Before calling any write tools (shift_machine_orders, update_order_dates, update_order_status), you MUST:
1. Use get_orders to see the current state
2. Clearly describe to the user exactly what changes you plan to make
3. Wait for them to explicitly confirm (e.g. "yes", "go ahead", "confirm") before executing writes

Dates are always in YYYY-MM-DD format.`;

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
];

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
        toolUseBlocks.map(async block => ({
          type: "tool_result",
          tool_use_id: block.id,
          content: JSON.stringify(await executeAITool(block.name, block.input)),
        }))
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

    res.json({ ok: true, reply, messages: currentMessages });
  } catch (e) {
    console.error("AI chat error:", e);
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
