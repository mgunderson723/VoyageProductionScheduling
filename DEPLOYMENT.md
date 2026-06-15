# Voyage Production Scheduling — Self-Hosted Deployment Guide

Audience: Voyage Foods IT team, migrating this app off Railway onto an internal server.

This repo contains **two services** that ship together:

1. **`voyage-scheduling`** (root) — the main production planning app. Calendar, MRP, pipeline, yield analytics, error reporting, AI chat, bug reports. Used directly by Operations + Planning team.
2. **`cin7-proxy/`** (subfolder) — a read-only Cin7 API gateway. Holds the Cin7 master key server-side, hands each internal user (BI / analysts) their own Bearer token via env vars. Only forwards GET requests.

They're independent: one doesn't call the other. Both happen to live in the same repo because they were deployed together on Railway.

---

## 1. Prerequisites

| Requirement | Version |
|---|---|
| Docker Engine | 24+ (or any recent version) |
| Docker Compose | v2 (`docker compose`, not the old `docker-compose`) |
| OS | Linux strongly preferred; macOS works for dev |
| Inbound HTTPS | Required — Google Apps Script POSTs the daily inventory file to the scheduling app |
| Outbound HTTPS | Required — the scheduling app calls Cin7 + Anthropic API |
| Persistent storage | Required — both services need a writable volume |
| Reverse proxy with TLS | Strongly recommended (nginx, Caddy, Traefik, etc.) |

Both services run on Node 18 Alpine (no system dependencies beyond Docker).

---

## 2. Get the code

```bash
# Option A: clone from GitHub (preferred — easy to pull updates later)
git clone https://github.com/mgunderson723/VoyageProductionScheduling.git
cd VoyageProductionScheduling

# Option B: unzip a downloaded snapshot
unzip VoyageProductionScheduling-main.zip
cd VoyageProductionScheduling-main
```

You should see:

```
├── server.js                    ← main app entry point
├── public/                      ← frontend (HTML, JS, CSS)
├── Dockerfile                   ← main app container build
├── docker-compose.yml           ← orchestrates BOTH services
├── package.json
├── cin7-proxy/                  ← second service
│   ├── server.js
│   ├── Dockerfile
│   ├── package.json
│   └── scripts/
├── DEPLOYMENT.md                ← this file
└── data/                        ← runtime data (empty at first; populated from Railway)
```

---

## 3. Set up the data volumes

The runtime state (orders, users, BOMs, audit log, inventory snapshots, etc.) lives in JSON files inside `/app/data` in each container. We mount that to host-side directories so data survives container restarts.

```bash
mkdir -p data/voyage data/proxy
chown -R 1000:1000 data/   # node user inside the alpine image
chmod -R 750 data/
```

The compose file maps:
- `./data/voyage` → `voyage-scheduling:/app/data`
- `./data/proxy` → `cin7-proxy:/app/data`

**Critical:** if these are ephemeral (e.g., container restart wipes them), you'll lose ALL application state including user accounts, the audit log, and the production schedule. Pick durable storage and back it up.

---

## 4. Set up environment variables

Each service has its own `.env` file. **Both files contain secrets and should NEVER be committed to git.** They've been provided to IT separately via secure channel.

### `.env.voyage` (main app — at repo root)

```
PORT=3000
NODE_ENV=production

# Cin7 Core API credentials (master keys — keep secret)
CIN7_ACCOUNT_ID=<value provided separately>
CIN7_APPLICATION_KEY=<value provided separately>

# Apps Script webhook secret (must match what's set in the Google
# Apps Script that pushes the daily inventory XLSX)
INVENTORY_SYNC_SECRET=<value provided separately>

# HMAC secret for user session cookies. If unset, sessions invalidate
# on every restart. Generate with: openssl rand -hex 32
SESSION_SECRET=<value provided separately>

# Anthropic API key for the AI chat tab. Optional — if unset, chat
# returns a 503 and the rest of the app still works.
ANTHROPIC_API_KEY=<value provided separately>

# Optional — defaults to 90 if unset.
# CIN7_SYNC_DAYS=90
```

### `.env.proxy` (cin7-proxy — at repo root)

```
PORT=3000
NODE_ENV=production

# Same Cin7 master keys as the main app (the proxy holds them server-side)
CIN7_ACCOUNT_ID=<value provided separately>
CIN7_APPLICATION_KEY=<value provided separately>

# Per-user proxy tokens — one env var per authorized internal user.
# Format: PROXY_TOKEN_<USERNAME>=vf_<random_string>
# The suffix after PROXY_TOKEN_ is the username label that appears in audit logs.
# To revoke a user: delete their env var and restart the proxy.
PROXY_TOKEN_<USERNAME1>=<value provided separately>
PROXY_TOKEN_<USERNAME2>=<value provided separately>
# ...repeat per user
```

Make sure file permissions are tight:

```bash
chmod 600 .env.voyage .env.proxy
```

---

## 5. Build and start

```bash
docker compose build
docker compose up -d
```

Verify both containers come up:

```bash
docker compose ps
# Should show 2 services, both "Up" with healthcheck "healthy"

docker compose logs -f voyage-scheduling
# Look for: "[auth] SESSION_SECRET not set" warning → set it in .env.voyage
# Look for: "[Cin7 OnHand] Nightly sync scheduled at 06:30 UTC"
# Look for: "Server listening on port 3000"

docker compose logs -f cin7-proxy
# Look for: "[proxy] loaded N token(s) from PROXY_TOKEN_* env vars"
# Look for: "[proxy] listening on 3000"
```

Quick health checks from the host:

```bash
curl http://localhost:3000/health        # main app — should return JSON
curl http://localhost:3001/healthz       # proxy — should return { ok: true, tokens: N }
```

---

## 6. Restore data from Railway

**Do this AFTER you've confirmed the empty deployment works**, but BEFORE you point users + the Apps Script webhook at the new server.

Matt will export the data from Railway and hand IT a `voyage-data-backup.tgz` tarball. To import:

```bash
# Stop the main app so we don't write to data/ while restoring
docker compose stop voyage-scheduling

# Extract into the volume — the tarball is structured as `data/*.json`
tar xzf voyage-data-backup.tgz -C ./data/voyage --strip-components=1

# Sanity check — should see vf_orders.json, vf_users.json, etc.
ls -la ./data/voyage/

# Fix ownership in case extraction changed it
chown -R 1000:1000 ./data/voyage/

# Restart and verify
docker compose start voyage-scheduling
docker compose logs --tail 50 voyage-scheduling
```

Open the app in a browser. You should see the production schedule, users, and audit log all restored.

**Note:** the cin7-proxy's data dir contains only its audit log (`proxy_audit.jsonl`). Matt may or may not want to migrate this — it's historical, not load-bearing. Confirm before importing.

---

## 7. Point external integrations at the new server

The scheduling app has one inbound integration that needs to be updated:

### Google Apps Script webhook URL

A Google Apps Script bound to a Drive folder POSTs the daily Cin7 inventory XLSX to:

**Old (Railway):** `https://voyageproductionscheduling-production.up.railway.app/api/cin7/inventory-movements`

**New (your server):** `https://<your-new-domain>/api/cin7/inventory-movements`

Matt updates this in the Apps Script editor — IT doesn't need to touch it. But the Apps Script will continue posting to the OLD Railway URL until Matt makes the change, so coordinate the cutover.

### Verifying the webhook works

After Matt updates the Apps Script:

1. Wait for the next morning's automatic run (typically 06:00–07:00 UTC), or have Matt trigger the Apps Script manually
2. Check the main app logs: `docker compose logs voyage-scheduling | grep gdrive-sync`
3. You should see: `[gdrive-sync] Ingested <filename> — N rows · N new-window SKUs · ...`

---

## 8. Reverse proxy + TLS

The Node apps listen on plain HTTP on ports 3000 and 3001. **Do not expose these directly to the internet.** Put a reverse proxy in front with valid TLS.

### Minimal nginx example

```nginx
# Main scheduling app
server {
    listen 443 ssl http2;
    server_name voyage-scheduling.voyagefoods.internal;
    ssl_certificate     /etc/ssl/certs/voyage-scheduling.crt;
    ssl_certificate_key /etc/ssl/private/voyage-scheduling.key;

    client_max_body_size 12M;   # daily inventory XLSX is ~5–10 MB

    location / {
        proxy_pass http://localhost:3000;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}

# Cin7 read-only proxy
server {
    listen 443 ssl http2;
    server_name cin7-proxy.voyagefoods.internal;
    ssl_certificate     /etc/ssl/certs/cin7-proxy.crt;
    ssl_certificate_key /etc/ssl/private/cin7-proxy.key;

    location / {
        proxy_pass http://localhost:3001;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

The main app has `app.set("trust proxy", true)` so `X-Forwarded-For` is correctly honored. The proxy does too.

### Firewall rules

- **Inbound HTTPS (443)** open to the internet — specifically allowing Google Apps Script IP ranges if you can constrain it. If not, the `INVENTORY_SYNC_SECRET` header is the primary defense.
- **Outbound HTTPS (443)** open — the app calls `inventory.dearsystems.com` (Cin7), `api.anthropic.com`, and optionally Google's APIs if you ever add Sheets integration.
- **No inbound access** needed on ports 3000/3001 from outside — the reverse proxy is the only consumer.

---

## 9. Operational notes

### Backups

The `data/voyage/` and `data/proxy/` directories are the source of truth for all runtime state. Snapshot them regularly:

```bash
# Daily backup example — run from cron
tar czf /backups/voyage-$(date +%Y-%m-%d).tgz -C /path/to/repo data/
```

Rotate backups per your retention policy. The blobs are small (typically <100 MB total) so daily snapshots are cheap.

### Logs

- `docker compose logs voyage-scheduling --tail 100 -f` — main app
- `docker compose logs cin7-proxy --tail 100 -f` — proxy
- The proxy also writes `data/proxy/proxy_audit.jsonl` — one JSON line per request, useful for compliance / forensics

### Updates

To pull a new version of the code:

```bash
git pull origin main
docker compose build
docker compose up -d
```

The data volumes survive rebuilds since they're mounted from the host.

### Scheduled jobs

The main app runs three `node-cron` jobs internally:

| Time (UTC) | Job |
|---|---|
| 06:30 | Cin7 on-hand inventory sync |
| 06:45 | Cin7 product cost sync |
| 07:00 | Cin7 production-run sync (powers Error Reporting + Yield tabs) |

These start automatically when the container boots. No additional cron config needed on the host.

---

## 10. Troubleshooting

### "Session cookie invalid" / users keep getting logged out
- `SESSION_SECRET` is not set in `.env.voyage`. The app falls back to a random per-startup secret, so every restart invalidates every session.
- Generate one: `openssl rand -hex 32`, add to `.env.voyage`, restart.

### Apps Script webhook returns 401
- `INVENTORY_SYNC_SECRET` in `.env.voyage` doesn't match the value stored in the Apps Script. Verify both.

### Cin7 sync logs "401 Unauthorized"
- `CIN7_ACCOUNT_ID` or `CIN7_APPLICATION_KEY` is wrong or has been rotated. Verify in Cin7 admin (`https://inventory.dearsystems.com`).

### AI chat tab says "AI assistant is not configured"
- `ANTHROPIC_API_KEY` is not set in `.env.voyage`. This is optional — the rest of the app works without it.

### Cin7 proxy returns 401 for a user
- That user's `PROXY_TOKEN_<USERNAME>` env var isn't set, or the user is passing the wrong token. Check the `/healthz` endpoint — it returns the count of loaded tokens.

### Data files vanish on restart
- The host-side `data/voyage/` and `data/proxy/` directories aren't being mounted correctly into the containers. Inspect: `docker inspect voyage-scheduling | grep -A5 Mounts`. Fix the compose file or volume permissions.

---

## 11. Cutover checklist

When you're ready to flip from Railway to the new server:

- [ ] Both services running on new server, healthchecks green
- [ ] Reverse proxy + TLS working, new URLs reachable from inside corporate network
- [ ] Test login as a known admin user works (post-data-restore)
- [ ] Test the MRP run, Calendar, and Yield tabs render data correctly
- [ ] Matt updates the Apps Script webhook URL to the new server
- [ ] Wait one morning cycle and confirm the inventory sync lands in the new app
- [ ] Communicate the new URL to all users
- [ ] Run both Railway + new server in parallel for ~1 week
- [ ] Matt decommissions the Railway services
- [ ] First post-cutover data backup taken successfully

---

## 12. Contact

Code-level questions or unexpected behaviors after cutover: Matt Gunderson (m.gunderson@voyagefoods.com).
