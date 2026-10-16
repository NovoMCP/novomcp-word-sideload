# NovoMCP Word add-in (sideload)

Molecular intelligence for Word — configurable engine URL, manifest sideload distribution.

This is the **sideload/OSS variant** of the NovoMCP Word add-in. It differs from the [production build](https://github.com/novomcp/novomcp-word-addin) in three ways:

1. **Configurable engine URL** — defaults to `http://localhost:8018` (your local NovoMCP engine); the taskpane's "Engine URL" setting points it at a DGX Spark on your network (via the dev-server proxy) or any self-hosted deployment. Persisted across sessions in the taskpane's `localStorage`.
2. **Engine proxy for other machines** — the dev server proxies `https://localhost:3000/engine` to `NOVOMCP_ENGINE_URL`, so a plain-http engine on the LAN is reachable without mixed-content blocking. (`<AppDomains>` also lists `localhost` / `127.0.0.1`; that governs taskpane navigation, not engine calls.)
3. **No Office Store submission** — sideload the manifest directly via `office-addin-debugging start`. Runs from a local dev server (self-signed cert via `office-addin-dev-certs`). No AppSource review, no CDN, no store account.

## Install

Prerequisites: Node 18+, Word desktop (macOS 16.86+ / Windows M365).

```bash
git clone https://github.com/novomcp/novomcp-word-sideload.git
cd novomcp-word-sideload
npm install
```

### Before you start: trust a local HTTPS certificate (once per machine)

Word only loads add-ins over HTTPS, and this sideload serves its taskpane from your own machine at `https://localhost:3000`. So **every user needs a locally trusted development certificate, once per computer**:

```bash
npx office-addin-dev-certs install --days 365
```

- It creates a certificate authority **on your machine only** and trusts it in your keychain (macOS) or certificate store (Windows). You'll be asked for your password. The CA only vouches for `localhost` on your own computer, so it isn't shared and can't be shared. Each user and each machine runs this once.
- **It expires.** The tool's default is **30 days**; `--days 365` gives a year. When it lapses, Word shows *"The content is blocked because it isn't signed by a valid security certificate"*. See [Sideload troubleshooting](#sideload-troubleshooting) for the one-line fix.
- Check it at any time: `npx office-addin-dev-certs verify`.
- The Office Store version of the add-in doesn't need this. It loads from a hosted HTTPS address with a public certificate.

Full guide (both add-ins, remote engines, troubleshooting): [docs.novomcp.com/sideload-add-ins](https://docs.novomcp.com/sideload-add-ins/).

### Run it

Then use two terminals:

```bash
# Terminal 1: writes dist/manifest.xml, then serves the taskpane on https://localhost:3000
npm run dev

# Terminal 2: sideloads dist/manifest.xml into Word and launches Word
npm run start
```

`npm run start` runs `office-addin-debugging start dist/manifest.xml`. It only sideloads the manifest and launches Word. It does **not** install the cert, start the dev server, or build anything, so `npm run dev` (or `npm run build`) must have written `dist/manifest.xml` first.

Once Word opens: **Home → NovoMCP → Open NovoMCP**. The taskpane appears.

If Word doesn't launch automatically (macOS quirk), open Word manually and use **Insert → My Add-ins → Shared Folder** to select the sideloaded manifest.

> **IMPORTANT: `npm run dev` must stay running.** The taskpane is served from the Vite dev server at `https://localhost:3000`. If you close the terminal running `npm run dev`, the next time you click **Open NovoMCP** in Word you'll get an **"Add-in Error / Sorry, we can't load the add-in. Please make sure you have network and/or Internet connectivity"** message. Restart `npm run dev` to fix.
>
> This is not a bug in the extension — it's how Office add-ins work. The manifest references an HTTPS URL for the taskpane HTML; that URL must be reachable whenever the taskpane is open.

## Manual sideload (fallback path)

If `npm run start` misbehaves on your system:

1. `npm run build` (builds the dist/ + manifest)
2. `office-addin-debugging start-debug-mode dist/manifest.xml` (installs the cert without launching Word)
3. Sideload the manifest manually — on macOS drop `dist/manifest.xml` into `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` and relaunch Word
4. **Keep the dev server running separately:** `npx vite --port 3000` (or `npm run dev`)

Word will then find both the manifest AND the taskpane HTML. Skip step 4 and you'll hit the "network connectivity" error above.

## Configure

The taskpane opens on the **Connect your account** screen. Leave the API key blank for a single-user engine (`LocalAuthGate` accepts any bearer token).

**Option A — an engine on this machine:**
1. Start the engine (from the [engine repo](https://github.com/novomcp/novomcp)): `docker compose up -d` — see the [quickstart](https://docs.novomcp.com/quickstart/)
2. Leave **Engine URL** at the default, `http://localhost:8018`
3. Click **Connect**

**Option B — a DGX Spark (or any plain-http engine) on your network:**

The taskpane is an HTTPS page, so Word blocks plain-http calls to any machine other than localhost (mixed content). The add-in's dev server proxies `https://localhost:3000/engine` to the engine instead:

1. Bring the engine up on the Spark: [Run the whole engine on an NVIDIA DGX Spark](https://docs.novomcp.com/novo-station-dgx-spark/) (`docker compose -f compose.spark.yml up -d`)
2. Start the add-in's dev server pointed at it: `NOVOMCP_ENGINE_URL=http://spark-xxxx.local:8018 npm run dev` (it serves the taskpane too, so keep it running)
3. Set **Engine URL** to `https://localhost:3000/engine`
4. Click **Connect**

**Option C — a self-hosted engine behind HTTPS:**
1. Set **Engine URL** to your engine's URL (e.g. `https://novomcp.your-lab.edu`)
2. Paste any API key your engine expects (or leave blank for `LocalAuthGate`)
3. Click **Connect**

Engine URL persists across Word sessions on this device.

## Use

Once connected:
- **Highlight a SMILES string** in your manuscript → taskpane shows ADMET properties, compliance status, drug-likeness scores
- **Scan Document** button (Home ribbon) — finds every SMILES in the current doc and profiles them in batch
- Results include a funnel_id you can carry into any MCP-connected AI assistant (Claude, Cursor, etc.) to continue the analysis

## Development

```bash
npm run dev              # vite dev server only (no Word sideload)
npm run typecheck        # tsc --noEmit
npm run build            # production build into dist/
npm run validate-manifest # sanity-check manifest.xml against Office schema
npm run start            # sideload dist/manifest.xml + launch Word (run `npm run dev` first)
npm run stop             # stop the sideload session
npm run package          # build + zip dist/ (for sharing)
```

## Sideload troubleshooting

**"The content is blocked because it isn't signed by a valid security certificate"** (or "Certificate not trusted"):

The local dev certificate is missing, untrusted, or **expired** (the default lifetime is 30 days). Check the expiry date:
```bash
openssl x509 -in ~/.office-addin-dev-certs/localhost.crt -noout -enddate
```
Replace it with a fresh one (asks for your password):
```bash
npx office-addin-dev-certs uninstall && npx office-addin-dev-certs install --days 365
```
Then restart `npm run dev`, which only reads the certificate at startup, and click **Restart** in Word's add-in pane (or reopen Word).

**Word can't reach `https://localhost:3000`:**
- Confirm the dev server is running (`curl -k https://localhost:3000/taskpane.html`)
- Confirm the manifest is sideloaded: **Insert → My Add-ins → Shared Folder**

**Taskpane loads but "Cannot connect to NovoMCP engine":**
- Confirm the engine URL you set is reachable from your browser
- Local engine: `curl http://localhost:8018/health` should return `{"status":"healthy",...}`
- Engine on another machine: use the `/engine` proxy (Option B) — a direct `http://<host>:8018` Engine URL is blocked by Word as mixed content. Check the proxy with `curl -k https://localhost:3000/engine/health`

## What's the same as the production build?

Everything user-visible: SMILES detection, taskpane UI, tool catalog, funnel/audit conventions. Bug fixes usually cherry-pick cleanly between the production repo and this one.

## What's different from the production build?

| | Production | Sideload |
|---|---|---|
| Distribution | Office Store / AppSource | Manifest sideload (`office-addin-debugging`) |
| Hosting | `addin.novomcp.com` CDN | `https://localhost:3000` (dev server) |
| Engine URL | Hardcoded `https://api.novomcp.com` | Configurable, defaults to `http://localhost:8018` |
| `<AppDomains>` | Production hosts only | + `http://localhost` + `http://127.0.0.1` |
| Manifest `<Id>` | Locked production GUID | Fresh sideload-only GUID (coexists with production) |
| DisplayName | "NovoMCP" | "NovoMCP (Sideload)" |
| API key required? | Yes (`nmcp_*` for hosted API) | Optional (blank works for a single-user engine) |

## Support

- **NovoMCP OSS engine:** https://github.com/novomcp/novomcp
- **Docs:** https://docs.novomcp.com/
- **Issues:** file against this repo for add-in bugs, against `novomcp/novomcp` for engine bugs.

## License

Apache-2.0. See `LICENSE`.
