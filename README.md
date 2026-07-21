# NovoMCP Word add-in (sideload)

Molecular intelligence for Word — configurable engine URL, manifest sideload distribution.

This is the **sideload/OSS variant** of the NovoMCP Word add-in. It differs from the [production build](https://github.com/novomcp/novomcp-word-addin) in three ways:

1. **Configurable engine URL** — defaults to `https://api.novomcp.com`, but the taskpane surfaces an "Engine URL" setting so you can point at `http://localhost:8018` (your local NovoMCP OSS engine) or any self-hosted deployment. Persisted across sessions via `Office.context.roamingSettings`.
2. **`localhost` in `<AppDomains>`** — the manifest whitelists `http://localhost`, `https://localhost`, and `http://127.0.0.1` so the taskpane can reach a local engine.
3. **No Office Store submission** — sideload the manifest directly via `office-addin-debugging start`. Runs from a local dev server (self-signed cert via `office-addin-dev-certs`). No AppSource review, no CDN, no store account.

## Install

Prerequisites: Node 18+, Word desktop (macOS 16.86+ / Windows M365).

```bash
git clone https://github.com/novomcp/novomcp-word-sideload.git
cd novomcp-word-sideload
npm install
npm run start
```

`npm run start` runs `office-addin-debugging`, which:
1. Installs a self-signed cert for `localhost:3000` (first run only — you may be prompted)
2. Boots the Vite dev server on `https://localhost:3000` — **this serves the taskpane HTML that Word loads**
3. Sideloads the manifest into Word
4. Launches Word with the add-in loaded

Once Word opens: **Home → NovoMCP → Open NovoMCP**. The taskpane appears.

If Word doesn't launch automatically (macOS quirk), open Word manually and use **Insert → My Add-ins → Shared Folder** to select the sideloaded manifest.

> **IMPORTANT: `npm run start` must stay running.** The taskpane is served from the local dev server at `https://localhost:3000`. If you close the terminal running `npm run start`, the next time you click **Open NovoMCP** in Word you'll get an **"Add-in Error / Sorry, we can't load the add-in. Please make sure you have network and/or Internet connectivity"** message. Restart the dev server (or run `npm run start` again) to fix.
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

The taskpane opens on the **Connect your account** screen. Three paths:

**Option A — talk to your local OSS engine:**
1. Start the engine locally (from the [OSS repo](https://github.com/novomcp/novomcp)): `python main_https.py`
2. Expand **Self-hosted or local NovoMCP engine (advanced)**
3. Set **Engine URL** to `http://localhost:8018`
4. Leave the API key blank (or type any string — `LocalAuthGate` accepts any bearer token)
5. Click **Connect**

**Option B — talk to the hosted API:**
1. Sign up at [app.novomcp.com/signup](https://app.novomcp.com/signup)
2. Copy your `nmcp_*` key from [app.novomcp.com/keys](https://app.novomcp.com/keys)
3. Paste into the taskpane's **Novo API key** field
4. Click **Connect**

**Option C — talk to your own self-hosted engine:**
1. Set **Engine URL** to your engine's public URL (e.g. `https://novomcp.your-lab.edu`)
2. Paste any API key your engine expects
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
npm run start            # build + sideload + launch Word
npm run stop             # stop the sideload session
npm run package          # build + zip dist/ (for sharing)
```

## Sideload troubleshooting

**"Certificate not trusted" error:**
```bash
npx office-addin-dev-certs install
```
Restart Word after installing the cert.

**Word can't reach `https://localhost:3000`:**
- Confirm the dev server is running (`curl -k https://localhost:3000/taskpane.html`)
- Confirm the manifest is sideloaded: **Insert → My Add-ins → Shared Folder**

**Taskpane loads but "Cannot connect to NovoMCP engine":**
- Confirm the engine URL you set is reachable from your browser
- Local engine: `curl http://localhost:8018/health` should return `{"status":"healthy",...}`

## What's the same as the production build?

Everything user-visible: SMILES detection, taskpane UI, tool catalog, funnel/audit conventions. Bug fixes usually cherry-pick cleanly between the production repo and this one.

## What's different from the production build?

| | Production | Sideload |
|---|---|---|
| Distribution | Office Store / AppSource | Manifest sideload (`office-addin-debugging`) |
| Hosting | `addin.novomcp.com` CDN | `https://localhost:3000` (dev server) |
| Engine URL | Hardcoded `https://api.novomcp.com` | Configurable, defaults to hosted |
| `<AppDomains>` | Production hosts only | + `http://localhost` + `http://127.0.0.1` |
| Manifest `<Id>` | Locked production GUID | Fresh sideload-only GUID (coexists with production) |
| DisplayName | "NovoMCP" | "NovoMCP (Sideload)" |
| API key required? | Yes (`nmcp_*` for hosted API) | Optional (blank works for local engine) |

## Support

- **NovoMCP OSS engine:** https://github.com/novomcp/novomcp
- **Docs:** https://github.com/novomcp/novomcp/tree/main/docs
- **Issues:** file against this repo for add-in bugs, against `novomcp/novomcp` for engine bugs.

## License

Apache-2.0. See `LICENSE`.
