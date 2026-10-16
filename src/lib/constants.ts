/**
 * Single source of truth for cross-cutting client constants.
 *
 * SIDELOAD BUILD — the sideload/OSS variant of the NovoMCP Word add-in.
 * Differs from the production build (Office Store, addin.novomcp.com CDN,
 * points at api.novomcp.com) in three ways:
 *
 *   1. API_BASE resolves via getApiBase() at call time. Reads the user's
 *      engine URL from localStorage (persists across documents and
 *      sessions, like the rest of lib/storage.ts), falls back to a local
 *      engine at http://localhost:8018. (Not Office.context.roamingSettings:
 *      that's an Outlook-only API and is undefined in Word.)
 *   2. The taskpane fetches the engine directly. Plain-http engines on
 *      other machines (e.g. a DGX Spark on the LAN) are blocked as mixed
 *      content, so the dev server proxies https://localhost:3000/engine to
 *      NOVOMCP_ENGINE_URL (see vite.config.ts). AppDomains govern taskpane
 *      navigation, not fetch.
 *   3. SourceLocation resolves to the local dev server (office-addin-
 *      debugging start) instead of the CDN. Users sideload the manifest
 *      directly from this repo — no Office Store submission.
 *
 * SURFACE_TAG is sent as the X-Novo-Surface header on every API call —
 * namespaces per-surface sessions for the audit log.
 */
export const SURFACE_TAG = 'word-addin-sideload-v1';
export const UA_PREFIX = 'NovoMCP-WordAddin-Sideload/0.2.0';

// Default engine URL — used when the user hasn't set a custom one.
// Defaults to a local self-hosted engine; point at a hosted engine (with a
// key) via the settings UI if you prefer.
export const API_BASE_DEFAULT = 'http://localhost:8018';

// Backwards-compat alias for callers that still import API_BASE directly.
// New code should call getApiBase() so runtime settings take effect.
export const API_BASE = API_BASE_DEFAULT;

export const STORAGE_KEYS = {
  novoKey: 'novo.key',
  computeKey: 'novo.computeKey',
  user: 'novo.user',
  smilesCache: 'novo.smilesCache',
  currentFunnelId: 'novo.currentFunnelId',
  pendingSmiles: 'novo.pendingSmiles',
  apiBase: 'novo.apiBase',
} as const;

export const SMILES_CACHE_TTL_MS = 24 * 60 * 60 * 1000;

/**
 * Resolve the engine URL at call time. Reads the user's setting from
 * localStorage (per-origin, persists across Word documents and sessions),
 * falls back to API_BASE_DEFAULT. Handles storage being unavailable (unit
 * tests, private contexts) by returning the default. Trailing slashes are
 * stripped so ${base}${path} produces a clean URL.
 */
export function getApiBase(): string {
  try {
    const stored = localStorage.getItem(STORAGE_KEYS.apiBase);
    if (typeof stored === 'string' && stored.trim().length > 0) {
      return stored.trim().replace(/\/$/, '');
    }
  } catch {
    // fall through to default
  }
  return API_BASE_DEFAULT;
}

/** Persist the engine URL override; null, empty, or the default clears it. */
export function setApiBase(base: string | null): void {
  try {
    const v = (base ?? '').trim();
    if (!v || v === API_BASE_DEFAULT) localStorage.removeItem(STORAGE_KEYS.apiBase);
    else localStorage.setItem(STORAGE_KEYS.apiBase, v);
  } catch {
    // storage unavailable: getApiBase() falls back to the default
  }
}
