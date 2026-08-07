/**
 * Single source of truth for cross-cutting client constants.
 *
 * SIDELOAD BUILD — the sideload/OSS variant of the NovoMCP Word add-in.
 * Differs from the production build (Office Store, addin.novomcp.com CDN,
 * points at api.novomcp.com) in three ways:
 *
 *   1. API_BASE resolves via getApiBase() at call time. Reads from
 *      Office.context.roamingSettings first (so the user's engine URL
 *      setting persists across sessions), falls back to the hosted API.
 *   2. manifest.xml AppDomains include http://localhost:8018 so a local
 *      NovoMCP engine is reachable from the taskpane.
 *   3. SourceLocation resolves to the local dev server (office-addin-
 *      debugging start) instead of the CDN. Users sideload the manifest
 *      directly from this repo — no Office Store submission.
 *
 * SURFACE_TAG is sent as the X-Novo-Surface header on every API call —
 * namespaces per-surface sessions for the audit log.
 */
export const SURFACE_TAG = 'word-addin-sideload-v1';
export const UA_PREFIX = 'NovoMCP-WordAddin-Sideload/0.1.0';

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
 * Resolve the engine URL at call time. Reads Office.context.roamingSettings
 * first (persists across Word sessions on this device), falls back to
 * API_BASE_DEFAULT. Handles the case where Office isn't available (unit
 * tests, dev preview) by returning the default. Trailing slashes are
 * stripped so ${base}${path} produces a clean URL.
 */
export function getApiBase(): string {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const settings = (globalThis as any).Office?.context?.roamingSettings;
    const stored = settings?.get?.(STORAGE_KEYS.apiBase);
    if (typeof stored === 'string' && stored.trim().length > 0) {
      return stored.trim().replace(/\/$/, '');
    }
  } catch {
    // fall through to default
  }
  return API_BASE_DEFAULT;
}
