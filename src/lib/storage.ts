/**
 * Typed storage wrapper. Word add-ins don't have chrome.storage; we use
 * localStorage (per-origin, persists across documents and sessions).
 *
 * Design parity with the Chrome extension's storage helper: same key
 * names, same auth.* surface area. Lib code that calls auth.getNovoKey()
 * is identical between surfaces.
 */
import { STORAGE_KEYS } from './constants';
import type { NovoUser } from '../types';

function get<T>(key: string): T | undefined {
  try {
    const raw = localStorage.getItem(key);
    if (raw == null) return undefined;
    return JSON.parse(raw) as T;
  } catch {
    return undefined;
  }
}

function set(key: string, value: unknown): void {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch { /* quota exceeded — non-fatal */ }
}

function remove(key: string): void {
  try { localStorage.removeItem(key); } catch { /* noop */ }
}

export const auth = {
  getNovoKey: () => get<string>(STORAGE_KEYS.novoKey),
  setNovoKey: (key: string) => set(STORAGE_KEYS.novoKey, key),
  clearNovoKey: () => remove(STORAGE_KEYS.novoKey),

  getComputeKey: () => get<string>(STORAGE_KEYS.computeKey),
  setComputeKey: (key: string) => set(STORAGE_KEYS.computeKey, key),
  clearComputeKey: () => remove(STORAGE_KEYS.computeKey),

  getUser: () => get<NovoUser>(STORAGE_KEYS.user),
  setUser: (user: NovoUser) => set(STORAGE_KEYS.user, user),
  clearUser: () => remove(STORAGE_KEYS.user),

  signOut: () => {
    remove(STORAGE_KEYS.novoKey);
    remove(STORAGE_KEYS.computeKey);
    remove(STORAGE_KEYS.user);
  },
};
