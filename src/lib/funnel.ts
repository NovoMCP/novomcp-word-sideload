/**
 * Funnel ID tracking — captures the server-resolved funnel_id from every
 * successful tool call so cross-surface CTAs can deep-link the user
 * forward to any AI assistant connected to NovoMCP, or to the dashboard
 * audit log.
 *
 * Storage: localStorage (per-origin), shared across documents and sessions.
 */
import { STORAGE_KEYS } from './constants';
import type { ApiResponse } from '../types';

export function recordFunnelId(response: ApiResponse<unknown>): void {
  const funnelId = response?.usage?.funnel_id;
  if (!funnelId) return;
  try { localStorage.setItem(STORAGE_KEYS.currentFunnelId, funnelId); } catch { /* noop */ }
}

export function getCurrentFunnelId(): string | null {
  try { return localStorage.getItem(STORAGE_KEYS.currentFunnelId); } catch { return null; }
}

export function clearFunnelId(): void {
  try { localStorage.removeItem(STORAGE_KEYS.currentFunnelId); } catch { /* noop */ }
}

/** Dashboard URL with the funnel pre-selected and auto-expanded. */
export function dashboardUrl(funnelId: string): string {
  return `https://app.novomcp.com/audit/pipelines?funnel_id=${encodeURIComponent(funnelId)}`;
}

/**
 * Continue-in-AI-assistant prompt template. Copy-to-clipboard handoff
 * since AI assistants don't have a deep-link receiver. The user pastes
 * the prompt into Claude / ChatGPT / Gemini / NovoWorkbench and the
 * conversation picks up via get_funnel_audit.
 */
export function aiHandoffPrompt(funnelId: string, smiles?: string): string {
  const lines = [
    `Continue NovoMCP funnel ${funnelId}.`,
    smiles ? `Last molecule looked at: ${smiles}` : '',
    `Load the prior audit with get_funnel_audit, then continue analysis from where the Word add-in left off.`,
  ].filter(Boolean);
  return lines.join('\n');
}
