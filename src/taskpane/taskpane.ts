/**
 * Taskpane — auth + Office.js selection bridge + tabbed analysis.
 *
 * Architecture parity with the Chrome extension's side panel:
 *   - SMILES sourced from the host (Word selection here, page DOM there)
 *   - Profile renders eagerly from get_molecule_profile
 *   - ADMET / Compliance lazy-load on tab switch (cached per SMILES)
 *   - Similar runs explicitly via "Search neighbors" button (5 credits)
 *   - Advanced gates on a Compute key; 4 sync tools with friendly errors
 *   - Cross-surface CTAs (dashboard deep-link + AI-assistant copy-prompt)
 *     bind to the funnel_id captured from the latest tool response.
 */
import { auth } from '../lib/storage';
import {
  probeKey, validateKey,
  getMoleculeProfile, predictAdmet, checkComplianceDeep,
  searchSimilar,
  predictPka, predictSolubility, predictBde, predictFrontierOrbitals,
  ApiError$,
} from '../lib/api';
import { getApiBase } from '../lib/constants';
import { findSmilesCandidates, isLikelySmiles } from '../lib/smiles';
import { readSelection, onSelectionChange, insertTableAtSelection, type InsertableTable } from '../lib/word';
import { readAllTables, insertCellComment, type ParsedTable } from '../lib/tables';
import { detectSchema, findDiscrepancies, extractServerValues, type Discrepancy, type TableSchema } from '../lib/validator';
import { normalize as normalizeAdmet, classificationColor, type AdmetCategory } from '../lib/admet';
import { getCurrentFunnelId, dashboardUrl, aiHandoffPrompt } from '../lib/funnel';
import type { NovoUser, ApiResponse, MoleculeProfile } from '../types';

const $ = <T extends HTMLElement>(id: string): T => {
  const el = document.getElementById(id);
  if (!el) throw new Error(`#${id} not found`);
  return el as T;
};

type AppState = 'loading' | 'onboarding' | 'connected';
const states: Record<AppState, HTMLElement> = {
  loading: $<HTMLElement>('state-loading'),
  onboarding: $<HTMLElement>('state-onboarding'),
  connected: $<HTMLElement>('state-connected'),
};

function show(state: AppState): void {
  for (const [name, el] of Object.entries(states)) el.hidden = name !== state;
}

function showError(message: string): void {
  const el = $<HTMLElement>('error-msg');
  el.textContent = message;
  el.hidden = false;
}
function clearError(): void {
  const el = $<HTMLElement>('error-msg');
  el.textContent = '';
  el.hidden = true;
}

// ─── Auth flows ────────────────────────────────────────────────────────

function isLocalEngineUrl(url: string): boolean {
  return /^https?:\/\/(localhost|127\.0\.0\.1)/i.test(url);
}

function renderConnected(_user: NovoUser | null): void {
  // Sideload build: connected state shows Engine URL + Mode (local /
  // hosted). The hosted-product fields (email, tier, credits, compute-key
  // status) don't apply in local single-user mode.
  const base = getApiBase();
  const local = isLocalEngineUrl(base);
  $<HTMLElement>('user-engine').textContent = base;
  $<HTMLElement>('user-mode').innerHTML = local
    ? '<span class="badge ok">local single-user</span>'
    : '<span class="badge">hosted / self-hosted</span>';

  show('connected');

  if (!selectionWired) {
    wireSelectionBridge();
    setupTabs();
    setupAdvancedHandlers();
    setupSimilarHandlers();
    setupScanHandlers();
    setupInsertHandlers();
    selectionWired = true;
    void refreshFromCurrentSelection();
  }
  void refreshAdvancedGate();
}

async function bootstrap(): Promise<void> {
  show('loading');
  const key = auth.getNovoKey();
  if (!key) { show('onboarding'); return; }
  try {
    const user = await validateKey();
    auth.setUser(user);
    renderConnected(user);
  } catch (err) {
    if (err instanceof ApiError$ && err.isUnauthorized) {
      auth.signOut();
      show('onboarding');
      showError('Saved key is no longer valid. Please reconnect.');
      return;
    }
    const cached = auth.getUser();
    if (cached) renderConnected(cached);
    else { show('onboarding'); showError('Unable to reach NovoMCP. Check your connection.'); }
  }
}

async function handleConnect(event: Event): Promise<void> {
  event.preventDefault();
  clearError();
  const apiBase = ($<HTMLInputElement>('api-base')).value.trim() || 'http://localhost:8018';
  const novoKey = ($<HTMLInputElement>('novo-key')).value.trim();

  // Persist engine URL first so probeKey targets the right engine.
  const { STORAGE_KEYS, API_BASE_DEFAULT } = await import('../lib/constants');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const settings = (globalThis as any).Office?.context?.roamingSettings;
  if (settings) {
    if (apiBase && apiBase !== API_BASE_DEFAULT) {
      settings.set(STORAGE_KEYS.apiBase, apiBase);
    } else {
      settings.remove(STORAGE_KEYS.apiBase);
    }
    await new Promise<void>((resolve) => settings.saveAsync(() => resolve()));
  }

  const local = isLocalEngineUrl(apiBase);
  if (!local && novoKey && !novoKey.startsWith('nmcp_')) {
    showError('For hosted engines the API key must start with nmcp_. Get one at app.novomcp.com/keys, or use a local engine URL.');
    return;
  }
  if (!local && !novoKey) {
    showError('Hosted engines require an API key. Add one, or point at a local engine URL.');
    return;
  }

  const btn = $<HTMLButtonElement>('connect-btn');
  btn.disabled = true; btn.textContent = 'Connecting…';
  const effectiveKey = novoKey || 'local-dev';
  try {
    const user = await probeKey(effectiveKey);
    auth.setNovoKey(effectiveKey);
    auth.setUser(user);
    renderConnected(user);
  } catch (err) {
    // Local engines return a shape probeKey doesn't recognize — but if the
    // engine responded at all, treat it as connected. Only reject on 401
    // (bad key on hosted API) or network failure.
    if (local && err instanceof ApiError$ && !err.isUnauthorized) {
      auth.setNovoKey(effectiveKey);
      renderConnected(null);
      return;
    }
    const msg = err instanceof ApiError$ && err.isUnauthorized
      ? 'Key rejected. Check the key and engine URL.'
      : err instanceof ApiError$
        ? `Could not connect (${err.status}): ${err.message}`
        : 'Network error. Is the engine running at that URL?';
    showError(msg);
  } finally {
    btn.disabled = false; btn.textContent = 'Connect';
  }
}

function handleSignOut(): void {
  auth.signOut();
  // Also clear the stored engine URL override so the next Connect starts
  // fresh from the default.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const settings = (globalThis as any).Office?.context?.roamingSettings;
  if (settings) {
    settings.remove('novo.apiBase');
    try { settings.saveAsync(() => {}); } catch { /* ignore */ }
  }
  ($<HTMLInputElement>('novo-key')).value = '';
  ($<HTMLInputElement>('api-base')).value = 'http://localhost:8018';
  currentSmiles = null;
  $<HTMLElement>('active-block').hidden = true;
  $<HTMLElement>('empty-active').hidden = false;
  show('onboarding');
}

// ─── Office.js selection bridge ────────────────────────────────────────

let selectionWired = false;
let selectionDebounce: number | null = null;
let currentSmiles: string | null = null;

function wireSelectionBridge(): void {
  onSelectionChange(() => {
    if (selectionDebounce != null) window.clearTimeout(selectionDebounce);
    selectionDebounce = window.setTimeout(() => { void refreshFromCurrentSelection(); }, 250);
  });
}

async function refreshFromCurrentSelection(): Promise<void> {
  let snap;
  try { snap = await readSelection(); } catch { return; }

  if (!snap.hasText) { paintSelectionEmpty(); return; }
  const candidates = findSmilesCandidates(snap.text);
  if (candidates.length === 0) { paintSelectionNoMatch(); return; }
  paintSelectionCandidates(candidates.map((c) => c.smiles));
}

function paintSelectionEmpty(): void {
  $<HTMLElement>('selection-empty').hidden = false;
  $<HTMLElement>('selection-candidates').hidden = true;
  $<HTMLElement>('selection-none').hidden = true;
  $<HTMLElement>('selection-summary-count').textContent = '';
}

function paintSelectionNoMatch(): void {
  $<HTMLElement>('selection-empty').hidden = true;
  $<HTMLElement>('selection-candidates').hidden = true;
  $<HTMLElement>('selection-none').hidden = false;
  $<HTMLElement>('selection-summary-count').textContent = '0';
}

function paintSelectionCandidates(smilesList: string[]): void {
  const list = $<HTMLElement>('selection-candidates');
  list.innerHTML = '';
  const unique = Array.from(new Set(smilesList)).slice(0, 10);
  for (const smiles of unique) {
    const li = document.createElement('li');
    li.dataset['smiles'] = smiles;
    if (smiles === currentSmiles) li.classList.add('active');
    li.innerHTML = `<span class="candidate-smiles"></span><span class="candidate-arrow">→</span>`;
    li.querySelector<HTMLElement>('.candidate-smiles')!.textContent = smiles;
    li.addEventListener('click', () => { void load(smiles); });
    list.appendChild(li);
  }
  $<HTMLElement>('selection-empty').hidden = true;
  $<HTMLElement>('selection-candidates').hidden = false;
  $<HTMLElement>('selection-none').hidden = true;
  $<HTMLElement>('selection-summary-count').textContent = String(unique.length);
  // Auto-expand the details when there's something to click
  const details = $<HTMLDetailsElement>('selection-section');
  if (!currentSmiles) details.open = true;
}

// ─── Profile load ──────────────────────────────────────────────────────

const profileCache = new Map<string, ApiResponse<MoleculeProfile>>();
const admetCache = new Map<string, Record<AdmetCategory, ReturnType<typeof normalizeAdmet>[AdmetCategory]>>();
const complianceCache = new Map<string, Record<string, unknown>>();
const similarCache = new Map<string, Array<Record<string, unknown>>>();
const admetInFlight = new Set<string>();
const complianceInFlight = new Set<string>();

async function load(smiles: string): Promise<void> {
  if (!smiles) return;
  currentSmiles = smiles;
  $<HTMLElement>('empty-active').hidden = true;
  $<HTMLElement>('error-block').hidden = true;
  $<HTMLElement>('active-block').hidden = true;
  $<HTMLElement>('loading-block').hidden = false;
  resetAdvancedResults();
  resetSimilarResults();
  // Reset insert-button states — re-enabled by their respective render
  // paths once data lands.
  setInsertEnabled('profile', false);
  setInsertEnabled('admet', false);
  setInsertEnabled('compliance', false);
  setInsertEnabled('similar', false);
  switchTab('profile');

  try {
    const r = await getMoleculeProfile(smiles);
    if (smiles !== currentSmiles) return; // user picked another in the meantime
    renderProfile(smiles, r);
  } catch (e) {
    renderProfileError(smiles, e instanceof Error ? e.message : 'Network error');
  } finally {
    $<HTMLElement>('loading-block').hidden = true;
  }

  // Mark the active candidate
  document.querySelectorAll<HTMLElement>('.candidate-list li').forEach((li) => {
    li.classList.toggle('active', li.dataset['smiles'] === smiles);
  });
}

function renderProfileError(smiles: string, message: string): void {
  const el = $<HTMLElement>('error-block');
  el.textContent = `${smiles}: ${message}`;
  el.hidden = false;
  $<HTMLElement>('active-block').hidden = true;
}

function renderProfile(smiles: string, response: ApiResponse<MoleculeProfile>): void {
  profileCache.set(smiles, response);
  setInsertEnabled('profile', true);
  setInsertEnabled('admet', false);
  setInsertEnabled('compliance', false);
  setInsertEnabled('similar', false);
  const p = response.result;
  $<HTMLElement>('active-smiles').textContent = smiles;

  // Probe both naming conventions: enriched-DB vs computed-on-demand.
  const props = (p.properties || {}) as Record<string, unknown>;
  const pick = (...keys: string[]): unknown => {
    for (const k of keys) if (props[k] != null) return props[k];
    return undefined;
  };
  const propRows: Array<[string, string]> = [];
  pushIfPresent(propRows, 'CID', props['cid'], formatInt);
  pushIfPresent(propRows, 'Formula', props['molecular_formula'], (v) => escapeText(String(v)));
  pushIfPresent(propRows, 'MW', props['molecular_weight'], (v) => `${formatNum(v)} g/mol`);
  pushIfPresent(propRows, 'LogP', props['logp'], formatNum);
  pushIfPresent(propRows, 'TPSA', props['tpsa'], formatNum);
  pushIfPresent(propRows, 'QED', props['qed'], formatNum);
  pushIfPresent(propRows, 'Drug-likeness', props['drug_likeness'], formatNum);
  pushIfPresent(propRows, 'Synth. accessibility', props['synthetic_accessibility'], formatNum);
  pushIfPresent(propRows, 'Complexity', props['complexity'], formatNum);
  pushIfPresent(propRows, 'Fsp³', props['fsp3'], formatNum);
  pushIfPresent(propRows, 'HBD', pick('hbd_count', 'hbd'), formatInt);
  pushIfPresent(propRows, 'HBA', pick('hba_count', 'hba'), formatInt);
  pushIfPresent(propRows, 'Rot bonds', pick('rotatable_bond_count', 'rotatable_bonds'), formatInt);
  pushIfPresent(propRows, 'Heavy atoms', props['heavy_atom_count'], formatInt);
  pushIfPresent(propRows, 'Aromatic rings', pick('aromatic_ring_count', 'aromatic_rings'), formatInt);
  pushIfPresent(propRows, 'Aromatic atoms', props['aromatic_atom_count'], formatInt);
  pushIfPresent(propRows, 'Lipinski violations', props['lipinski_violations'], formatInt);
  if (propRows.length === 0) propRows.push(['—', 'No properties returned']);
  renderGrid('profile-properties', propRows);

  resetAdmetTab();
  resetComplianceTab();
  if (document.getElementById('panel-admet')?.classList.contains('active')) void ensureAdmetLoaded(smiles);
  if (document.getElementById('panel-compliance')?.classList.contains('active')) void ensureComplianceLoaded(smiles);

  const cost = response.usage?.credits ?? 0;
  const remaining = response.usage?.credits_remaining;
  const meta = $<HTMLElement>('profile-meta');
  meta.innerHTML = '';
  meta.appendChild(metaPill(p.in_database ? 'cached profile' : 'computed on demand'));
  meta.appendChild(metaPill(cost === 0 ? 'free lookup' : `${cost} credit${cost === 1 ? '' : 's'}`));
  if (typeof remaining === 'number') meta.appendChild(metaPill(`${Math.floor(remaining).toLocaleString()} credits left`));

  void renderCrossSurface(smiles);
  $<HTMLElement>('active-block').hidden = false;
}

// ─── Tabs ──────────────────────────────────────────────────────────────

function setupTabs(): void {
  document.querySelectorAll<HTMLButtonElement>('.tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tabName = btn.dataset['tab'];
      if (!tabName || btn.disabled) return;
      switchTab(tabName);
    });
  });
}

function switchTab(name: string): void {
  document.querySelectorAll<HTMLElement>('.tab').forEach((b) => {
    b.setAttribute('aria-selected', b.dataset['tab'] === name ? 'true' : 'false');
  });
  document.querySelectorAll<HTMLElement>('.tab-panel').forEach((p) => {
    p.classList.toggle('active', p.id === `panel-${name}`);
  });
  if (currentSmiles) {
    if (name === 'admet') void ensureAdmetLoaded(currentSmiles);
    else if (name === 'compliance') void ensureComplianceLoaded(currentSmiles);
  }
  if (name === 'advanced') void refreshAdvancedGate();
}

// ─── ADMET tab ─────────────────────────────────────────────────────────

const ADMET_CATEGORIES: Array<[AdmetCategory, string, string]> = [
  ['absorption', 'A', 'Absorption'],
  ['distribution', 'D', 'Distribution'],
  ['metabolism', 'M', 'Metabolism'],
  ['excretion', 'E', 'Excretion'],
  ['toxicity', 'T', 'Toxicity'],
];

function resetAdmetTab(): void { $<HTMLElement>('profile-admet').innerHTML = ''; }

async function ensureAdmetLoaded(smiles: string): Promise<void> {
  if (admetCache.has(smiles)) { paintAdmet(admetCache.get(smiles)!); return; }
  if (admetInFlight.has(smiles)) return;
  admetInFlight.add(smiles);
  const el = $<HTMLElement>('profile-admet');
  el.innerHTML = `<div class="loading-state">
    <div class="loading-skeleton"></div>
    <div class="loading-skeleton" style="width: 80%;"></div>
    <div class="loading-skeleton" style="width: 65%;"></div>
  </div>`;
  try {
    const r = await predictAdmet(smiles);
    if (smiles !== currentSmiles) return;
    const normalized = normalizeAdmet(r.result as Record<string, unknown>);
    admetCache.set(smiles, normalized);
    paintAdmet(normalized);
    setInsertEnabled('admet', true);
    void renderCrossSurface(smiles);
  } catch (e) {
    el.innerHTML = `<p class="muted" style="font-size:12px;color:var(--error);">${escapeText(e instanceof Error ? e.message : 'ADMET fetch failed')}</p>`;
  } finally {
    admetInFlight.delete(smiles);
  }
}

function paintAdmet(normalized: Record<AdmetCategory, ReturnType<typeof normalizeAdmet>[AdmetCategory]>): void {
  const el = $<HTMLElement>('profile-admet');
  el.innerHTML = '';
  for (const [cat, letter, label] of ADMET_CATEGORIES) {
    const fields = normalized[cat];
    if (!fields || fields.length === 0) continue;
    const divider = document.createElement('div');
    divider.className = 'admet-divider';
    divider.innerHTML = `
      <span class="admet-divider-letter">${escapeText(letter)}</span>
      <span class="admet-divider-label">${escapeText(label)}</span>
      <span class="admet-divider-count">${fields.length}</span>
      <span class="admet-divider-line"></span>
    `;
    el.appendChild(divider);
    for (const f of fields) {
      const color = classificationColor(f.classification);
      const isProb = f.value >= 0 && f.value <= 1 && !f.unit;
      const valueDisplay = isProb
        ? `<span class="bar-track"><span class="bar-fill" style="width:${(f.value * 100).toFixed(0)}%;background:${color};"></span></span>
           <span class="bar-num" style="color:${color};">${f.value.toFixed(2)}</span>`
        : `<span class="bar-num bar-num-wide" style="color:${color};">${f.value.toFixed(f.unit ? 2 : 3)}${f.unit ? ` ${escapeText(f.unit)}` : ''}</span>`;
      const row = document.createElement('div');
      row.className = 'admet-field';
      row.innerHTML = `
        <div class="admet-field-title">${escapeText(f.label)}</div>
        <div class="admet-field-metric">
          ${valueDisplay}
          <span class="admet-class" style="color:${color};border-color:color-mix(in srgb,${color} 35%,transparent);">${escapeText(f.classification)}</span>
        </div>
      `;
      el.appendChild(row);
    }
  }
}

// ─── Compliance tab ────────────────────────────────────────────────────

function resetComplianceTab(): void { $<HTMLElement>('profile-compliance').innerHTML = ''; }

async function ensureComplianceLoaded(smiles: string): Promise<void> {
  if (complianceCache.has(smiles)) { paintCompliance(complianceCache.get(smiles)!); return; }
  if (complianceInFlight.has(smiles)) return;
  complianceInFlight.add(smiles);
  const el = $<HTMLElement>('profile-compliance');
  el.innerHTML = `<div class="loading-state">
    <div class="loading-skeleton"></div>
    <div class="loading-skeleton" style="width: 70%;"></div>
  </div>`;
  try {
    const r = await checkComplianceDeep(smiles);
    if (smiles !== currentSmiles) return;
    const result = r.result as Record<string, unknown>;
    complianceCache.set(smiles, result);
    paintCompliance(result);
    setInsertEnabled('compliance', true);
    void renderCrossSurface(smiles);
  } catch (e) {
    el.innerHTML = `<p class="muted" style="font-size:12px;color:var(--error);">${escapeText(e instanceof Error ? e.message : 'Compliance fetch failed')}</p>`;
  } finally {
    complianceInFlight.delete(smiles);
  }
}

function paintCompliance(result: Record<string, unknown>): void {
  const el = $<HTMLElement>('profile-compliance');
  el.innerHTML = '';

  const overall = String(result['overall_status'] ?? '');
  const baseCompliance = (result['base_compliance'] || {}) as Record<string, unknown>;
  const ctxCompliance = (result['context_compliance'] || {}) as Record<string, unknown>;

  const verdictEl = document.createElement('div');
  verdictEl.className = 'compliance-verdict';
  verdictEl.innerHTML = `
    <span class="badge ${verdictBadgeClass(overall)}" style="font-size:11px;padding:3px 10px;">${escapeText(overall || 'unknown')}</span>
    <span style="font-size:11px;color:var(--text-muted);margin-left:8px;">pharmaceutical · US</span>
  `;
  el.appendChild(verdictEl);

  const baseRows: Array<[string, string]> = [];
  const baseStatus = (baseCompliance['status'] as string) ?? null;
  if (baseStatus) baseRows.push(['Base status', `<span class="badge ${complianceBadgeClass(baseStatus)}">${escapeText(baseStatus)}</span>`]);
  pushFlag(baseRows, 'DEA controlled', baseCompliance['is_dea_controlled']);
  pushFlag(baseRows, 'FDA banned', baseCompliance['is_fda_banned']);
  pushFlag(baseRows, 'CWC scheduled', baseCompliance['is_cwc_scheduled']);
  pushFlag(baseRows, 'EPA PBT', baseCompliance['is_epa_pbt']);
  pushFlag(baseRows, 'EU REACH banned', baseCompliance['is_eu_reach_banned']);
  pushFlag(baseRows, 'Scaffold match', baseCompliance['is_scaffold_match']);
  pushFlag(baseRows, 'Whitelisted', baseCompliance['is_whitelisted'], 'ok');
  if (typeof baseCompliance['faves_flag_count'] === 'number' && (baseCompliance['faves_flag_count'] as number) > 0) {
    baseRows.push(['Flag count', `<span class="badge warn">${baseCompliance['faves_flag_count']}</span>`]);
  }
  if (baseRows.length > 0) appendSection(el, 'Regulatory flags', () => {
    const dl = document.createElement('dl');
    dl.className = 'props-grid';
    for (const [label, value] of baseRows) {
      const dt = document.createElement('dt'); dt.textContent = label;
      const dd = document.createElement('dd'); dd.innerHTML = value;
      dl.appendChild(dt); dl.appendChild(dd);
    }
    return dl;
  });

  const risk = result['risk_assessment'] as Record<string, unknown> | undefined;
  if (risk && Object.keys(risk).length > 0) appendSection(el, 'Risk assessment', () => renderKvBlock(risk));

  const pathway = result['regulatory_pathway'] as Record<string, unknown> | string | undefined;
  if (pathway) appendSection(el, 'Regulatory pathway', () => {
    const block = document.createElement('div');
    block.className = 'compliance-block';
    if (typeof pathway === 'string') block.textContent = pathway;
    else block.appendChild(renderKvBlock(pathway));
    return block;
  });

  const recs = result['recommendations'];
  if (Array.isArray(recs) && recs.length > 0) appendSection(el, 'Recommendations', () => {
    const ul = document.createElement('ul');
    ul.className = 'compliance-recs';
    for (const r of recs.slice(0, 8)) {
      const li = document.createElement('li');
      li.textContent = String(r);
      ul.appendChild(li);
    }
    return ul;
  });

  const dims = ctxCompliance['dimensions'] as Record<string, unknown> | undefined;
  if (dims && Object.keys(dims).length > 0) appendSection(el, 'Dimensions', () => {
    const wrap = document.createElement('div');
    wrap.className = 'faves-dims';
    for (const [name, dim] of Object.entries(dims)) {
      const status = String((dim as Record<string, unknown>)?.['status'] ?? '');
      const cls = status === 'PASS' ? 'ok' : status === 'WARN' ? 'warn' : status === 'FAIL' ? 'err' : '';
      const dimEl = document.createElement('div');
      dimEl.className = 'faves-dim';
      dimEl.innerHTML = `<div class="faves-name">${escapeText(humanizeKey(name))}</div><span class="badge ${cls}">${escapeText(status || '—')}</span>`;
      wrap.appendChild(dimEl);
    }
    return wrap;
  });
}

function appendSection(parent: HTMLElement, title: string, builder: () => HTMLElement): void {
  const heading = document.createElement('h4');
  heading.textContent = title;
  heading.style.cssText = 'margin: 14px 0 6px 0; font-size: 11px; letter-spacing: 0.05em; text-transform: uppercase; color: var(--text-muted); font-weight: 500;';
  parent.appendChild(heading);
  parent.appendChild(builder());
}

function renderKvBlock(obj: Record<string, unknown>): HTMLElement {
  const dl = document.createElement('dl');
  dl.className = 'props-grid';
  for (const [k, v] of Object.entries(obj)) {
    if (v == null || v === '') continue;
    if (Array.isArray(v) && v.length === 0) continue;
    const dt = document.createElement('dt'); dt.textContent = humanizeKey(k);
    const dd = document.createElement('dd'); dd.appendChild(renderValue(v));
    dl.appendChild(dt); dl.appendChild(dd);
  }
  return dl;
}

function renderValue(v: unknown): Node {
  if (v == null || v === '') return document.createTextNode('—');
  if (typeof v === 'boolean') {
    const span = document.createElement('span');
    span.className = `badge ${v ? 'warn' : ''}`;
    span.textContent = v ? 'yes' : 'no';
    return span;
  }
  if (typeof v === 'number') return document.createTextNode(formatNum(v));
  if (typeof v === 'string') return document.createTextNode(v);
  if (Array.isArray(v)) {
    if (v.length === 0) return document.createTextNode('—');
    if (v.every((x) => typeof x !== 'object' || x === null)) {
      return document.createTextNode(v.map((x) => String(x)).join(', '));
    }
    const wrap = document.createElement('div');
    wrap.className = 'kv-array';
    for (const item of v.slice(0, 8)) {
      if (item && typeof item === 'object') {
        const card = document.createElement('div');
        card.className = 'kv-array-item';
        card.appendChild(renderKvBlock(item as Record<string, unknown>));
        wrap.appendChild(card);
      } else {
        const li = document.createElement('div');
        li.textContent = String(item);
        wrap.appendChild(li);
      }
    }
    return wrap;
  }
  if (typeof v === 'object') return renderKvBlock(v as Record<string, unknown>);
  return document.createTextNode(String(v));
}

function pushFlag(rows: Array<[string, string]>, label: string, value: unknown, trueClass: 'ok' | 'warn' | 'err' = 'warn'): void {
  if (value === true) rows.push([label, `<span class="badge ${trueClass}">yes</span>`]);
  else if (value === false) rows.push([label, '<span class="badge">no</span>']);
}

function verdictBadgeClass(s: string): string {
  const u = s.toUpperCase();
  if (u === 'PROCEED' || u === 'PASS' || u === 'CLEAR') return 'ok';
  if (u === 'CAUTION' || u === 'CONDITIONAL' || u === 'WARN' || u === 'REVIEW_REQUIRED') return 'warn';
  if (u === 'STOP' || u === 'BLOCKED' || u === 'FAIL' || u === 'REJECTED') return 'err';
  return '';
}

function complianceBadgeClass(status: string): string {
  const s = status.toLowerCase();
  if (s === 'clear' || s === 'ok' || s === 'pass') return 'ok';
  if (s === 'controlled' || s === 'flagged' || s === 'warn') return 'warn';
  if (s === 'blocked' || s === 'fail' || s === 'rejected') return 'err';
  return '';
}

// ─── Similar tab ───────────────────────────────────────────────────────

let similarThreshold = 0.7;

function setupSimilarHandlers(): void {
  $<HTMLButtonElement>('similar-run').addEventListener('click', () => { void runSimilar(); });
  document.querySelectorAll<HTMLButtonElement>('.similar-thresh').forEach((b) => {
    b.addEventListener('click', () => {
      const v = Number(b.dataset['similarThreshold']);
      if (!isFinite(v)) return;
      similarThreshold = v;
      document.querySelectorAll<HTMLElement>('.similar-thresh').forEach((x) => x.classList.toggle('active', x === b));
    });
  });
}

function resetSimilarResults(): void {
  const out = $<HTMLElement>('similar-results');
  out.innerHTML = '';
  const btn = $<HTMLButtonElement>('similar-run');
  btn.disabled = false; btn.textContent = 'Search neighbors';
}

async function runSimilar(): Promise<void> {
  if (!currentSmiles) return;
  const out = $<HTMLElement>('similar-results');
  const btn = $<HTMLButtonElement>('similar-run');
  btn.disabled = true; btn.textContent = 'Searching…';
  out.innerHTML = '<p class="muted" style="font-size:12px;">Searching for neighbors…</p>';
  try {
    const r = await searchSimilar(currentSmiles, { top_k: 10, min_similarity: similarThreshold });
    out.innerHTML = renderSimilarResults(r.result);
    void renderCrossSurface(currentSmiles);
  } catch (e) {
    out.innerHTML = `<p class="muted" style="color: var(--error); font-size: 12px;">${escapeText(e instanceof Error ? e.message : 'Network error')}</p>`;
  } finally {
    btn.disabled = false; btn.textContent = 'Re-run';
  }
}

function renderSimilarResults(data: unknown): string {
  const d = (data || {}) as Record<string, unknown>;
  const rows: Array<Record<string, unknown>> =
    (Array.isArray(d['results']) ? d['results'] :
     Array.isArray(d['similar_molecules']) ? d['similar_molecules'] :
     Array.isArray(d['matches']) ? d['matches'] :
     Array.isArray(data) ? (data as Array<Record<string, unknown>>) : []) as Array<Record<string, unknown>>;
  if (currentSmiles) similarCache.set(currentSmiles, rows);
  setInsertEnabled('similar', rows.length > 0);
  if (rows.length === 0) {
    return `<p class="muted" style="font-size:12px;">No neighbors at threshold ${similarThreshold.toFixed(1)}. Try a lower threshold.</p>`;
  }
  const items = rows.slice(0, 20).map((row) => {
    const smiles = String(row['smiles'] ?? '');
    const sim = (row['similarity'] ?? row['tanimoto'] ?? row['score']) as number | undefined;
    const cid = row['cid'];
    const mw = row['molecular_weight'];
    const qed = row['qed'];
    const logp = row['logp'] ?? row['xlogp'];
    const status = (row['compliance_status'] ?? row['status']) as string | undefined;
    const cidLink = cid != null ? `<a href="https://pubchem.ncbi.nlm.nih.gov/compound/${escapeText(String(cid))}" target="_blank" rel="noopener" class="cid">CID ${escapeText(String(cid))}</a>` : '';
    const simLabel = typeof sim === 'number' ? `<span class="badge">${(sim * 100).toFixed(0)}% Tc</span>` : '';
    const statusBadge = status ? `<span class="badge ${complianceBadgeClass(status)}">${escapeText(status)}</span>` : '';
    const meta = [
      typeof mw === 'number' ? `MW ${formatNum(mw)}` : '',
      typeof logp === 'number' ? `LogP ${formatNum(logp)}` : '',
      typeof qed === 'number' ? `QED ${formatNum(qed)}` : '',
    ].filter(Boolean).join(' · ');
    return `<div class="similar-row" data-smiles="${escapeText(smiles)}">
      <div class="similar-head">${simLabel}${statusBadge}${cidLink}</div>
      <div class="similar-smiles">${escapeText(smiles)}</div>
      ${meta ? `<div class="similar-meta">${meta}</div>` : ''}
    </div>`;
  }).join('');
  setTimeout(() => {
    document.querySelectorAll<HTMLElement>('.similar-row').forEach((el) => {
      el.addEventListener('click', () => {
        const s = el.dataset['smiles'];
        if (s) { void load(s); switchTab('profile'); }
      });
    });
  }, 0);
  return items;
}

// ─── Advanced tab ──────────────────────────────────────────────────────

type ComputeTool = 'predict_pka' | 'predict_solubility' | 'predict_bde' | 'predict_frontier_orbitals';

const TOOL_LABEL: Record<ComputeTool, string> = {
  predict_pka: 'pKa',
  predict_solubility: 'Solubility',
  predict_bde: 'BDE',
  predict_frontier_orbitals: 'Frontier orbitals',
};

const TOOL_FN: Record<ComputeTool, (smiles: string) => Promise<ApiResponse<unknown>>> = {
  predict_pka: predictPka,
  predict_solubility: predictSolubility,
  predict_bde: predictBde,
  predict_frontier_orbitals: predictFrontierOrbitals,
};

async function refreshAdvancedGate(): Promise<void> {
  // Sideload build: always show the advanced tools. If the engine doesn't
  // have NOVOMCP_QM_URL / NOVOMCP_PROPERTIES_URL wired, individual tool
  // calls return structured 503 responses; the per-card error UX handles
  // that gracefully. No gate at the tab level — users can see what's
  // available and hit a clear "compute service not configured" error
  // if they try to use one that isn't wired.
  const lock = $<HTMLElement>('advanced-lock');
  const locked = $<HTMLElement>('advanced-locked');
  const tools = $<HTMLElement>('advanced-tools');
  lock.hidden = true;
  locked.hidden = true;
  tools.hidden = false;
}

function setupAdvancedHandlers(): void {
  document.querySelectorAll<HTMLElement>('.compute-card').forEach((card) => {
    const tool = card.dataset['tool'] as ComputeTool;
    const btn = card.querySelector<HTMLButtonElement>('button[data-action="run"]');
    if (!btn) return;
    btn.addEventListener('click', () => { void runComputeTool(tool, card); });
  });
}

function resetAdvancedResults(): void {
  document.querySelectorAll<HTMLElement>('.compute-card').forEach((card) => {
    const tool = card.dataset['tool'] as ComputeTool;
    const body = card.querySelector<HTMLElement>('[data-result]');
    if (!body) return;
    body.innerHTML = defaultHint(tool);
    const btn = card.querySelector<HTMLButtonElement>('button[data-action="run"]');
    if (btn) { btn.disabled = false; btn.textContent = 'Run'; }
  });
}

function defaultHint(tool: ComputeTool): string {
  switch (tool) {
    case 'predict_pka': return 'Click run to predict ionization sites and pKa values.';
    case 'predict_solubility': return 'Aqueous solubility (logS) with category labels.';
    case 'predict_bde': return 'BDE per bond — flag the weakest links.';
    case 'predict_frontier_orbitals': return 'HOMO, LUMO, and gap (eV).';
  }
}

async function runComputeTool(tool: ComputeTool, card: HTMLElement): Promise<void> {
  if (!currentSmiles) return;
  const btn = card.querySelector<HTMLButtonElement>('button[data-action="run"]')!;
  const body = card.querySelector<HTMLElement>('[data-result]')!;
  btn.disabled = true; btn.textContent = 'Running…';
  body.innerHTML = '<span class="muted">Running</span>';
  try {
    const r = await TOOL_FN[tool](currentSmiles);
    body.innerHTML = renderComputeResult(tool, r.result);
    if (currentSmiles) void renderCrossSurface(currentSmiles);
  } catch (e) {
    const raw = e instanceof Error ? e.message : 'failed';
    body.innerHTML = `<span class="muted" style="color: var(--error);">${escapeText(friendlyComputeError(tool, raw))}</span>`;
  } finally {
    btn.disabled = false; btn.textContent = 'Re-run';
  }
}

function friendlyComputeError(tool: ComputeTool, raw: string): string {
  const label = TOOL_LABEL[tool] ?? 'Calculation';
  const codeMatch = raw.match(/\((\d{3})\)/);
  const status = codeMatch ? Number(codeMatch[1]) : null;
  let detail = '';
  const jsonStart = raw.indexOf('{');
  if (jsonStart >= 0) {
    try {
      const parsed = JSON.parse(raw.slice(jsonStart));
      if (parsed && typeof parsed === 'object' && 'detail' in parsed) {
        detail = String((parsed as Record<string, unknown>)['detail'] ?? '');
      }
    } catch {}
  }
  if (detail) {
    detail = detail.replace(/[—-]\s*check\s+\w+_URL\s+configuration[\s\S]*$/i, '').trim();
    detail = detail.replace(/for:\s*[A-Za-z0-9@+\-\[\]()=#$/\\.]{20,}/, 'for this molecule');
    if (detail) return detail;
  }
  if (status && status >= 500) return `${label} service is temporarily unavailable. Try again in a moment.`;
  if (status === 422 || status === 400) return `${label} couldn't process this molecule. Try a different scaffold.`;
  if (status === 401) return `${label} requires a valid Compute (ncmcp_) key. Add or update it from sign-in.`;
  if (status === 402) return `${label} needs more credits.`;
  if (status === 429) return `${label} is rate-limited right now. Try again shortly.`;
  const trimmed = raw.length > 140 ? raw.slice(0, 140) + '…' : raw;
  return `${label}: ${trimmed}`;
}

function renderComputeResult(tool: ComputeTool, data: unknown): string {
  const d = (data || {}) as Record<string, unknown>;
  switch (tool) {
    case 'predict_pka': {
      const values = (d['pka_values'] as unknown[]) || [];
      const groups = (d['ionizable_groups'] as unknown[]) || [];
      const interp = d['interpretation'] as string | undefined;
      if (values.length === 0 || (groups.length === 1 && groups[0] === 'none_detected')) {
        return `<span class="muted">No ionizable groups detected. Molecule expected to be neutral across physiological pH range.</span>`;
      }
      const rows = values.slice(0, 6).map((v, i) => {
        const site = groups[i] ?? `site ${i + 1}`;
        return `<dt>${escapeText(String(site))}</dt><dd>pKa ${formatNum(v)}</dd>`;
      }).join('');
      return `<dl class="props-grid">${rows}</dl>${interp ? `<div class="hint" style="margin-top:6px;">${escapeText(interp)}</div>` : ''}`;
    }
    case 'predict_solubility': {
      const rows: string[] = [];
      if (typeof d['logS'] === 'number') rows.push(`<dt>logS</dt><dd>${formatNum(d['logS'])}</dd>`);
      if (typeof d['solubility_mg_ml'] === 'number') rows.push(`<dt>Solubility</dt><dd>${formatNum(d['solubility_mg_ml'])} mg/mL</dd>`);
      if (d['category']) rows.push(`<dt>Class</dt><dd>${escapeText(String(d['category']))}</dd>`);
      if (d['temperature']) rows.push(`<dt>Temperature</dt><dd>${escapeText(String(d['temperature']))}</dd>`);
      if (rows.length === 0) return '<span class="muted">No solubility data returned.</span>';
      return `<dl class="props-grid">${rows.join('')}</dl>`;
    }
    case 'predict_bde': {
      const weakest = d['weakest_bond'] as Record<string, unknown> | null | undefined;
      const interp = d['interpretation'] as string | undefined;
      const bondCount = d['bond_count'];
      const rows: string[] = [];
      if (weakest && weakest['atoms'] != null) {
        rows.push(`<dt>Weakest bond</dt><dd>${escapeText(String(weakest['atoms']))}</dd>`);
        if (typeof weakest['bde_kcal_mol'] === 'number') rows.push(`<dt>BDE</dt><dd>${formatNum(weakest['bde_kcal_mol'])} kcal/mol</dd>`);
      }
      if (typeof bondCount === 'number') rows.push(`<dt>Bonds analyzed</dt><dd>${bondCount}</dd>`);
      if (rows.length === 0) {
        const bonds = (d['bonds'] as Array<Record<string, unknown>>) || [];
        if (bonds.length === 0) return '<span class="muted">No BDE data returned.</span>';
        const top = bonds.slice().sort((a, b) => Number(a['bde_kcal_mol']) - Number(b['bde_kcal_mol'])).slice(0, 3);
        rows.push(...top.map((b) => `<dt>${escapeText(String(b['atoms'] ?? '—'))}</dt><dd>${formatNum(b['bde_kcal_mol'])} kcal/mol</dd>`));
      }
      return `<dl class="props-grid">${rows.join('')}</dl>${interp ? `<div class="hint" style="margin-top:6px;">${escapeText(interp)}</div>` : ''}`;
    }
    case 'predict_frontier_orbitals': {
      const candidates: Array<[string, string]> = [
        ['HOMO', 'homo_ev'], ['HOMO', 'homo'],
        ['LUMO', 'lumo_ev'], ['LUMO', 'lumo'],
        ['Gap', 'gap_ev'], ['Gap', 'gap'], ['Gap', 'homo_lumo_gap'],
        ['Dipole', 'dipole_debye'], ['Dipole', 'dipole'],
      ];
      const seen = new Set<string>();
      const rows: string[] = [];
      for (const [label, key] of candidates) {
        if (seen.has(label)) continue;
        const v = d[key];
        if (typeof v === 'number') { rows.push(`<dt>${label}</dt><dd>${formatNum(v)} eV</dd>`); seen.add(label); }
      }
      if (rows.length === 0) {
        for (const [k, v] of Object.entries(d)) {
          if (k === 'smiles' || k === 'method' || k === 'units') continue;
          if (typeof v === 'number') rows.push(`<dt>${humanizeKey(k)}</dt><dd>${formatNum(v)}</dd>`);
        }
      }
      if (rows.length === 0) return '<span class="muted">No orbital data returned.</span>';
      return `<dl class="props-grid">${rows.join('')}</dl>`;
    }
  }
}

// ─── Cross-surface CTAs ───────────────────────────────────────────────

async function renderCrossSurface(smiles: string): Promise<void> {
  const funnelId = getCurrentFunnelId();
  const wrap = $<HTMLElement>('cross-surface');
  if (!funnelId) { wrap.hidden = true; return; }
  wrap.hidden = false;

  $<HTMLAnchorElement>('cta-dashboard').href = dashboardUrl(funnelId);

  const aiBtn = $<HTMLButtonElement>('cta-ai-assistant');
  aiBtn.onclick = async () => {
    try {
      const text = aiHandoffPrompt(funnelId, smiles);
      await navigator.clipboard.writeText(text);
      const label = $<HTMLElement>('cta-ai-label');
      const arrow = $<HTMLElement>('cta-ai-arrow');
      const originalLabel = label.textContent;
      const originalArrow = arrow.textContent;
      label.textContent = 'Copied — paste into your AI assistant';
      arrow.textContent = '✓';
      aiBtn.classList.add('copied');
      window.setTimeout(() => {
        label.textContent = originalLabel;
        arrow.textContent = originalArrow;
        aiBtn.classList.remove('copied');
      }, 2000);
    } catch {
      const label = $<HTMLElement>('cta-ai-label');
      const original = label.textContent;
      label.textContent = 'Clipboard blocked — paste manually from below';
      window.setTimeout(() => { label.textContent = original; }, 2000);
    }
  };

  $<HTMLElement>('cta-funnel-id').textContent = funnelId;
}

// ─── Helpers ──────────────────────────────────────────────────────────

function pushIfPresent(rows: Array<[string, string]>, label: string, value: unknown, fmt: (v: unknown) => string): void {
  if (value == null || value === '') return;
  if (typeof value === 'number' && !isFinite(value)) return;
  const formatted = fmt(value);
  if (formatted === '—' || formatted === '') return;
  rows.push([label, formatted]);
}

function renderGrid(id: string, rows: Array<[string, string]>): void {
  const el = $<HTMLElement>(id);
  el.innerHTML = '';
  for (const [label, value] of rows) {
    const dt = document.createElement('dt'); dt.textContent = label;
    const dd = document.createElement('dd'); dd.innerHTML = value;
    el.appendChild(dt); el.appendChild(dd);
  }
}

function metaPill(text: string): HTMLElement {
  const span = document.createElement('span');
  span.className = 'badge';
  span.textContent = text;
  return span;
}

function formatNum(v: unknown): string {
  if (typeof v !== 'number' || !isFinite(v)) return '—';
  return v.toFixed(2);
}

function formatInt(v: unknown): string {
  if (typeof v !== 'number' || !isFinite(v)) return '—';
  return Math.round(v).toString();
}

function humanizeKey(k: string): string {
  return k
    .replace(/_/g, ' ')
    .replace(/\bcyp\b/gi, 'CYP')
    .replace(/\bbbb\b/gi, 'BBB')
    .replace(/\bherg\b/gi, 'hERG')
    .replace(/\bdili\b/gi, 'DILI')
    .replace(/\bhia\b/gi, 'HIA')
    .replace(/\bpgp\b/gi, 'P-gp');
}

function escapeText(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ─── Insert-as-table — drop tab data into Word at cursor ─────────────

type InsertKind = 'profile' | 'admet' | 'compliance' | 'similar';

function setupInsertHandlers(): void {
  document.querySelectorAll<HTMLButtonElement>('.insert-table-btn').forEach((btn) => {
    btn.addEventListener('click', () => { void handleInsert(btn); });
  });
}

function setInsertEnabled(kind: InsertKind, enabled: boolean): void {
  const btn = document.querySelector<HTMLButtonElement>(`.insert-table-btn[data-insert="${kind}"]`);
  if (!btn) return;
  btn.disabled = !enabled;
  btn.classList.remove('inserted');
  btn.textContent = 'Insert as Word table';
}

async function handleInsert(btn: HTMLButtonElement): Promise<void> {
  const kind = btn.dataset['insert'] as InsertKind | undefined;
  if (!kind || !currentSmiles) return;
  const table = buildTable(kind, currentSmiles);
  if (!table || table.rows.length === 0) return;

  btn.disabled = true;
  btn.textContent = 'Inserting…';
  try {
    await insertTableAtSelection(table);
    btn.classList.add('inserted');
    btn.textContent = '✓ Inserted at cursor';
    window.setTimeout(() => {
      btn.classList.remove('inserted');
      btn.textContent = 'Insert as Word table';
      btn.disabled = false;
    }, 1800);
  } catch (e) {
    btn.textContent = `Failed — ${e instanceof Error ? e.message.slice(0, 60) : 'unknown error'}`;
    window.setTimeout(() => {
      btn.textContent = 'Insert as Word table';
      btn.disabled = false;
    }, 3000);
  }
}

function buildTable(kind: InsertKind, smiles: string): InsertableTable | null {
  switch (kind) {
    case 'profile': return buildProfileTable(smiles);
    case 'admet': return buildAdmetTable(smiles);
    case 'compliance': return buildComplianceTable(smiles);
    case 'similar': return buildSimilarTable(smiles);
  }
}

function buildProfileTable(smiles: string): InsertableTable | null {
  const profile = profileCache.get(smiles);
  if (!profile) return null;
  const props = (profile.result.properties || {}) as Record<string, unknown>;
  const pick = (...keys: string[]): unknown => {
    for (const k of keys) if (props[k] != null) return props[k];
    return undefined;
  };
  const fmtNum = (v: unknown): string => typeof v === 'number' && isFinite(v) ? v.toFixed(2) : '—';
  const fmtInt = (v: unknown): string => typeof v === 'number' && isFinite(v) ? String(Math.round(v)) : '—';
  const rows: string[][] = [];
  const push = (label: string, value: string): void => { if (value && value !== '—') rows.push([label, value]); };
  push('SMILES', smiles);
  push('CID', fmtInt(props['cid']));
  push('Formula', String(props['molecular_formula'] ?? ''));
  push('MW (g/mol)', fmtNum(props['molecular_weight']));
  push('LogP', fmtNum(props['logp']));
  push('TPSA', fmtNum(props['tpsa']));
  push('QED', fmtNum(props['qed']));
  push('Drug-likeness', fmtNum(props['drug_likeness']));
  push('Synth. accessibility', fmtNum(props['synthetic_accessibility']));
  push('Complexity', fmtNum(props['complexity']));
  push('Fsp³', fmtNum(props['fsp3']));
  push('HBD', fmtInt(pick('hbd_count', 'hbd')));
  push('HBA', fmtInt(pick('hba_count', 'hba')));
  push('Rotatable bonds', fmtInt(pick('rotatable_bond_count', 'rotatable_bonds')));
  push('Heavy atoms', fmtInt(props['heavy_atom_count']));
  push('Aromatic rings', fmtInt(pick('aromatic_ring_count', 'aromatic_rings')));
  push('Lipinski violations', fmtInt(props['lipinski_violations']));
  return { headers: ['Property', 'Value'], rows, caption: `NovoMCP profile · ${smiles}` };
}

function buildAdmetTable(smiles: string): InsertableTable | null {
  const normalized = admetCache.get(smiles);
  if (!normalized) return null;
  const rows: string[][] = [];
  const ORDER: Array<[AdmetCategory, string]> = [
    ['absorption', 'Absorption'],
    ['distribution', 'Distribution'],
    ['metabolism', 'Metabolism'],
    ['excretion', 'Excretion'],
    ['toxicity', 'Toxicity'],
  ];
  for (const [cat, label] of ORDER) {
    const fields = normalized[cat];
    if (!fields || fields.length === 0) continue;
    for (const f of fields) {
      const value = f.value.toFixed(f.unit ? 2 : 3) + (f.unit ? ` ${f.unit}` : '');
      rows.push([label, f.label, value, f.classification]);
    }
  }
  if (rows.length === 0) return null;
  return {
    headers: ['Category', 'Property', 'Value', 'Classification'],
    rows,
    caption: `NovoMCP ADMET · ${smiles}`,
  };
}

function buildComplianceTable(smiles: string): InsertableTable | null {
  const result = complianceCache.get(smiles);
  if (!result) return null;
  const rows: string[][] = [];
  const overall = String(result['overall_status'] ?? '');
  if (overall) rows.push(['Verdict', 'Pharmaceutical · US', overall]);

  const base = (result['base_compliance'] || {}) as Record<string, unknown>;
  const baseStatus = base['status'];
  if (typeof baseStatus === 'string') rows.push(['Base', 'Status', baseStatus]);
  for (const [k, label] of [
    ['is_dea_controlled', 'DEA controlled'],
    ['is_fda_banned', 'FDA banned'],
    ['is_cwc_scheduled', 'CWC scheduled'],
    ['is_epa_pbt', 'EPA PBT'],
    ['is_eu_reach_banned', 'EU REACH banned'],
    ['is_scaffold_match', 'Scaffold match'],
    ['is_whitelisted', 'Whitelisted'],
  ] as Array<[string, string]>) {
    const v = base[k];
    if (typeof v === 'boolean') rows.push(['Base', label, v ? 'Yes' : 'No']);
  }
  if (typeof base['faves_flag_count'] === 'number') {
    rows.push(['Base', 'Flag count', String(base['faves_flag_count'])]);
  }

  const recs = result['recommendations'];
  if (Array.isArray(recs)) {
    for (let i = 0; i < Math.min(recs.length, 6); i++) {
      rows.push(['Recommendation', `#${i + 1}`, String(recs[i])]);
    }
  }

  if (rows.length === 0) return null;
  return {
    headers: ['Section', 'Field', 'Value'],
    rows,
    caption: `NovoMCP compliance · ${smiles}`,
  };
}

function buildSimilarTable(smiles: string): InsertableTable | null {
  const rows = similarCache.get(smiles);
  if (!rows || rows.length === 0) return null;

  // Defensive field probing — neighbor responses sometimes wrap properties
  // in a nested `properties` dict; sometimes they're flat at the row level;
  // sometimes (isotope-labeled or chain-concatenated SMILES that came from
  // outside the indexed PubChem snapshot) the property fields aren't
  // present at all.
  const pickNum = (row: Record<string, unknown>, ...keys: string[]): number | undefined => {
    const props = (row['properties'] || {}) as Record<string, unknown>;
    for (const k of keys) {
      if (typeof row[k] === 'number') return row[k] as number;
      if (typeof props[k] === 'number') return props[k] as number;
    }
    return undefined;
  };
  const pickStr = (row: Record<string, unknown>, ...keys: string[]): string | undefined => {
    for (const k of keys) {
      if (typeof row[k] === 'string' && row[k]) return row[k] as string;
    }
    // Boolean compliance fields → derived label
    if (row['is_dea_controlled'] === true || row['is_fda_banned'] === true || row['is_cwc_scheduled'] === true) return 'controlled';
    if (row['is_whitelisted'] === true) return 'whitelisted';
    if (row['is_scaffold_match'] === true) return 'flagged';
    return undefined;
  };

  type Col = { header: string; render: (row: Record<string, unknown>) => string };
  const allCols: Col[] = [
    { header: 'Tc',     render: (r) => { const s = pickNum(r, 'similarity', 'tanimoto', 'score'); return typeof s === 'number' ? `${(s * 100).toFixed(0)}%` : ''; } },
    { header: 'CID',    render: (r) => r['cid'] != null ? String(r['cid']) : '' },
    { header: 'SMILES', render: (r) => String(r['smiles'] ?? '') },
    { header: 'MW',     render: (r) => { const v = pickNum(r, 'molecular_weight', 'mw'); return typeof v === 'number' ? v.toFixed(2) : ''; } },
    { header: 'LogP',   render: (r) => { const v = pickNum(r, 'xlogp', 'logp'); return typeof v === 'number' ? v.toFixed(2) : ''; } },
    { header: 'QED',    render: (r) => { const v = pickNum(r, 'qed', 'drug_likeness'); return typeof v === 'number' ? v.toFixed(2) : ''; } },
    { header: 'TPSA',   render: (r) => { const v = pickNum(r, 'tpsa'); return typeof v === 'number' ? v.toFixed(2) : ''; } },
    { header: 'Status', render: (r) => pickStr(r, 'compliance_status', 'status') ?? '' },
  ];

  // Render every cell once so we know which columns are entirely empty —
  // those get dropped so the inserted Word table doesn't ship blank
  // columns for properties the search response didn't include.
  const cells: string[][] = rows.map((row) => allCols.map((col) => col.render(row)));
  const keepCol: boolean[] = allCols.map((_, c) => cells.some((row) => (row[c] ?? '') !== ''));
  const cols = allCols.filter((_, c) => keepCol[c]);
  const finalRows = cells.map((row) => row.filter((_, c) => keepCol[c]));

  // Always keep the three core columns even if (somehow) empty — Tc/CID/SMILES
  // are the identifying minimum for a neighbor table. If even those are
  // empty there's nothing useful to insert.
  if (cols.length === 0 || finalRows.length === 0) return null;

  return {
    headers: cols.map((c) => c.header),
    rows: finalRows,
    caption: `NovoMCP neighbors of ${smiles}`,
  };
}

// ─── Compound-table scan ──────────────────────────────────────────────

interface ScanResult {
  table: ParsedTable;
  schema: TableSchema | null;
  /** Reason for skipping when schema is null. */
  skipReason?: string;
  rowCount: number;
  validatedCount: number;
  discrepancies: Discrepancy[];
}

let pendingDiscrepancies: Discrepancy[] = [];

function setupScanHandlers(): void {
  $<HTMLButtonElement>('scan-btn').addEventListener('click', () => { void runScan(); });
  $<HTMLButtonElement>('scan-close').addEventListener('click', () => closeScan());
  $<HTMLButtonElement>('scan-apply-btn').addEventListener('click', () => { void applyComments(); });
}

function openScan(): void {
  $<HTMLElement>('scan-block').hidden = false;
  $<HTMLElement>('scan-progress').hidden = true;
  $<HTMLElement>('scan-summary').hidden = true;
  $<HTMLElement>('scan-actions').hidden = true;
  $<HTMLElement>('scan-error').hidden = true;
  $<HTMLElement>('scan-tables').innerHTML = '';
  // Hide the single-molecule UI while a scan is open — the two flows
  // share screen real estate but are mentally distinct.
  $<HTMLElement>('active-block').hidden = true;
  $<HTMLElement>('empty-active').hidden = true;
}

function closeScan(): void {
  $<HTMLElement>('scan-block').hidden = true;
  pendingDiscrepancies = [];
  // Restore the single-molecule view if there's an active SMILES,
  // otherwise show the empty hint.
  if (currentSmiles) $<HTMLElement>('active-block').hidden = false;
  else $<HTMLElement>('empty-active').hidden = false;
}

async function runScan(): Promise<void> {
  openScan();
  const progressBlock = $<HTMLElement>('scan-progress');
  const progressText = $<HTMLElement>('scan-progress-text');
  const progressFill = $<HTMLElement>('scan-progress-fill');
  progressBlock.hidden = false;
  progressText.textContent = 'Reading tables…';
  progressFill.style.width = '8%';

  let tables: ParsedTable[];
  try {
    tables = await readAllTables();
  } catch (e) {
    showScanError(e instanceof Error ? e.message : 'Failed to read tables');
    return;
  }

  if (tables.length === 0) {
    progressBlock.hidden = true;
    paintScanSummary({ tablesScanned: 0, validated: 0, discrepancies: 0, skipped: 0 });
    $<HTMLElement>('scan-tables').innerHTML = '<p class="muted" style="font-size: 12px;">No tables found in this document.</p>';
    return;
  }

  // Identify which tables have SMILES schemas. Skip the rest.
  const work: ScanResult[] = tables.map((t) => {
    const schema = detectSchema(t);
    if (!schema) {
      return { table: t, schema: null, skipReason: 'No SMILES column detected', rowCount: t.cells.length, validatedCount: 0, discrepancies: [] };
    }
    // rowCount = compound (data) rows, excluding the header — so the "N rows"
    // label matches the validated-compound count.
    return { table: t, schema, rowCount: Math.max(0, t.cells.length - schema.dataStart), validatedCount: 0, discrepancies: [] };
  });

  // Collect every distinct SMILES across all detected schemas. Dedupe to
  // avoid paying credits twice for the same molecule appearing in
  // multiple tables.
  const smilesToFetch = new Set<string>();
  for (const w of work) {
    if (!w.schema) continue;
    for (let r = w.schema.dataStart; r < w.table.cells.length; r++) {
      const row = w.table.cells[r];
      if (!row) continue;
      const s = (row[w.schema.smilesCol] ?? '').trim();
      // Only fetch/count cells that are actually SMILES — guards against a
      // header label ("SMILES"/"Structure") or stray text being counted as a
      // compound when header detection didn't strip row 0.
      if (s && isLikelySmiles(s)) smilesToFetch.add(s);
    }
  }

  const total = smilesToFetch.size;
  if (total === 0) {
    progressBlock.hidden = true;
    paintScanResults(work, { tablesScanned: tables.length, validated: 0, discrepancies: 0, skipped: work.filter((w) => !w.schema).length });
    return;
  }

  progressText.textContent = `Validating ${total} compound${total === 1 ? '' : 's'}…`;
  progressFill.style.width = '15%';

  // Concurrent profile fetches with a small in-flight cap. Browsers
  // throttle ~6 connections per origin, but we cap explicitly so the
  // progress bar updates smoothly across the batch.
  const profiles = new Map<string, ApiResponse<MoleculeProfile> | null>();
  const list = Array.from(smilesToFetch);
  const CONCURRENCY = 6;
  let done = 0;

  async function worker(): Promise<void> {
    while (true) {
      const next = list.shift();
      if (!next) return;
      try {
        const r = await getMoleculeProfile(next);
        profiles.set(next, r);
      } catch {
        profiles.set(next, null);
      } finally {
        done++;
        const pct = 15 + Math.floor((done / total) * 80);
        progressFill.style.width = `${pct}%`;
        progressText.textContent = `Validated ${done} of ${total}…`;
      }
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, total) }, worker));

  // Wire profiles back into per-table discrepancies.
  for (const w of work) {
    if (!w.schema) continue;
    const serverByRow = new Map<number, ReturnType<typeof extractServerValues>>();
    for (let r = w.schema.dataStart; r < w.table.cells.length; r++) {
      const row = w.table.cells[r];
      if (!row) continue;
      const s = (row[w.schema.smilesCol] ?? '').trim();
      if (!s || !isLikelySmiles(s)) continue;
      const profile = profiles.get(s);
      if (!profile) continue;
      serverByRow.set(r, extractServerValues(s, profile));
      w.validatedCount++;
    }
    w.discrepancies = findDiscrepancies(w.table, w.schema, serverByRow);
  }

  progressBlock.hidden = true;
  paintScanResults(work, {
    tablesScanned: tables.length,
    validated: work.reduce((acc, w) => acc + w.validatedCount, 0),
    discrepancies: work.reduce((acc, w) => acc + w.discrepancies.length, 0),
    skipped: work.filter((w) => !w.schema).length,
  });
}

function paintScanSummary(stats: { tablesScanned: number; validated: number; discrepancies: number; skipped: number }): void {
  const summary = $<HTMLElement>('scan-summary');
  summary.hidden = false;
  summary.innerHTML = `
    <span class="stat"><strong>${stats.tablesScanned}</strong> table${stats.tablesScanned === 1 ? '' : 's'}</span>
    <span class="stat"><strong>${stats.validated}</strong> validated</span>
    <span class="stat"><strong>${stats.discrepancies}</strong> discrepanc${stats.discrepancies === 1 ? 'y' : 'ies'}</span>
    ${stats.skipped > 0 ? `<span class="stat muted">${stats.skipped} skipped</span>` : ''}
  `;
}

function paintScanResults(work: ScanResult[], stats: { tablesScanned: number; validated: number; discrepancies: number; skipped: number }): void {
  paintScanSummary(stats);

  const wrap = $<HTMLElement>('scan-tables');
  wrap.innerHTML = '';
  pendingDiscrepancies = [];

  for (const w of work) {
    const div = document.createElement('div');
    div.className = 'scan-table';

    const head = document.createElement('div');
    head.className = 'scan-table-head';
    head.innerHTML = `
      <span class="scan-table-name">Table ${w.table.index + 1}</span>
      <span class="scan-table-stat">${w.rowCount} row${w.rowCount === 1 ? '' : 's'}</span>
    `;
    div.appendChild(head);

    if (!w.schema) {
      const skip = document.createElement('div');
      skip.className = 'scan-table-skipped';
      skip.textContent = w.skipReason || 'Skipped — no SMILES column detected.';
      div.appendChild(skip);
    } else if (w.discrepancies.length === 0) {
      const clean = document.createElement('div');
      clean.className = 'scan-table-clean';
      clean.textContent = `✓ ${w.validatedCount} compound${w.validatedCount === 1 ? '' : 's'} match the table values.`;
      div.appendChild(clean);
    } else {
      for (const d of w.discrepancies) {
        pendingDiscrepancies.push(d);
        const row = document.createElement('div');
        row.className = 'scan-discrepancy';
        row.innerHTML = `
          <div class="discrepancy-row">
            <span class="row-num">Row ${d.rowIndex}</span>
            <span class="discrepancy-prop">${escapeText(d.property.toUpperCase())}</span>
          </div>
          <div class="discrepancy-values">claimed ${d.claimed.toFixed(2)} · computed ${d.expected.toFixed(2)} (Δ ${(Math.abs(d.claimed - d.expected)).toFixed(2)} > ±${d.tolerance})</div>
          <div class="discrepancy-smiles">${escapeText(d.smiles)}</div>
        `;
        div.appendChild(row);
      }
    }

    wrap.appendChild(div);
  }

  if (pendingDiscrepancies.length > 0) {
    const apply = $<HTMLElement>('scan-actions');
    apply.hidden = false;
    const btn = $<HTMLButtonElement>('scan-apply-btn');
    btn.disabled = false;
    btn.textContent = `Insert ${pendingDiscrepancies.length} margin comment${pendingDiscrepancies.length === 1 ? '' : 's'}`;
  }
}

async function applyComments(): Promise<void> {
  const btn = $<HTMLButtonElement>('scan-apply-btn');
  btn.disabled = true;
  btn.textContent = 'Inserting…';
  let inserted = 0;
  for (const d of pendingDiscrepancies) {
    try {
      await insertCellComment(d.tableIndex, d.rowIndex, d.colIndex, d.comment);
      inserted++;
    } catch {
      // Continue on failure — the user gets a partial result rather
      // than a hard halt; remaining discrepancies stay in the panel.
    }
  }
  btn.textContent = `Inserted ${inserted} comment${inserted === 1 ? '' : 's'}`;
  pendingDiscrepancies = [];
}

function showScanError(message: string): void {
  const el = $<HTMLElement>('scan-error');
  el.textContent = message;
  el.hidden = false;
  $<HTMLElement>('scan-progress').hidden = true;
}

// ─── Boot ─────────────────────────────────────────────────────────────

Office.onReady(() => {
  $<HTMLFormElement>('onboarding-form').addEventListener('submit', handleConnect);
  $<HTMLButtonElement>('signout-btn').addEventListener('click', handleSignOut);
  void bootstrap();
});
