/**
 * Compound-table validator.
 *
 * Pure logic, no Office.js / network dependencies — testable in
 * isolation. Given a parsed table, identifies the SMILES column and
 * any property columns, then for each row produces a Discrepancy
 * record when the table's claimed value differs from the server-
 * computed value beyond a per-property tolerance.
 *
 * Spec: Novo_Dist_Play.md §4 ("the real value-add — compound table validation").
 */
import { isLikelySmiles } from './smiles';
import type { ParsedTable } from './tables';

// ─── Header patterns ──────────────────────────────────────────────────

/**
 * Header substrings that identify the SMILES column. Case-insensitive.
 * Per the spec — "SMILES, Canonical SMILES, Structure, Cpd, Compound, Mol".
 * Order matters only for display; column detection takes the first hit.
 */
const SMILES_HEADERS = ['canonical smiles', 'smiles', 'structure', 'compound', 'cpd', 'mol'];

/**
 * Property column patterns. Each entry maps header substrings to a
 * canonical key + tolerance for cross-checking against API values.
 * Tolerances reflect the real-world spread between different
 * computation methods (cLogP vs xLogP, RDKit MW rounding, etc.).
 */
const PROPERTY_HEADERS: Array<{
  key: 'mw' | 'logp' | 'qed' | 'tpsa';
  patterns: string[];
  tolerance: number;
  unitSuffix?: string;
}> = [
  { key: 'mw',   patterns: ['mol weight', 'molecular weight', 'mw', 'mol wt', 'g/mol'], tolerance: 1.0 },
  { key: 'logp', patterns: ['logp', 'log p', 'xlogp', 'clogp', 'lipophil'],            tolerance: 0.5 },
  { key: 'qed',  patterns: ['qed', 'drug-like', 'druglike'],                            tolerance: 0.05 },
  { key: 'tpsa', patterns: ['tpsa', 'polar surface'],                                   tolerance: 2.0 },
];

// ─── Schema detection ─────────────────────────────────────────────────

export interface TableSchema {
  smilesCol: number;
  propertyCols: Partial<Record<'mw' | 'logp' | 'qed' | 'tpsa', number>>;
  /** Row index where data starts (0 if no header, 1 if header row was found). */
  dataStart: number;
}

export function detectSchema(table: ParsedTable): TableSchema | null {
  const headerRow = table.hasHeader ? table.cells[0] : null;
  if (!headerRow) {
    // Headerless tables: scan rows column-by-column for a column where
    // ≥3 cells are likely SMILES. Not common in pubs but supported.
    return detectByRowSampling(table);
  }

  const lower = headerRow.map((h) => h.toLowerCase());

  // SMILES column — try header match first
  let smilesCol = -1;
  for (const pat of SMILES_HEADERS) {
    const idx = lower.findIndex((h) => h.includes(pat));
    if (idx >= 0) { smilesCol = idx; break; }
  }
  // If header doesn't match, fall back to row-sampling (e.g. column
  // labeled "Cmpd #" with SMILES inside).
  if (smilesCol < 0) {
    const sampled = detectByRowSampling(table);
    return sampled;
  }

  // Property columns — first header pattern hit wins
  const propertyCols: TableSchema['propertyCols'] = {};
  for (const { key, patterns } of PROPERTY_HEADERS) {
    for (const pat of patterns) {
      const idx = lower.findIndex((h) => h.includes(pat));
      if (idx >= 0) { propertyCols[key] = idx; break; }
    }
  }

  return { smilesCol, propertyCols, dataStart: 1 };
}

function detectByRowSampling(table: ParsedTable): TableSchema | null {
  if (table.cells.length === 0) return null;
  const colCount = table.cells[0]?.length ?? 0;
  if (colCount === 0) return null;

  // For each column, count rows where the cell parses as a likely SMILES.
  let bestCol = -1;
  let bestHits = 0;
  for (let c = 0; c < colCount; c++) {
    let hits = 0;
    const sample = Math.min(table.cells.length, 12);
    for (let r = 0; r < sample; r++) {
      const v = (table.cells[r]?.[c] ?? '').trim();
      if (v && isLikelySmiles(v)) hits++;
    }
    if (hits > bestHits) { bestHits = hits; bestCol = c; }
  }

  // Need at least 3 valid SMILES hits per the spec ("a column where ≥3
  // rows parse as valid SMILES").
  if (bestHits < 3) return null;
  // Row 0 is a header when its cell in the SMILES column is not itself a valid
  // SMILES (e.g. the literal label "SMILES"/"Structure"). Derive dataStart from
  // that directly — the numeric header heuristic in tables.ts can't tell for
  // SMILES + text-only tables (no numeric column), which would otherwise count
  // the header row as a compound.
  const row0Smiles = (table.cells[0]?.[bestCol] ?? '').trim();
  const dataStart = isLikelySmiles(row0Smiles) ? 0 : 1;
  return { smilesCol: bestCol, propertyCols: {}, dataStart };
}

// ─── Discrepancy detection ────────────────────────────────────────────

export interface ServerValues {
  smiles: string;
  mw?: number;
  logp?: number;
  qed?: number;
  tpsa?: number;
}

export interface Discrepancy {
  tableIndex: number;
  rowIndex: number;
  smiles: string;
  /** Which property column flagged. */
  property: 'mw' | 'logp' | 'qed' | 'tpsa';
  /** Column index for the flagged property — used to locate the cell
   *  for insertCellComment(). */
  colIndex: number;
  claimed: number;
  expected: number;
  tolerance: number;
  /** Pre-formatted comment string ready for insertCellComment. */
  comment: string;
}

export function findDiscrepancies(
  table: ParsedTable,
  schema: TableSchema,
  serverByRow: Map<number, ServerValues>,
): Discrepancy[] {
  const out: Discrepancy[] = [];

  for (let r = schema.dataStart; r < table.cells.length; r++) {
    const row = table.cells[r];
    if (!row) continue;
    const server = serverByRow.get(r);
    if (!server) continue;

    for (const { key, tolerance } of PROPERTY_HEADERS) {
      const colIndex = schema.propertyCols[key];
      if (colIndex == null) continue;
      const cellText = (row[colIndex] ?? '').trim();
      if (!cellText) continue;

      const claimed = parseNumeric(cellText);
      if (claimed == null) continue;
      // Type-narrowed lookup — `key` is a literal union over the
      // numeric fields of ServerValues, but TS can't see through a
      // generic Record cast because ServerValues also has smiles:string.
      let expected: number | undefined;
      switch (key) {
        case 'mw': expected = server.mw; break;
        case 'logp': expected = server.logp; break;
        case 'qed': expected = server.qed; break;
        case 'tpsa': expected = server.tpsa; break;
      }
      if (typeof expected !== 'number') continue;

      const delta = Math.abs(claimed - expected);
      if (delta <= tolerance) continue;

      out.push({
        tableIndex: table.index,
        rowIndex: r,
        smiles: server.smiles,
        property: key,
        colIndex,
        claimed,
        expected,
        tolerance,
        comment: formatComment(key, claimed, expected, tolerance),
      });
    }
  }

  return out;
}

function parseNumeric(s: string): number | null {
  // Accept "180.16", "180.16 g/mol", "1.31 (calc)", "≈ 1.3". Strip trailing
  // units and parenthetical annotations; match the leading number.
  const m = s.match(/-?\d+(\.\d+)?/);
  if (!m) return null;
  const n = Number(m[0]);
  return isFinite(n) ? n : null;
}

function formatComment(
  key: 'mw' | 'logp' | 'qed' | 'tpsa',
  claimed: number,
  expected: number,
  tolerance: number,
): string {
  const labels: Record<typeof key, string> = {
    mw: 'MW', logp: 'LogP', qed: 'QED', tpsa: 'TPSA',
  };
  const claimedStr = key === 'qed' ? claimed.toFixed(2) : claimed.toFixed(2);
  const expectedStr = key === 'qed' ? expected.toFixed(2) : expected.toFixed(2);
  return `NovoMCP: ${labels[key]} discrepancy. Table claims ${claimedStr}; computed ${expectedStr} (tolerance ±${tolerance}).`;
}

// ─── Server-value extraction ──────────────────────────────────────────

/**
 * Pull the relevant fields out of a get_molecule_profile response into
 * the typed shape the validator expects. Probes both naming conventions
 * (hbd_count vs hbd, etc.) the same way the renderer does.
 */
export function extractServerValues(smiles: string, response: { result: { properties?: Record<string, unknown> } }): ServerValues {
  const props = response.result.properties || {};
  const num = (...keys: string[]): number | undefined => {
    for (const k of keys) {
      const v = props[k];
      if (typeof v === 'number' && isFinite(v)) return v;
    }
    return undefined;
  };
  return {
    smiles,
    mw: num('molecular_weight'),
    logp: num('logp', 'xlogp'),
    qed: num('qed'),
    tpsa: num('tpsa'),
  };
}
