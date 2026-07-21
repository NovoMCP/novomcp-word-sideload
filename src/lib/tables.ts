/**
 * Office.js table reading.
 *
 * Word.run requires explicit load() declarations before sync, so this
 * module batches a single round-trip per scan: list all tables, load
 * their values, return as plain string matrices the rest of the
 * pipeline can work with without holding context references.
 *
 * Out of scope for v1 (per Novo_Dist_Play.md §4): tables where
 * structures are embedded as MOL/SDF OLE objects or inline images
 * rather than text. Those need chemical OCR or a binary-blob MOL
 * parser — v2 feature.
 */

import { isLikelySmiles } from './smiles';

export interface ParsedTable {
  /** Stable index in document order — used to find the live table again
   *  when we need to insertComment() into a specific cell. */
  index: number;
  /** Plain-text values, [row][col]. Headers (if any) are row 0. */
  cells: string[][];
  /** True when the table has a header row (Word's tableStyles often
   *  expose this; we infer from layout when the API doesn't). */
  hasHeader: boolean;
}

/**
 * Read every table in the document body as a string matrix. Returns
 * empty array when the document has no tables.
 */
export async function readAllTables(): Promise<ParsedTable[]> {
  return Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load('items');
    await context.sync();

    const parsed: ParsedTable[] = [];

    // Word doesn't expose a single bulk `values` property, so we batch
    // the loads cell-by-cell — still one round-trip thanks to the
    // proxy-object load + single sync at the end.
    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];
      if (!table) continue;
      table.rows.load('items');
    }
    await context.sync();

    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];
      if (!table) continue;
      for (const row of table.rows.items) {
        row.cells.load('items');
      }
    }
    await context.sync();

    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];
      if (!table) continue;
      for (const row of table.rows.items) {
        for (const cell of row.cells.items) {
          // `cell.value` is preferred but not in all Word versions; fall
          // back to body.text. Both are loaded in this same sync.
          cell.body.load('text');
        }
      }
    }
    await context.sync();

    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];
      if (!table) continue;
      const matrix: string[][] = [];
      for (const row of table.rows.items) {
        const rowVals: string[] = [];
        for (const cell of row.cells.items) {
          rowVals.push((cell.body.text || '').trim());
        }
        matrix.push(rowVals);
      }
      // Heuristic: assume the first row is a header when the second row
      // contains at least one cell with a number, and the first row's
      // cells are mostly non-numeric. Conservative — we'd rather miss a
      // header detection than corrupt data alignment.
      const hasHeader = looksLikeHeader(matrix);
      parsed.push({ index: i, cells: matrix, hasHeader });
    }

    return parsed;
  });
}

function looksLikeHeader(matrix: string[][]): boolean {
  if (matrix.length < 2) return false;
  const first = matrix[0];
  const second = matrix[1];
  if (!first || !second) return false;
  // A first row with no valid SMILES sitting above a data row that DOES have
  // one is a label header — even when no column is numeric (SMILES + text-only
  // tables, which the numeric signal below misses).
  if (!first.some(isLikelySmiles) && second.some(isLikelySmiles)) return true;
  const firstNumeric = first.filter(isNumericish).length;
  const secondNumeric = second.filter(isNumericish).length;
  return firstNumeric < second.length / 2 && secondNumeric > 0;
}

function isNumericish(s: string): boolean {
  if (!s) return false;
  // Strip trailing units / spaces then attempt parse
  const cleaned = s.replace(/[a-zA-Z%/]+$/, '').trim();
  return /^-?\d+(\.\d+)?$/.test(cleaned);
}

/**
 * Insert a margin comment into a specific cell. Used by the validator
 * to flag discrepancies between table-claimed and server-computed values.
 */
export async function insertCellComment(
  tableIndex: number,
  rowIndex: number,
  colIndex: number,
  comment: string,
): Promise<void> {
  await Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load('items');
    await context.sync();

    const table = tables.items[tableIndex];
    if (!table) return;
    table.rows.load('items');
    await context.sync();

    const row = table.rows.items[rowIndex];
    if (!row) return;
    row.cells.load('items');
    await context.sync();

    const cell = row.cells.items[colIndex];
    if (!cell) return;

    // Word.Range.insertComment() lands a tracked margin comment that
    // the author can accept/reject. Body.insertComment isn't on every
    // Word API channel, so we route through an explicit Range.
    cell.body.getRange('Whole').insertComment(comment);
    await context.sync();
  });
}
