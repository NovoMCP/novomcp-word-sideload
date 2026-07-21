/**
 * Office.js helpers — read selection / paragraphs / tables from Word.
 *
 * The taskpane runs in an iframe inside Word; all document reads go
 * through Word.run() which manages the request batch / sync cycle.
 * Selection-change events come from the older Office.context.document
 * handler API (the Word.run "context" object is request-scoped, not
 * event-scoped).
 */

export interface SelectionSnapshot {
  /** Raw text contained in the user's current selection. Empty when the
   *  selection collapses to a caret position. */
  text: string;
  /** True when the selection has at least one character. */
  hasText: boolean;
}

/** Read the current selection's text via Word.run. Empty string on no selection. */
export async function readSelection(): Promise<SelectionSnapshot> {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load('text');
    await context.sync();
    const text = range.text || '';
    return { text, hasText: text.trim().length > 0 };
  });
}

/** Read the entire body text. Used by the W4 document scanner; here it's
 *  a fallback when the user invokes the "Scan Document" ribbon button
 *  while the taskpane is open. */
export async function readBodyText(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load('text');
    await context.sync();
    return body.text || '';
  });
}

export interface InsertableTable {
  /** First row, rendered as the table header. */
  headers: string[];
  /** Body rows. Cell count must equal headers.length per row. */
  rows: string[][];
  /** Optional caption inserted as a paragraph above the table. */
  caption?: string;
}

/**
 * Insert a formatted table at the user's current selection. The table
 * appears AFTER the selection (cursor position) so existing content
 * isn't overwritten. Values are stringified by the caller — this
 * helper doesn't format numbers, dates, or anything else.
 *
 * Word applies the document's default table style. Users restyle
 * post-insertion if they want a different look.
 */
export async function insertTableAtSelection(table: InsertableTable): Promise<void> {
  const rowCount = table.rows.length + 1; // +1 for header
  const colCount = table.headers.length;
  if (rowCount === 1 || colCount === 0) return;
  // Validate row widths — silently truncate or pad rather than throw,
  // since the caller's data may have inconsistent column counts.
  const values: string[][] = [table.headers.slice(0, colCount)];
  for (const row of table.rows) {
    const padded = Array.from({ length: colCount }, (_, i) => row[i] ?? '');
    values.push(padded);
  }

  await Word.run(async (context) => {
    const range = context.document.getSelection();
    if (table.caption) {
      // Insert the caption as a paragraph BEFORE the table; cursor
      // ends up between the two.
      range.insertParagraph(table.caption, 'After');
    }
    range.insertTable(rowCount, colCount, 'After', values);
    await context.sync();
  });
}

/**
 * Subscribe to selection-change events. Returns an unsubscribe function.
 * The event fires whenever the user moves their cursor or extends/shrinks
 * the selection — which is high-frequency, so debouncing in the handler
 * is the caller's job.
 */
export function onSelectionChange(handler: () => void): () => void {
  const wrappedHandler = (): void => { handler(); };
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    wrappedHandler,
  );
  return () => {
    Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      { handler: wrappedHandler },
    );
  };
}
