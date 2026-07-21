/**
 * Ribbon command handlers — invoked by Word for ExecuteFunction
 * controls in the manifest. The "Scan Document" button lands here in
 * W4 once the compound-table validator is wired up; for now it's a
 * stub that opens the taskpane so the manifest validates and the
 * ribbon control is functional.
 *
 * Functions registered here MUST be associated via Office.actions.associate
 * and the FunctionName MUST match the manifest <FunctionName> exactly.
 */

Office.onReady(() => {
  Office.actions.associate('scanDocument', scanDocument);
});

async function scanDocument(event: Office.AddinCommands.Event): Promise<void> {
  // W4 will replace this with the actual compound-table walker:
  //   - context.document.body.tables — find tables with SMILES columns
  //   - concurrent get_molecule_profile against rows
  //   - cross-check MW / LogP / QED against API values
  //   - Word.range.insertComment() for discrepancies
  //
  // For W1, just open the taskpane so the user has a path to manual
  // single-compound analysis. event.completed() is required — Word will
  // hang the ribbon if a command handler doesn't call it.
  try {
    // No public API to programmatically open the taskpane from a command
    // function in Word desktop yet (pinned for Outlook only). Best we can
    // do is signal completion; the user opens the taskpane themselves
    // via the "Open NovoMCP" ribbon button.
  } finally {
    event.completed();
  }
}
