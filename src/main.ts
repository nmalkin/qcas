/**
 * Called when some cell in the spreadsheet has been changed
 */
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const sheet = e.range.getSheet();
  const code = isCodeSheet(sheet);
  if (code) {
    replaceShortcutCodes(code, e);
  }

  if (FINAL_CODES_PATTERN.exec(sheet.getName())) {
    updateConflictColors(e);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Coding Assistant')
    .addItem('Show codebook', 'showCodebook')
    .addItem('Find conflicts', 'findConflicts')
    .addSubMenu(
      ui
        .createMenu('Compute Kupper-Hafner agreement')
        .addItem('Infer codebook', 'computeKupperHafnerInfer')
        .addItem('Referencing codebook', 'computeKupperHafnerReference')
    )
    .addItem("Compute Cohen's kappa", 'computeCohensKappa_')
    .addItem(
      "Compute Cohen's kappa (multi-code)",
      'computeCohensKappaNonExclusive_'
    )
    .addToUi();
}
