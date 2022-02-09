/**
 * Called when some cell in the spreadsheet has been changed
 */
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  Logger.log('edit received');

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
  SpreadsheetApp.getUi()
    .createMenu('Coding Assistant')
    .addItem('Show codebook', 'showCodebook')
    .addItem('Find conflicts', 'findConflicts')
    .addItem('Compute Kupper-Hafner agreement', 'computeKupperHafner')
    .addToUi();
}
