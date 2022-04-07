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
        .createMenu('Compute inter-rater reliability')
        .addSubMenu(
          ui
            .createMenu('One code per response')
            .addItem("Compute Cohen's kappa", 'computeCohensKappa_')
            .addItem("Compute Krippendorff's alpha", 'computeKrippendorff_')
        )
        .addSubMenu(
          ui
            .createMenu('Multiple codes per response')
            .addItem(
              'Compute Kupper-Hafner agreement',
              'computeKupperHafnerInfer'
            )
            .addItem(
              "Compute Cohen's kappa (multi-code)",
              'computeCohensKappaNonExclusive_'
            )
        )
    )
    .addItem('About', 'showAbout_')
    .addToUi();
}
