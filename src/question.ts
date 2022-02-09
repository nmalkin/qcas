/**
 * Get the question ID for the currently selected question
 */
function getCurrentQuestionCodeOrError_(): string {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Is the current sheet a coding sheet?
  const question = isCodeSheet(sheet);
  if (question !== null) {
    return question;
  }

  // Else, is the current sheet a codebook sheet?
  const match = CODEBOOK_PATTERN.exec(sheet.getName());
  if (match) {
    return match[1];
  }

  // Otherwise, I really have no idea what sheet this is.
  throw new QcasError(
    `can't figure out which question sheet ${sheet.getName()} is associated with`
  );
}
