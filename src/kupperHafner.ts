function CONCORDANCE(
  cellA: Cell,
  cellB: Cell,
  questionId: string
): number | null {
  const codeListA = getCodesInCell_(cellA);
  const codeListB = getCodesInCell_(cellB);

  if (codeListA.length === 0 && codeListB.length === 0) {
    return null;
  }

  // Get the flags from the codebook, so we know which entries to ignore
  const codesAndFlags = getCodesAndFlags(questionId);

  // Let the variables a_i, and b_i, denote the numbers of attributes for the i-th unit chosen by raters A and B, respectively
  let a_i = 0;
  let b_i = 0;
  for (let i = 0; i < codeListB.length; i++) {
    const el = codeListB[i];
    if (codesAndFlags.codes.includes(el)) {
      b_i++;
    }
  }

  // Let the random variable Xi denote the number of elements common to the sets A_i and B_i
  let x_i = 0;
  for (let i = 0; i < codeListA.length; i++) {
    const el = codeListA[i];
    if (codesAndFlags.codes.includes(el)) {
      a_i++;

      if (codeListB.includes(el)) {
        x_i++;
      }
    } else if (!codesAndFlags.flags.includes(el)) {
      throw 'Not recognized as either code or flag: ' + el;
    }
  }

  if (a_i === 0 && b_i === 0) {
    return null;
  }

  // The observed proportion of concordance is pi_hat_i = x_i / max(a_i, b_i)
  const pi_hat_i = x_i / Math.max(a_i, b_i);
  return pi_hat_i;
}

function MINCOUNT(cellA: Cell, cellB: Cell, questionId: string): number | null {
  const codeListA = getCodesInCell_(cellA);
  const codeListB = getCodesInCell_(cellB);

  let a_i = codeListA.length;
  let b_i = codeListB.length;

  if (a_i === 0 && b_i === 0) {
    return null;
  }

  // Don't count flags
  const codesAndFlags = getCodesAndFlags(questionId);
  for (let i = 0; i < codeListA.length; i++) {
    const el = codeListA[i];
    if (codesAndFlags.flags.includes(el)) {
      a_i--;
    }
  }
  for (let i = 0; i < codeListB.length; i++) {
    const el = codeListB[i];
    if (codesAndFlags.flags.includes(el)) {
      b_i--;
    }
  }

  const minCount = Math.min(a_i, b_i);
  return minCount;
}

function computeKupperHafner() {
  // Check that the selected range is valid
  const currentSelection =
    SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  // TODO: may need separate validRange function - or just a new name
  if (!validRangeForConflicts(currentSelection)) {
    showConflictInstructions();
    return;
  }

  // Insert new columns (after the selected ones)
  const newColumnIndex = insertColumns(currentSelection, 2, [
    'Concordance',
    'MinCount',
  ]);

  // Get handles to the columns with the codes
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const questionId = isCodeSheet(currentSheet);
  const leftColumn = currentSelection.getColumn();
  const rightColumn = currentSelection.getLastColumn();

  let currentRow = FIRST_ROW;
  // For each code:
  while (currentRow <= currentSheet.getLastRow()) {
    const leftCell = currentSheet
      .getRange(currentRow, leftColumn)
      .getA1Notation();
    const rightCell = currentSheet
      .getRange(currentRow, rightColumn)
      .getA1Notation();

    const concordanceString =
      '=CONCORDANCE(' + leftCell + ',' + rightCell + ',"' + questionId + '")';
    const minCountString =
      '=MINCOUNT(' + leftCell + ',' + rightCell + ',"' + questionId + '")';

    const outputRange = currentSheet.getRange(currentRow, newColumnIndex, 1, 2);
    outputRange.setValues([[concordanceString, minCountString]]);

    currentRow++;
  }

  const concordanceColumn = currentSheet
    .getRange(FIRST_ROW, newColumnIndex, currentRow - FIRST_ROW, 1)
    .getA1Notation();
  const piHat =
    '=SUM(' + concordanceColumn + ')/COUNT(' + concordanceColumn + ')';

  const codebook = getCodesAndFlags(questionId).codes;
  const minCountColumn = currentSheet
    .getRange(FIRST_ROW, newColumnIndex + 1, currentRow - FIRST_ROW, 1)
    .getA1Notation();
  const pi_0 =
    '=SUM(' +
    minCountColumn +
    ')/(COUNT(' +
    minCountColumn +
    ')*' +
    codebook.length +
    ')';

  const outputRange = currentSheet.getRange(currentRow, newColumnIndex, 3, 2);
  const piHatRange = outputRange.getCell(1, 2).getA1Notation();
  const pi0Range = outputRange.getCell(2, 2).getA1Notation();

  const concordance =
    '=(' + piHatRange + '-' + pi0Range + ')/(1-' + pi0Range + ')';

  outputRange.setValues([
    ['pi-hat', piHat],
    ['pi0', pi_0],
    ['Kupper-Hafner concordance', concordance],
  ]);
}
