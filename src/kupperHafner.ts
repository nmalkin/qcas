/**
 * Compute concordance between the pairs of codes in the given range
 * @param cells
 * @param questionId the question id, used to look up the codebook. Codes listed as "flag"s in that codebook will not be included in calculation. If omitted, the codebook will be inferred from the contents of the range (i.e., only codes that appear in it will be used in the counts), but the function will be unable to distinguish between codes and flags.
 * @returns
 */
function CONCORDANCE(
  cells: CellRange,
  questionId?: string
): Array<Array<Cell | null>> {
  const codesAndFlags: CodesAndFlags = questionId
    ? getCodesAndFlags(questionId) // Get the flags from the codebook, so we know which entries to ignore
    : { codes: allUniqueCodesInRange_(cells), flags: [] }; // Infer the codebook by taking all unique codes in the range

  const computeConcordanceForCells = (
    cellA: Cell,
    cellB: Cell
  ): number | null => {
    const codeListA = getCodesInCell_(cellA);
    const codeListB = getCodesInCell_(cellB);

    if (codeListA.length === 0 && codeListB.length === 0) {
      return null;
    }

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
  };

  return cells.map((row: Cell[], i: number) => {
    if (row.length != 2) {
      throw new QcasError(
        `expecting two cells in each input row, but found ${row.length} in row ${i}`
      );
    }

    const [cellA, cellB] = row;

    return [computeConcordanceForCells(cellA, cellB)];
  });
}

function MINCOUNT(
  cells: CellRange,
  questionId?: string
): Array<Array<Cell | null>> {
  const codesAndFlags: CodesAndFlags = questionId
    ? getCodesAndFlags(questionId) // Get the flags from the codebook, so we know which entries to ignore
    : { codes: allUniqueCodesInRange_(cells), flags: [] }; // Infer the codebook by taking all unique codes in the range

  const computeMinCountForCells = (cellA: Cell, cellB: Cell): number | null => {
    const codeListA = getCodesInCell_(cellA);
    const codeListB = getCodesInCell_(cellB);

    let a_i = codeListA.length;
    let b_i = codeListB.length;

    if (a_i === 0 && b_i === 0) {
      return null;
    }

    // Don't count flags
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
  };

  return cells.map((row: Cell[], i: number) => {
    if (row.length != 2) {
      throw new QcasError(
        `expecting two cells in each input row, but found ${row.length} in row ${i}`
      );
    }

    const [cellA, cellB] = row;

    return [computeMinCountForCells(cellA, cellB)];
  });
}

function computeKupperHafner_(inferCodebook: boolean) {
  const currentSelection =
    SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  // Check that the selected range is valid
  if (!validRangeForConflicts(currentSelection)) {
    showConflictInstructions();
    return;
  }

  // Insert new columns (after the selected ones)
  const newColumnIndex = insertColumns(currentSelection, 2, [
    'Concordance',
    'MinCount',
  ]);

  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const questionId = inferCodebook ? null : isCodeSheet(currentSheet);
  if (!inferCodebook && questionId == null) {
    throw new QcasError(
      "couldn't determine question associated with currently opened sheet"
    );
  }

  // Get handles to the columns with the codes
  const leftColumn = currentSelection.getColumn();
  const rightColumn = currentSelection.getLastColumn();
  const lastRow = getLastRowInColumn_(currentSheet, leftColumn);
  const topLeft = currentSheet.getRange(FIRST_ROW, leftColumn).getA1Notation();
  const bottomRight = currentSheet
    .getRange(lastRow, rightColumn)
    .getA1Notation();
  const codeRangeA1 = `${topLeft}:${bottomRight}`;

  // Populate columns with concordance and mincount values
  const concordanceString = inferCodebook
    ? `=CONCORDANCE(${codeRangeA1})`
    : `=CONCORDANCE(${codeRangeA1},"${questionId}")`;
  const minCountString = inferCodebook
    ? `=MINCOUNT(${codeRangeA1})`
    : `=MINCOUNT(${codeRangeA1},"${questionId}")`;

  const newColumnOutputRange = currentSheet.getRange(
    FIRST_ROW,
    newColumnIndex,
    1,
    2
  );
  newColumnOutputRange.setValues([[concordanceString, minCountString]]);

  // Compute summary statistics at the bottom of the new columns
  const concordanceColumn = currentSheet
    .getRange(FIRST_ROW, newColumnIndex, lastRow + 1 - FIRST_ROW, 1)
    .getA1Notation();
  const piHat =
    '=SUM(' + concordanceColumn + ')/COUNT(' + concordanceColumn + ')';

  const codebookSize: string = inferCodebook
    ? `COUNTUNIQUECODES(${codeRangeA1})`
    : // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
      getCodesAndFlags(questionId!).codes.length.toString();
  const minCountColumn = currentSheet
    .getRange(FIRST_ROW, newColumnIndex + 1, lastRow + 1 - FIRST_ROW, 1)
    .getA1Notation();
  const pi_0 =
    '=SUM(' +
    minCountColumn +
    ')/(COUNT(' +
    minCountColumn +
    ')*' +
    codebookSize +
    ')';

  const summaryOutputRange = currentSheet.getRange(
    lastRow + 1,
    newColumnIndex,
    3,
    2
  );
  const piHatRange = summaryOutputRange.getCell(1, 2).getA1Notation();
  const pi0Range = summaryOutputRange.getCell(2, 2).getA1Notation();

  const concordance =
    '=(' + piHatRange + '-' + pi0Range + ')/(1-' + pi0Range + ')';

  summaryOutputRange.setValues([
    ['pi-hat', piHat],
    ['pi0', pi_0],
    ['Kupper-Hafner concordance', concordance],
  ]);
}

function computeKupperHafnerReference() {
  computeKupperHafner_(false);
}

function computeKupperHafnerInfer() {
  computeKupperHafner_(true);
}
