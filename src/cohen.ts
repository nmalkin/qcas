function validateCellHasOneCode_(cell: Cell): void {
  if (cell.toString().indexOf(',') != -1) {
    throw new QcasError(`cell has more than one code: ${cell}`);
  }
}

/**
 * Returns agreement for two cells that can only have one code
 * @returns 1 if both cells have the same code, 0 if they don't
 * @throws if either cell has more than one code
 */
function twoCodesAgree_(cellA: Cell, cellB: Cell): number {
  validateCellHasOneCode_(cellA);
  validateCellHasOneCode_(cellB);

  return cellA == cellB ? 1 : 0;
}

/**
 * Returns agreement between the pairs of codes in the given range
 * @returns 1 if both cells have the same code, 0 if they don't
 * @throws if either cell has more than one code
 */
function CODESAGREE2(cells: CellRange): CellRange {
  if (!isRange_(cells)) {
    throw new QcasError(
      `input must be range with two columns of codes, got ${cells}`
    );
  }

  return cells.map((row: Cell[], i: number) => {
    if (row.length != 2) {
      throw new QcasError(
        `expecting two cells in each input row, but found ${row.length} in row ${i}`
      );
    }

    const [cellA, cellB] = row;

    return [twoCodesAgree_(cellA, cellB)];
  });
}

/**
 * Compute estimated probability of chance agreement, using Cohen's method
 * @param cells
 * @returns
 */
function COHEN_PROBABILITY(cells: CellRange): number {
  if (!isRange_(cells)) {
    throw new QcasError(
      `input must be range with two columns of codes, got ${cells}`
    );
  }

  const responseCount = cells.length;
  if (responseCount == 0) {
    throw new QcasError('no cells in range'); // this shouldn't even be possible
  }

  const codeCounts: Record<string, [number, number]> = {};

  cells.forEach((row: Cell[], i: number) => {
    if (row.length != 2) {
      throw new QcasError(
        `expecting two cells in each input row, but found ${row.length} in row ${i}`
      );
    }

    let [codeA, codeB] = row;
    validateCellHasOneCode_(codeA);
    validateCellHasOneCode_(codeB);

    codeA = codeA.toString();
    codeB = codeB.toString();

    if (!(codeA in codeCounts)) {
      codeCounts[codeA] = [0, 0];
    }
    if (!(codeB in codeCounts)) {
      codeCounts[codeB] = [0, 0];
    }

    codeCounts[codeA][0]++;
    codeCounts[codeB][1]++;
  });

  const probabilitySum = Object.keys(codeCounts).reduce((sum, code) => {
    return sum + codeCounts[code][0] * codeCounts[code][1];
  }, 0);

  const probabilityEstimate = probabilitySum / (responseCount * responseCount);
  return probabilityEstimate;
}

function computeCohensKappa_() {
  const currentSelection =
    SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  // Check that the selected range is valid
  if (!validRangeForConflicts(currentSelection)) {
    showConflictInstructions();
    return;
  }

  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const questionId = isCodeSheet(currentSheet);
  if (questionId == null) {
    throw new QcasError(
      "couldn't determine question associated with currently opened sheet"
    );
  }

  // Get handles to the columns with the codes
  const leftColumn = currentSelection.getColumn();
  const rightColumn = currentSelection.getLastColumn();
  const lastRow = currentSheet.getLastRow();
  const topLeft = currentSheet.getRange(FIRST_ROW, leftColumn).getA1Notation();
  const bottomRight = currentSheet
    .getRange(lastRow, rightColumn)
    .getA1Notation();
  const codeRangeA1 = `${topLeft}:${bottomRight}`;

  // Populate columns with concordance and mincount values
  const agreementFormula = `=CODESAGREE2(${codeRangeA1})`;

  // Insert new columns (after the selected ones)
  const newColumnIndex = insertColumns(currentSelection, 2, [
    '', // Leaving one empty to fit the labels in the summary step
    'Agreement',
  ]);
  const newColumnOutputRange = currentSheet.getRange(
    FIRST_ROW,
    newColumnIndex,
    1,
    2
  );
  newColumnOutputRange.setValues([['', agreementFormula]]);

  // Compute summary statistics at the bottom of the new columns
  const agreementColumn = currentSheet
    .getRange(FIRST_ROW, newColumnIndex, lastRow + 1 - FIRST_ROW, 1)
    .getA1Notation();
  const observedAgreement =
    '=SUM(' + agreementColumn + ')/COUNT(' + agreementColumn + ')';

  const chanceAgreementFormula = `=COHEN_PROBABILITY(${codeRangeA1})`;

  const summaryOutputRange = currentSheet.getRange(
    lastRow + 1,
    newColumnIndex,
    3,
    2
  );
  const observedAgreementRange = summaryOutputRange
    .getCell(1, 2)
    .getA1Notation();
  const chanceAgreementRange = summaryOutputRange.getCell(2, 2).getA1Notation();

  const irr =
    '=(' +
    observedAgreementRange +
    '-' +
    chanceAgreementRange +
    ')/(1-' +
    chanceAgreementRange +
    ')';

  summaryOutputRange.setValues([
    ['observed agreement', observedAgreement],
    ['chance agremeent', chanceAgreementFormula],
    ["Cohen's kappa", irr],
  ]);
}
