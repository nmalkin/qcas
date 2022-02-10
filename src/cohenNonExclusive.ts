/**
 * Return count of how many codes each cell has in common
 */
function commonCodeCount_(cellA: Cell, cellB: Cell): number {
  const codesA = getCodesInCell_(cellA);
  const codesB = getCodesInCell_(cellB);

  const commonCount = codesA.reduce(
    (count, code) => count + (codesB.includes(code) ? 1 : 0),
    0
  );
  return commonCount;
}

/**
 * Return how many of the codes in the paired cells agree
 * @returns count of how many codes in each cell agree
 * @throws if either cell has more than one code
 */
function CODESAGREECOUNT(cells: CellRange): CellRange {
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

    return [commonCodeCount_(cellA, cellB)];
  });
}
/**
 * Count the maximum number of codes used by either of the coders
 */
function maxCodeCount_(cellA: Cell, cellB: Cell): number {
  const codesA = getCodesInCell_(cellA);
  const codesB = getCodesInCell_(cellB);

  return Math.max(codesA.length, codesB.length);
}

/**
 * Count the maximum number of codes used by either of the coders
 * @param cells
 * @returns
 */
function MAXCOUNT(cells: CellRange): Array<Array<Cell | null>> {
  return cells.map((row: Cell[], i: number) => {
    if (row.length != 2) {
      throw new QcasError(
        `expecting two cells in each input row, but found ${row.length} in row ${i}`
      );
    }

    const [cellA, cellB] = row;

    return [maxCodeCount_(cellA, cellB)];
  });
}

/**
 * Compute estimated probability of chance agreement, using Cohen's method
 * @param cells
 * @returns
 */
function COHEN_PROBABILITY_MULTIPLE(cells: CellRange): number {
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

    const [cellA, cellB] = row;

    const codesA = getCodesInCell_(cellA);
    const codesB = getCodesInCell_(cellB);

    codesA.forEach((codeA) => {
      if (!(codeA in codeCounts)) {
        codeCounts[codeA] = [0, 0];
      }
      codeCounts[codeA][0]++;
    });

    codesB.forEach((codeB) => {
      if (!(codeB in codeCounts)) {
        codeCounts[codeB] = [0, 0];
      }

      codeCounts[codeB][1]++;
    });
  });

  const probabilitySum = Object.keys(codeCounts).reduce((sum, code) => {
    return sum + codeCounts[code][0] * codeCounts[code][1];
  }, 0);

  const probabilityEstimate = probabilitySum / (responseCount * responseCount);
  return probabilityEstimate;
}

function computeCohensKappaNonExclusive_() {
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
  const lastRow = getLastRowInColumn_(currentSheet, leftColumn);
  const topLeft = currentSheet.getRange(FIRST_ROW, leftColumn).getA1Notation();
  const bottomRight = currentSheet
    .getRange(lastRow, rightColumn)
    .getA1Notation();
  const codeRangeA1 = `${topLeft}:${bottomRight}`;

  const agreementFormula = `=CODESAGREECOUNT(${codeRangeA1})`;
  const countFormula = `=MAXCOUNT(${codeRangeA1})`;

  // Insert new columns (after the selected ones)
  const newColumnIndex = insertColumns(currentSelection, 2, [
    'Agreement',
    'Max count',
  ]);
  const newColumnOutputRange = currentSheet.getRange(
    FIRST_ROW,
    newColumnIndex,
    1,
    2
  );
  newColumnOutputRange.setValues([[agreementFormula, countFormula]]);

  // Compute summary statistics at the bottom of the new columns
  const agreementColumn = currentSheet
    .getRange(FIRST_ROW, newColumnIndex, lastRow + 1 - FIRST_ROW, 1)
    .getA1Notation();
  const countColumn = currentSheet
    .getRange(FIRST_ROW, newColumnIndex + 1, lastRow + 1 - FIRST_ROW, 1)
    .getA1Notation();
  const observedAgreement =
    '=SUM(' + agreementColumn + ')/SUM(' + countColumn + ')';

  const chanceAgreementFormula = `=COHEN_PROBABILITY_MULTIPLE(${codeRangeA1})`;

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
    ["(fake) Cohen's kappa", irr],
  ]);
}
