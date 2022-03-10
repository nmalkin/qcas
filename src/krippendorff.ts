/*
Functions for computing Krippendorff's Alpha
Based on author's explanations at:
https://repository.upenn.edu/cgi/viewcontent.cgi?article=1043&context=asc_papers
*/

function allUniqueCellsInRange_(cells: CellRange): Cell[] {
  const allCodes = new Set<Cell>();
  cells.forEach((row) =>
    row.forEach((cell) => {
      validateCellHasOneCode_(cell);
      if (notEmpty(cell) && cell != '') {
        allCodes.add(cell);
      }
    })
  );
  return Array.from(allCodes);
}

/**
 * Return both permutations of every 2-combination in the array
 * @param arr
 * @returns every pair of two possible values, twice
 */
function allPairs_<T>(arr: T[]): [T, T][] {
  const pairs: [T, T][] = [];
  for (let i = 0; i < arr.length; i++) {
    for (let j = 0; j < arr.length; j++) {
      if (i != j) {
        pairs.push([arr[i], arr[j]]);
      }
    }
  }
  return pairs;
}

/**
 * Generate the coincidence matrix for the codes in the given range.
 * This is a matrix that shows how often each pair of codes co-occurs.
 * It is used in (one method of) calculating Krippendorff's alpha.
 *
 * @customfunction
 * @param range
 * @returns
 */
function COINCIDENCE_MATRIX(range: CellRange): CellRange {
  const codeList: (number | string)[] = validateCells_(
    allUniqueCellsInRange_(range)
  ).filter(notEmpty);
  const counter: Record<string, Record<string, number>> = {};
  codeList.forEach((iCode) => {
    const innerCounter: Record<string, number> = {};
    codeList.forEach((jCode) => {
      innerCounter[jCode] = 0;
    });

    counter[iCode] = innerCounter;
  });

  range.forEach((row) => {
    const rowWithoutMissingData = validateCells_(row.filter((el) => el != ''));
    const pairs = allPairs_(rowWithoutMissingData);
    pairs.forEach((pair) => {
      const [a, b] = pair;
      counter[a][b] += 1 / (rowWithoutMissingData.length - 1);
    });
  });

  const output: CellRange = codeList.map((iCode) => {
    const outputRow: Cell[] = codeList.map((jCode) => {
      return counter[iCode][jCode];
    });
    outputRow.unshift(iCode);
    return outputRow;
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const columnLabels = ([''] as any).concat(codeList);
  output.unshift(columnLabels);

  return output;
}

/**
 * For each number in the input range, multiply it by each subsequent value (one at a time), then return the sum of those products.
 * @customfunction
 * @param range
 * @returns
 */
function PRODUCTSUM(range: CellRange): number {
  const values: number[] = range
    .flatMap((a) => a)
    .map((v) => cellAsNumberOrError_(v));

  let sum = 0;
  for (let i = 0; i < values.length; i++) {
    for (let j = i + 1; j < values.length; j++) {
      const product = values[i] * values[j];
      sum += product;
    }
  }
  return sum;
}

function computeKrippendorff_() {
  const currentSelection =
    SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  // Check that the selected range is valid
  if (currentSelection == null || currentSelection.getWidth() <= 1) {
    showGenericInstructions();
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get handles to the columns with the codes
  const leftColumn = currentSelection.getColumn();
  const rightColumn = currentSelection.getLastColumn();
  const numInputColumns = rightColumn - leftColumn + 1;
  const lastRow = Math.max(
    ...range(leftColumn, rightColumn).map((column) =>
      getLastRowInColumn_(sheet, column)
    )
  );
  const numInputRows = lastRow - FIRST_ROW + 1;
  const topLeft = sheet.getRange(FIRST_ROW, leftColumn).getA1Notation();
  const bottomRight = sheet.getRange(lastRow, rightColumn).getA1Notation();
  const codeRangeA1 = `${topLeft}:${bottomRight}`;

  // Insert new columns for the values-by-units matrix
  const codeRange = sheet.getRange(codeRangeA1);
  const codeCount = COUNTUNIQUECODES(codeRange.getValues());
  const newColumnIndex = insertColumns(currentSelection, codeCount + 3, [
    '',
    'Ratings / item - 1',
    `=TRANSPOSE(LISTUNIQUECODES(${codeRangeA1}))`,
  ]);
  const lastColumnAddedIndex = newColumnIndex + codeCount + 2;

  // Populate code counts for the first row
  const codeCountRange = sheet
    .getRange(FIRST_ROW, newColumnIndex + 2, 1, codeCount)
    .getA1Notation();
  const codeNameRange = sheet
    .getRange(HEADER_ROW, newColumnIndex + 2, 1, codeCount)
    .getA1Notation()
    .replace(
      new RegExp(HEADER_ROW.toString(), 'g'),
      '$' + HEADER_ROW.toString()
    ); // fix the row
  const ratingsRange = sheet
    .getRange(FIRST_ROW, leftColumn, 1, numInputColumns)
    .getA1Notation();

  const sumFormula = `=SUM(${codeCountRange}) - 1`;
  const countFormula = `=COUNTCODE(${codeNameRange},${ratingsRange})`;
  const productSumFormula = `=PRODUCTSUM(${codeCountRange})`;

  sheet
    .getRange(FIRST_ROW, newColumnIndex + 1, 1, 2)
    .setValues([[sumFormula, countFormula]]);
  sheet
    .getRange(FIRST_ROW, lastColumnAddedIndex, 1, 1)
    .setValues([[productSumFormula]]);

  // Autofill remaining rows based on the first one
  const firstRowRange = sheet.getRange(
    FIRST_ROW,
    newColumnIndex + 1,
    1,
    codeCount + 2
  );
  const allRowsRange = sheet.getRange(
    FIRST_ROW,
    newColumnIndex + 1,
    numInputRows + 1,
    codeCount + 2
  );
  firstRowRange.autoFill(
    allRowsRange,
    SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
  );

  // Compute summary statistics at the bottom of the new columns
  const sumRange = sheet
    .getRange(FIRST_ROW, newColumnIndex + 1, numInputRows, 1)
    .getA1Notation();
  const sumCell = sheet
    .getRange(lastRow + 1, newColumnIndex + 1, 1, 1)
    .getA1Notation();
  const productSumRange = sheet
    .getRange(FIRST_ROW, lastColumnAddedIndex, numInputRows, 1)
    .getA1Notation();
  const productSumCell = sheet
    .getRange(lastRow + 1, lastColumnAddedIndex, 1, 1)
    .getA1Notation();
  const observedDisagreementCell = sheet
    .getRange(lastRow + 2, newColumnIndex + 1, 1, 1)
    .getA1Notation();
  const expectedDisagreementCell = sheet
    .getRange(lastRow + 3, newColumnIndex + 1, 1, 1)
    .getA1Notation();

  const observedDisagreementFormula = `=SUMPRODUCT(1/MAXEACH(1,${sumRange}),${productSumRange})`;
  const expectedDisagreementFormula = `=${productSumCell}/${sumCell}`;
  const alphaFormula = `=1-${observedDisagreementCell}/${expectedDisagreementCell}`;

  sheet.getRange(lastRow + 2, newColumnIndex, 3, 2).setValues([
    ['observed disagreement', observedDisagreementFormula],
    ['expected disagreement', expectedDisagreementFormula],
    ["Krippendorff's alpha", alphaFormula],
  ]);

  // Add row with pairable counts
  const sumRangeFixed = sumRange.replace(new RegExp('([A-Z])', 'g'), '$$$1');
  const firstCountRange = sheet
    .getRange(FIRST_ROW, newColumnIndex + 2, numInputRows, 1)
    .getA1Notation();
  const pairableCountsFormula = `=SUMIF(${sumRangeFixed},">0",${firstCountRange})`;

  const firstSumIfCell = sheet.getRange(lastRow + 1, newColumnIndex + 2, 1, 1);
  const sumIfRange = sheet.getRange(
    lastRow + 1,
    newColumnIndex + 2,
    1,
    codeCount
  );

  firstSumIfCell.setValues([[pairableCountsFormula]]);
  firstSumIfCell.autoFill(
    sumIfRange,
    SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
  );
}
