function countCodeInOneCell_(codeToCount: Cell, cellContents: Cell): number {
  const codes = getCodesInCell_(cellContents);
  return codes.reduce(
    (count, code) => (code == codeToCount ? count + 1 : count),
    0
  );
}

function countCodeInRange_(codeToCount: Cell, cells: CellRange): number {
  return cells.reduce(
    (count: number, currentRow) =>
      count +
      currentRow.reduce(
        (count: number, currentCell: Cell) =>
          count + countCodeInOneCell_(codeToCount, currentCell),
        0
      ),
    0
  );
}

function countCodeInCellOrRange_(
  codeToCount: Cell,
  cells: CellOrRange
): number {
  return isRange_(cells)
    ? countCodeInRange_(codeToCount, cells)
    : countCodeInOneCell_(codeToCount, cells);
}

function COUNTCODE(codeToCount: Cell, rangeToCount: CellOrRange): number;
function COUNTCODE(
  codeToCount: CellRange,
  rangeToCount: CellOrRange
): CellRange;
/**
 * Count how many times the code(s) in the first argument appears in the cell(s) from the second argument
 *
 * @customfunction
 * @param codeToCount cell or range with code(s) to count
 * @param rangeToCount cell(s) to search
 */
function COUNTCODE(
  codeToCount: CellOrRange,
  rangeToCount: CellOrRange
): CellOrRange {
  if (isRange_(codeToCount)) {
    return codeToCount.map((row: Cell[]) =>
      row.map((cell: Cell) => countCodeInCellOrRange_(cell, rangeToCount))
    );
  } else {
    return countCodeInCellOrRange_(codeToCount, rangeToCount);
  }
}

/**
 * Create a range with all the final codes, followed by the number of times that code appeared in the specified range
 * @customfunction
 * @param {Array<Array<string>>} input
 */
function COUNTCODEBOOK(input: CellOrRange): CellRange {
  const codes = getFinalCodeList_(getCurrentQuestionCodeOrError_());
  const counter: Record<string, number> = {};
  codes.forEach((code) => {
    counter[code] = COUNTCODE(code, input);
  });
  const output = Object.keys(counter).map((code) => [code, counter[code]]);

  return output;
}

function allUniqueCodesInRange_(cells: CellRange): string[] {
  const allCodes = new Set<string>();
  cells.forEach((row) =>
    row.forEach((cell) =>
      getCodesInCell_(cell).forEach((code) => allCodes.add(code))
    )
  );
  return Array.from(allCodes);
}

/**
 * Return a cell for each unique code in the given range
 * @customfunction
 * @param cells
 * @returns
 */
function LISTUNIQUECODES(cells: CellRange): string[] {
  return allUniqueCodesInRange_(cells);
}

/**
 * Count the number of unique codes appearing in these cells
 * @param cells
 * @returns
 * @customfunction
 */
function COUNTUNIQUECODES(cells: CellRange): number {
  return allUniqueCodesInRange_(cells).length;
}
