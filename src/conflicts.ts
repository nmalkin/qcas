/**
 * Check if the given range is a valid selection for conflict resolution
 */
function validRangeForConflicts(
  range: GoogleAppsScript.Spreadsheet.Range | null
): range is GoogleAppsScript.Spreadsheet.Range {
  return range != null && range.getWidth() == 2;
}

function showConflictInstructions(): void {
  const message =
    'To start conflict resolution, please select ' +
    'the two columns that contain the codes to be resolved.';
  showAlert(message);
}

/**
 * Insert conflict columns after the ones in the specified range
 * @param {Range} range
 * @return {Integer} the index of the first newly created column
 */
function insertConflictColumns(
  range: GoogleAppsScript.Spreadsheet.Range
): number {
  return insertColumns(range, 2, ['final', 'status']);
}

interface CodeDiff {
  both: string[];
  onlyA: string[];
  onlyB: string[];
}

/**
 * Return object with commonalities & differences of the two arrays
 *
 * @param flags ignore any differences in flags
 */
function computeDiff(a: string[], b: string[], flags: string[]): CodeDiff {
  const both = [],
    onlyA = [],
    onlyB = [];

  // Check whether each value in A is in B
  for (let i = 0; i < a.length; i++) {
    const el = a[i];
    if (b.includes(el)) {
      both.push(el);
    } else if (flags.includes(el)) {
      both.push(el);
    } else {
      onlyA.push(el);
    }
  }

  // Now check whether each value in B is in A
  for (let i = 0; i < b.length; i++) {
    const el = b[i];
    if (a.includes(el)) {
      // Don't need to add to "both" because first loop already covered it
    } else if (flags.includes(el)) {
      both.push(el);
    } else {
      onlyB.push(el);
    }
  }

  return {
    both: both,
    onlyA: onlyA,
    onlyB: onlyB,
  };
}

/**
 * Given a diff object (@see computeDiff) return string representation of it
 */
function formatDiff(diff: CodeDiff): string {
  let str = '';
  if (diff.both.length > 0) {
    str += diff.both.join(CODES_SEPARATOR);
  }
  if (diff.onlyA.length > 0) {
    str += '\n<' + diff.onlyA.join(CODES_SEPARATOR);
  }
  if (diff.onlyB.length > 0) {
    str += '\n>' + diff.onlyB.join(CODES_SEPARATOR);
  }
  return str;
}

function cellDifferences(leftCell: string, rightCell: string) {
  // Get the codes
  const leftValues = leftCell.split(CODES_SEPARATOR);
  const rightValues = rightCell.split(CODES_SEPARATOR);

  // Find commonalities and differences
  const question = isCodeSheet(SpreadsheetApp.getActiveSheet());
  if (question == null)
    throw new QcasError("current sheet isn't recognized as a coding sheet");
  const flags = getCodebook(question, true);
  const diff = computeDiff(leftValues, rightValues, flags);
  return diff;
}

/**
 * 
 * @param cellA 
 * @param cellB 
 * @returns 
 * @customfunction
 */
function CODES_AGREE(cellA: Cell, cellB: Cell) {
  // Check if any (real) differences remain
  const diff = cellDifferences(cellA.toString(), cellB.toString());
  let status;
  const difference = diff.onlyA.concat(diff.onlyB);
  if (difference.length == 0) {
    status = 'agree';
  } else {
    status = 'conflict';
  }
  return status;
}

/**
 * Look for conflicts between codes in two columns.
 * Write the union of the codes into a new column.
 * If there's disagreement, flag it (in a new column).
 */
function findConflicts() {
  // Check that we can actually compute conflicts for this range.
  const currentSelection =
    SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!validRangeForConflicts(currentSelection)) {
    showConflictInstructions();
    return;
  }

  // Insert new columns (after the selected ones) to hold conflict information.
  const newColumnIndex = insertConflictColumns(currentSelection);

  // Get handles to the columns with the codes to be resolved
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const leftColumn = currentSelection.getColumn();
  const rightColumn = currentSelection.getLastColumn();

  let currentRow = 2;
  // For each code:
  while (currentRow <= currentSheet.getLastRow()) {
    const leftCell = currentSheet.getRange(currentRow, leftColumn);
    const rightCell = currentSheet.getRange(currentRow, rightColumn);

    const leftCellValue: Cell = leftCell.getValue();
    const rightCellValue: Cell = rightCell.getValue();
    // TODO: the profiler says the getValue call is expensive. Replace it with
    // getValues outside the loop.

    const diff = cellDifferences(leftCellValue.toString(), rightCellValue.toString());
    const diffStr = formatDiff(diff);
    const agreementCommand =
      '=CODES_AGREE(' +
      leftCell.getA1Notation() +
      ',' +
      rightCell.getA1Notation() +
      ')';

    // Write the results
    const outputRange = currentSheet.getRange(currentRow, newColumnIndex, 1, 2);
    outputRange.setValues([[diffStr, agreementCommand]]);

    if (CODES_AGREE(leftCellValue, rightCellValue) == 'conflict') {
      outputRange.setBackgrounds([['yellow', 'white']]);
    }

    currentRow++;
  }
}

/**
 * Clear the color in the given cell if it doesn't have conflicts anymore
 */
function updateConflictColors(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const currentColor = e.range.getBackground();
  // Only handle cells that were marked as conflicted (using yellow color)
  if (currentColor == '#ffff00') {
    if (e.value.indexOf('<') === -1 && e.value.indexOf('>') === -1) {
      // No more conflict!
      e.range.setBackground('white');
    }
  }
}
