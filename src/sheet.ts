/**
 * Determine whether the sheet is used for coding
 * and so whether we should perform substitution on it
 *
 * @param {*Sheet} sheet
 * @return the name of the question being coded, or null if there is none
 */
function isCodeSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const sheetName = sheet.getName();
  const match = CODING_PATTERN.exec(sheetName);
  if (match) {
    return match[1];
  } else {
    return null;
  }
}

/**
 * Return specified sheet
 */
function getSheet(sheetName: string) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet == null) {
    showAlert("Couldn't find a sheet with the name " + sheetName);
  }
  return sheet;
}

/**
 * Given a sheet, return its first (header) row as a Range
 */
function getSheetHeader(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn());
}

/**
 * Return the given array with trailing empty values removed
 */
function truncateEmptyArray(arr: string[]): string[] {
  let next = arr.pop();
  while (next == '') {
    next = arr.pop();
  }

  // Put the last (non-empty) value back
  if (next) {
    arr.push(next);
  }

  return arr;
}

/**
 * Given a Range, return a flattened array with all its values.
 */
function getAllValues(range: GoogleAppsScript.Spreadsheet.Range) {
  const data = [];
  const values = range.getValues();
  for (let i = 0; i < range.getHeight(); i++) {
    for (let j = 0; j < range.getWidth(); j++) {
      const value = values[i][j];
      data.push(value);
    }
  }

  return truncateEmptyArray(data);
}

/**
 * Given a sheet and a column name, return that column's number
 */
function getColumnNumberByName(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  name: string
) {
  const header = getAllValues(getSheetHeader(sheet));
  for (let i = 0; i < header.length; i++) {
    if (header[i] === name) {
      return i + 1; // add 1 because columns are 1-indexed :-/
    }
  }

  return -1;
}

/**
 * Insert the given number of columns after the ones in the specified range
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {Integer} howMany how many columns to insert
 * @param {Array} names the names to put in the first row of the new columns
 * @return {Integer} the index of the first newly created column
 */
function insertColumns(
  range: GoogleAppsScript.Spreadsheet.Range,
  howMany: number,
  names: string[]
): number {
  // Figure out where to put the columns
  const position = range.getLastColumn();

  // Insert the columns
  const sheet = range.getSheet();
  sheet.insertColumnsAfter(position, howMany);

  // Give them appropriate titles
  const newColumnIndex = position + 1;
  const header = sheet.getRange(1, newColumnIndex, 1, names.length);
  header.setValues([names]);

  return newColumnIndex;
}

function getLastRowInColumn_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  column: number
): number {
  let last = sheet.getLastRow();
  const rows = sheet
    .getRange(FIRST_ROW, column, last - FIRST_ROW + 1)
    .getValues();
  let next = rows.pop();
  while (next && !next[0]) {
    last--;
    next = rows.pop();
  }
  return last;
}
