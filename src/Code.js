/*
Copyright (C) 2019 N. Malkin

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

function alert(message) {
  let ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

/**
 * Return specified sheet
 */
function getSheet(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet == null) {
    alert("Couldn't find a sheet with the name " + sheetName);
  }
  return sheet;
}

/**
 * Given a sheet, return its first (header) row as a Range
 */
function getSheetHeader(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn());
}

/**
 * Return the given array with trailing empty values removed
 */
function truncateEmptyArray(arr) {
  let next = arr.pop();
  while (next == '') {
    next = arr.pop();
  }

  // Put the last (non-empty) value back
  arr.push(next);

  return arr;
}

/**
 * Given a Range, return a flattened array with all its values.
 */
function getAllValues(range) {
  let data = [];
  let values = range.getValues();
  for (let i = 0; i < range.getHeight(); i++) {
    for (let j = 0; j < range.getWidth(); j++) {
      let value = values[i][j];
      data.push(value);
    }
  }

  return truncateEmptyArray(data);
}

/**
 * Given a sheet and a column name, return that column's number
 */
function getColumnNumberByName(sheet, name) {
  let header = getAllValues(getSheetHeader(sheet));
  for (let i = 0; i < header.length; i++) {
    if (header[i] === name) {
      return i + 1; // add 1 because columns are 1-indexed :-/
    }
  }

  return -1;
}

/**
 * Return an object with all codes and flags in the codebook
 *
 * @param question the name of the question, used in the sheet title
 */
function getCodesAndFlags(question) {
  let codebookSheetName = question + '_codebook';
  let sheet = getSheet(codebookSheetName);

  // Find the range where the relevant codebook columns are located
  let codeColumn = getColumnNumberByName(sheet, CODEBOOK_HEADER_CODE);
  let typeColumn = getColumnNumberByName(sheet, CODEBOOK_HEADER_TYPE);
  let firstColumn = Math.min(codeColumn, typeColumn);
  let lastColumn = Math.max(codeColumn, typeColumn);
  let range = sheet.getRange(
    FIRST_ROW,
    firstColumn,
    sheet.getLastRow() - 1,
    lastColumn - firstColumn + 1
  );

  let values = range.getValues();
  let codes = [],
    flags = [];
  for (let i = 0; i < range.getHeight(); i++) {
    let code = values[i][codeColumn - firstColumn];
    if (code === '') {
      // Tolerate holes in codebook
      continue;
    }

    let type = values[i][typeColumn - firstColumn];
    if (type === '') {
      // If no type is specified, assume it's a code
      type = CODEBOOK_TYPE_CODE;
    }

    if (type === CODEBOOK_TYPE_CODE) {
      codes.push(code);
    } else if (type === CODEBOOK_TYPE_FLAG) {
      flags.push(code);
    } else {
      alert('Unrecognized code type ' + type + ' in codebook ' + question);
      break;
    }
  }

  return {
    codes: codes,
    flags: flags,
  };
}

/**
 * Return an array of all codes in the codebook
 *
 * @param question the name of the question, used in the sheet title
 * @param {boolean} flagsOnly if true, only return the flags
 */
function getCodebook(question, flagsOnly) {
  let codesAndFlags = getCodesAndFlags(question);

  if (flagsOnly) {
    return codesAndFlags.flags;
  } else {
    return codesAndFlags.codes.concat(codesAndFlags.flags);
  }
}

/**
 * Replace shortcuts with full codes for the given change event
 *
 * @param {string} question
 * @param {event} e
 */
function replaceShortcutCodes(question, e) {
  // Check that we're dealing with only 1 cell
  let range = e.range;
  if (range.getWidth() > 1 || range.getHeight() > 1) {
    return;
  }

  // Check that the values we're substiting
  let value = e.value;
  Logger.log('checking value %s', value);
  let re = /^[0-9 ]+$/;
  if (!re.test(value)) {
    return;
  }

  Logger.log('looking up values');

  let codebook = getCodebook(question);
  let values = value.split(' ');
  let codes = values.map(function (value) {
    let index = parseInt(value) - 1;

    if (index < 0 || index >= codebook.length) {
      return '?';
    }

    let code = codebook[index];
    return code;
  });

  Logger.log('performing substitution');

  let newValue = codes.join(CODES_SEPARATOR);
  range.setValue(newValue);
}

/**
 * Check if the given range is a valid selection for conflict resolution
 */
function validRangeForConflicts(range) {
  return range.getWidth() == 2;
}

function showConflictInstructions() {
  let message =
    'To start conflict resolution, please select ' +
    'the two columns that contain the codes to be resolved.';
  alert(message);
}

/**
 * Insert 2 columns after the ones in the specified range
 * @param {Range} range
 * @param {Integer} howMany how many columns to insert
 * @param {Array} names the names to put in the first row of the new columns
 * @return {Integer} the index of the first newly created column
 */
function insertColumns(range, howMany, names) {
  // Figure out where to put the columns
  let position = range.getLastColumn();

  // Insert the columns
  let sheet = range.getSheet();
  sheet.insertColumnsAfter(position, howMany);

  // Give them appropriate titles
  let newColumnIndex = position + 1;
  let header = sheet.getRange(1, newColumnIndex, 1, howMany);
  header.setValues([names]);

  return newColumnIndex;
}

/**
 * Insert conflict columns after the ones in the specified range
 * @param {Range} range
 * @return {Integer} the index of the first newly created column
 */
function insertConflictColumns(range) {
  return insertColumns(range, 2, ['final', 'status']);
}

/**
 * Return object with commonalities & differences of the two arrays
 *
 * @param flags ignore any differences in flags
 */
function computeDiff(a, b, flags) {
  let both = [],
    onlyA = [],
    onlyB = [];

  // Check whether each value in A is in B
  for (let i = 0; i < a.length; i++) {
    let el = a[i];
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
    let el = b[i];
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
function formatDiff(diff) {
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

function cellDifferences(leftCell, rightCell) {
  // Get the codes
  let leftValues = leftCell.split(CODES_SEPARATOR);
  let rightValues = rightCell.split(CODES_SEPARATOR);

  // Find commonalities and differences
  let flags = getCodebook(isCodeSheet(SpreadsheetApp.getActiveSheet()), true);
  let diff = computeDiff(leftValues, rightValues, flags);
  return diff;
}

function CODES_AGREE(cellA, cellB) {
  // Check if any (real) differences remain
  let diff = cellDifferences(cellA, cellB);
  let status;
  let difference = diff.onlyA.concat(diff.onlyB);
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
  let currentSelection = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!validRangeForConflicts(currentSelection)) {
    showConflictInstructions();
    return;
  }

  // Insert new columns (after the selected ones) to hold conflict information.
  let newColumnIndex = insertConflictColumns(currentSelection);

  // Get handles to the columns with the codes to be resolved
  let currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let leftColumn = currentSelection.getColumn();
  let rightColumn = currentSelection.getLastColumn();

  let currentRow = 2;
  // For each code:
  while (currentRow <= currentSheet.getLastRow()) {
    let leftCell = currentSheet.getRange(currentRow, leftColumn);
    let rightCell = currentSheet.getRange(currentRow, rightColumn);

    let leftCellValue = leftCell.getValue();
    let rightCellValue = rightCell.getValue();
    // TODO: the profiler says the getValue call is expensive. Replace it with
    // getValues outside the loop.

    let diff = cellDifferences(leftCellValue, rightCellValue);
    let diffStr = formatDiff(diff);
    let agreementCommand =
      '=CODES_AGREE(' +
      leftCell.getA1Notation() +
      ',' +
      rightCell.getA1Notation() +
      ')';

    // Write the results
    let outputRange = currentSheet.getRange(currentRow, newColumnIndex, 1, 2);
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
function updateConflictColors(e) {
  let currentColor = e.range.getBackground();
  // Only handle cells that were marked as conflicted (using yellow color)
  if (currentColor == '#ffff00') {
    if (e.value.indexOf('<') === -1 && e.value.indexOf('>') === -1) {
      // No more conflict!
      e.range.setBackground('white');
    }
  }
}

/**
 * Determine whether the sheet is used for coding
 * and so whether we should perform substitution on it
 *
 * @param {*Sheet} sheet
 * @return the name of the question being coded, or null if there is none
 */
function isCodeSheet(sheet) {
  let sheetName = sheet.getName();
  let match = CODING_PATTERN.exec(sheetName);
  if (match) {
    return match[1];
  } else {
    return null;
  }
}

/**
 * Called when some cell in the spreadsheet has been changed
 */
function onEdit(e) {
  Logger.log('edit received');

  let sheet = e.range.getSheet();
  let code = isCodeSheet(sheet);
  if (code) {
    replaceShortcutCodes(code, e);
  }

  if (FINAL_CODES_PATTERN.exec(sheet.getName())) {
    updateConflictColors(e);
  }
}

function showCodebook() {
  let html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('Coding Assistant');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Get the question ID for the currently selected question
 */
function getCurrentQuestionCode() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Is the current sheet a coding sheet?
  let question = isCodeSheet(sheet);
  if (question !== null) {
    return question;
  }

  // Else, is the current sheet a codebook sheet?
  let match = CODEBOOK_PATTERN.exec(sheet.getName());
  if (match) {
    return match[1];
  }

  // Otherwise, I really have no idea what sheet this is.
  return null;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Coding Assistant')
    .addItem('Show codebook', 'showCodebook')
    .addItem('Find conflicts', 'findConflicts')
    .addItem('Compute Kupper-Hafner agreement', 'computeKupperHafner')
    .addToUi();
}
