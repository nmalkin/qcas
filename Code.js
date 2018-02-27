/*
Copyright (C) 2018 N. Malkin

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

var CODEBOOK_HEADER_CODE = 'Code';
var CODEBOOK_HEADER_TYPE = 'Type';
var CODEBOOK_TYPE_CODE = 'code';
var CODEBOOK_TYPE_FLAG = 'flag';
var CODEBOOK_PATTERN = /(\w+)_codebook/;
var CODING_PATTERN = /(\w+)_codes(_\w+)?/;
var FINAL_CODES_PATTERN = /(\w+)_codes_final/;

var FIRST_ROW = 2; // Assuming a header, row 2 is always the first row.

function alert(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

/**
 * Return specified sheet
 */
function getSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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
  var next = arr.pop();
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
  var data = [];
  var values = range.getValues();
  for (var i = 0; i < range.getHeight(); i++) {
    for (var j = 0; j < range.getWidth(); j++) {
      var value = values[i][j];
      data.push(value);
    }
  }

  return truncateEmptyArray(data);
}

/**
 * Given a sheet and a column name, return that column's number
 */
function getColumnNumberByName(sheet, name) {
  var header = getAllValues(getSheetHeader(sheet));
  for (var i = 0; i < header.length; i++) {
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
 * @param {boolean} flagsOnly if true, only return the flags
 */
function getCodesAndFlags(question) {
  var codebookSheetName = question + '_codebook';
  var sheet = getSheet(codebookSheetName);

  // Find the range where the relevant codebook columns are located
  var codeColumn = getColumnNumberByName(sheet, CODEBOOK_HEADER_CODE);
  var typeColumn = getColumnNumberByName(sheet, CODEBOOK_HEADER_TYPE);
  var firstColumn = Math.min(codeColumn, typeColumn);
  var lastColumn = Math.max(codeColumn, typeColumn);
  var range = sheet.getRange(
    FIRST_ROW,
    firstColumn,
    sheet.getLastRow() - 1,
    lastColumn - firstColumn + 1
  );

  var values = range.getValues();
  var codes = [],
    flags = [];
  for (var i = 0; i < range.getHeight(); i++) {
    var code = values[i][codeColumn - firstColumn];
    if (code === '') {
      // Tolerate holes in codebook
      continue;
    }

    var type = values[i][typeColumn - firstColumn];
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
    flags: flags
  };
}

/**
 * Return an array of all codes in the codebook
 *
 * @param question the name of the question, used in the sheet title
 * @param {boolean} flagsOnly if true, only return the flags
 */
function getCodebook(question, flagsOnly) {
  var codesAndFlags = getCodesAndFlags(question);

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
  var range = e.range;
  if (range.getWidth() > 1 || range.getHeight() > 1) {
    return;
  }

  // Check that the values we're substiting
  var value = e.value;
  Logger.log('checking value %s', value);
  var re = /^[0-9 ]+$/;
  if (!re.test(value)) {
    return;
  }

  Logger.log('looking up values');

  var codebook = getCodebook(question);
  var values = value.split(' ');
  var codes = values.map(function(value) {
    var index = parseInt(value) - 1;

    if (index < 0 || index >= codebook.length) {
      return '?';
    }

    var code = codebook[index];
    return code;
  });

  Logger.log('performing substitution');

  var newValue = codes.join(',');
  range.setValue(newValue);
}

/**
 * Check if the given range is a valid selection for conflict resolution
 */
function validRangeForConflicts(range) {
  return range.getWidth() == 2;
}

function showConflictInstructions() {
  var message =
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
  var position = range.getLastColumn();

  // Insert the columns
  var sheet = range.getSheet();
  sheet.insertColumnsAfter(position, howMany);

  // Give them appropriate titles
  var newColumnIndex = position + 1;
  var header = sheet.getRange(1, newColumnIndex, 1, howMany);
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
  var both = [],
    onlyA = [],
    onlyB = [];

  // Check whether each value in A is in B
  for (var i = 0; i < a.length; i++) {
    var el = a[i];
    if (b.includes(el)) {
      both.push(el);
    } else if (flags.includes(el)) {
      both.push(el);
    } else {
      onlyA.push(el);
    }
  }

  // Now check whether each value in B is in A
  for (var i = 0; i < b.length; i++) {
    var el = b[i];
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
    onlyB: onlyB
  };
}

/**
 * Given a diff object (@see computeDiff) return string representation of it
 */
function formatDiff(diff) {
  var str = '';
  if (diff.both.length > 0) {
    str += diff.both.join(',');
  }
  if (diff.onlyA.length > 0) {
    str += '\n<' + diff.onlyA.join(',');
  }
  if (diff.onlyB.length > 0) {
    str += '\n>' + diff.onlyB.join(',');
  }
  return str;
}

/**
 * Look for conflicts between codes in two columns.
 * Write the union of the codes into a new column.
 * If there's disagreement, flag it (in a new column).
 */
function findConflicts() {
  // Check that we can actually compute conflicts for this range.
  var currentSelection = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!validRangeForConflicts(currentSelection)) {
    showConflictInstructions();
    return;
  }

  // Insert new columns (after the selected ones) to hold conflict information.
  var newColumnIndex = insertConflictColumns(currentSelection);

  // Get handles to the columns with the codes to be resolved
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var leftColumn = currentSelection.getColumn();
  var rightColumn = currentSelection.getLastColumn();

  var currentRow = 2;
  // For each code:
  while (currentRow <= currentSheet.getLastRow()) {
    var leftCell = currentSheet.getRange(currentRow, leftColumn).getValue();
    var rightCell = currentSheet.getRange(currentRow, rightColumn).getValue();
    // TODO: the profiler says the getValue call is expensive. Replace it with
    // getValues outside the loop.

    // Get the codes
    var leftValues = leftCell.split(',');
    var rightValues = rightCell.split(',');

    // Find commonalities and differences
    var flags = getCodebook(isCodeSheet(SpreadsheetApp.getActiveSheet()), true);
    var diff = computeDiff(leftValues, rightValues, flags);
    var diffStr = formatDiff(diff);

    // Check if any (real) differences remain
    var status;
    var difference = diff.onlyA.concat(diff.onlyB);
    if (difference.length == 0) {
      status = 'agree';
    } else {
      status = 'conflict';
    }

    // Write the results
    var outputRange = currentSheet.getRange(currentRow, newColumnIndex, 1, 2);
    outputRange.setValues([[diffStr, status]]);

    if (status == 'conflict') {
      outputRange.setBackgrounds([['yellow', 'white']]);
    }

    currentRow++;
  }
}

/**
 * Clear the color in the given cell if it doesn't have conflicts anymore
 */
function updateConflictColors(e) {
  var currentColor = e.range.getBackground();
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
  var sheetName = sheet.getName();
  var match = CODING_PATTERN.exec(sheetName);
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

  var sheet = e.range.getSheet();
  var code = isCodeSheet(sheet);
  if (code) {
    replaceShortcutCodes(code, e);
  }

  if (FINAL_CODES_PATTERN.exec(sheet.getName())) {
    updateConflictColors(e);
  }
}

function showCodebook() {
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('Coding Assistant');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Get the question ID for the currently selected question
 */
function getCurrentQuestionCode() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Is the current sheet a coding sheet?
  var question = isCodeSheet(sheet);
  if (question !== null) {
    return question;
  }

  // Else, is the current sheet a codebook sheet?
  var match = CODEBOOK_PATTERN.exec(sheet.getName());
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
