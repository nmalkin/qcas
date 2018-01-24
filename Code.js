var CODEBOOK_HEADER_CODES = 'Codes';
var CODEBOOK_HEADER_FLAGS = 'Flags';
var CODEBOOK_PATTERN = /(\w+)_codebook/;
var CODING_COLUMN = 2;
var CODING_PATTERN = /(\w+)_codes(_\w+)?/;

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
 * Given a sheet and a column name, return that column's values
 */
function getColumnByName(sheet, name) {
  var columnNumber = getColumnNumberByName(sheet, name);
  if (columnNumber === -1) {
    Logger.log('Invalid column name');
    return [];
  }

  var range = sheet.getRange(2, columnNumber, sheet.getLastRow() - 1, 1);
  return getAllValues(range);
}

/**
 * Return an array of all codes in the codebook
 *
 * @param question the name of the question, used in the sheet title
 * @param {boolean} flagsOnly if true, only return the flags
 */
function getCodebook(question, flagsOnly) {
  var codebookSheetName = question + '_codebook';
  var sheet = getSheet(codebookSheetName);
  var codes = getColumnByName(sheet, CODEBOOK_HEADER_CODES);
  var flags = getColumnByName(sheet, CODEBOOK_HEADER_FLAGS);

  if (flagsOnly) {
    return flags;
  } else {
    return codes.concat(flags);
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

  // Check that the modification is in the right column
  var column = range.getColumn();
  if (column != CODING_COLUMN) {
    // TODO: also check that it's the right sheet
    return;
  }

  // Check that the values we're substiting
  var value = e.value;
  Logger.log('checking value %s', value);
  var re = /[0-9 ]+/;
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
 * @return {Integer} the index of the first newly created column
 */
function insertConflictColumns(range) {
  // Figure out where to put the columns
  var position = range.getLastColumn();

  // Insert the columns
  var sheet = range.getSheet();
  sheet.insertColumnsAfter(position, 2);

  // Give them appropriate titles
  var newColumnIndex = position + 1;
  var header = sheet.getRange(1, newColumnIndex, 1, 2);
  header.setValues([['final', 'status']]);

  return newColumnIndex;
}

/**
 * Compute the union of two arrays
 */
function findUnion(a, b) {
  var c = a.slice();

  for (var i = 0; i < b.length; i++) {
    var el = b[i];
    if (!c.includes(el)) {
      c.push(el);
    }
  }

  return c;
}

/**
 * Return the elements not in both arrays
 */
function findDifference(a, b) {
  var diff = [];

  // Loop over both array, checking that each value isn't contained in the
  // other one.
  // This implementation is suboptimal, but is used for API compatibility and
  // simplicity.
  for (var i = 0; i < a.length; i++) {
    var el = a[i];
    if (!b.includes(el)) {
      diff.push(el);
    }
  }

  for (var i = 0; i < b.length; i++) {
    var el = b[i];
    if (!a.includes(el)) {
      diff.push(el);
    }
  }

  return diff;
}

// https://tc39.github.io/ecma262/#sec-array.prototype.includes
if (!Array.prototype.includes) {
  Object.defineProperty(Array.prototype, 'includes', {
    value: function(searchElement, fromIndex) {
      if (this == null) {
        throw new TypeError('"this" is null or not defined');
      }

      // 1. Let O be ? ToObject(this value).
      var o = Object(this);

      // 2. Let len be ? ToLength(? Get(O, "length")).
      var len = o.length >>> 0;

      // 3. If len is 0, return false.
      if (len === 0) {
        return false;
      }

      // 4. Let n be ? ToInteger(fromIndex).
      //    (If fromIndex is undefined, this step produces the value 0.)
      var n = fromIndex | 0;

      // 5. If n ≥ 0, then
      //  a. Let k be n.
      // 6. Else n < 0,
      //  a. Let k be len + n.
      //  b. If k < 0, let k be 0.
      var k = Math.max(n >= 0 ? n : len - Math.abs(n), 0);

      function sameValueZero(x, y) {
        return (
          x === y ||
          (typeof x === 'number' &&
            typeof y === 'number' &&
            isNaN(x) &&
            isNaN(y))
        );
      }

      // 7. Repeat, while k < len
      while (k < len) {
        // a. Let elementK be the result of ? Get(O, ! ToString(k)).
        // b. If SameValueZero(searchElement, elementK) is true, return true.
        if (sameValueZero(o[k], searchElement)) {
          return true;
        }
        // c. Increase k by 1.
        k++;
      }

      // 8. Return false
      return false;
    }
  });
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
  while (currentRow < currentSheet.getLastRow()) {
    var leftCell = currentSheet.getRange(currentRow, leftColumn).getValue();
    var rightCell = currentSheet.getRange(currentRow, rightColumn).getValue();
    // TODO: the profiler says the getValue call is expensive. Replace it with
    // getValues outside the loop.

    // Get the codes
    var leftValues = leftCell.split(',');
    var rightValues = rightCell.split(',');

    // We'll output the union of the codes
    var union = findUnion(leftValues, rightValues);
    var finalValue = union.join(',');

    // Look for any differences between the codes
    var difference = findDifference(leftValues, rightValues);

    // But if the differences are only in flags ("disagreeable" codes), ignore them
    var flags = getCodebook(isCodeSheet(SpreadsheetApp.getActiveSheet()), true);
    for (var i = 0; i < difference.length; i++) {
      var currentDifference = difference[i];
      if (flags.includes(currentDifference)) {
        difference.splice(i, 1);
      }
    }

    // Check if any (real) differences remain
    var status;
    if (difference.length == 0) {
      status = 'agree';
    } else {
      status = 'conflict';
    }

    // Write the results
    var outputRange = currentSheet.getRange(currentRow, newColumnIndex, 1, 2);
    outputRange.setValues([[finalValue, status]]);

    if (status == 'conflict') {
      outputRange.setBackground('yellow');
    }

    currentRow++;
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
}

function showCodebook() {
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('Coding Helper');
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
    .createMenu('Coding Helper')
    .addItem('Show codebook', 'showCodebook')
    .addItem('Find conflicts', 'findConflicts')
    .addToUi();
}
