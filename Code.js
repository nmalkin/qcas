var CODEBOOK_SHEET = 'laws_codebook';
var CODEBOOK_HEADER = 'Code';
var CODING_COLUMN = 2;

/**
 * Return specified sheet
 */
function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * Given a sheet, return its first (header) row as a Range
 */
function getSheetHeader(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn());
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

  return data;
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
    alert('Invalid column name');
  }

  var range = sheet.getRange(2, columnNumber, sheet.getLastRow() - 1, 1);
  return getAllValues(range);
}

/**
 * Return an array of all codes in the codebook
 */
function getCodebook() {
  return getColumnByName(getSheet(CODEBOOK_SHEET), CODEBOOK_HEADER);
}

function replaceShortcutCodes(e) {
  Logger.log('edit received');

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

  var codebook = getCodebook();
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
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
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

function findConflicts() {
  var currentSelection = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!validRangeForConflicts(currentSelection)) {
    showConflictInstructions();
    return;
  }

  var newColumnIndex = insertConflictColumns(currentSelection);

  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var leftColumn = currentSelection.getColumn();
  var rightColumn = currentSelection.getLastColumn();

  var currentRow = 2;
  while (currentRow < currentSheet.getLastRow()) {
    var leftCell = currentSheet.getRange(currentRow, leftColumn).getValue();
    var rightCell = currentSheet.getRange(currentRow, rightColumn).getValue();
    // TODO: the profiler says the getValue call is expensive. Replace it with
    // getValues outside the loop.

    var leftValues = leftCell.split(',');
    var rightValues = rightCell.split(',');

    var union = findUnion(leftValues, rightValues);
    var finalValue = union.join(',');

    var difference = findDifference(leftValues, rightValues);
    var status;
    if (difference.length == 0) {
      status = 'agree';
    } else {
      status = 'conflict';
    }

    var outputRange = currentSheet.getRange(currentRow, newColumnIndex, 1, 2);
    outputRange.setValues([[finalValue, status]]);

    if (status == 'conflict') {
      outputRange.setBackground('yellow');
    }

    currentRow++;
  }
}

function onEdit(e) {
  replaceShortcutCodes(e);
}

function showCodebook() {
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('Coding Helper');
  SpreadsheetApp.getUi().showSidebar(html);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Coding Helper')
    .addItem('Show codebook', 'showCodebook')
    .addItem('Find conflicts', 'findConflicts')
    .addToUi();
}
