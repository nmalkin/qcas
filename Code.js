var CODEBOOK_SHEET = 'laws_codebook';
var CODEBOOK_HEADER = 'Code';
var CODING_SHEET = 'laws_codes';
var CODING_COLUMN = 2;

function TEST() {
  Logger.log(getCodebook());
}

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

function splitOnSemicolons(data) {
  var newData = [];
  data.forEach(function(item) {
    newData = newData.concat(item.split(';'));
  });
  return newData;
}

function filterEmpty(array) {
  return array.filter(function(value) {
    return value != '';
  });
}

Array.prototype.unique = function() {
  var arr = [];
  for (var i = 0; i < this.length; i++) {
    if (arr.indexOf(this[i]) == -1) {
      arr.push(this[i]);
    }
  }
  return arr;
};

function getData() {
  var data = splitOnSemicolons(filterEmpty(getColumn(1)))
    .sort()
    .unique();
  return data;
}

function showCodebook() {
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('Coding Helper');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .showSidebar(html);
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .createMenu('Coding Helper')
    .addItem('Show codebook', 'showCodebook')
    .addItem('Find conflicts', 'findConflicts')
    .addToUi();
}

function onEdit(e) {
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
