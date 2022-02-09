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

function showConflictInstructions() {
  let message =
    'To start conflict resolution, please select ' +
    'the two columns that contain the codes to be resolved.';
  alert(message);
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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Coding Assistant')
    .addItem('Show codebook', 'showCodebook')
    .addItem('Find conflicts', 'findConflicts')
    .addItem('Compute Kupper-Hafner agreement', 'computeKupperHafner')
    .addToUi();
}
