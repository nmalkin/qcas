const CODEBOOK_HEADER_FINAL = 'Code - final';
const CODEBOOK_SHEET_NAME = (questionId) => questionId + '_codebook';

class QcasError extends Error {
  constructor(message) {
    super(message);
  }
}

/**
 * Return specified sheet
 */
function getSheetOrError_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet == null) {
    throw QcasError("Couldn't find a sheet with the name " + sheetName);
  }
  return sheet;
}

function getCodebookSheet_(questionId) {
  const codebookSheetName = CODEBOOK_SHEET_NAME(questionId);
  const sheet = getSheetOrError_(codebookSheetName);
  return sheet;
}

/**
 * Return an object mapping all codes to their final values
 *
 * @param question the name of the question, used in the sheet title
 */
function getCodeToFinalNameMapping_(question) {
  const sheet = getCodebookSheet_(question);

  // Find the range where the relevant codebook columns are located
  const codeNameColumn = getColumnNumberByName(sheet, CODEBOOK_HEADER_CODE);
  const finalNameColumn = getColumnNumberByName(sheet, CODEBOOK_HEADER_FINAL);

  const firstColumn = Math.min(codeNameColumn, finalNameColumn);
  const lastColumn = Math.max(codeNameColumn, finalNameColumn);
  const range = sheet.getRange(
    FIRST_ROW,
    firstColumn,
    sheet.getLastRow() - 1,
    lastColumn - firstColumn + 1
  );
  const values = range.getValues();

  const codeNameMappings = {};

  for (let i = 0; i < range.getHeight(); i++) {
    const codeName = values[i][codeNameColumn - firstColumn];
    if (codeName === '') {
      // Tolerate holes in codebook
      continue;
    }

    let finalName = values[i][finalNameColumn - firstColumn];
    if (finalName === '') {
      // If no final name is specified, keep the original
      finalName = codeName;
    }

    codeNameMappings[codeName] = finalName;
  }

  return codeNameMappings;
}

/**
 *
 * @param {string | Array<Array<string>>} input
 * @return original code names returned to input
 * @customfunction
 */
function FINALNAMES(input) {
  const questionId = getCurrentQuestionCode();
  const mappings = getCodeToFinalNameMapping_(questionId);

  const mapCellContents = (cellContents) => {
    const codeList = getCodeList(cellContents);

    const renamedCodes = codeList.map((code) => {
      if (!(code in mappings)) {
        throw QcasError(`ERROR: ${code} not found in codebook`);
      }

      return mappings[code];
    });

    return renamedCodes.join(',');
  };

  if (Array.isArray(input)) {
    return input.map((row) => row.map((cell) => mapCellContents(cell)));
  } else {
    return mapCellContents(input);
  }
}
