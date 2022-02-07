type Cell = string;
type CellRange = Array<Array<string>>;
type CellOrRange = Cell | CellRange;

const CODEBOOK_HEADER_FINAL = 'Code - final';
const CODEBOOK_SHEET_NAME = (questionId: string) => questionId + '_codebook';

/**
 * Return specified sheet
 */
function getSheetOrError_(sheetName: string) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet == null) {
    throw new QcasError("Couldn't find a sheet with the name " + sheetName);
  }
  return sheet;
}

function getCodebookSheet_(questionId: string) {
  const codebookSheetName = CODEBOOK_SHEET_NAME(questionId);
  const sheet = getSheetOrError_(codebookSheetName);
  return sheet;
}

/**
 * Return an object mapping all codes to their final values
 *
 * @param question the name of the question, used in the sheet title
 */
function getCodeToFinalNameMapping_(question: string) {
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

  const codeNameMappings: Record<string, string> = {};

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
function FINALNAMES(input: CellOrRange) {
  const questionId: string | null = getCurrentQuestionCode();
  if (!questionId) {
    throw new QcasError(
      "couldn't determine which codebook current sheet is associated with"
    );
  }

  const mappings = getCodeToFinalNameMapping_(questionId);

  const mapCellContents = (cellContents: string) => {
    const codeList = getCodeList(cellContents);

    const renamedCodes = codeList.map((code: string) => {
      if (!(code in mappings)) {
        throw new QcasError(`ERROR: ${code} not found in codebook`);
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
