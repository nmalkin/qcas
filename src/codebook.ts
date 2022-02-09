/**
 * Return sheet with given name
 * @throws if no sheet with given name is found
 */
function getSheetOrError_(
  sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet == null) {
    throw new QcasError("Couldn't find a sheet with the name " + sheetName);
  }
  return sheet;
}

/**
 * Return codebook sheet for the given question
 * @param questionId
 * @returns
 */
function getCodebookSheet_(
  questionId: string
): GoogleAppsScript.Spreadsheet.Sheet {
  const codebookSheetName = CODEBOOK_SHEET_NAME(questionId);
  const sheet = getSheetOrError_(codebookSheetName);
  return sheet;
}

/**
 * Return an object mapping all codes to their final values
 *
 * @param question the name of the question, used in the sheet title
 */
function getCodeToFinalNameMapping_(question: string): Record<string, string> {
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
 * Get a deduplicated list of all final names for the codes in the codebook
 *
 * @param questionId
 */
function getFinalCodeList_(questionId: string): string[] {
  const mappings = getCodeToFinalNameMapping_(questionId);
  const allFinalNames = Object.values(mappings);
  return allFinalNames.unique();
}

/**
 *
 * @param {string | Array<Array<string>>} input
 * @return original code names returned to input
 * @customfunction
 */
function FINALNAMES(input: CellOrRange): CellOrRange {
  const questionId: string | null = getCurrentQuestionCode();
  if (!questionId) {
    throw new QcasError(
      "couldn't determine which codebook current sheet is associated with"
    );
  }

  const mappings = getCodeToFinalNameMapping_(questionId);

  const mapCellContents = (cellContents: Cell) => {
    const codeList = getCodeList(cellContents.toString());

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
