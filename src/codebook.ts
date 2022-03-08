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

interface CodesAndFlags {
  codes: string[];
  flags: string[];
}

/**
 * Return an object with all codes and flags in the codebook
 *
 * @param question the name of the question, used in the sheet title
 */
function getCodesAndFlags(question: string): CodesAndFlags {
  const sheet = getCodebookSheet_(question);

  // Find the range where the relevant codebook columns are located
  const codeColumn = getColumnNumberByName(sheet, CODEBOOK_HEADER_CODE);
  const typeColumn = getColumnNumberByName(sheet, CODEBOOK_HEADER_TYPE);
  const firstColumn = Math.min(codeColumn, typeColumn);
  const lastColumn = Math.max(codeColumn, typeColumn);
  const range = sheet.getRange(
    FIRST_ROW,
    firstColumn,
    sheet.getLastRow() - 1,
    lastColumn - firstColumn + 1
  );

  const values = range.getValues();
  const codes = [],
    flags = [];
  for (let i = 0; i < range.getHeight(); i++) {
    const code = values[i][codeColumn - firstColumn];
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
      showAlert('Unrecognized code type ' + type + ' in codebook ' + question);
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
function getCodebook(question: string, flagsOnly?: boolean): string[] {
  const codesAndFlags = getCodesAndFlags(question);

  if (flagsOnly) {
    return codesAndFlags.flags;
  } else {
    return codesAndFlags.codes.concat(codesAndFlags.flags);
  }
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
 * Replace codes with their final names
 * @param {string | Array<Array<string>>} input
 * @return final names for all codes
 * @customfunction
 */
function FINALNAMES(input: CellOrRange): CellOrRange {
  const questionId = getCurrentQuestionCodeOrError_();

  const mappings = getCodeToFinalNameMapping_(questionId);

  const mapCellContents = (
    cellContents: Cell,
    rowNumber?: number,
    columnNumber?: number
  ) => {
    const codeList = getCodesInCell_(cellContents);

    const renamedCodes = codeList.map((code: string) => {
      if (!(code in mappings)) {
        throw new QcasError(
          `not in codebook: ${code}` +
            (rowNumber != undefined ? ` in row ${rowNumber + 1}` : '') +
            (columnNumber != undefined ? `, column ${columnNumber + 1}` : '')
        );
      }

      return mappings[code];
    });

    return renamedCodes.unique().join(CODES_SEPARATOR);
  };

  if (isRange_(input)) {
    return input.map((row, rowNumber) =>
      row.map((cell, columnNumber) =>
        mapCellContents(cell, rowNumber, columnNumber)
      )
    );
  } else {
    return mapCellContents(input);
  }
}

/**
 * Filter flags from given code cells
 * @param {string | Array<Array<string>>} input
 * @return input codes but with flags removed
 * @customfunction
 */
function FILTERFLAGS(input: CellOrRange): CellOrRange {
  const questionId = getCurrentQuestionCodeOrError_();

  const codesAndFlags = getCodesAndFlags(questionId);
  const allCodes = new Set(codesAndFlags.codes);
  const allFlags = new Set(codesAndFlags.flags);

  const filterCellContents = (
    cellContents: Cell,
    rowNumber?: number,
    columnNumber?: number
  ) => {
    const codeList = getCodesInCell_(cellContents);

    const filteredCodes = codeList.filter((code: string) => {
      if (allCodes.has(code)) {
        return true;
      } else if (allFlags.has(code)) {
        return false;
      } else {
        throw new QcasError(
          `not in codebook: ${code}` +
            (rowNumber != undefined ? ` in row ${rowNumber + 1}` : '') +
            (columnNumber != undefined ? `, column ${columnNumber + 1}` : '')
        );
      }
    });

    return filteredCodes.join(CODES_SEPARATOR);
  };

  if (isRange_(input)) {
    return input.map((row, rowNumber) =>
      row.map((cell, columnNumber) =>
        filterCellContents(cell, rowNumber, columnNumber)
      )
    );
  } else {
    return filterCellContents(input);
  }
}
