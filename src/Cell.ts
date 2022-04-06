type Cell = number | string | Date;
type NonDateCell = number | string;

type CellRange = Array<Array<Cell>>;
type CellOrRange = Cell | CellRange;

function isRange_(input: CellOrRange): input is CellRange {
  return Array.isArray(input);
}

function cellIsEmpty_(cell: Cell): boolean {
  return cell == '';
}

function filterEmpty_(array: string[]) {
  return array.filter(function (value) {
    return !cellIsEmpty_(value);
  });
}

/**
 * Return true if all given cells are empty
 * @param cells
 * @returns
 */
function cellsAreEmpty_(cells: Cell[]): boolean {
  return cells.reduce<boolean>((acc, cell) => acc && cellIsEmpty_(cell), true);
}

function getCodesInCell_(cell: Cell): string[] {
  return filterEmpty_(
    cell
      .toString()
      .split(CODES_SEPARATOR)
      .map((str) => str.trim())
  ).unique();
}

function cellIsDate_(cell: Cell): cell is NonDateCell {
  return !(cell instanceof Date);
}

function validateCells_(cells: Cell[]): NonDateCell[] {
  return cells.map((cell) => {
    if (cell instanceof Date) {
      throw new QcasError(`code ${cell} is unexpected type: Date`);
    }
    return cell;
  });
}

function cellAsNumberOrError_(cell: Cell): number {
  if (isNumber(cell)) {
    return cell;
  }

  const value = Number(cell.toString());
  if (isNumber(value)) {
    return value;
  }

  throw new QcasError(`expected number but got ${cell}`);
}
