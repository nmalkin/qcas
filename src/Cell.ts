type Cell = number | string | Date;
type CellRange = Array<Array<Cell>>;
type CellOrRange = Cell | CellRange;

function isRange_(input: CellOrRange): input is CellRange {
  return Array.isArray(input);
}

function filterEmpty_(array: string[]) {
  return array.filter(function (value) {
    return value != '';
  });
}

function getCodesInCell_(cell: Cell): string[] {
  return filterEmpty(cell.toString().split(CODES_SEPARATOR)).unique();
}
