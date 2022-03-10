function isString(data: unknown): data is string {
  return typeof data === 'string';
}

function isNumber(value: unknown): value is number {
  return (
    value != undefined && value != null && value !== '' && !isNaN(Number(value))
  );
}

/**
 * via https://stackoverflow.com/a/46700791
 */
function notEmpty<TValue>(value: TValue | null | undefined): value is TValue {
  if (value === null || value === undefined) return false;
  const testDummy: TValue = value;
  return true;
}

/**
 * For each value in the input range, return the maximum of that or the given number
 * @customfunction
 * @param value
 * @param range
 * @returns
 */
function MAXEACH(value: number, range: CellRange): CellRange {
  return range.map((row) =>
    row.map((cell) => {
      const n = Number(cell);
      return isNaN(n) ? NaN : Math.max(value, n);
    })
  );
}

function range(start: number, stop: number, step = 1): number[] {
  return Array(stop - start)
    .fill(start)
    .map((x, y) => x + y * step);
}
