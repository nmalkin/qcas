function isString(data: unknown): data is string {
  return typeof data === 'string';
}

/**
 * via https://stackoverflow.com/a/46700791
 */
function notEmpty<TValue>(value: TValue | null | undefined): value is TValue {
  if (value === null || value === undefined) return false;
  const testDummy: TValue = value;
  return true;
}
