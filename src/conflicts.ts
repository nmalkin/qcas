/**
 * Check if the given range is a valid selection for conflict resolution
 */
function validRangeForConflicts(
  range: GoogleAppsScript.Spreadsheet.Range | null
): range is GoogleAppsScript.Spreadsheet.Range {
  return range != null && range.getWidth() == 2;
}
