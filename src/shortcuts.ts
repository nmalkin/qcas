/**
 * Replace shortcuts with full codes for the given change event
 *
 * @param {string} question
 * @param {event} e
 */
function replaceShortcutCodes(
  question: string,
  e: GoogleAppsScript.Events.SheetsOnEdit
) {
  // Check that we're dealing with only 1 cell
  const range = e.range;
  if (range.getWidth() > 1 || range.getHeight() > 1) {
    return;
  }

  // Check that the values we're substituting are only numbers
  const value = e.value;
  const re = /^[0-9 ]+$/;
  if (!re.test(value)) {
    return;
  }

  const codebook = getCodebook(question);
  const values = value.split(' ');
  const codes = values.map(function (value) {
    const index = parseInt(value) - 2;

    if (index < 0 || index >= codebook.length) {
      return '?';
    }

    const code = codebook[index];
    return code;
  });

  const newValue = codes.join(CODES_SEPARATOR);
  range.setValue(newValue);
}
