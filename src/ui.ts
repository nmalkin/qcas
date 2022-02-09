// eslint-disable-next-line @typescript-eslint/no-unused-vars
function alert(message: string) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}
