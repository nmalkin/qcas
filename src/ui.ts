// eslint-disable-next-line @typescript-eslint/no-unused-vars
function showAlert(message: string) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}
