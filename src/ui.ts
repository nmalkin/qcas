// eslint-disable-next-line @typescript-eslint/no-unused-vars
function showAlert(message: string) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

function showGenericInstructions(): void {
  const message =
    'To start, please select the two columns that contain the codes .';
  showAlert(message);
}
