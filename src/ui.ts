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

function html_(content: string): string {
  return `<!DOCTYPE html>
  <html>
  
  <head>
    <base target="_top">
    <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
  </head>
  <body>
  ${content}
  </body>
  </html>
`;
}

function showAbout_() {
  var htmlOutput = HtmlService.createHtmlOutput(
    html_(
      `<h1>QCAS ${VERSION}</h1>
    <p>For more information and documentation, please visit
    <a href='https://github.com/nmalkin/qcas'>https://github.com/nmalkin/qcas</a></p>`
    )
  )
    .setWidth(200)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About QCAS');
}
