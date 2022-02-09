function showCodebook() {
  const html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('Coding Assistant');
  SpreadsheetApp.getUi().showSidebar(html);
}
