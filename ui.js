function onOpen() {
  createCustomMenu();
}

function createSidebar() {
  const html = HtmlService.createTemplateFromFile("sideBar").evaluate().setTitle("Import excel");
  SpreadsheetApp.getUi().showSidebar(html);
}

function createCustomMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Custom')
    .addItem('Import excel', 'createSidebar')
    .addToUi();
}