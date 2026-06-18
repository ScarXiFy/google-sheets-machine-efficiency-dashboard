function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Machine Dashboard')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Load Sample Data', 'loadSampleData')
    .addToUi();
}
