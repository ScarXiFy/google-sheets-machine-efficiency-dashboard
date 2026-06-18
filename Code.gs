function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Machine Dashboard')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Load Sample Data', 'loadSampleData')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Generate Corrective Actions', 'generateCorrectiveActions')
    .addItem('Generate Daily Summary', 'generateDailySummary')
    .addItem('Export Dashboard PDF', 'exportDashboardPdf')
    .addItem('Export Corrective Actions PDF', 'exportCorrectiveActionsPdf')
    .addToUi();
}
