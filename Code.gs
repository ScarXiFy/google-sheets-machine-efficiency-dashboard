function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Machine Dashboard')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Load Sample Data', 'loadSampleData')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Generate Corrective Actions', 'generateCorrectiveActions')
    .addItem('Generate Daily Summary', 'generateDailySummary')
    .addItem('Generate Historical Trends', 'generateHistoricalTrends')
    .addItem('Setup Data Validation', 'setupDataValidation')
    .addItem('Audit Production Data', 'auditProductionData')
    .addItem('Send Daily Report', 'sendDailyReport')
    .addItem('Send Weekly Report', 'sendWeeklyReport')
    .addItem('View Report History', 'viewReportHistory')
    .addItem('Export Dashboard PDF', 'exportDashboardPdf')
    .addItem('Export Corrective Actions PDF', 'exportCorrectiveActionsPdf')
    .addToUi();
}
