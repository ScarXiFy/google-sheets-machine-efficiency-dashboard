var DAILY_SUMMARY_SHEET_NAME = 'Daily Summary';
var REPORT_DASHBOARD_SHEET_NAME = 'Dashboard';
var REPORT_DAILY_PRODUCTION_SHEET_NAME = 'DailyProduction';
var REPORT_CORRECTIVE_ACTIONS_SHEET_NAME = 'Corrective Actions';

function generateDailySummary() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = spreadsheet.getSheetByName(REPORT_DASHBOARD_SHEET_NAME);
  var productionSheet = spreadsheet.getSheetByName(REPORT_DAILY_PRODUCTION_SHEET_NAME);

  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('Dashboard sheet was not found. Please run Setup Sheets first.');
    return;
  }

  if (!productionSheet) {
    SpreadsheetApp.getUi().alert('DailyProduction sheet was not found. Please run Setup Sheets first.');
    return;
  }

  var productionRows = getProductionRows_(productionSheet);
  var filters = readDashboardFilters(dashboardSheet);
  var validationResult = validateDashboardFilters(filters, productionRows);

  if (!validationResult.isValid) {
    SpreadsheetApp.getUi().alert(validationResult.message);
    return;
  }

  var normalizedFilters = validationResult.filters;
  var filteredRows = applyProductionFilters(productionRows, normalizedFilters);
  var dashboardData = calculateDashboardData_(filteredRows);
  var summarySheet = getOrCreateReportSheet_(spreadsheet, DAILY_SUMMARY_SHEET_NAME);

  writeDailySummary_(summarySheet, dashboardData, normalizedFilters);

  SpreadsheetApp.getUi().alert('Daily Summary generated.');
}

function exportDashboardPdf() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = spreadsheet.getSheetByName(REPORT_DASHBOARD_SHEET_NAME);

  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('Dashboard sheet was not found. Please run Setup Sheets first.');
    return;
  }

  var fileName = 'Machine_Efficiency_Dashboard_' + getReportTimestamp_() + '.pdf';
  createPdfExport_(spreadsheet, dashboardSheet, fileName, true);

  SpreadsheetApp.getUi().alert('Dashboard PDF exported successfully.\n\nFile name: ' + fileName);
}

function exportCorrectiveActionsPdf() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var correctiveActionsSheet = spreadsheet.getSheetByName(REPORT_CORRECTIVE_ACTIONS_SHEET_NAME);

  if (!correctiveActionsSheet) {
    SpreadsheetApp.getUi().alert('Corrective Actions sheet was not found. Please run Setup Sheets first.');
    return;
  }

  if (correctiveActionsSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Corrective Actions sheet has no report rows to export.');
    return;
  }

  var fileName = 'Corrective_Actions_Report_' + getReportTimestamp_() + '.pdf';
  createPdfExport_(spreadsheet, correctiveActionsSheet, fileName, true);

  SpreadsheetApp.getUi().alert('Corrective Actions PDF exported successfully.\n\nFile name: ' + fileName);
}

function createPdfExport_(spreadsheet, sheet, fileName, useLandscape) {
  var exportUrl = spreadsheet.getUrl().replace(/edit.*$/, '') +
    'export?format=pdf' +
    '&gid=' + sheet.getSheetId() +
    '&size=letter' +
    '&portrait=' + (!useLandscape) +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=false' +
    '&gridlines=false' +
    '&fzr=false' +
    '&top_margin=0.50' +
    '&bottom_margin=0.50' +
    '&left_margin=0.50' +
    '&right_margin=0.50';
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });
  var blob = response.getBlob().setName(fileName);

  return DriveApp.createFile(blob);
}

function getCurrentFilterSummary_(filters) {
  return [
    ['Date From', formatReportDateValue_(filters.dateFrom)],
    ['Date To', formatReportDateValue_(filters.dateTo)],
    ['Section', filters.section || 'All'],
    ['Machine', filters.machine || 'All']
  ];
}

function getOrCreateReportSheet_(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet) {
    return sheet;
  }

  return spreadsheet.insertSheet(sheetName);
}

function writeDailySummary_(sheet, dashboardData, filters) {
  var machineSummary = getMachinePerformanceSummary_(dashboardData.machineRows);
  var summaryRows = [
    ['Report Date', formatReportDateValue_(new Date())],
    ['Applied Filters', ''],
    ['Date From', formatReportDateValue_(filters.dateFrom)],
    ['Date To', formatReportDateValue_(filters.dateTo)],
    ['Section', filters.section || 'All'],
    ['Machine', filters.machine || 'All'],
    ['', ''],
    ['Total Target Output', dashboardData.totalTargetOutput],
    ['Total Actual Output', dashboardData.totalActualOutput],
    ['Overall Efficiency %', dashboardData.overallEfficiency],
    ['Average Downtime', dashboardData.averageDowntime],
    ['Total Machines', dashboardData.totalMachines],
    ['Total Sections', dashboardData.totalSections],
    ['', ''],
    ['Top Performing Machine', machineSummary.top],
    ['Lowest Performing Machine', machineSummary.lowest]
  ];

  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:D2').merge().setValue('DAILY MACHINE EFFICIENCY SUMMARY');
  sheet.getRange(4, 1, summaryRows.length, 2).setValues(summaryRows);
  formatDailySummarySheet_(sheet);
}

function formatDailySummarySheet_(sheet) {
  sheet
    .getRange('A1:D2')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet
    .getRange('A4:B19')
    .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');
  sheet
    .getRange('A4:A19')
    .setFontWeight('bold')
    .setFontColor('#334e68')
    .setBackground('#f8fafc');
  sheet.getRange('B13').setNumberFormat('0.00%');
  sheet.getRange('B11:B12').setNumberFormat('#,##0');
  sheet.getRange('B14').setNumberFormat('#,##0.00');
  sheet.getRange('B15:B16').setNumberFormat('#,##0');
  sheet.setRowHeights(1, 2, 34);
  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
}

function getMachinePerformanceSummary_(machineRows) {
  if (!machineRows || machineRows.length === 0) {
    return {
      top: 'No matching machine data',
      lowest: 'No matching machine data'
    };
  }

  var lowestMachine = machineRows[0];
  var topMachine = machineRows[machineRows.length - 1];

  return {
    top: topMachine[0] + ' (' + formatReportPercent_(topMachine[1]) + ')',
    lowest: lowestMachine[0] + ' (' + formatReportPercent_(lowestMachine[1]) + ')'
  };
}

function formatReportDateValue_(value) {
  if (!value) {
    return 'All';
  }

  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatReportPercent_(value) {
  return Utilities.formatString('%.2f%%', (Number(value) || 0) * 100);
}

function getReportTimestamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
}
