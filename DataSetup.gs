var SHEET_NAMES = {
  dailyProduction: 'DailyProduction',
  dashboard: 'Dashboard',
  correctiveActions: 'Corrective Actions',
  rcaConfig: 'RCA Config'
};

var DAILY_PRODUCTION_HEADERS = [
  'Date',
  'Section',
  'Machine',
  'Target Output',
  'Actual Output',
  'Downtime Minutes',
  'Remarks'
];

var CORRECTIVE_ACTION_HEADERS = [
  'Date',
  'Section',
  'Machine',
  'Target Output',
  'Actual Output',
  'Efficiency %',
  'Downtime Minutes',
  'Issue',
  'Possible Cause',
  'Recommended Corrective Action'
];

var RCA_CONFIG_HEADERS = [
  'Issue',
  'Possible Cause',
  'Recommended Corrective Action'
];

var DEFAULT_RCA_CONFIG_ROWS = [
  [
    'Low Efficiency',
    'Machine running below target',
    'Check machine setup and operator workflow'
  ],
  [
    'High Downtime',
    'Frequent machine stoppage',
    'Inspect machine condition and downtime logs'
  ],
  [
    'No Output',
    'Machine produced zero output',
    'Verify machine availability and production entry'
  ]
];

var SAMPLE_PRODUCTION_ROWS = [
  [new Date(2026, 5, 1), 'Assembly', 'Machine A1', 500, 470, 18, 'Minor setup delay'],
  [new Date(2026, 5, 2), 'Assembly', 'Machine A2', 520, 515, 5, 'Normal run'],
  [new Date(2026, 5, 3), 'Cutting', 'Machine C1', 450, 390, 42, 'Blade change needed'],
  [new Date(2026, 5, 4), 'Cutting', 'Machine C2', 460, 455, 8, 'Normal run'],
  [new Date(2026, 5, 5), 'Packaging', 'Machine P1', 600, 580, 12, 'Short material wait'],
  [new Date(2026, 5, 6), 'Packaging', 'Machine P2', 610, 0, 90, 'Machine unavailable'],
  [new Date(2026, 5, 7), 'Inspection', 'Machine I1', 350, 340, 7, 'Normal run'],
  [new Date(2026, 5, 8), 'Inspection', 'Machine I2', 360, 300, 35, 'Sensor alignment issue'],
  [new Date(2026, 5, 9), 'Finishing', 'Machine F1', 420, 415, 10, 'Normal run'],
  [new Date(2026, 5, 10), 'Finishing', 'Machine F2', 430, 385, 28, 'Operator workflow check']
];

function setupSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var dailyProductionSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.dailyProduction);
  var dashboardSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.dashboard);
  var correctiveActionsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.correctiveActions);
  var rcaConfigSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.rcaConfig);

  resetSheet_(dailyProductionSheet);
  resetSheet_(dashboardSheet);
  resetSheet_(correctiveActionsSheet);
  resetSheet_(rcaConfigSheet);

  setHeaders_(dailyProductionSheet, DAILY_PRODUCTION_HEADERS);
  setHeaders_(correctiveActionsSheet, CORRECTIVE_ACTION_HEADERS);
  setHeaders_(rcaConfigSheet, RCA_CONFIG_HEADERS);

  rcaConfigSheet
    .getRange(2, 1, DEFAULT_RCA_CONFIG_ROWS.length, RCA_CONFIG_HEADERS.length)
    .setValues(DEFAULT_RCA_CONFIG_ROWS);

  formatDailyProductionSheet_(dailyProductionSheet);
  formatCorrectiveActionsSheet_(correctiveActionsSheet);
  formatRcaConfigSheet_(rcaConfigSheet);
  formatEmptySheet_(dashboardSheet);

  SpreadsheetApp.getUi().alert('Sheets setup finished.');
}

function loadSampleData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dailyProductionSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.dailyProduction);

  setHeaders_(dailyProductionSheet, DAILY_PRODUCTION_HEADERS);
  clearRowsBelowHeader_(dailyProductionSheet);

  dailyProductionSheet
    .getRange(2, 1, SAMPLE_PRODUCTION_ROWS.length, DAILY_PRODUCTION_HEADERS.length)
    .setValues(SAMPLE_PRODUCTION_ROWS);

  formatDailyProductionSheet_(dailyProductionSheet);

  SpreadsheetApp.getUi().alert('Sample data loaded.');
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet) {
    return sheet;
  }

  return spreadsheet.insertSheet(sheetName);
}

function resetSheet_(sheet) {
  sheet.clearContents();
  sheet.setFrozenRows(1);
}

function setHeaders_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function clearRowsBelowHeader_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), DAILY_PRODUCTION_HEADERS.length);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
  }
}

function formatDailyProductionSheet_(sheet) {
  formatSheet_(sheet, DAILY_PRODUCTION_HEADERS.length);

  sheet.getRange('A:A').setNumberFormat('mmm d, yyyy');
  sheet.getRange('D:F').setNumberFormat('#,##0');
}

function formatCorrectiveActionsSheet_(sheet) {
  formatSheet_(sheet, CORRECTIVE_ACTION_HEADERS.length);

  sheet.getRange('A:A').setNumberFormat('mmm d, yyyy');
  sheet.getRange('D:E').setNumberFormat('#,##0');
  sheet.getRange('F:F').setNumberFormat('0.00%');
  sheet.getRange('G:G').setNumberFormat('#,##0');
}

function formatRcaConfigSheet_(sheet) {
  formatSheet_(sheet, RCA_CONFIG_HEADERS.length);
}

function formatEmptySheet_(sheet) {
  sheet.setFrozenRows(1);
}

function formatSheet_(sheet, headerCount) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headerCount).setFontWeight('bold');
  sheet.autoResizeColumns(1, headerCount);
}
