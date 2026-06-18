var REPORT_AUTOMATION_SHEET_NAMES = {
  settings: 'Report Settings',
  history: 'Report History',
  dailyProduction: 'DailyProduction',
  dashboard: 'Dashboard',
  correctiveActions: 'Corrective Actions',
  rcaConfig: 'RCA Config',
  dailySummary: 'Daily Summary',
  historicalTrends: 'Historical Trends'
};

var REPORT_SETTINGS_HEADERS = [
  'Setting',
  'Value'
];

var REPORT_HISTORY_HEADERS = [
  'Timestamp',
  'Report Type',
  'Recipient',
  'Status',
  'File Name'
];

var DEFAULT_REPORT_SETTINGS_ROWS = [
  ['Report Recipient', 'manager@example.com'],
  ['CC Recipient', ''],
  ['Daily Report Enabled', 'No'],
  ['Weekly Report Enabled', 'No'],
  ['Last Report Sent', '']
];

var REPORT_RECIPIENT_SETTING = 'Report Recipient';
var REPORT_CC_SETTING = 'CC Recipient';
var DAILY_REPORT_ENABLED_SETTING = 'Daily Report Enabled';
var WEEKLY_REPORT_ENABLED_SETTING = 'Weekly Report Enabled';
var LAST_REPORT_SENT_SETTING = 'Last Report Sent';
var REPORT_PLACEHOLDER_RECIPIENT = 'manager@example.com';

function sendDailyReport() {
  sendConfiguredReport_('Daily', false);
}

function sendWeeklyReport() {
  sendConfiguredReport_('Weekly', false);
}

function runDailyScheduledReport() {
  sendConfiguredReport_('Daily', true);
}

function runWeeklyScheduledReport() {
  sendConfiguredReport_('Weekly', true);
}

function viewReportHistory() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = getOrCreateReportHistorySheet_(spreadsheet);

  formatReportHistorySheet_(historySheet);
  spreadsheet.setActiveSheet(historySheet);
}

function sendConfiguredReport_(reportType, isScheduledRun) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = getOrCreateReportSettingsSheet_(spreadsheet);
  var historySheet = getOrCreateReportHistorySheet_(spreadsheet);

  formatReportSettingsSheet_(settingsSheet);
  formatReportHistorySheet_(historySheet);

  var settings = getReportSettings_(settingsSheet);
  var enabledSetting = reportType === 'Weekly' ? WEEKLY_REPORT_ENABLED_SETTING : DAILY_REPORT_ENABLED_SETTING;

  if (isScheduledRun && !isReportEnabled_(settings[enabledSetting])) {
    appendReportHistory_(historySheet, reportType, settings[REPORT_RECIPIENT_SETTING] || '', 'Skipped: report disabled', '');
    return;
  }

  var validation = validateReportSendRequest_(spreadsheet, settings);

  if (!validation.isValid) {
    appendReportHistory_(historySheet, reportType, settings[REPORT_RECIPIENT_SETTING] || '', 'Failed: ' + validation.message, '');
    notifyReportUser_(validation.message, isScheduledRun);
    return;
  }

  try {
    var reportPackage = buildReportPackage_(spreadsheet, reportType);
    var subject = 'Machine Efficiency Report - ' + formatReportAutomationDate_(new Date());
    var body = buildReportEmailBody_(reportType, reportPackage.dashboardData, {
      correctiveActionCount: reportPackage.correctiveActionCount,
      trendDayCount: reportPackage.trendDayCount
    });
    var mailOptions = {
      attachments: reportPackage.files.map(function(file) {
        return file.getBlob();
      }),
      name: 'Machine Efficiency Dashboard'
    };
    var ccRecipient = String(settings[REPORT_CC_SETTING] || '').trim();

    if (ccRecipient) {
      mailOptions.cc = ccRecipient;
    }

    MailApp.sendEmail(settings[REPORT_RECIPIENT_SETTING], subject, body, mailOptions);

    var fileNames = reportPackage.files.map(function(file) {
      return file.getName();
    }).join(', ');

    appendReportHistory_(historySheet, reportType, settings[REPORT_RECIPIENT_SETTING], 'Sent', fileNames);
    updateLastReportSent_(settingsSheet, new Date());

    if (!isScheduledRun) {
      notifyReportUser_(reportType + ' report sent to ' + settings[REPORT_RECIPIENT_SETTING] + '.', isScheduledRun);
    }
  } catch (error) {
    appendReportHistory_(historySheet, reportType, settings[REPORT_RECIPIENT_SETTING] || '', 'Failed: ' + error.message, '');
    notifyReportUser_('Report could not be sent. ' + error.message, isScheduledRun);
  }
}

function buildReportPackage_(spreadsheet, reportType) {
  var dashboardSheet = spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.dashboard);
  var productionSheet = spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.dailyProduction);
  var correctiveActionsSheet = getOrCreateCorrectiveActionsSheet_(spreadsheet);
  var summarySheet = getOrCreateReportSheet_(spreadsheet, REPORT_AUTOMATION_SHEET_NAMES.dailySummary);
  var productionRows = getProductionRows_(productionSheet);
  var filters = getAutomationReportFilters_(dashboardSheet, productionRows);
  var filteredRows = applyProductionFilters(productionRows, filters);
  var dashboardData = calculateDashboardData_(filteredRows);
  var correctiveActionCount = rebuildCorrectiveActionsForReport_(spreadsheet);
  var trendDayCount = getHistoricalTrendDayCount_(spreadsheet);
  var timestamp = getReportTimestamp_();

  writeDailySummary_(summarySheet, dashboardData, filters);
  SpreadsheetApp.flush();

  var summaryFile = createPdfExport_(spreadsheet, summarySheet, 'Daily_Summary_Report_' + timestamp + '.pdf', false);
  var dashboardFile = createPdfExport_(spreadsheet, dashboardSheet, 'Machine_Efficiency_Dashboard_' + timestamp + '.pdf', true);
  var correctiveFile = createPdfExport_(spreadsheet, correctiveActionsSheet, 'Corrective_Actions_Report_' + timestamp + '.pdf', true);

  return {
    dashboardData: dashboardData,
    correctiveActionCount: correctiveActionCount,
    trendDayCount: trendDayCount,
    files: [
      summaryFile,
      dashboardFile,
      correctiveFile
    ]
  };
}

function validateReportSendRequest_(spreadsheet, settings) {
  var recipient = String(settings[REPORT_RECIPIENT_SETTING] || '').trim();

  if (!isValidReportRecipient_(recipient)) {
    return {
      isValid: false,
      message: 'Report Recipient is missing or still uses manager@example.com. Update Report Settings before sending.'
    };
  }

  if (!spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.dashboard)) {
    return {
      isValid: false,
      message: 'Dashboard sheet was not found. Run Setup Sheets and Refresh Dashboard first.'
    };
  }

  if (!spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.dailyProduction)) {
    return {
      isValid: false,
      message: 'DailyProduction sheet was not found. Run Setup Sheets first.'
    };
  }

  return {
    isValid: true,
    message: ''
  };
}

function getAutomationReportFilters_(dashboardSheet, productionRows) {
  try {
    var filters = readDashboardFilters(dashboardSheet);
    var validationResult = validateDashboardFilters(filters, productionRows);

    if (validationResult.isValid) {
      return validationResult.filters;
    }
  } catch (error) {
    return {
      dateFrom: null,
      dateTo: null,
      section: 'All',
      machine: 'All'
    };
  }

  return {
    dateFrom: null,
    dateTo: null,
    section: 'All',
    machine: 'All'
  };
}

function rebuildCorrectiveActionsForReport_(spreadsheet) {
  var productionSheet = spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.dailyProduction);
  var rcaConfigSheet = spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.rcaConfig);
  var correctiveActionsSheet = getOrCreateCorrectiveActionsSheet_(spreadsheet);
  var productionRows = getCorrectiveActionProductionRows_(productionSheet);
  var rcaConfig = getRcaConfigMap_(rcaConfigSheet);
  var correctiveActionRows = buildCorrectiveActionRows_(productionRows, rcaConfig);

  prepareCorrectiveActionsSheet_(correctiveActionsSheet);

  if (correctiveActionRows.length > 0) {
    correctiveActionsSheet
      .getRange(2, 1, correctiveActionRows.length, CA_HEADERS.length)
      .setValues(correctiveActionRows);
  }

  formatCorrectiveActionsOutput_(correctiveActionsSheet);

  return correctiveActionRows.length;
}

function getHistoricalTrendDayCount_(spreadsheet) {
  var trendsSheet = spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.historicalTrends);

  if (!trendsSheet) {
    return 0;
  }

  return Math.max(trendsSheet.getLastRow() - HISTORICAL_TRENDS_DATA_START_ROW + 1, 0);
}

function buildReportEmailBody_(reportType, dashboardData, reportDetails) {
  var machineSummary = getReportMachineSummary_(dashboardData.machineRows);
  var lines = [
    'Hello,',
    '',
    'Attached is the latest Machine Efficiency Report.',
    '',
    'Summary:',
    '',
    '- Report Type: ' + reportType,
    '- Overall Efficiency: ' + formatReportPercent_(dashboardData.overallEfficiency),
    '- Average Downtime: ' + formatReportNumber_(dashboardData.averageDowntime) + ' minutes',
    '- Total Machines: ' + dashboardData.totalMachines,
    '- Total Sections: ' + dashboardData.totalSections,
    '- Corrective Actions: ' + reportDetails.correctiveActionCount,
    '- Historical Trend Days: ' + reportDetails.trendDayCount,
    '- Top Performing Machine: ' + machineSummary.top,
    '- Lowest Performing Machine: ' + machineSummary.lowest,
    '',
    'Please review the attached reports.',
    '',
    'Regards,',
    'Machine Efficiency Dashboard'
  ];

  return lines.join('\n');
}

function getReportMachineSummary_(machineRows) {
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

function getOrCreateReportSettingsSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.settings);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(REPORT_AUTOMATION_SHEET_NAMES.settings);
  }

  ensureReportSettingsRows_(sheet);

  return sheet;
}

function ensureReportSettingsRows_(sheet) {
  var currentValues = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, REPORT_SETTINGS_HEADERS.length).getValues()
    : [];
  var settings = buildReportSettingsMap_(currentValues);

  sheet.getRange(1, 1, 1, REPORT_SETTINGS_HEADERS.length).setValues([REPORT_SETTINGS_HEADERS]);

  DEFAULT_REPORT_SETTINGS_ROWS.forEach(function(defaultRow) {
    if (typeof settings[defaultRow[0]] === 'undefined') {
      sheet.appendRow(defaultRow);
    }
  });
}

function getOrCreateReportHistorySheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(REPORT_AUTOMATION_SHEET_NAMES.history);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(REPORT_AUTOMATION_SHEET_NAMES.history);
  }

  sheet.getRange(1, 1, 1, REPORT_HISTORY_HEADERS.length).setValues([REPORT_HISTORY_HEADERS]);

  return sheet;
}

function getReportSettings_(settingsSheet) {
  var lastRow = settingsSheet.getLastRow();

  if (lastRow < 2) {
    return {};
  }

  return buildReportSettingsMap_(settingsSheet.getRange(2, 1, lastRow - 1, REPORT_SETTINGS_HEADERS.length).getValues());
}

function buildReportSettingsMap_(settingsRows) {
  var settings = {};

  settingsRows.forEach(function(row) {
    var setting = String(row[0] || '').trim();

    if (setting) {
      settings[setting] = row[1];
    }
  });

  return settings;
}

function appendReportHistory_(historySheet, reportType, recipient, status, fileName) {
  historySheet.appendRow([
    new Date(),
    reportType,
    recipient,
    status,
    fileName
  ]);
  formatReportHistorySheet_(historySheet);
}

function updateLastReportSent_(settingsSheet, sentAt) {
  var lastRow = settingsSheet.getLastRow();
  var settings = lastRow > 1
    ? settingsSheet.getRange(2, 1, lastRow - 1, REPORT_SETTINGS_HEADERS.length).getValues()
    : [];

  for (var rowOffset = 0; rowOffset < settings.length; rowOffset += 1) {
    if (settings[rowOffset][0] === LAST_REPORT_SENT_SETTING) {
      settingsSheet.getRange(rowOffset + 2, 2).setValue(formatReportAutomationDateTime_(sentAt));
      return;
    }
  }

  settingsSheet.appendRow([LAST_REPORT_SENT_SETTING, formatReportAutomationDateTime_(sentAt)]);
}

function formatReportSettingsSheet_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, REPORT_SETTINGS_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, REPORT_SETTINGS_HEADERS.length);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 260);
}

function formatReportHistorySheet_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, REPORT_HISTORY_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.getRange('A:A').setNumberFormat('mmm d, yyyy h:mm AM/PM');
  sheet.autoResizeColumns(1, REPORT_HISTORY_HEADERS.length);
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 420);
}

function isReportEnabled_(value) {
  return String(value || '').trim().toLowerCase() === 'yes';
}

function isValidReportRecipient_(value) {
  var recipient = String(value || '').trim();

  return recipient !== '' &&
    recipient !== REPORT_PLACEHOLDER_RECIPIENT &&
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient);
}

function formatReportAutomationDate_(value) {
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatReportAutomationDateTime_(value) {
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function formatReportNumber_(value) {
  return Utilities.formatString('%.2f', Number(value) || 0);
}

function notifyReportUser_(message, isScheduledRun) {
  if (isScheduledRun) {
    return;
  }

  SpreadsheetApp.getUi().alert(message);
}
