var CA_DAILY_PRODUCTION_SHEET_NAME = 'DailyProduction';
var CA_CORRECTIVE_ACTIONS_SHEET_NAME = 'Corrective Actions';
var CA_RCA_CONFIG_SHEET_NAME = 'RCA Config';

var CA_DAILY_PRODUCTION_COLUMNS = {
  date: 0,
  section: 1,
  machine: 2,
  targetOutput: 3,
  actualOutput: 4,
  downtimeMinutes: 5
};

var CA_HEADERS = [
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

var CA_FALLBACK_POSSIBLE_CAUSE = 'Needs review';
var CA_FALLBACK_RECOMMENDED_ACTION = 'Investigate production record and machine condition';

function generateCorrectiveActions() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productionSheet = spreadsheet.getSheetByName(CA_DAILY_PRODUCTION_SHEET_NAME);
  var rcaConfigSheet = spreadsheet.getSheetByName(CA_RCA_CONFIG_SHEET_NAME);
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

  SpreadsheetApp.getUi().alert(
    correctiveActionRows.length + ' corrective action row(s) generated.'
  );
}

function getOrCreateCorrectiveActionsSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(CA_CORRECTIVE_ACTIONS_SHEET_NAME);

  if (sheet) {
    return sheet;
  }

  return spreadsheet.insertSheet(CA_CORRECTIVE_ACTIONS_SHEET_NAME);
}

function getCorrectiveActionProductionRows_(productionSheet) {
  if (!productionSheet) {
    return [];
  }

  var lastRow = productionSheet.getLastRow();
  var lastColumn = productionSheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  return productionSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
}

function getRcaConfigMap_(rcaConfigSheet) {
  var rcaConfig = {};

  if (!rcaConfigSheet) {
    return rcaConfig;
  }

  var lastRow = rcaConfigSheet.getLastRow();
  var lastColumn = rcaConfigSheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return rcaConfig;
  }

  var configRows = rcaConfigSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  configRows.forEach(function(row) {
    var issue = String(row[0] || '').trim();

    if (issue) {
      rcaConfig[issue] = {
        possibleCause: row[1] || CA_FALLBACK_POSSIBLE_CAUSE,
        recommendedAction: row[2] || CA_FALLBACK_RECOMMENDED_ACTION
      };
    }
  });

  return rcaConfig;
}

function buildCorrectiveActionRows_(productionRows, rcaConfig) {
  var correctiveActionRows = [];

  productionRows.forEach(function(row) {
    var targetOutput = Number(row[CA_DAILY_PRODUCTION_COLUMNS.targetOutput]) || 0;
    var actualOutput = Number(row[CA_DAILY_PRODUCTION_COLUMNS.actualOutput]) || 0;
    var downtimeMinutes = Number(row[CA_DAILY_PRODUCTION_COLUMNS.downtimeMinutes]) || 0;
    var efficiency = calculateCorrectiveActionEfficiency_(actualOutput, targetOutput);

    if (efficiency >= 0.9) {
      return;
    }

    var issue = getCorrectiveActionIssue_(actualOutput, downtimeMinutes, efficiency);
    var rcaDetails = rcaConfig[issue] || {
      possibleCause: CA_FALLBACK_POSSIBLE_CAUSE,
      recommendedAction: CA_FALLBACK_RECOMMENDED_ACTION
    };

    correctiveActionRows.push([
      row[CA_DAILY_PRODUCTION_COLUMNS.date],
      row[CA_DAILY_PRODUCTION_COLUMNS.section],
      row[CA_DAILY_PRODUCTION_COLUMNS.machine],
      targetOutput,
      actualOutput,
      efficiency,
      downtimeMinutes,
      issue,
      rcaDetails.possibleCause,
      rcaDetails.recommendedAction
    ]);
  });

  return correctiveActionRows;
}

function calculateCorrectiveActionEfficiency_(actualOutput, targetOutput) {
  if (targetOutput <= 0) {
    return 0;
  }

  return actualOutput / targetOutput;
}

function getCorrectiveActionIssue_(actualOutput, downtimeMinutes, efficiency) {
  if (actualOutput === 0) {
    return 'No Output';
  }

  if (downtimeMinutes >= 30) {
    return 'High Downtime';
  }

  if (efficiency < 0.9) {
    return 'Low Efficiency';
  }

  return '';
}

function prepareCorrectiveActionsSheet_(sheet) {
  sheet.getRange(1, 1, 1, CA_HEADERS.length).setValues([CA_HEADERS]);
  clearCorrectiveActionRowsBelowHeader_(sheet);
}

function clearCorrectiveActionRowsBelowHeader_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), CA_HEADERS.length);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
  }
}

function formatCorrectiveActionsOutput_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, CA_HEADERS.length).setFontWeight('bold');
  sheet.getRange('A:A').setNumberFormat('mmm d, yyyy');
  sheet.getRange('D:E').setNumberFormat('#,##0');
  sheet.getRange('F:F').setNumberFormat('0.00%');
  sheet.getRange('G:G').setNumberFormat('#,##0');
  sheet.autoResizeColumns(1, CA_HEADERS.length);
  sheet.setColumnWidth(6, 95);
  sheet.setColumnWidth(7, 135);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 220);
  sheet.setColumnWidth(10, 350);
}
