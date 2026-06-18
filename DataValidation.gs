var DATA_VALIDATION_SHEET_NAMES = {
  dailyProduction: 'DailyProduction',
  dashboard: 'Dashboard',
  correctiveActions: 'Corrective Actions',
  dailySummary: 'Daily Summary',
  rcaConfig: 'RCA Config',
  dataAudit: 'Data Audit'
};

var DATA_VALIDATION_HEADERS = [
  'Date',
  'Section',
  'Machine',
  'Target Output',
  'Actual Output',
  'Downtime Minutes',
  'Remarks'
];

var DATA_AUDIT_HEADERS = [
  'Row Number',
  'Issue Type',
  'Description',
  'Suggested Fix'
];

var PRODUCTION_SECTIONS = [
  'Assembly',
  'Cutting',
  'Packaging',
  'Inspection',
  'Finishing'
];

var PRODUCTION_MACHINES = [
  'Machine A1',
  'Machine A2',
  'Machine C1',
  'Machine F1',
  'Machine F2',
  'Machine I1',
  'Machine I2',
  'Machine P1',
  'Machine P2',
  'Machine P3'
];

var DATA_VALIDATION_COLUMNS = {
  date: 1,
  section: 2,
  machine: 3,
  targetOutput: 4,
  actualOutput: 5,
  downtimeMinutes: 6
};

var DATA_VALIDATION_REQUIRED_COLUMNS = [
  DATA_VALIDATION_COLUMNS.date,
  DATA_VALIDATION_COLUMNS.section,
  DATA_VALIDATION_COLUMNS.machine,
  DATA_VALIDATION_COLUMNS.targetOutput,
  DATA_VALIDATION_COLUMNS.actualOutput,
  DATA_VALIDATION_COLUMNS.downtimeMinutes
];

var DATA_VALIDATION_DATA_START_ROW = 2;
var DATA_VALIDATION_MANAGED_PROTECTION_DESCRIPTION = 'Machine Dashboard managed protection';

function setupDataValidation() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productionSheet = spreadsheet.getSheetByName(DATA_VALIDATION_SHEET_NAMES.dailyProduction);

  if (!productionSheet) {
    SpreadsheetApp.getUi().alert('DailyProduction sheet was not found. Please run Setup Sheets first.');
    return;
  }

  ensureProductionHeaders_(productionSheet);
  applyProductionDataValidation_(productionSheet);
  applyProductionConditionalFormatting_(productionSheet);
  formatProductionInputSheet_(productionSheet);
  protectManagedSheets_(spreadsheet);

  SpreadsheetApp.getUi().alert('Data validation, data quality checks, and sheet protection setup finished.');
}

function auditProductionData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productionSheet = spreadsheet.getSheetByName(DATA_VALIDATION_SHEET_NAMES.dailyProduction);
  var auditSheet = getOrCreateDataAuditSheet_(spreadsheet);

  prepareDataAuditSheet_(auditSheet);

  if (!productionSheet) {
    writeAuditRows_(auditSheet, [[
      '',
      'Missing Sheet',
      'DailyProduction sheet was not found.',
      'Run Setup Sheets, then run Setup Data Validation.'
    ]]);
    SpreadsheetApp.getUi().alert('Data audit finished with 1 issue.');
    return;
  }

  var lastRow = productionSheet.getLastRow();

  if (lastRow < DATA_VALIDATION_DATA_START_ROW) {
    writeAuditRows_(auditSheet, [[
      '',
      'Empty Data',
      'DailyProduction has no production records.',
      'Enter production records starting on row 2.'
    ]]);
    SpreadsheetApp.getUi().alert('Data audit finished with 1 issue.');
    return;
  }

  var productionRows = productionSheet
    .getRange(DATA_VALIDATION_DATA_START_ROW, 1, lastRow - 1, DATA_VALIDATION_HEADERS.length)
    .getValues();
  var auditRows = buildProductionAuditRows_(productionRows);

  if (auditRows.length > 0) {
    writeAuditRows_(auditSheet, auditRows);
  }

  SpreadsheetApp.getUi().alert('Data audit finished with ' + auditRows.length + ' issue(s).');
}

function applyProductionDataValidation_(sheet) {
  var rowCount = getValidationRowCount_(sheet);
  var dateRule = SpreadsheetApp
    .newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Enter a valid production date.')
    .build();
  var sectionRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(PRODUCTION_SECTIONS, true)
    .setAllowInvalid(false)
    .setHelpText('Select a valid production section.')
    .build();
  var machineRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(PRODUCTION_MACHINES, true)
    .setAllowInvalid(false)
    .setHelpText('Select a valid production machine.')
    .build();
  var targetRule = buildWholeNumberRule_('D2', '>0', 'Enter a positive production target.');
  var actualRule = buildWholeNumberRule_('E2', '>=0', 'Enter actual production output.');
  var downtimeRule = buildWholeNumberRule_('F2', '>=0', 'Enter downtime in minutes.');

  sheet.getRange(DATA_VALIDATION_DATA_START_ROW, DATA_VALIDATION_COLUMNS.date, rowCount, 1).setDataValidation(dateRule);
  sheet.getRange(DATA_VALIDATION_DATA_START_ROW, DATA_VALIDATION_COLUMNS.section, rowCount, 1).setDataValidation(sectionRule);
  sheet.getRange(DATA_VALIDATION_DATA_START_ROW, DATA_VALIDATION_COLUMNS.machine, rowCount, 1).setDataValidation(machineRule);
  sheet.getRange(DATA_VALIDATION_DATA_START_ROW, DATA_VALIDATION_COLUMNS.targetOutput, rowCount, 1).setDataValidation(targetRule);
  sheet.getRange(DATA_VALIDATION_DATA_START_ROW, DATA_VALIDATION_COLUMNS.actualOutput, rowCount, 1).setDataValidation(actualRule);
  sheet.getRange(DATA_VALIDATION_DATA_START_ROW, DATA_VALIDATION_COLUMNS.downtimeMinutes, rowCount, 1).setDataValidation(downtimeRule);
}

function buildWholeNumberRule_(cellReference, comparison, helpText) {
  return SpreadsheetApp
    .newDataValidation()
    .requireFormulaSatisfied('=AND(ISNUMBER(' + cellReference + '),' + cellReference + '=INT(' + cellReference + '),' + cellReference + comparison + ')')
    .setAllowInvalid(false)
    .setHelpText(helpText)
    .build();
}

function applyProductionConditionalFormatting_(sheet) {
  var rowCount = getValidationRowCount_(sheet);
  var qualityRange = sheet.getRange(DATA_VALIDATION_DATA_START_ROW, 1, rowCount, DATA_VALIDATION_HEADERS.length);
  var existingRules = sheet.getConditionalFormatRules().filter(function(rule) {
    return !isManagedProductionRule_(rule, sheet);
  });
  var missingDataRule = SpreadsheetApp
    .newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(COUNTA($A2:$G2)>0,OR($A2="",$B2="",$C2="",$D2="",$E2="",$F2=""))')
    .setBackground('#f4cccc')
    .setRanges([qualityRange])
    .build();
  var overTargetRule = SpreadsheetApp
    .newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($D2),ISNUMBER($E2),$E2>$D2)')
    .setBackground('#fff2cc')
    .setRanges([qualityRange])
    .build();

  existingRules.push(missingDataRule);
  existingRules.push(overTargetRule);
  sheet.setConditionalFormatRules(existingRules);
}

function isManagedProductionRule_(rule, sheet) {
  var ranges = rule.getRanges();
  var condition = rule.getBooleanCondition();
  var criteriaValues = condition ? condition.getCriteriaValues() : [];
  var formula = criteriaValues.length > 0 ? String(criteriaValues[0]) : '';

  if (
    formula !== '=OR($A2="",$B2="",$C2="",$D2="",$E2="",$F2="")' &&
    formula !== '=AND(COUNTA($A2:$G2)>0,OR($A2="",$B2="",$C2="",$D2="",$E2="",$F2=""))' &&
    formula !== '=AND(ISNUMBER($D2),ISNUMBER($E2),$E2>$D2)'
  ) {
    return false;
  }

  return ranges.some(function(range) {
    return range.getSheet().getName() === sheet.getName() &&
      range.getRow() === DATA_VALIDATION_DATA_START_ROW &&
      range.getColumn() === 1 &&
      range.getNumColumns() === DATA_VALIDATION_HEADERS.length;
  });
}

function formatProductionInputSheet_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, DATA_VALIDATION_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#d9eaf7')
    .setFontColor('#102a43')
    .setHorizontalAlignment('center');
  sheet.getRange('A:A').setNumberFormat('mmm d, yyyy');
  sheet.getRange('D:F').setNumberFormat('#,##0');
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), DATA_VALIDATION_HEADERS.length)
    .setVerticalAlignment('middle');
  sheet.autoResizeColumns(1, DATA_VALIDATION_HEADERS.length);
}

function protectManagedSheets_(spreadsheet) {
  [
    DATA_VALIDATION_SHEET_NAMES.dashboard,
    DATA_VALIDATION_SHEET_NAMES.correctiveActions,
    DATA_VALIDATION_SHEET_NAMES.dailySummary
  ].forEach(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);

    if (sheet) {
      protectSheetForScriptUse_(sheet);
    }
  });
}

function protectSheetForScriptUse_(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var protection = protections.length > 0 ? protections[0] : sheet.protect();

  protection
    .setDescription(DATA_VALIDATION_MANAGED_PROTECTION_DESCRIPTION + ': ' + sheet.getName())
    .setWarningOnly(false);

  try {
    protection.setDomainEdit(false);
  } catch (error) {
    // Domain edit settings are only available for some Google Workspace accounts.
  }

  protection.getEditors().forEach(function(editor) {
    try {
      protection.removeEditor(editor);
    } catch (error) {
      // The effective owner cannot always be removed from a protection.
    }
  });
}

function buildProductionAuditRows_(productionRows) {
  var auditRows = [];

  productionRows.forEach(function(row, rowOffset) {
    var sheetRowNumber = DATA_VALIDATION_DATA_START_ROW + rowOffset;
    var targetOutput = row[DATA_VALIDATION_COLUMNS.targetOutput - 1];
    var actualOutput = row[DATA_VALIDATION_COLUMNS.actualOutput - 1];
    var downtimeMinutes = row[DATA_VALIDATION_COLUMNS.downtimeMinutes - 1];

    DATA_VALIDATION_REQUIRED_COLUMNS.forEach(function(columnNumber) {
      if (isBlankValue_(row[columnNumber - 1])) {
        auditRows.push([
          sheetRowNumber,
          'Missing Field',
          DATA_VALIDATION_HEADERS[columnNumber - 1] + ' is blank.',
          'Enter a value for ' + DATA_VALIDATION_HEADERS[columnNumber - 1] + '.'
        ]);
      }
    });

    if (!isBlankValue_(row[DATA_VALIDATION_COLUMNS.date - 1]) && !isValidDateValue_(row[DATA_VALIDATION_COLUMNS.date - 1])) {
      auditRows.push([
        sheetRowNumber,
        'Invalid Date',
        'Date is not a valid date.',
        'Enter a valid production date.'
      ]);
    }

    addNumberAuditRows_(auditRows, sheetRowNumber, 'Target Output', targetOutput, true);
    addNumberAuditRows_(auditRows, sheetRowNumber, 'Actual Output', actualOutput, false);
    addNumberAuditRows_(auditRows, sheetRowNumber, 'Downtime Minutes', downtimeMinutes, false);

    if (isWholeNumber_(targetOutput) && isWholeNumber_(actualOutput) && Number(actualOutput) > Number(targetOutput)) {
      auditRows.push([
        sheetRowNumber,
        'Actual Above Target',
        'Actual Output is greater than Target Output.',
        'Confirm the production entry or update the target.'
      ]);
    }
  });

  return auditRows;
}

function addNumberAuditRows_(auditRows, rowNumber, fieldName, value, mustBeGreaterThanZero) {
  if (isBlankValue_(value)) {
    return;
  }

  if (!isWholeNumber_(value)) {
    auditRows.push([
      rowNumber,
      'Invalid Number',
      fieldName + ' must be a whole number.',
      'Replace it with a whole number.'
    ]);
    return;
  }

  if (mustBeGreaterThanZero && Number(value) <= 0) {
    auditRows.push([
      rowNumber,
      'Invalid Number',
      fieldName + ' must be greater than zero.',
      'Enter a positive production target.'
    ]);
    return;
  }

  if (!mustBeGreaterThanZero && Number(value) < 0) {
    auditRows.push([
      rowNumber,
      'Negative Value',
      fieldName + ' cannot be negative.',
      'Enter zero or a positive whole number.'
    ]);
  }
}

function getOrCreateDataAuditSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(DATA_VALIDATION_SHEET_NAMES.dataAudit);

  if (sheet) {
    return sheet;
  }

  return spreadsheet.insertSheet(DATA_VALIDATION_SHEET_NAMES.dataAudit);
}

function prepareDataAuditSheet_(sheet) {
  sheet.clear();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, DATA_AUDIT_HEADERS.length).setValues([DATA_AUDIT_HEADERS]);
  sheet.getRange(1, 1, 1, DATA_AUDIT_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, DATA_AUDIT_HEADERS.length);
}

function writeAuditRows_(sheet, auditRows) {
  sheet.getRange(2, 1, auditRows.length, DATA_AUDIT_HEADERS.length).setValues(auditRows);
  sheet.getRange(2, 1, auditRows.length, DATA_AUDIT_HEADERS.length).setVerticalAlignment('middle');
  sheet.autoResizeColumns(1, DATA_AUDIT_HEADERS.length);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 320);
}

function ensureProductionHeaders_(sheet) {
  var currentHeaders = sheet.getRange(1, 1, 1, DATA_VALIDATION_HEADERS.length).getValues()[0];
  var hasAnyHeader = currentHeaders.some(function(header) {
    return !isBlankValue_(header);
  });

  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, DATA_VALIDATION_HEADERS.length).setValues([DATA_VALIDATION_HEADERS]);
  }
}

function getValidationRowCount_(sheet) {
  return Math.max(sheet.getMaxRows() - DATA_VALIDATION_DATA_START_ROW + 1, 1);
}

function isBlankValue_(value) {
  return value === '' || value === null || typeof value === 'undefined';
}

function isValidDateValue_(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

function isWholeNumber_(value) {
  return typeof value === 'number' && isFinite(value) && Math.floor(value) === value;
}
