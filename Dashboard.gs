var DASHBOARD_SHEET_NAME = 'Dashboard';
var DAILY_PRODUCTION_SHEET_NAME = 'DailyProduction';

var DAILY_PRODUCTION_COLUMNS = {
  date: 0,
  section: 1,
  machine: 2,
  targetOutput: 3,
  actualOutput: 4,
  downtimeMinutes: 5
};

function refreshDashboard() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productionSheet = spreadsheet.getSheetByName(DAILY_PRODUCTION_SHEET_NAME);
  var dashboardSheet = getOrCreateDashboardSheet_(spreadsheet);
  var productionRows = getProductionRows_(productionSheet);
  var dashboardFilters = readDashboardFilters(dashboardSheet);
  var validationResult = validateDashboardFilters(dashboardFilters, productionRows);

  if (!validationResult.isValid) {
    SpreadsheetApp.getUi().alert(validationResult.message);
    return;
  }

  dashboardFilters = validationResult.filters;

  var filteredProductionRows = applyProductionFilters(productionRows, dashboardFilters);
  var dashboardData = calculateDashboardData_(filteredProductionRows);

  prepareDashboardSheet_(dashboardSheet);

  writeDashboardHeader_(dashboardSheet);
  writeDashboardFilters(dashboardSheet, dashboardFilters, productionRows);
  writeKpiCards_(dashboardSheet, dashboardData);
  writeMachineEfficiencyTable_(dashboardSheet, dashboardData.machineRows);
  formatDashboard_(dashboardSheet, dashboardData);
  refreshDashboardCharts(dashboardSheet, filteredProductionRows, dashboardData);

  SpreadsheetApp.getUi().alert('Dashboard refreshed.');
}

function getOrCreateDashboardSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME);

  if (sheet) {
    return sheet;
  }

  return spreadsheet.insertSheet(DASHBOARD_SHEET_NAME);
}

function getProductionRows_(productionSheet) {
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

function calculateDashboardData_(productionRows) {
  var totalTargetOutput = 0;
  var totalActualOutput = 0;
  var totalDowntime = 0;
  var machines = {};
  var sections = {};

  productionRows.forEach(function(row) {
    var section = row[DAILY_PRODUCTION_COLUMNS.section];
    var machine = row[DAILY_PRODUCTION_COLUMNS.machine];
    var targetOutput = Number(row[DAILY_PRODUCTION_COLUMNS.targetOutput]) || 0;
    var actualOutput = Number(row[DAILY_PRODUCTION_COLUMNS.actualOutput]) || 0;
    var downtimeMinutes = Number(row[DAILY_PRODUCTION_COLUMNS.downtimeMinutes]) || 0;

    totalTargetOutput += targetOutput;
    totalActualOutput += actualOutput;
    totalDowntime += downtimeMinutes;

    if (section) {
      sections[section] = true;
    }

    if (machine) {
      if (!machines[machine]) {
        machines[machine] = {
          targetOutput: 0,
          actualOutput: 0
        };
      }

      machines[machine].targetOutput += targetOutput;
      machines[machine].actualOutput += actualOutput;
    }
  });

  var machineRows = Object.keys(machines).map(function(machine) {
    var machineTotals = machines[machine];
    var efficiency = calculateEfficiency_(machineTotals.actualOutput, machineTotals.targetOutput);

    return [
      machine,
      efficiency,
      getMachineStatus_(efficiency)
    ];
  });

  machineRows.sort(function(firstMachine, secondMachine) {
    return firstMachine[1] - secondMachine[1];
  });

  return {
    totalTargetOutput: totalTargetOutput,
    totalActualOutput: totalActualOutput,
    overallEfficiency: calculateEfficiency_(totalActualOutput, totalTargetOutput),
    averageDowntime: productionRows.length > 0 ? totalDowntime / productionRows.length : 0,
    totalMachines: Object.keys(machines).length,
    totalSections: Object.keys(sections).length,
    machineRows: machineRows
  };
}

function calculateEfficiency_(actualOutput, targetOutput) {
  if (targetOutput <= 0) {
    return 0;
  }

  return actualOutput / targetOutput;
}

function getMachineStatus_(efficiency) {
  if (efficiency >= 0.9) {
    return 'Normal';
  }

  if (efficiency >= 0.8) {
    return 'Warning';
  }

  return 'Critical';
}

function readDashboardFilters(dashboardSheet) {
  return {
    dateFromValue: dashboardSheet.getRange('B4').getValue(),
    dateToValue: dashboardSheet.getRange('D4').getValue(),
    section: dashboardSheet.getRange('F4').getValue(),
    machine: dashboardSheet.getRange('H4').getValue()
  };
}

function validateDashboardFilters(filters, productionRows) {
  var dateFrom = parseDashboardDate_(filters.dateFromValue);
  var dateTo = parseDashboardDate_(filters.dateToValue);

  if (dateFrom.invalid) {
    return {
      isValid: false,
      message: 'Date From is invalid. Please enter a valid date.'
    };
  }

  if (dateTo.invalid) {
    return {
      isValid: false,
      message: 'Date To is invalid. Please enter a valid date.'
    };
  }

  if (dateFrom.date && dateTo.date && dateFrom.date.getTime() > dateTo.date.getTime()) {
    return {
      isValid: false,
      message: 'Date From cannot be later than Date To.'
    };
  }

  return {
    isValid: true,
    filters: {
      dateFrom: dateFrom.date,
      dateTo: dateTo.date,
      section: normalizeFilterOption_(filters.section, getUniqueProductionValues_(productionRows, DAILY_PRODUCTION_COLUMNS.section)),
      machine: normalizeFilterOption_(filters.machine, getUniqueProductionValues_(productionRows, DAILY_PRODUCTION_COLUMNS.machine))
    }
  };
}

function applyProductionFilters(productionRows, filters) {
  return productionRows.filter(function(row) {
    var rowDate = parseDashboardDate_(row[DAILY_PRODUCTION_COLUMNS.date]).date;
    var rowSection = row[DAILY_PRODUCTION_COLUMNS.section];
    var rowMachine = row[DAILY_PRODUCTION_COLUMNS.machine];

    if (filters.dateFrom && (!rowDate || rowDate.getTime() < filters.dateFrom.getTime())) {
      return false;
    }

    if (filters.dateTo && (!rowDate || rowDate.getTime() > filters.dateTo.getTime())) {
      return false;
    }

    if (filters.section !== 'All' && rowSection !== filters.section) {
      return false;
    }

    if (filters.machine !== 'All' && rowMachine !== filters.machine) {
      return false;
    }

    return true;
  });
}

function writeDashboardFilters(dashboardSheet, filters, productionRows) {
  dashboardSheet.getRange('A4').setValue('Date From');
  dashboardSheet.getRange('B4').setValue(filters.dateFrom || '');
  dashboardSheet.getRange('C4').setValue('Date To');
  dashboardSheet.getRange('D4').setValue(filters.dateTo || '');
  dashboardSheet.getRange('E4').setValue('Section');
  dashboardSheet.getRange('F4').setValue(filters.section);
  dashboardSheet.getRange('G4').setValue('Machine');
  dashboardSheet.getRange('H4').setValue(filters.machine);

  buildFilterDropdowns(dashboardSheet, productionRows);
  formatDashboardFilters_(dashboardSheet);
}

function buildFilterDropdowns(dashboardSheet, productionRows) {
  var sectionOptions = ['All'].concat(getUniqueProductionValues_(productionRows, DAILY_PRODUCTION_COLUMNS.section));
  var machineOptions = ['All'].concat(getUniqueProductionValues_(productionRows, DAILY_PRODUCTION_COLUMNS.machine));
  var sectionRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(sectionOptions, true)
    .setAllowInvalid(false)
    .build();
  var machineRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(machineOptions, true)
    .setAllowInvalid(false)
    .build();

  dashboardSheet.getRange('F4').setDataValidation(sectionRule);
  dashboardSheet.getRange('H4').setDataValidation(machineRule);
}

function parseDashboardDate_(value) {
  if (!value) {
    return {
      date: null,
      invalid: false
    };
  }

  var parsedDate = value instanceof Date ? value : new Date(value);

  if (isNaN(parsedDate.getTime())) {
    return {
      date: null,
      invalid: true
    };
  }

  return {
    date: new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate()),
    invalid: false
  };
}

function normalizeFilterOption_(value, validOptions) {
  var normalizedValue = String(value || '').trim();

  if (!normalizedValue || normalizedValue === 'All') {
    return 'All';
  }

  return validOptions.indexOf(normalizedValue) >= 0 ? normalizedValue : 'All';
}

function getUniqueProductionValues_(productionRows, columnIndex) {
  var values = {};

  productionRows.forEach(function(row) {
    var value = row[columnIndex];

    if (value) {
      values[value] = true;
    }
  });

  return Object.keys(values).sort();
}

function prepareDashboardSheet_(dashboardSheet) {
  dashboardSheet.setFrozenRows(0);
  dashboardSheet
    .getRange(1, 1, dashboardSheet.getMaxRows(), dashboardSheet.getMaxColumns())
    .breakApart();
  dashboardSheet.clear();
  dashboardSheet.setHiddenGridlines(true);
}

function writeDashboardHeader_(dashboardSheet) {
  dashboardSheet.getRange('A1:L2').merge().setValue('MACHINE EFFICIENCY DASHBOARD');
}

function writeKpiCards_(dashboardSheet, dashboardData) {
  var cards = [
    {
      label: 'Total Target Output',
      value: dashboardData.totalTargetOutput,
      range: 'A6:B9',
      numberFormat: '#,##0'
    },
    {
      label: 'Total Actual Output',
      value: dashboardData.totalActualOutput,
      range: 'C6:D9',
      numberFormat: '#,##0'
    },
    {
      label: 'Overall Efficiency %',
      value: dashboardData.overallEfficiency,
      range: 'E6:F9',
      numberFormat: '0.00%'
    },
    {
      label: 'Average Downtime',
      value: dashboardData.averageDowntime,
      range: 'G6:H9',
      numberFormat: '#,##0.00'
    },
    {
      label: 'Total Machines',
      value: dashboardData.totalMachines,
      range: 'I6:J9',
      numberFormat: '#,##0'
    },
    {
      label: 'Total Sections',
      value: dashboardData.totalSections,
      range: 'K6:L9',
      numberFormat: '#,##0'
    }
  ];

  cards.forEach(function(card) {
    dashboardSheet.getRange(card.range).breakApart();
    dashboardSheet.getRange(card.range).setBorder(true, true, true, true, true, true);
    dashboardSheet.getRange(card.range).setVerticalAlignment('middle');

    var cardStart = dashboardSheet.getRange(card.range);
    var labelRange = dashboardSheet.getRange(
      cardStart.getRow(),
      cardStart.getColumn(),
      2,
      cardStart.getNumColumns()
    );
    var valueRange = dashboardSheet.getRange(
      cardStart.getRow() + 2,
      cardStart.getColumn(),
      2,
      cardStart.getNumColumns()
    );

    labelRange.merge().setValue(card.label);
    valueRange.merge().setValue(card.value).setNumberFormat(card.numberFormat);
  });
}

function writeMachineEfficiencyTable_(dashboardSheet, machineRows) {
  dashboardSheet.getRange('B12:K12').merge().setValue('MACHINE EFFICIENCY STATUS');
  dashboardSheet.getRange('B13:E13').merge().setValue('Machine');
  dashboardSheet.getRange('F13:H13').merge().setValue('Efficiency %');
  dashboardSheet.getRange('I13:K13').merge().setValue('Status');

  if (machineRows.length > 0) {
    machineRows.forEach(function(machineRow, rowOffset) {
      var rowNumber = 14 + rowOffset;

      dashboardSheet.getRange(rowNumber, 2, 1, 4).merge().setValue(machineRow[0]);
      dashboardSheet.getRange(rowNumber, 6, 1, 3).merge().setValue(machineRow[1]);
      dashboardSheet.getRange(rowNumber, 9, 1, 3).merge().setValue(machineRow[2]);
    });
  }
}

function formatDashboard_(dashboardSheet, dashboardData) {
  formatDashboardHeader_(dashboardSheet);
  formatKpiCards_(dashboardSheet, dashboardData.overallEfficiency);
  formatMachineTable_(dashboardSheet, dashboardData.machineRows.length);
  applyDashboardLayout_(dashboardSheet);
}

function formatDashboardFilters_(dashboardSheet) {
  dashboardSheet
    .getRange('A4:H4')
    .setBackground('#f8fafc')
    .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');

  dashboardSheet
    .getRangeList(['A4', 'C4', 'E4', 'G4'])
    .setFontWeight('bold')
    .setFontColor('#334e68')
    .setHorizontalAlignment('center');

  dashboardSheet
    .getRangeList(['B4', 'D4', 'F4', 'H4'])
    .setBackground('#ffffff')
    .setFontColor('#102a43')
    .setHorizontalAlignment('center');

  dashboardSheet.getRangeList(['B4', 'D4']).setNumberFormat('mmm d, yyyy');
}

function formatDashboardHeader_(dashboardSheet) {
  dashboardSheet
    .getRange('A1:L2')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function formatKpiCards_(dashboardSheet, overallEfficiency) {
  var cardRanges = ['A6:B9', 'C6:D9', 'E6:F9', 'G6:H9', 'I6:J9', 'K6:L9'];
  var labelRanges = ['A6:B7', 'C6:D7', 'E6:F7', 'G6:H7', 'I6:J7', 'K6:L7'];
  var valueRanges = ['A8:B9', 'C8:D9', 'E8:F9', 'G8:H9', 'I8:J9', 'K8:L9'];

  cardRanges.forEach(function(cardRange) {
    dashboardSheet
      .getRange(cardRange)
      .setBackground('#f8fafc')
      .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID);
  });

  labelRanges.forEach(function(labelRange) {
    dashboardSheet
      .getRange(labelRange)
      .setFontWeight('bold')
      .setFontColor('#334e68')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  valueRanges.forEach(function(valueRange) {
    dashboardSheet
      .getRange(valueRange)
      .setFontWeight('bold')
      .setFontColor('#102a43')
      .setFontSize(19)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  var efficiencyColors = getEfficiencyCardColors_(overallEfficiency);
  dashboardSheet
    .getRange('E6:F9')
    .setBackground(efficiencyColors.background)
    .setFontColor(efficiencyColors.font);
  dashboardSheet.getRange('E6:F7').setFontColor(efficiencyColors.font);
  dashboardSheet.getRange('E8:F9').setFontColor(efficiencyColors.font);
}

function getEfficiencyCardColors_(overallEfficiency) {
  if (overallEfficiency >= 0.9) {
    return {
      background: '#2e7d32',
      font: '#ffffff'
    };
  }

  if (overallEfficiency >= 0.8) {
    return {
      background: '#f9ab00',
      font: '#1f2933'
    };
  }

  return {
    background: '#c62828',
    font: '#ffffff'
  };
}

function formatMachineTable_(dashboardSheet, machineCount) {
  dashboardSheet
    .getRange('B12:K12')
    .setBackground('#334e68')
    .setFontColor('#ffffff')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  dashboardSheet
    .getRange('B13:K13')
    .setBackground('#486581')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#bcccdc', SpreadsheetApp.BorderStyle.SOLID);

  if (machineCount > 0) {
    var dataRange = dashboardSheet.getRange(14, 2, machineCount, 10);

    dataRange
      .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('middle');

    dashboardSheet.getRange(14, 6, machineCount, 3).setNumberFormat('0.00%');
    dashboardSheet.getRange(14, 2, machineCount, 10).setHorizontalAlignment('center');
    applyMachineTableBanding_(dashboardSheet, machineCount);
    applyStatusColorCoding_(dashboardSheet, machineCount);
  }
}

function applyMachineTableBanding_(dashboardSheet, machineCount) {
  for (var rowOffset = 0; rowOffset < machineCount; rowOffset += 1) {
    var background = rowOffset % 2 === 0 ? '#ffffff' : '#f8fafc';
    dashboardSheet.getRange(14 + rowOffset, 2, 1, 10).setBackground(background);
  }
}

function applyStatusColorCoding_(dashboardSheet, machineCount) {
  var statuses = dashboardSheet.getRange(14, 9, machineCount, 1).getValues();

  statuses.forEach(function(statusRow, rowOffset) {
    var status = statusRow[0];
    var statusCell = dashboardSheet.getRange(14 + rowOffset, 9, 1, 3);

    if (status === 'Normal') {
      statusCell.setBackground('#2e7d32').setFontColor('#ffffff').setFontWeight('bold');
      return;
    }

    if (status === 'Warning') {
      statusCell.setBackground('#f9ab00').setFontColor('#1f2933').setFontWeight('bold');
      return;
    }

    if (status === 'Critical') {
      statusCell.setBackground('#c62828').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
}

function applyDashboardLayout_(dashboardSheet) {
  dashboardSheet.setFrozenRows(0);
  dashboardSheet.setRowHeights(1, 2, 38);
  dashboardSheet.setRowHeight(4, 30);
  dashboardSheet.setRowHeight(5, 18);
  dashboardSheet.setRowHeights(6, 4, 34);
  dashboardSheet.setRowHeight(11, 20);
  dashboardSheet.setRowHeight(12, 34);
  dashboardSheet.setRowHeight(13, 30);
  dashboardSheet.setColumnWidths(1, 12, 95);
  dashboardSheet.setColumnWidth(1, 110);
  dashboardSheet.setColumnWidth(2, 110);
  dashboardSheet.setColumnWidth(3, 110);
  dashboardSheet.setColumnWidth(4, 110);
  dashboardSheet.setColumnWidth(5, 120);
  dashboardSheet.setColumnWidth(6, 110);
  dashboardSheet.setColumnWidth(7, 115);
  dashboardSheet.setColumnWidth(8, 115);
  dashboardSheet.setColumnWidth(9, 110);
  dashboardSheet.setColumnWidth(10, 110);
  dashboardSheet.setColumnWidth(11, 105);
  dashboardSheet.setColumnWidth(12, 105);
}
