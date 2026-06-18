var DASHBOARD_SHEET_NAME = 'Dashboard';
var DAILY_PRODUCTION_SHEET_NAME = 'DailyProduction';

var DAILY_PRODUCTION_COLUMNS = {
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
  var dashboardData = calculateDashboardData_(productionRows);

  prepareDashboardSheet_(dashboardSheet);

  writeDashboardHeader_(dashboardSheet);
  writeKpiCards_(dashboardSheet, dashboardData);
  writeMachineEfficiencyTable_(dashboardSheet, dashboardData.machineRows);
  formatDashboard_(dashboardSheet, dashboardData);

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

function prepareDashboardSheet_(dashboardSheet) {
  dashboardSheet.setFrozenRows(0);
  dashboardSheet
    .getRange(1, 1, dashboardSheet.getMaxRows(), dashboardSheet.getMaxColumns())
    .breakApart();
  dashboardSheet.clear();
  dashboardSheet.setHiddenGridlines(true);
}

function writeDashboardHeader_(dashboardSheet) {
  dashboardSheet.getRange('A1:F2').merge().setValue('MACHINE EFFICIENCY DASHBOARD');
}

function writeKpiCards_(dashboardSheet, dashboardData) {
  var cards = [
    {
      label: 'Total Target Output',
      value: dashboardData.totalTargetOutput,
      range: 'A4:B5',
      valueRange: 'A5:B5',
      numberFormat: '#,##0'
    },
    {
      label: 'Total Actual Output',
      value: dashboardData.totalActualOutput,
      range: 'C4:D5',
      valueRange: 'C5:D5',
      numberFormat: '#,##0'
    },
    {
      label: 'Overall Efficiency %',
      value: dashboardData.overallEfficiency,
      range: 'E4:F5',
      valueRange: 'E5:F5',
      numberFormat: '0.00%'
    },
    {
      label: 'Average Downtime',
      value: dashboardData.averageDowntime,
      range: 'A7:B8',
      valueRange: 'A8:B8',
      numberFormat: '#,##0.00'
    },
    {
      label: 'Total Machines',
      value: dashboardData.totalMachines,
      range: 'C7:D8',
      valueRange: 'C8:D8',
      numberFormat: '#,##0'
    },
    {
      label: 'Total Sections',
      value: dashboardData.totalSections,
      range: 'E7:F8',
      valueRange: 'E8:F8',
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
      1,
      cardStart.getNumColumns()
    );
    var valueRange = dashboardSheet.getRange(card.valueRange);

    labelRange.merge().setValue(card.label);
    valueRange.merge().setValue(card.value).setNumberFormat(card.numberFormat);
  });
}

function writeMachineEfficiencyTable_(dashboardSheet, machineRows) {
  dashboardSheet.getRange('A11:C11').merge().setValue('MACHINE EFFICIENCY STATUS');
  dashboardSheet.getRange(12, 1, 1, 3).setValues([
    ['Machine', 'Efficiency %', 'Status']
  ]);

  if (machineRows.length > 0) {
    dashboardSheet.getRange(13, 1, machineRows.length, 3).setValues(machineRows);
  }
}

function formatDashboard_(dashboardSheet, dashboardData) {
  formatDashboardHeader_(dashboardSheet);
  formatKpiCards_(dashboardSheet, dashboardData.overallEfficiency);
  formatMachineTable_(dashboardSheet, dashboardData.machineRows.length);
  applyDashboardLayout_(dashboardSheet);
}

function formatDashboardHeader_(dashboardSheet) {
  dashboardSheet
    .getRange('A1:F2')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function formatKpiCards_(dashboardSheet, overallEfficiency) {
  var cardRanges = ['A4:B5', 'C4:D5', 'E4:F5', 'A7:B8', 'C7:D8', 'E7:F8'];
  var labelRanges = ['A4:B4', 'C4:D4', 'E4:F4', 'A7:B7', 'C7:D7', 'E7:F7'];
  var valueRanges = ['A5:B5', 'C5:D5', 'E5:F5', 'A8:B8', 'C8:D8', 'E8:F8'];

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
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  valueRanges.forEach(function(valueRange) {
    dashboardSheet
      .getRange(valueRange)
      .setFontWeight('bold')
      .setFontColor('#102a43')
      .setFontSize(16)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  var efficiencyColors = getEfficiencyCardColors_(overallEfficiency);
  dashboardSheet
    .getRange('E4:F5')
    .setBackground(efficiencyColors.background)
    .setFontColor(efficiencyColors.font);
  dashboardSheet.getRange('E4:F4').setFontColor(efficiencyColors.font);
  dashboardSheet.getRange('E5:F5').setFontColor(efficiencyColors.font);
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
    .getRange('A11:C11')
    .setBackground('#334e68')
    .setFontColor('#ffffff')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  dashboardSheet
    .getRange('A12:C12')
    .setBackground('#486581')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#bcccdc', SpreadsheetApp.BorderStyle.SOLID);

  if (machineCount > 0) {
    var dataRange = dashboardSheet.getRange(13, 1, machineCount, 3);

    dataRange
      .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('middle');

    dashboardSheet.getRange(13, 2, machineCount, 1).setNumberFormat('0.00%');
    dashboardSheet.getRange(13, 2, machineCount, 1).setHorizontalAlignment('center');
    dashboardSheet.getRange(13, 3, machineCount, 1).setHorizontalAlignment('center');
    applyMachineTableBanding_(dashboardSheet, machineCount);
    applyStatusColorCoding_(dashboardSheet, machineCount);
  }
}

function applyMachineTableBanding_(dashboardSheet, machineCount) {
  for (var rowOffset = 0; rowOffset < machineCount; rowOffset += 1) {
    var background = rowOffset % 2 === 0 ? '#ffffff' : '#f8fafc';
    dashboardSheet.getRange(13 + rowOffset, 1, 1, 3).setBackground(background);
  }
}

function applyStatusColorCoding_(dashboardSheet, machineCount) {
  var statuses = dashboardSheet.getRange(13, 3, machineCount, 1).getValues();

  statuses.forEach(function(statusRow, rowOffset) {
    var status = statusRow[0];
    var statusCell = dashboardSheet.getRange(13 + rowOffset, 3);

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
  dashboardSheet.setRowHeights(1, 2, 34);
  dashboardSheet.setRowHeights(4, 2, 34);
  dashboardSheet.setRowHeights(7, 2, 34);
  dashboardSheet.setRowHeight(10, 18);
  dashboardSheet.setRowHeight(11, 32);
  dashboardSheet.setRowHeight(12, 28);
  dashboardSheet.setColumnWidths(1, 6, 135);
  dashboardSheet.setColumnWidth(1, 180);
  dashboardSheet.setColumnWidth(2, 105);
  dashboardSheet.setColumnWidth(3, 150);
  dashboardSheet.setColumnWidth(4, 105);
  dashboardSheet.setColumnWidth(5, 150);
  dashboardSheet.setColumnWidth(6, 105);
}
