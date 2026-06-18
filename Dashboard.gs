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
  dashboardSheet.getRange('A1:L2').merge().setValue('MACHINE EFFICIENCY DASHBOARD');
}

function writeKpiCards_(dashboardSheet, dashboardData) {
  var cards = [
    {
      label: 'Total Target Output',
      value: dashboardData.totalTargetOutput,
      range: 'A4:B7',
      numberFormat: '#,##0'
    },
    {
      label: 'Total Actual Output',
      value: dashboardData.totalActualOutput,
      range: 'C4:D7',
      numberFormat: '#,##0'
    },
    {
      label: 'Overall Efficiency %',
      value: dashboardData.overallEfficiency,
      range: 'E4:F7',
      numberFormat: '0.00%'
    },
    {
      label: 'Average Downtime',
      value: dashboardData.averageDowntime,
      range: 'G4:H7',
      numberFormat: '#,##0.00'
    },
    {
      label: 'Total Machines',
      value: dashboardData.totalMachines,
      range: 'I4:J7',
      numberFormat: '#,##0'
    },
    {
      label: 'Total Sections',
      value: dashboardData.totalSections,
      range: 'K4:L7',
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
  dashboardSheet.getRange('B10:K10').merge().setValue('MACHINE EFFICIENCY STATUS');
  dashboardSheet.getRange('B11:E11').merge().setValue('Machine');
  dashboardSheet.getRange('F11:H11').merge().setValue('Efficiency %');
  dashboardSheet.getRange('I11:K11').merge().setValue('Status');

  if (machineRows.length > 0) {
    machineRows.forEach(function(machineRow, rowOffset) {
      var rowNumber = 12 + rowOffset;

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
  var cardRanges = ['A4:B7', 'C4:D7', 'E4:F7', 'G4:H7', 'I4:J7', 'K4:L7'];
  var labelRanges = ['A4:B5', 'C4:D5', 'E4:F5', 'G4:H5', 'I4:J5', 'K4:L5'];
  var valueRanges = ['A6:B7', 'C6:D7', 'E6:F7', 'G6:H7', 'I6:J7', 'K6:L7'];

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
    .getRange('E4:F7')
    .setBackground(efficiencyColors.background)
    .setFontColor(efficiencyColors.font);
  dashboardSheet.getRange('E4:F5').setFontColor(efficiencyColors.font);
  dashboardSheet.getRange('E6:F7').setFontColor(efficiencyColors.font);
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
    .getRange('B10:K10')
    .setBackground('#334e68')
    .setFontColor('#ffffff')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  dashboardSheet
    .getRange('B11:K11')
    .setBackground('#486581')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#bcccdc', SpreadsheetApp.BorderStyle.SOLID);

  if (machineCount > 0) {
    var dataRange = dashboardSheet.getRange(12, 2, machineCount, 10);

    dataRange
      .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('middle');

    dashboardSheet.getRange(12, 6, machineCount, 3).setNumberFormat('0.00%');
    dashboardSheet.getRange(12, 2, machineCount, 10).setHorizontalAlignment('center');
    applyMachineTableBanding_(dashboardSheet, machineCount);
    applyStatusColorCoding_(dashboardSheet, machineCount);
  }
}

function applyMachineTableBanding_(dashboardSheet, machineCount) {
  for (var rowOffset = 0; rowOffset < machineCount; rowOffset += 1) {
    var background = rowOffset % 2 === 0 ? '#ffffff' : '#f8fafc';
    dashboardSheet.getRange(12 + rowOffset, 2, 1, 10).setBackground(background);
  }
}

function applyStatusColorCoding_(dashboardSheet, machineCount) {
  var statuses = dashboardSheet.getRange(12, 9, machineCount, 1).getValues();

  statuses.forEach(function(statusRow, rowOffset) {
    var status = statusRow[0];
    var statusCell = dashboardSheet.getRange(12 + rowOffset, 9, 1, 3);

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
  dashboardSheet.setRowHeights(4, 4, 34);
  dashboardSheet.setRowHeight(9, 20);
  dashboardSheet.setRowHeight(10, 34);
  dashboardSheet.setRowHeight(11, 30);
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
