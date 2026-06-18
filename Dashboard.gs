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

  dashboardSheet.clearContents();

  writeKpis_(dashboardSheet, dashboardData);
  writeMachineEfficiencyTable_(dashboardSheet, dashboardData.machineRows);
  formatDashboard_(dashboardSheet, dashboardData.machineRows.length);

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

function writeKpis_(dashboardSheet, dashboardData) {
  dashboardSheet.getRange('A1').setValue('Machine Efficiency Dashboard');

  dashboardSheet.getRange(3, 1, 6, 2).setValues([
    ['Total Target Output', dashboardData.totalTargetOutput],
    ['Total Actual Output', dashboardData.totalActualOutput],
    ['Overall Efficiency %', dashboardData.overallEfficiency],
    ['Average Downtime', dashboardData.averageDowntime],
    ['Total Machines', dashboardData.totalMachines],
    ['Total Sections', dashboardData.totalSections]
  ]);
}

function writeMachineEfficiencyTable_(dashboardSheet, machineRows) {
  dashboardSheet.getRange(12, 1, 1, 3).setValues([
    ['Machine', 'Efficiency %', 'Status']
  ]);

  if (machineRows.length > 0) {
    dashboardSheet.getRange(13, 1, machineRows.length, 3).setValues(machineRows);
  }
}

function formatDashboard_(dashboardSheet, machineCount) {
  dashboardSheet.getRange('A1').setFontWeight('bold');
  dashboardSheet.getRange('A3:A8').setFontWeight('bold');
  dashboardSheet.getRange('A12:C12').setFontWeight('bold');

  dashboardSheet.getRange('B5').setNumberFormat('0.00%');
  dashboardSheet.getRange('B3:B4').setNumberFormat('#,##0');
  dashboardSheet.getRange('B6').setNumberFormat('#,##0.00');
  dashboardSheet.getRange('B7:B8').setNumberFormat('#,##0');

  if (machineCount > 0) {
    dashboardSheet.getRange(13, 2, machineCount, 1).setNumberFormat('0.00%');
  }

  dashboardSheet.autoResizeColumns(1, 3);
}
