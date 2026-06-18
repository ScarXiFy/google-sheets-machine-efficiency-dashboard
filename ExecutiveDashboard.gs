var EXECUTIVE_DASHBOARD_SHEET_NAME = 'Executive Dashboard';
var EXECUTIVE_DASHBOARD_DAILY_PRODUCTION_SHEET_NAME = 'DailyProduction';
var EXECUTIVE_DASHBOARD_DASHBOARD_SHEET_NAME = 'Dashboard';
var EXECUTIVE_DASHBOARD_CORRECTIVE_ACTIONS_SHEET_NAME = 'Corrective Actions';

var EXECUTIVE_HELPER_START_COLUMN = 16;
var EXECUTIVE_HELPER_COLUMN_COUNT = 5;

function refreshExecutiveDashboard() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productionSheet = spreadsheet.getSheetByName(EXECUTIVE_DASHBOARD_DAILY_PRODUCTION_SHEET_NAME);
  var dashboardSheet = spreadsheet.getSheetByName(EXECUTIVE_DASHBOARD_DASHBOARD_SHEET_NAME);
  var executiveSheet = getOrCreateExecutiveDashboardSheet_(spreadsheet);

  prepareExecutiveDashboardSheet_(executiveSheet);

  if (!productionSheet) {
    writeExecutiveDashboardMessage_(executiveSheet, 'DailyProduction sheet was not found. Please run Setup Sheets first.');
    SpreadsheetApp.getUi().alert('DailyProduction sheet was not found. Please run Setup Sheets first.');
    return;
  }

  var productionRows = getProductionRows_(productionSheet);

  if (productionRows.length === 0) {
    writeExecutiveDashboardMessage_(executiveSheet, 'No DailyProduction records were found.');
    SpreadsheetApp.getUi().alert('No DailyProduction records were found.');
    return;
  }

  var filterResult = getExecutiveDashboardFilterResult_(dashboardSheet, productionRows);
  var summary = buildExecutiveSummary_(
    filterResult.rows,
    getExecutiveCorrectiveActionCount_(spreadsheet)
  );

  writeExecutiveDashboard_(executiveSheet, summary, filterResult.message);
  formatExecutiveDashboard_(executiveSheet, summary);
  writeExecutiveChartHelpers_(executiveSheet, summary);

  SpreadsheetApp.flush();
  buildExecutiveCharts_(executiveSheet, summary);
  executiveSheet.hideColumns(EXECUTIVE_HELPER_START_COLUMN, EXECUTIVE_HELPER_COLUMN_COUNT);

  SpreadsheetApp.getUi().alert('Executive Dashboard refreshed.');
}

function getOrCreateExecutiveDashboardSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(EXECUTIVE_DASHBOARD_SHEET_NAME);

  if (sheet) {
    return sheet;
  }

  return spreadsheet.insertSheet(EXECUTIVE_DASHBOARD_SHEET_NAME);
}

function prepareExecutiveDashboardSheet_(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });

  sheet.showColumns(EXECUTIVE_HELPER_START_COLUMN, EXECUTIVE_HELPER_COLUMN_COUNT);
  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .breakApart();
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(0);
}

function writeExecutiveDashboardMessage_(sheet, message) {
  sheet.getRange('A1:L2').merge().setValue('EXECUTIVE DASHBOARD');
  sheet.getRange('A4:L4').merge().setValue(message);
  sheet.getRange('A1:L2')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('A4:L4')
    .setBackground('#fff2cc')
    .setFontColor('#7f6000')
    .setFontWeight('bold');
}

function getExecutiveDashboardFilterResult_(dashboardSheet, productionRows) {
  if (!dashboardSheet) {
    return {
      rows: productionRows,
      message: 'Filters: Dashboard sheet not found. Showing all DailyProduction rows.'
    };
  }

  try {
    var dashboardFilters = readDashboardFilters(dashboardSheet);
    var validationResult = validateDashboardFilters(dashboardFilters, productionRows);

    if (!validationResult.isValid) {
      return {
        rows: productionRows,
        message: 'Filters: Invalid Dashboard filters ignored. Showing all DailyProduction rows.'
      };
    }

    return {
      rows: applyProductionFilters(productionRows, validationResult.filters),
      message: buildExecutiveFilterMessage_(validationResult.filters)
    };
  } catch (error) {
    return {
      rows: productionRows,
      message: 'Filters: Dashboard filters unavailable. Showing all DailyProduction rows.'
    };
  }
}

function buildExecutiveSummary_(productionRows, correctiveActionCount) {
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

    if (machine) {
      if (!machines[machine]) {
        machines[machine] = {
          targetOutput: 0,
          actualOutput: 0,
          downtimeMinutes: 0
        };
      }

      machines[machine].targetOutput += targetOutput;
      machines[machine].actualOutput += actualOutput;
      machines[machine].downtimeMinutes += downtimeMinutes;
    }

    if (section) {
      if (!sections[section]) {
        sections[section] = {
          targetOutput: 0,
          actualOutput: 0,
          downtimeMinutes: 0,
          recordCount: 0
        };
      }

      sections[section].targetOutput += targetOutput;
      sections[section].actualOutput += actualOutput;
      sections[section].downtimeMinutes += downtimeMinutes;
      sections[section].recordCount += 1;
    }
  });

  var machineRows = buildExecutiveMachineRows_(machines);
  var sectionRows = buildExecutiveSectionRows_(sections);
  var statusCounts = countExecutiveMachineStatuses_(machineRows);

  return {
    overallEfficiency: calculateExecutiveEfficiency_(totalActualOutput, totalTargetOutput),
    totalTargetOutput: totalTargetOutput,
    totalActualOutput: totalActualOutput,
    averageDowntime: productionRows.length > 0 ? totalDowntime / productionRows.length : 0,
    totalCorrectiveActions: correctiveActionCount || 0,
    criticalMachinesCount: statusCounts.Critical,
    warningMachinesCount: statusCounts.Warning,
    normalMachinesCount: statusCounts.Normal,
    topMachines: rankExecutiveRows_(machineRows.slice().sort(function(first, second) {
      return second.efficiency - first.efficiency || first.machine.localeCompare(second.machine);
    })).slice(0, 5),
    bottomMachines: rankExecutiveRows_(machineRows.slice().sort(function(first, second) {
      return first.efficiency - second.efficiency || first.machine.localeCompare(second.machine);
    })).slice(0, 5),
    sectionRanking: rankExecutiveRows_(sectionRows.slice().sort(function(first, second) {
      return second.efficiency - first.efficiency || first.section.localeCompare(second.section);
    })).slice(0, 5),
    downtimeRanking: rankExecutiveRows_(machineRows.slice().sort(function(first, second) {
      return second.totalDowntimeMinutes - first.totalDowntimeMinutes || first.machine.localeCompare(second.machine);
    })).slice(0, 5)
  };
}

function buildExecutiveMachineRows_(machines) {
  return Object.keys(machines).map(function(machine) {
    var machineTotals = machines[machine];
    var efficiency = calculateExecutiveEfficiency_(machineTotals.actualOutput, machineTotals.targetOutput);

    return {
      machine: machine,
      efficiency: efficiency,
      status: getExecutiveMachineStatus_(efficiency),
      totalDowntimeMinutes: machineTotals.downtimeMinutes
    };
  });
}

function buildExecutiveSectionRows_(sections) {
  return Object.keys(sections).map(function(section) {
    var sectionTotals = sections[section];

    return {
      section: section,
      efficiency: calculateExecutiveEfficiency_(sectionTotals.actualOutput, sectionTotals.targetOutput),
      averageDowntime: sectionTotals.recordCount > 0 ? sectionTotals.downtimeMinutes / sectionTotals.recordCount : 0
    };
  });
}

function countExecutiveMachineStatuses_(machineRows) {
  var counts = {
    Critical: 0,
    Warning: 0,
    Normal: 0
  };

  machineRows.forEach(function(machineRow) {
    counts[machineRow.status] += 1;
  });

  return counts;
}

function rankExecutiveRows_(rows) {
  return rows.map(function(row, index) {
    var rankedRow = {};

    Object.keys(row).forEach(function(key) {
      rankedRow[key] = row[key];
    });

    rankedRow.rank = index + 1;

    return rankedRow;
  });
}

function getExecutiveCorrectiveActionCount_(spreadsheet) {
  var correctiveActionsSheet = spreadsheet.getSheetByName(EXECUTIVE_DASHBOARD_CORRECTIVE_ACTIONS_SHEET_NAME);

  if (!correctiveActionsSheet || correctiveActionsSheet.getLastRow() < 2) {
    return 0;
  }

  return correctiveActionsSheet.getLastRow() - 1;
}

function writeExecutiveDashboard_(sheet, summary, filterMessage) {
  sheet.getRange('A1:L2').merge().setValue('EXECUTIVE DASHBOARD');
  sheet.getRange('A3:L3').merge().setValue(filterMessage);

  writeExecutiveKpiCards_(sheet, summary);
  writeExecutiveMachineTables_(sheet, summary);
  writeExecutiveSectionTables_(sheet, summary);
}

function writeExecutiveKpiCards_(sheet, summary) {
  var cards = [
    {label: 'Overall Efficiency %', value: summary.overallEfficiency, range: 'A5:C7', format: '0.00%'},
    {label: 'Total Target Output', value: summary.totalTargetOutput, range: 'D5:F7', format: '#,##0'},
    {label: 'Total Actual Output', value: summary.totalActualOutput, range: 'G5:I7', format: '#,##0'},
    {label: 'Average Downtime', value: summary.averageDowntime, range: 'J5:L7', format: '#,##0.00'},
    {label: 'Total Corrective Actions', value: summary.totalCorrectiveActions, range: 'A9:C11', format: '#,##0'},
    {label: 'Critical Machines', value: summary.criticalMachinesCount, range: 'D9:F11', format: '#,##0'},
    {label: 'Warning Machines', value: summary.warningMachinesCount, range: 'G9:I11', format: '#,##0'},
    {label: 'Normal Machines', value: summary.normalMachinesCount, range: 'J9:L11', format: '#,##0'}
  ];

  cards.forEach(function(card) {
    var cardRange = sheet.getRange(card.range);
    var labelRange = sheet.getRange(cardRange.getRow(), cardRange.getColumn(), 1, cardRange.getNumColumns());
    var valueRange = sheet.getRange(cardRange.getRow() + 1, cardRange.getColumn(), 2, cardRange.getNumColumns());

    cardRange.breakApart();
    labelRange.merge().setValue(card.label);
    valueRange.merge().setValue(card.value).setNumberFormat(card.format);
  });
}

function writeExecutiveMachineTables_(sheet, summary) {
  sheet.getRange('A13:C13').merge().setValue('TOP 5 MACHINES BY EFFICIENCY');
  sheet.getRange('A14:C14').setValues([['Rank', 'Machine', 'Efficiency %']]);
  writeExecutiveRows_(sheet, 15, 1, summary.topMachines.map(function(row) {
    return [row.rank, row.machine, row.efficiency];
  }), 3);

  sheet.getRange('E13:H13').merge().setValue('BOTTOM 5 MACHINES BY EFFICIENCY');
  sheet.getRange('E14:H14').setValues([['Rank', 'Machine', 'Efficiency %', 'Status']]);
  writeExecutiveRows_(sheet, 15, 5, summary.bottomMachines.map(function(row) {
    return [row.rank, row.machine, row.efficiency, row.status];
  }), 4);
}

function writeExecutiveSectionTables_(sheet, summary) {
  sheet.getRange('A23:D23').merge().setValue('SECTION RANKING');
  sheet.getRange('A24:D24').setValues([['Rank', 'Section', 'Efficiency %', 'Average Downtime']]);
  writeExecutiveRows_(sheet, 25, 1, summary.sectionRanking.map(function(row) {
    return [row.rank, row.section, row.efficiency, row.averageDowntime];
  }), 4);

  sheet.getRange('F23:H23').merge().setValue('DOWNTIME RANKING');
  sheet.getRange('F24:H24').setValues([['Rank', 'Machine', 'Total Downtime Minutes']]);
  writeExecutiveRows_(sheet, 25, 6, summary.downtimeRanking.map(function(row) {
    return [row.rank, row.machine, row.totalDowntimeMinutes];
  }), 3);
}

function writeExecutiveRows_(sheet, startRow, startColumn, rows, columnCount) {
  if (rows.length === 0) {
    sheet.getRange(startRow, startColumn, 1, columnCount).setValues([buildEmptyExecutiveRow_(columnCount)]);
    return;
  }

  sheet.getRange(startRow, startColumn, rows.length, columnCount).setValues(rows);
}

function buildEmptyExecutiveRow_(columnCount) {
  var row = [];

  for (var index = 0; index < columnCount; index += 1) {
    row.push(index === 1 ? 'No data' : '');
  }

  return row;
}

function formatExecutiveDashboard_(sheet, summary) {
  sheet.getRange('A1:L2')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontSize(22)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('A3:L3')
    .setBackground('#f8fafc')
    .setFontColor('#334e68')
    .setFontStyle('italic')
    .setVerticalAlignment('middle');

  formatExecutiveKpiCards_(sheet);
  formatExecutiveTables_(sheet, summary);
  applyExecutiveLayout_(sheet);
}

function formatExecutiveKpiCards_(sheet) {
  var cardRanges = ['A5:C7', 'D5:F7', 'G5:I7', 'J5:L7', 'A9:C11', 'D9:F11', 'G9:I11', 'J9:L11'];

  cardRanges.forEach(function(cardRange) {
    sheet.getRange(cardRange)
      .setBackground('#f8fafc')
      .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('middle');
  });

  sheet.getRangeList(['A5:C5', 'D5:F5', 'G5:I5', 'J5:L5', 'A9:C9', 'D9:F9', 'G9:I9', 'J9:L9'])
    .setFontWeight('bold')
    .setFontColor('#334e68')
    .setHorizontalAlignment('center');
  sheet.getRangeList(['A6:C7', 'D6:F7', 'G6:I7', 'J6:L7', 'A10:C11', 'D10:F11', 'G10:I11', 'J10:L11'])
    .setFontWeight('bold')
    .setFontSize(18)
    .setFontColor('#102a43')
    .setHorizontalAlignment('center');
}

function formatExecutiveTables_(sheet, summary) {
  var headerRanges = ['A13:C13', 'E13:H13', 'A23:D23', 'F23:H23'];
  var tableHeaderRanges = ['A14:C14', 'E14:H14', 'A24:D24', 'F24:H24'];

  headerRanges.forEach(function(range) {
    sheet.getRange(range)
      .setBackground('#334e68')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  tableHeaderRanges.forEach(function(range) {
    sheet.getRange(range)
      .setBackground('#486581')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  formatExecutiveTableBody_(sheet, 15, 1, Math.max(summary.topMachines.length, 1), 3);
  formatExecutiveTableBody_(sheet, 15, 5, Math.max(summary.bottomMachines.length, 1), 4);
  formatExecutiveTableBody_(sheet, 25, 1, Math.max(summary.sectionRanking.length, 1), 4);
  formatExecutiveTableBody_(sheet, 25, 6, Math.max(summary.downtimeRanking.length, 1), 3);
  applyExecutiveStatusColors_(sheet, 15, summary.bottomMachines.length);
}

function formatExecutiveTableBody_(sheet, startRow, startColumn, rowCount, columnCount) {
  sheet.getRange(startRow, startColumn, rowCount, columnCount)
    .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
}

function applyExecutiveStatusColors_(sheet, startRow, rowCount) {
  if (rowCount < 1) {
    return;
  }

  var statuses = sheet.getRange(startRow, 8, rowCount, 1).getValues();

  statuses.forEach(function(statusRow, rowOffset) {
    var status = statusRow[0];
    var statusCell = sheet.getRange(startRow + rowOffset, 8);

    if (status === 'Normal') {
      statusCell.setBackground('#2e7d32').setFontColor('#ffffff').setFontWeight('bold');
    }

    if (status === 'Warning') {
      statusCell.setBackground('#f9ab00').setFontColor('#1f2933').setFontWeight('bold');
    }

    if (status === 'Critical') {
      statusCell.setBackground('#c62828').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
}

function applyExecutiveLayout_(sheet) {
  sheet.getRange('C15:C19').setNumberFormat('0.00%');
  sheet.getRange('G15:G19').setNumberFormat('0.00%');
  sheet.getRange('C25:C29').setNumberFormat('0.00%');
  sheet.getRange('D25:D29').setNumberFormat('#,##0.00');
  sheet.getRange('H25:H29').setNumberFormat('#,##0');
  sheet.setRowHeights(1, 2, 36);
  sheet.setRowHeight(3, 28);
  sheet.setRowHeights(5, 7, 34);
  sheet.autoResizeColumns(1, 12);
  sheet.setColumnWidths(1, 12, 105);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(10, 120);
}

function writeExecutiveChartHelpers_(sheet, summary) {
  var statusRows = [
    ['Status', 'Count'],
    ['Critical', summary.criticalMachinesCount],
    ['Warning', summary.warningMachinesCount],
    ['Normal', summary.normalMachinesCount]
  ];
  var sectionRows = [['Section', 'Efficiency %']].concat(summary.sectionRanking.map(function(row) {
    return [row.section, row.efficiency];
  }));
  var downtimeRows = [['Machine', 'Total Downtime Minutes']].concat(summary.downtimeRanking.map(function(row) {
    return [row.machine, row.totalDowntimeMinutes];
  }));

  sheet.getRange(1, EXECUTIVE_HELPER_START_COLUMN, statusRows.length, 2).setValues(statusRows);

  if (sectionRows.length > 1) {
    sheet.getRange(1, EXECUTIVE_HELPER_START_COLUMN + 2, sectionRows.length, 2).setValues(sectionRows);
  }

  if (downtimeRows.length > 1) {
    sheet.getRange(12, EXECUTIVE_HELPER_START_COLUMN + 2, downtimeRows.length, 2).setValues(downtimeRows);
  }

  sheet.getRange(2, EXECUTIVE_HELPER_START_COLUMN + 3, Math.max(sectionRows.length - 1, 1), 1).setNumberFormat('0.00%');
}

function buildExecutiveCharts_(sheet, summary) {
  if (summary.criticalMachinesCount + summary.warningMachinesCount + summary.normalMachinesCount > 0) {
    buildExecutiveMachineStatusChart_(sheet);
  }

  if (summary.sectionRanking.length > 0) {
    buildExecutiveSectionEfficiencyChart_(sheet, summary.sectionRanking.length + 1);
  }

  if (summary.downtimeRanking.length > 0) {
    buildExecutiveDowntimeChart_(sheet, summary.downtimeRanking.length + 1);
  }
}

function buildExecutiveMachineStatusChart_(sheet) {
  var chartRange = sheet.getRange(1, EXECUTIVE_HELPER_START_COLUMN, 4, 2);
  var chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(chartRange)
    .setNumHeaders(1)
    .setPosition(13, 10, 0, 0)
    .setOption('title', 'Machine Status Distribution')
    .setOption('height', 300)
    .setOption('width', 420)
    .setOption('colors', ['#c62828', '#f9ab00', '#2e7d32'])
    .build();

  sheet.insertChart(chart);
}

function buildExecutiveSectionEfficiencyChart_(sheet, sourceRowCount) {
  var chartRange = sheet.getRange(1, EXECUTIVE_HELPER_START_COLUMN + 2, sourceRowCount, 2);
  var chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setNumHeaders(1)
    .setPosition(32, 1, 0, 0)
    .setOption('title', 'Section Efficiency Ranking')
    .setOption('height', 320)
    .setOption('width', 560)
    .setOption('legend', {position: 'none'})
    .setOption('colors', ['#2e7d32'])
    .setOption('hAxis', {format: 'percent', minValue: 0})
    .build();

  sheet.insertChart(chart);
}

function buildExecutiveDowntimeChart_(sheet, sourceRowCount) {
  var chartRange = sheet.getRange(12, EXECUTIVE_HELPER_START_COLUMN + 2, sourceRowCount, 2);
  var chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setNumHeaders(1)
    .setPosition(32, 7, 0, 0)
    .setOption('title', 'Downtime Ranking')
    .setOption('height', 320)
    .setOption('width', 560)
    .setOption('legend', {position: 'none'})
    .setOption('colors', ['#486581'])
    .setOption('hAxis', {minValue: 0})
    .build();

  sheet.insertChart(chart);
}

function buildExecutiveFilterMessage_(filters) {
  return 'Filters: Date From ' + formatExecutiveFilterDate_(filters.dateFrom) +
    ', Date To ' + formatExecutiveFilterDate_(filters.dateTo) +
    ', Section ' + (filters.section || 'All') +
    ', Machine ' + (filters.machine || 'All');
}

function formatExecutiveFilterDate_(value) {
  if (!value) {
    return 'All';
  }

  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function calculateExecutiveEfficiency_(actualOutput, targetOutput) {
  if (targetOutput <= 0) {
    return 0;
  }

  return actualOutput / targetOutput;
}

function getExecutiveMachineStatus_(efficiency) {
  if (efficiency >= 0.9) {
    return 'Normal';
  }

  if (efficiency >= 0.8) {
    return 'Warning';
  }

  return 'Critical';
}
