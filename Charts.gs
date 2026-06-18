var MACHINE_CHART_SOURCE_START_ROW = 1;
var MACHINE_CHART_SOURCE_START_COLUMN = 16;
var SECTION_CHART_SOURCE_START_ROW = 1;
var SECTION_CHART_SOURCE_START_COLUMN = 19;
var CHART_HELPER_START_COLUMN = 16;
var CHART_HELPER_COLUMN_COUNT = 5;

function refreshDashboardCharts(dashboardSheet, productionRows, dashboardData) {
  clearDashboardCharts(dashboardSheet);

  var machineChartRows = writeMachineChartSource(dashboardSheet, dashboardData.machineRows);
  var sectionChartRows = writeSectionChartSource(dashboardSheet, productionRows);

  SpreadsheetApp.flush();

  if (machineChartRows > 1) {
    buildMachineEfficiencyChart(dashboardSheet, machineChartRows);
  }

  if (sectionChartRows > 1) {
    buildSectionEfficiencyChart(dashboardSheet, sectionChartRows);
  }

  dashboardSheet.hideColumns(CHART_HELPER_START_COLUMN, CHART_HELPER_COLUMN_COUNT);
}

function clearDashboardCharts(dashboardSheet) {
  dashboardSheet.getCharts().forEach(function(chart) {
    dashboardSheet.removeChart(chart);
  });

  dashboardSheet.showColumns(CHART_HELPER_START_COLUMN, CHART_HELPER_COLUMN_COUNT);
  dashboardSheet
    .getRange(1, CHART_HELPER_START_COLUMN, dashboardSheet.getMaxRows(), CHART_HELPER_COLUMN_COUNT)
    .clear();
}

function writeMachineChartSource(dashboardSheet, machineRows) {
  var sourceRows = [['Machine', 'Efficiency %']];

  machineRows.forEach(function(machineRow) {
    sourceRows.push([
      machineRow[0],
      machineRow[1]
    ]);
  });

  dashboardSheet
    .getRange(MACHINE_CHART_SOURCE_START_ROW, MACHINE_CHART_SOURCE_START_COLUMN, sourceRows.length, 2)
    .setValues(sourceRows);
  dashboardSheet
    .getRange(MACHINE_CHART_SOURCE_START_ROW + 1, MACHINE_CHART_SOURCE_START_COLUMN + 1, Math.max(sourceRows.length - 1, 1), 1)
    .setNumberFormat('0.00%');

  return sourceRows.length;
}

function writeSectionChartSource(dashboardSheet, productionRows) {
  var sourceRows = [['Section', 'Efficiency %']];
  var sectionTotals = {};

  productionRows.forEach(function(row) {
    var section = row[DAILY_PRODUCTION_COLUMNS.section];
    var targetOutput = Number(row[DAILY_PRODUCTION_COLUMNS.targetOutput]) || 0;
    var actualOutput = Number(row[DAILY_PRODUCTION_COLUMNS.actualOutput]) || 0;

    if (!section) {
      return;
    }

    if (!sectionTotals[section]) {
      sectionTotals[section] = {
        targetOutput: 0,
        actualOutput: 0
      };
    }

    sectionTotals[section].targetOutput += targetOutput;
    sectionTotals[section].actualOutput += actualOutput;
  });

  Object.keys(sectionTotals)
    .map(function(section) {
      var totals = sectionTotals[section];

      return [
        section,
        calculateEfficiency_(totals.actualOutput, totals.targetOutput)
      ];
    })
    .sort(function(firstSection, secondSection) {
      return firstSection[1] - secondSection[1];
    })
    .forEach(function(sectionRow) {
      sourceRows.push(sectionRow);
    });

  dashboardSheet
    .getRange(SECTION_CHART_SOURCE_START_ROW, SECTION_CHART_SOURCE_START_COLUMN, sourceRows.length, 2)
    .setValues(sourceRows);
  dashboardSheet
    .getRange(SECTION_CHART_SOURCE_START_ROW + 1, SECTION_CHART_SOURCE_START_COLUMN + 1, Math.max(sourceRows.length - 1, 1), 1)
    .setNumberFormat('0.00%');

  return sourceRows.length;
}

function buildMachineEfficiencyChart(dashboardSheet, sourceRowCount) {
  var chartRange = dashboardSheet.getRange(
    MACHINE_CHART_SOURCE_START_ROW,
    MACHINE_CHART_SOURCE_START_COLUMN,
    sourceRowCount,
    2
  );
  var chart = dashboardSheet
    .newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setNumHeaders(1)
    .setPosition(24, 1, 0, 0)
    .setOption('title', 'Machine Efficiency %')
    .setOption('legend', {position: 'none'})
    .setOption('height', 350)
    .setOption('width', 550)
    .setOption('colors', ['#486581'])
    .setOption('hAxis', {format: 'percent', minValue: 0})
    .build();

  dashboardSheet.insertChart(chart);
}

function buildSectionEfficiencyChart(dashboardSheet, sourceRowCount) {
  var chartRange = dashboardSheet.getRange(
    SECTION_CHART_SOURCE_START_ROW,
    SECTION_CHART_SOURCE_START_COLUMN,
    sourceRowCount,
    2
  );
  var chart = dashboardSheet
    .newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setNumHeaders(1)
    .setPosition(24, 7, 0, 0)
    .setOption('title', 'Section Efficiency %')
    .setOption('legend', {position: 'none'})
    .setOption('height', 350)
    .setOption('width', 500)
    .setOption('colors', ['#2e7d32'])
    .setOption('hAxis', {format: 'percent', minValue: 0})
    .build();

  dashboardSheet.insertChart(chart);
}
