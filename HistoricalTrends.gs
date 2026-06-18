var HISTORICAL_TRENDS_SHEET_NAME = 'Historical Trends';
var HISTORICAL_TRENDS_DAILY_PRODUCTION_SHEET_NAME = 'DailyProduction';
var HISTORICAL_TRENDS_DASHBOARD_SHEET_NAME = 'Dashboard';

var HISTORICAL_TRENDS_HEADERS = [
  'Date',
  'Total Target Output',
  'Total Actual Output',
  'Efficiency %',
  'Average Downtime',
  'Total Machines',
  'Total Sections'
];

var HISTORICAL_TRENDS_DATA_START_ROW = 4;
var HISTORICAL_TRENDS_COLUMN_COUNT = HISTORICAL_TRENDS_HEADERS.length;

function generateHistoricalTrends() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productionSheet = spreadsheet.getSheetByName(HISTORICAL_TRENDS_DAILY_PRODUCTION_SHEET_NAME);
  var dashboardSheet = spreadsheet.getSheetByName(HISTORICAL_TRENDS_DASHBOARD_SHEET_NAME);
  var trendsSheet = getOrCreateHistoricalTrendsSheet_(spreadsheet);

  prepareHistoricalTrendsSheet_(trendsSheet);

  if (!productionSheet) {
    writeHistoricalTrendsMessage_(trendsSheet, 'DailyProduction sheet was not found. Please run Setup Sheets first.');
    SpreadsheetApp.getUi().alert('DailyProduction sheet was not found. Please run Setup Sheets first.');
    return;
  }

  var productionRows = getProductionRows_(productionSheet);

  if (productionRows.length === 0) {
    writeHistoricalTrendsMessage_(trendsSheet, 'No DailyProduction records were found.');
    SpreadsheetApp.getUi().alert('No DailyProduction records were found.');
    return;
  }

  var filterResult = getHistoricalTrendFilterResult_(dashboardSheet, productionRows);
  var filteredRows = filterResult.rows;
  var chartContext = filterResult.filters || {};
  var trendRows = buildHistoricalTrendRows_(filteredRows);

  writeHistoricalTrendFilterNote_(trendsSheet, filterResult.message);

  if (trendRows.length === 0) {
    writeHistoricalTrendsMessage_(trendsSheet, 'No valid dated records matched the selected filters.');
    SpreadsheetApp.getUi().alert('No valid dated records matched the selected filters.');
    return;
  }

  writeHistoricalTrendRows_(trendsSheet, trendRows);
  formatHistoricalTrendsSheet_(trendsSheet, trendRows.length);

  SpreadsheetApp.flush();
  buildHistoricalTrendCharts_(trendsSheet, trendRows.length, chartContext);

  SpreadsheetApp.getUi().alert('Historical Trends generated.');
}

function getHistoricalTrendFilterResult_(dashboardSheet, productionRows) {
  if (!dashboardSheet) {
    return {
      rows: productionRows,
      filters: {},
      message: 'Filters: Dashboard sheet not found. Showing all DailyProduction rows.'
    };
  }

  try {
    var dashboardFilters = readDashboardFilters(dashboardSheet);
    var validationResult = validateDashboardFilters(dashboardFilters, productionRows);

    if (!validationResult.isValid) {
      return {
        rows: productionRows,
        filters: {},
        message: 'Filters: Invalid Dashboard filters ignored. Showing all DailyProduction rows.'
      };
    }

    return {
      rows: applyProductionFilters(productionRows, validationResult.filters),
      filters: validationResult.filters,
      message: buildHistoricalTrendFilterMessage_(validationResult.filters)
    };
  } catch (error) {
    return {
      rows: productionRows,
      filters: {},
      message: 'Filters: Dashboard filters unavailable. Showing all DailyProduction rows.'
    };
  }
}

function buildHistoricalTrendRows_(productionRows) {
  var dailyTotals = {};

  productionRows.forEach(function(row) {
    var parsedDate = parseHistoricalDate_(row[DAILY_PRODUCTION_COLUMNS.date]);

    if (!parsedDate) {
      return;
    }

    var dateKey = getHistoricalDateKey_(parsedDate);
    var targetOutput = Number(row[DAILY_PRODUCTION_COLUMNS.targetOutput]) || 0;
    var actualOutput = Number(row[DAILY_PRODUCTION_COLUMNS.actualOutput]) || 0;
    var downtimeMinutes = Number(row[DAILY_PRODUCTION_COLUMNS.downtimeMinutes]) || 0;
    var section = row[DAILY_PRODUCTION_COLUMNS.section];
    var machine = row[DAILY_PRODUCTION_COLUMNS.machine];

    if (!dailyTotals[dateKey]) {
      dailyTotals[dateKey] = {
        date: parsedDate,
        targetOutput: 0,
        actualOutput: 0,
        downtimeMinutes: 0,
        recordCount: 0,
        machines: {},
        sections: {}
      };
    }

    dailyTotals[dateKey].targetOutput += targetOutput;
    dailyTotals[dateKey].actualOutput += actualOutput;
    dailyTotals[dateKey].downtimeMinutes += downtimeMinutes;
    dailyTotals[dateKey].recordCount += 1;

    if (machine) {
      dailyTotals[dateKey].machines[machine] = true;
    }

    if (section) {
      dailyTotals[dateKey].sections[section] = true;
    }
  });

  return Object.keys(dailyTotals)
    .sort()
    .map(function(dateKey) {
      var totals = dailyTotals[dateKey];

      return [
        totals.date,
        totals.targetOutput,
        totals.actualOutput,
        calculateHistoricalEfficiency_(totals.actualOutput, totals.targetOutput),
        totals.recordCount > 0 ? totals.downtimeMinutes / totals.recordCount : 0,
        Object.keys(totals.machines).length,
        Object.keys(totals.sections).length
      ];
    });
}

function getOrCreateHistoricalTrendsSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(HISTORICAL_TRENDS_SHEET_NAME);

  if (sheet) {
    return sheet;
  }

  return spreadsheet.insertSheet(HISTORICAL_TRENDS_SHEET_NAME);
}

function prepareHistoricalTrendsSheet_(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });

  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.getRange(1, 1, 1, HISTORICAL_TRENDS_COLUMN_COUNT).merge().setValue('HISTORICAL TRENDS');
  sheet.getRange(3, 1, 1, HISTORICAL_TRENDS_COLUMN_COUNT).setValues([HISTORICAL_TRENDS_HEADERS]);
  sheet.setFrozenRows(3);
}

function writeHistoricalTrendFilterNote_(sheet, message) {
  sheet.getRange(2, 1, 1, HISTORICAL_TRENDS_COLUMN_COUNT).merge().setValue(message);
}

function writeHistoricalTrendsMessage_(sheet, message) {
  sheet.getRange(2, 1, 1, HISTORICAL_TRENDS_COLUMN_COUNT).merge().setValue(message);
  formatHistoricalTrendsSheet_(sheet, 0);
}

function writeHistoricalTrendRows_(sheet, trendRows) {
  sheet
    .getRange(HISTORICAL_TRENDS_DATA_START_ROW, 1, trendRows.length, HISTORICAL_TRENDS_COLUMN_COUNT)
    .setValues(trendRows);
}

function formatHistoricalTrendsSheet_(sheet, trendRowCount) {
  sheet
    .getRange(1, 1, 1, HISTORICAL_TRENDS_COLUMN_COUNT)
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet
    .getRange(2, 1, 1, HISTORICAL_TRENDS_COLUMN_COUNT)
    .setBackground('#f8fafc')
    .setFontColor('#334e68')
    .setFontStyle('italic')
    .setVerticalAlignment('middle');

  sheet
    .getRange(3, 1, 1, HISTORICAL_TRENDS_COLUMN_COUNT)
    .setBackground('#486581')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  if (trendRowCount > 0) {
    sheet
      .getRange(HISTORICAL_TRENDS_DATA_START_ROW, 1, trendRowCount, HISTORICAL_TRENDS_COLUMN_COUNT)
      .setBorder(true, true, true, true, true, true, '#d9e2ec', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('middle');
    sheet
      .getRange(HISTORICAL_TRENDS_DATA_START_ROW, 1, trendRowCount, 1)
      .setNumberFormat('mmm d, yyyy');
    sheet
      .getRange(HISTORICAL_TRENDS_DATA_START_ROW, 2, trendRowCount, 2)
      .setNumberFormat('#,##0');
    sheet
      .getRange(HISTORICAL_TRENDS_DATA_START_ROW, 4, trendRowCount, 1)
      .setNumberFormat('0.00%');
    sheet
      .getRange(HISTORICAL_TRENDS_DATA_START_ROW, 5, trendRowCount, 1)
      .setNumberFormat('#,##0.00');
    sheet
      .getRange(HISTORICAL_TRENDS_DATA_START_ROW, 6, trendRowCount, 2)
      .setNumberFormat('#,##0');
  }

  sheet.setRowHeight(1, 34);
  sheet.setRowHeight(2, 28);
  sheet.setRowHeight(3, 30);
  sheet.autoResizeColumns(1, HISTORICAL_TRENDS_COLUMN_COUNT);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 140);
}

function buildHistoricalTrendCharts_(sheet, trendRowCount, chartContext) {
  if (trendRowCount < 1) {
    return;
  }

  buildDailyEfficiencyTrendChart_(sheet, trendRowCount, chartContext);
  buildDailyDowntimeTrendChart_(sheet, trendRowCount, chartContext);
}

function buildDailyEfficiencyTrendChart_(sheet, trendRowCount, chartContext) {
  var dateRange = sheet.getRange(HISTORICAL_TRENDS_DATA_START_ROW, 1, trendRowCount, 1);
  var efficiencyRange = sheet.getRange(HISTORICAL_TRENDS_DATA_START_ROW, 4, trendRowCount, 1);
  var chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dateRange)
    .addRange(efficiencyRange)
    .setPosition(3, 9, 0, 0)
    .setOption('title', buildHistoricalTrendChartTitle_(chartContext, 'Efficiency Trend'))
    .setOption('legend', {position: 'bottom'})
    .setOption('height', 320)
    .setOption('width', 760)
    .setOption('colors', ['#2e7d32'])
    .setOption('curveType', 'function')
    .setOption('pointSize', 5)
    .setOption('dataLabel', 'value')
    .setOption('vAxis', {format: 'percent', minValue: 0, title: 'Efficiency %'})
    .setOption('hAxis', {format: 'MMM d'})
    .build();

  sheet.insertChart(chart);
}

function buildDailyDowntimeTrendChart_(sheet, trendRowCount, chartContext) {
  var dateRange = sheet.getRange(HISTORICAL_TRENDS_DATA_START_ROW, 1, trendRowCount, 1);
  var downtimeRange = sheet.getRange(HISTORICAL_TRENDS_DATA_START_ROW, 5, trendRowCount, 1);
  var chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dateRange)
    .addRange(downtimeRange)
    .setPosition(21, 9, 0, 0)
    .setOption('title', buildHistoricalTrendChartTitle_(chartContext, 'Downtime Minutes'))
    .setOption('legend', {position: 'bottom'})
    .setOption('height', 320)
    .setOption('width', 760)
    .setOption('colors', ['#f9ab00'])
    .setOption('curveType', 'function')
    .setOption('pointSize', 5)
    .setOption('dataLabel', 'value')
    .setOption('vAxis', {minValue: 0, title: 'Downtime Minutes'})
    .setOption('hAxis', {format: 'MMM d'})
    .build();

  sheet.insertChart(chart);
}

function buildHistoricalTrendChartTitle_(filters, metricName) {
  var section = filters && filters.section ? filters.section : 'All';
  var machine = filters && filters.machine ? filters.machine : 'All';

  if (machine !== 'All') {
    return machine + ' ' + metricName;
  }

  if (section !== 'All') {
    return section + ' ' + metricName;
  }

  return 'Daily ' + metricName;
}

function buildHistoricalTrendFilterMessage_(filters) {
  return 'Filters: Date From ' + formatHistoricalFilterDate_(filters.dateFrom) +
    ', Date To ' + formatHistoricalFilterDate_(filters.dateTo) +
    ', Section ' + (filters.section || 'All') +
    ', Machine ' + (filters.machine || 'All');
}

function formatHistoricalFilterDate_(value) {
  if (!value) {
    return 'All';
  }

  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function calculateHistoricalEfficiency_(actualOutput, targetOutput) {
  if (targetOutput <= 0) {
    return 0;
  }

  return actualOutput / targetOutput;
}

function parseHistoricalDate_(value) {
  if (!value) {
    return null;
  }

  var parsedDate = value instanceof Date ? value : new Date(value);

  if (isNaN(parsedDate.getTime())) {
    return null;
  }

  return new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate());
}

function getHistoricalDateKey_(value) {
  var year = value.getFullYear();
  var month = String(value.getMonth() + 1);
  var day = String(value.getDate());

  return year + '-' + month.padStart(2, '0') + '-' + day.padStart(2, '0');
}
