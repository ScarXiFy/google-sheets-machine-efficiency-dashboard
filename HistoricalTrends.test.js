const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('HistoricalTrends.gs', 'utf8');
const context = {
  console,
  Utilities: {
    formatDate(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
  },
  Session: {
    getScriptTimeZone() {
      return 'Asia/Manila';
    }
  },
  DAILY_PRODUCTION_COLUMNS: {
    date: 0,
    section: 1,
    machine: 2,
    targetOutput: 3,
    actualOutput: 4,
    downtimeMinutes: 5
  }
};

vm.createContext(context);
vm.runInContext(source, context);

const rows = [
  [new Date(2026, 5, 2), 'Assembly', 'Machine A1', 100, 90, 10],
  [new Date(2026, 5, 1), 'Cutting', 'Machine C1', 0, 50, 20],
  [new Date(2026, 5, 2), 'Assembly', 'Machine A2', 200, 180, 30],
  ['not a date', 'Assembly', 'Machine A3', 100, 100, 5]
];

const trendRows = context.buildHistoricalTrendRows_(rows);

assert.strictEqual(trendRows.length, 2);
assert.strictEqual(trendRows[0][0].getFullYear(), 2026);
assert.strictEqual(trendRows[0][0].getMonth(), 5);
assert.strictEqual(trendRows[0][0].getDate(), 1);
assert.strictEqual(trendRows[0][1], 0);
assert.strictEqual(trendRows[0][2], 50);
assert.strictEqual(trendRows[0][3], 0);
assert.strictEqual(trendRows[0][4], 20);
assert.strictEqual(trendRows[0][5], 1);
assert.strictEqual(trendRows[0][6], 1);

assert.strictEqual(trendRows[1][0].getDate(), 2);
assert.strictEqual(trendRows[1][1], 300);
assert.strictEqual(trendRows[1][2], 270);
assert.strictEqual(trendRows[1][3], 0.9);
assert.strictEqual(trendRows[1][4], 20);
assert.strictEqual(trendRows[1][5], 2);
assert.strictEqual(trendRows[1][6], 1);

assert.strictEqual(
  context.buildHistoricalTrendChartTitle_(
    {
      section: 'Packaging',
      machine: 'All'
    },
    'Efficiency Trend'
  ),
  'Packaging Efficiency Trend'
);
assert.strictEqual(
  context.buildHistoricalTrendChartTitle_(
    {
      section: 'All',
      machine: 'Machine P1'
    },
    'Downtime Minutes'
  ),
  'Machine P1 Downtime Minutes'
);
assert.strictEqual(
  context.buildHistoricalTrendChartTitle_(
    {
      section: 'All',
      machine: 'All'
    },
    'Efficiency Trend'
  ),
  'Daily Efficiency Trend'
);

console.log('Historical trend tests passed');
