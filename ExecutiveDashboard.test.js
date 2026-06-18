const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('ExecutiveDashboard.gs', 'utf8');
const context = {
  console,
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
  [new Date(2026, 5, 1), 'Assembly', 'Machine A1', 100, 95, 10],
  [new Date(2026, 5, 2), 'Assembly', 'Machine A2', 100, 85, 20],
  [new Date(2026, 5, 3), 'Packaging', 'Machine P1', 100, 70, 40],
  [new Date(2026, 5, 4), 'Packaging', 'Machine P2', 0, 50, 5],
  [new Date(2026, 5, 5), 'Cutting', 'Machine C1', 100, 100, 12],
  [new Date(2026, 5, 6), 'Finishing', 'Machine F1', 200, 170, 33]
];

const summary = context.buildExecutiveSummary_(rows, 4);

assert.strictEqual(summary.totalTargetOutput, 600);
assert.strictEqual(summary.totalActualOutput, 570);
assert.strictEqual(summary.criticalMachinesCount, 2);
assert.strictEqual(summary.warningMachinesCount, 2);
assert.strictEqual(summary.normalMachinesCount, 2);
assert.strictEqual(summary.totalCorrectiveActions, 4);

assert.strictEqual(
  JSON.stringify(
  summary.topMachines.slice(0, 3).map((row) => [row.rank, row.machine, row.efficiency]),
  ),
  JSON.stringify(
  [
    [1, 'Machine C1', 1],
    [2, 'Machine A1', 0.95],
    [3, 'Machine A2', 0.85]
  ]
  )
);

assert.strictEqual(
  JSON.stringify(
  summary.bottomMachines.slice(0, 3).map((row) => [row.rank, row.machine, row.efficiency, row.status]),
  ),
  JSON.stringify(
  [
    [1, 'Machine P2', 0, 'Critical'],
    [2, 'Machine P1', 0.7, 'Critical'],
    [3, 'Machine A2', 0.85, 'Warning']
  ]
  )
);

assert.strictEqual(
  JSON.stringify(
  summary.sectionRanking.slice(0, 2).map((row) => [row.rank, row.section, row.efficiency, row.averageDowntime]),
  ),
  JSON.stringify(
  [
    [1, 'Packaging', 1.2, 22.5],
    [2, 'Cutting', 1, 12]
  ]
  )
);

assert.strictEqual(
  JSON.stringify(
  summary.downtimeRanking.slice(0, 2).map((row) => [row.rank, row.machine, row.totalDowntimeMinutes]),
  ),
  JSON.stringify(
  [
    [1, 'Machine P1', 40],
    [2, 'Machine F1', 33]
  ]
  )
);

console.log('Executive dashboard tests passed');
