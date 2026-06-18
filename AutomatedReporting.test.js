const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('AutomatedReporting.gs', 'utf8');
const context = {
  console,
  Utilities: {
    formatDate(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    },
    formatString(format, value) {
      if (format === '%.2f%%') {
        return `${Number(value).toFixed(2)}%`;
      }

      if (format === '%.2f') {
        return Number(value).toFixed(2);
      }

      return String(value);
    }
  },
  Session: {
    getScriptTimeZone() {
      return 'Asia/Manila';
    }
  }
};

vm.createContext(context);
vm.runInContext(source, context);
context.formatReportPercent_ = function(value) {
  return `${((Number(value) || 0) * 100).toFixed(2)}%`;
};

const settings = context.buildReportSettingsMap_([
  ['Report Recipient', 'manager@example.com'],
  ['CC Recipient', 'director@example.com'],
  ['Daily Report Enabled', 'Yes']
]);

assert.strictEqual(settings['Report Recipient'], 'manager@example.com');
assert.strictEqual(settings['CC Recipient'], 'director@example.com');
assert.strictEqual(settings['Daily Report Enabled'], 'Yes');
assert.strictEqual(context.isReportEnabled_('Yes'), true);
assert.strictEqual(context.isReportEnabled_('No'), false);
assert.strictEqual(context.isValidReportRecipient_('manager@example.com'), false);
assert.strictEqual(context.isValidReportRecipient_('ops.manager@example.com'), true);

const dashboardData = {
  overallEfficiency: 0.8191,
  averageDowntime: 25.5,
  totalMachines: 10,
  totalSections: 5,
  machineRows: [
    ['Machine C1', 0.72, 'Critical'],
    ['Machine A1', 0.94, 'Normal']
  ]
};
const body = context.buildReportEmailBody_('Daily', dashboardData, {
  correctiveActionCount: 3,
  trendDayCount: 7
});

assert.ok(body.includes('Overall Efficiency: 81.91%'));
assert.ok(body.includes('Average Downtime: 25.50 minutes'));
assert.ok(body.includes('Total Machines: 10'));
assert.ok(body.includes('Total Sections: 5'));
assert.ok(body.includes('Corrective Actions: 3'));
assert.ok(body.includes('Top Performing Machine: Machine A1 (94.00%)'));
assert.ok(body.includes('Lowest Performing Machine: Machine C1 (72.00%)'));

console.log('Automated reporting tests passed');
