const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('Metadata.gs', 'utf8');
const context = {
  console
};

vm.createContext(context);
vm.runInContext(source, context);

const aboutText = context.buildAboutDashboardText_();

assert.ok(aboutText.includes('Machine Efficiency Dashboard'));
assert.ok(aboutText.includes('Version 1.0.0'));
assert.ok(aboutText.includes('KPI Dashboard'));
assert.ok(aboutText.includes('Corrective Actions'));
assert.ok(aboutText.includes('Historical Trends'));
assert.ok(aboutText.includes('Executive Dashboard'));
assert.ok(aboutText.includes('Reporting Automation'));
assert.ok(aboutText.includes('Author:'));

console.log('Metadata tests passed');
