var DASHBOARD_APP_NAME = 'Machine Efficiency Dashboard';
var DASHBOARD_APP_VERSION = '1.0.0';
var DASHBOARD_APP_AUTHOR = '[configurable]';

function showAbout() {
  SpreadsheetApp.getUi().alert(buildAboutDashboardText_());
}

function buildAboutDashboardText_() {
  return [
    DASHBOARD_APP_NAME,
    'Version ' + DASHBOARD_APP_VERSION,
    '',
    'Features:',
    '- KPI Dashboard',
    '- Corrective Actions',
    '- Historical Trends',
    '- Executive Dashboard',
    '- Reporting Automation',
    '',
    'Author:',
    DASHBOARD_APP_AUTHOR
  ].join('\n');
}
