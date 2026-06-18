# Architecture

## Data Flow

```text
DailyProduction
-> Dashboard
-> Corrective Actions
-> Daily Summary
-> Historical Trends
-> Executive Dashboard
-> Email Reports
-> PDF Reports
```

## Component Descriptions

### DailyProduction

`DailyProduction` is the source sheet for production records. Each row stores a production date, section, machine, target output, actual output, downtime minutes, and remarks.

This sheet is intentionally editable because it is the primary data-entry area.

### Dashboard

`Dashboard` summarizes production rows into KPI cards, machine status tables, and visual charts. It supports filters for date range, section, and machine.

The dashboard calculates overall efficiency, total output, average downtime, machine count, and section count.

### Corrective Actions

`Corrective Actions` stores generated action recommendations for records that need attention. It uses rule-based issue detection for no output, high downtime, and low efficiency.

The sheet depends on `DailyProduction` and `RCA Config`.

### Daily Summary

`Daily Summary` creates a concise management report from filtered dashboard data. It is designed for quick review and PDF export.

### Historical Trends

`Historical Trends` groups production performance by date. It shows daily target output, actual output, efficiency, downtime, machine count, and section count.

It also creates trend charts for efficiency and downtime.

### Executive Dashboard

`Executive Dashboard` provides a high-level management view. It includes KPI cards, status counts, top and bottom machine rankings, section ranking, downtime ranking, and executive charts.

It uses the current Dashboard filters when available and falls back to all production rows if needed.

### Email Reports

Email reporting is handled by `AutomatedReporting.gs`. It reads `Report Settings`, generates report PDFs, sends email attachments, updates `Last Report Sent`, and logs outcomes in `Report History`.

Scheduled helper functions are available, but triggers are not created automatically.

### PDF Reports

PDF export is handled by `Reports.gs`. It exports Dashboard, Corrective Actions, and Daily Summary sheets to Google Drive using Apps Script export URLs.

## Supporting Sheets

### RCA Config

Maps issue types to possible causes and recommended corrective actions.

### Report Settings

Stores configurable reporting values such as report recipient, CC recipient, daily report enabled, weekly report enabled, and last report sent.

### Report History

Stores timestamped report delivery outcomes.

### Data Audit

Stores detected data quality issues from `DailyProduction`, including missing fields, invalid dates, negative values, and actual output above target.
