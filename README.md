# Machine Efficiency Dashboard

## Overview

Machine Efficiency Dashboard is a Google Sheets and Apps Script-based machine efficiency monitoring system for tracking production performance, efficiency metrics, downtime, corrective actions, and management reporting.

The project is designed for production teams that need a simple spreadsheet-based workflow without a separate database or web application. It turns daily production records into KPI cards, machine status summaries, charts, corrective action recommendations, daily summaries, and PDF reports.

## Features

- KPI Dashboard
- Machine Efficiency Monitoring
- Section Efficiency Monitoring
- Corrective Action Generation
- RCA Configuration
- Dashboard Filters
- Daily Summary Reports
- PDF Export
- Charts and Visualizations

## System Workflow

```text
DailyProduction
-> Dashboard
-> Corrective Actions
-> Daily Summary
-> PDF Reports
```

Production data starts in the `DailyProduction` sheet. The dashboard summarizes that data into KPIs, machine status, and charts. Corrective actions can then be generated for machines below the expected efficiency level. A Daily Summary can be created for management review, and dashboard or corrective action sheets can be exported as PDF reports.

## Sheets Used

### DailyProduction

Purpose: Stores production records.

Columns:

- Date
- Section
- Machine
- Target Output
- Actual Output
- Downtime Minutes
- Remarks

### Dashboard

Purpose: Displays KPI cards, filter controls, machine performance status, and charts.

Dashboard filters are placed in row 4:

- `B4`: Date From
- `D4`: Date To
- `F4`: Section
- `H4`: Machine

Chart helper ranges are written on the Dashboard sheet:

- `P:Q`: Machine Efficiency chart data
- `S:T`: Section Efficiency chart data

### Corrective Actions

Purpose: Stores automatically generated corrective action recommendations for production records that need attention.

### RCA Config

Purpose: Stores issue-to-cause and issue-to-action mappings used by corrective action generation.

### Daily Summary

Purpose: Provides a management summary based on the currently selected Dashboard filters.

## KPI Definitions

### Total Target Output

The sum of all target output values for the selected production rows.

### Total Actual Output

The sum of all actual output values for the selected production rows.

### Overall Efficiency %

Formula:

```text
Actual Output / Target Output x 100
```

If target output is zero, efficiency is treated as `0%` to avoid division by zero.

### Average Downtime

The average downtime minutes across the selected production rows.

### Total Machines

The number of unique machines included in the selected production rows.

### Total Sections

The number of unique sections included in the selected production rows.

## Machine Status Rules

Machine status is based on machine efficiency:

| Status | Rule |
| --- | --- |
| Normal | `>= 90%` |
| Warning | `80%` to `89.99%` |
| Critical | `< 80%` |

## Corrective Action Rules

Corrective actions are generated for production records below the normal efficiency threshold.

Priority order:

1. No Output
2. High Downtime
3. Low Efficiency

### No Output

Used when actual output is zero.

### High Downtime

Used when downtime is high enough to indicate a stoppage or production issue.

### Low Efficiency

Used when output is below target and no higher-priority issue applies.

The issue is matched against `RCA Config` to determine the possible cause and recommended corrective action.

## Dashboard Filters

The Dashboard supports filtering by:

- Date Range
- Section
- Machine

If a filter is blank, it behaves as no filter. Section and Machine default to `All`.

Filters affect:

- KPI cards
- Machine Efficiency Status table
- Machine Efficiency chart
- Section Efficiency chart
- Daily Summary

## PDF Export

The Machine Dashboard menu includes two PDF export options:

- Export Dashboard PDF
- Export Corrective Actions PDF

PDF files are generated in Google Drive with timestamped filenames. The Dashboard PDF preserves dashboard formatting and charts. The Corrective Actions PDF exports the corrective action report sheet.

## Installation

1. Create a new Google Sheet.
2. Open **Extensions** > **Apps Script**.
3. Create script files matching this repository:
   - `Code.gs`
   - `DataSetup.gs`
   - `Dashboard.gs`
   - `CorrectiveActions.gs`
   - `Charts.gs`
   - `Reports.gs`
4. Copy each file's code into the matching Apps Script file.
5. Save the Apps Script project.
6. Run any menu function once from Apps Script if Google asks for authorization.
7. Reload the spreadsheet.
8. Open the **Machine Dashboard** menu.
9. Run **Setup Sheets**.
10. Run **Load Sample Data**.
11. Run **Refresh Dashboard**.

## Usage Guide

The spreadsheet menu is named **Machine Dashboard**.

### Setup Sheets

Creates or resets the required sheets:

- DailyProduction
- Dashboard
- Corrective Actions
- RCA Config

### Load Sample Data

Loads sample production records into the `DailyProduction` sheet.

### Refresh Dashboard

Rebuilds the Dashboard sheet using the current Dashboard filters. This updates KPI cards, the machine status table, and charts.

### Generate Corrective Actions

Builds the Corrective Actions sheet using DailyProduction data and RCA Config mappings.

### Generate Daily Summary

Creates or refreshes the Daily Summary sheet using the currently selected Dashboard filters.

### Export Dashboard PDF

Exports the Dashboard sheet to a timestamped PDF file in Google Drive.

### Export Corrective Actions PDF

Exports the Corrective Actions sheet to a timestamped PDF file in Google Drive.

## Example Workflow

1. Enter production data in `DailyProduction`.
2. Run **Refresh Dashboard**.
3. Review KPI cards, charts, and machine status.
4. Run **Generate Corrective Actions**.
5. Run **Generate Daily Summary**.
6. Export PDF reports for review or sharing.

## Portfolio Highlights

- Google Apps Script
- Google Sheets Automation
- Dashboard Design
- KPI Analytics
- Business Reporting
- Data Visualization
- Process Automation

## Limitations

- Designed for a single Google Sheets deployment.
- Production data is entered manually.
- Google Apps Script quotas may apply to PDF export and repeated script execution.
- RCA logic is sample rule-based logic and may need adjustment for real production environments.

## Future Enhancements

- Email reports
- Automated scheduling
- Trend analysis
- Machine ranking dashboard
- Historical reporting
- Multi-user support
