# Machine Efficiency Dashboard

## Overview

Machine Efficiency Dashboard is a Google Sheets and Apps Script-based machine efficiency monitoring system for tracking production performance, efficiency metrics, downtime, corrective actions, and management reporting.

The project is designed for production teams that need a simple spreadsheet-based workflow without a separate database or web application. It turns daily production records into KPI cards, machine status summaries, charts, corrective action recommendations, daily summaries, and PDF reports.

## Project Screenshots

Recommended screenshots for portfolio or demo use:

- Dashboard
- Corrective Actions
- Historical Trends
- Executive Dashboard
- Daily Summary

Store screenshots in a `screenshots/` folder if you want them to render directly in GitHub later.

## Features

- KPI Dashboard
- Machine Efficiency Monitoring
- Section Efficiency Monitoring
- Corrective Action Generation
- RCA Configuration
- Data Validation
- Data Audit
- Dashboard Filters
- Historical Trends
- Executive Dashboard
- Daily Summary Reports
- PDF Export
- Email Reporting Automation
- Report History
- Charts and Visualizations

## System Workflow

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

### Historical Trends

Purpose: Summarizes performance by production date and displays efficiency and downtime trends.

### Executive Dashboard

Purpose: Provides executive-level KPI cards, machine rankings, section rankings, downtime rankings, and status distribution charts.

### Data Audit

Purpose: Lists missing fields, invalid dates, negative values, and possible production entry errors.

### Report Settings

Purpose: Stores report recipient, CC recipient, enabled flags, and last report sent timestamp.

### Report History

Purpose: Tracks report delivery attempts, status, recipient, and generated file names.

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

## Skills Demonstrated

- Google Apps Script
- Spreadsheet Automation
- KPI Analytics
- Dashboard Design
- Data Validation
- Reporting Systems
- Business Process Automation
- Data Visualization
- PDF Generation
- Email Automation
- Technical Documentation

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
   - `DataValidation.gs`
   - `HistoricalTrends.gs`
   - `AutomatedReporting.gs`
   - `ExecutiveDashboard.gs`
   - `Metadata.gs`
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

### Setup Data Validation

Adds validation rules, dropdowns, formatting, data quality highlighting, and generated sheet protection.

### Audit Production Data

Creates or refreshes the Data Audit sheet with data quality issues found in DailyProduction.

### Generate Historical Trends

Creates or refreshes the Historical Trends sheet with daily performance summaries and trend charts.

### Refresh Executive Dashboard

Creates or refreshes the Executive Dashboard sheet for management review.

### Send Daily Report

Generates reports, emails PDF attachments, updates Report History, and updates Last Report Sent.

### Send Weekly Report

Generates a weekly-style management email report with KPI, trend, corrective action, and machine ranking context.

### View Report History

Opens the Report History sheet.

### About Dashboard

Displays project version and feature metadata.

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
6. Run **Generate Historical Trends**.
7. Run **Refresh Executive Dashboard**.
8. Export or email reports for review or sharing.

## Portfolio Highlights

- Google Apps Script
- Google Sheets Automation
- Dashboard Design
- KPI Analytics
- Business Reporting
- Data Visualization
- Process Automation
- Executive Analytics
- Email Automation

## Limitations

- Designed for a single Google Sheets deployment.
- Production data is entered manually.
- Google Apps Script quotas may apply to PDF export and repeated script execution.
- RCA logic is sample rule-based logic and may need adjustment for real production environments.

## Future Enhancements

- Trigger setup wizard for scheduled reports
- Monthly and quarterly reporting
- Screenshot gallery
- CSV import workflow
- Multi-user support
