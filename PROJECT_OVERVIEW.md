# Machine Efficiency Dashboard Project Overview

## Problem Statement

Production teams often track machine output, downtime, and corrective actions in separate spreadsheet tabs or manual reports. This makes it hard for managers to quickly see which machines are underperforming, where downtime is concentrated, and what actions should be taken.

The problem this project solves is turning daily production records into a clear management reporting system without requiring a separate database, web app, or paid business intelligence platform.

## Solution Overview

Machine Efficiency Dashboard is a Google Sheets and Apps Script system that converts production records into KPI dashboards, corrective action reports, historical trends, executive summaries, PDF exports, and automated email reports.

The spreadsheet remains the main user interface. Apps Script handles setup, validation, calculations, formatting, charts, reporting, and automation.

## Features

- Daily production data entry
- KPI dashboard with filters
- Machine and section efficiency tracking
- Corrective action generation
- RCA configuration sheet
- Data validation and data audit reporting
- Historical trend analytics
- Executive dashboard
- Daily summary report
- PDF export
- Email reporting automation
- Report settings and report history

## Technologies Used

- Google Sheets
- Google Apps Script
- JavaScript
- Spreadsheet formulas and formatting
- Google Drive PDF export
- Gmail-compatible Apps Script email sending
- Chart services in Google Sheets

## Architecture Overview

The project is organized by responsibility:

- `Code.gs` defines the custom spreadsheet menu.
- `DataSetup.gs` creates base sheets and sample data.
- `Dashboard.gs` calculates KPIs and renders the main dashboard.
- `Charts.gs` builds dashboard chart visuals.
- `CorrectiveActions.gs` generates recommended actions.
- `Reports.gs` creates summary reports and PDF exports.
- `DataValidation.gs` validates production input and audits data quality.
- `HistoricalTrends.gs` summarizes performance over time.
- `AutomatedReporting.gs` manages email reporting and report history.
- `ExecutiveDashboard.gs` creates executive-level summaries and charts.
- `Metadata.gs` displays application version and feature metadata.

## Future Improvements

- Add trigger setup utilities for scheduled reports.
- Add screenshot assets for the README.
- Add role-based sheet protection rules for larger teams.
- Add import templates for CSV production logs.
- Add unit tests for more pure helper functions.
- Add optional monthly and quarterly reports.
