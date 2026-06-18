# Machine Efficiency Dashboard User Guide

This guide is for production supervisors and manufacturing managers who need to monitor machine performance, review issues, and create reports.

## Getting Started

Open the Google Sheet and use the **Machine Dashboard** menu at the top of the spreadsheet.

If this is the first time using the file, run these menu items in order:

1. **Setup Sheets**
2. **Load Sample Data**
3. **Refresh Dashboard**

After setup, the main sheets you will use are:

- `DailyProduction`
- `Dashboard`
- `Corrective Actions`
- `Daily Summary`

## Entering Production Data

Enter daily production records in the `DailyProduction` sheet.

Each row should include:

- Date
- Section
- Machine
- Target Output
- Actual Output
- Downtime Minutes
- Remarks

Use one row per machine production record. Try to keep machine names and section names consistent so filters and summaries work clearly.

## Refreshing Dashboard

Use **Machine Dashboard** > **Refresh Dashboard**.

Refreshing the dashboard updates:

- KPI cards
- Machine efficiency table
- Machine efficiency chart
- Section efficiency chart

If filters are selected on the Dashboard sheet, the dashboard will only use matching production records.

## Using Dashboard Filters

The Dashboard filter row includes:

- Date From
- Date To
- Section
- Machine

Leave date fields blank if you do not want a date filter.

Set Section or Machine to `All` to include everything.

If Date From is later than Date To, the system will show a warning and stop the refresh. This prevents incorrect reports.

## Understanding KPIs

### Total Target Output

The total planned output for the selected records.

### Total Actual Output

The total actual production output for the selected records.

### Overall Efficiency %

Shows how close actual output was to target output.

Formula:

```text
Actual Output / Target Output x 100
```

### Average Downtime

The average downtime minutes for the selected records.

### Total Machines

The number of machines included in the selected data.

### Total Sections

The number of production sections included in the selected data.

## Understanding Machine Status

Machine status helps you quickly see which machines need attention.

| Status | Meaning |
| --- | --- |
| Normal | Machine efficiency is at least 90%. |
| Warning | Machine efficiency is from 80% to 89.99%. |
| Critical | Machine efficiency is below 80%. |

Focus on Critical machines first, then Warning machines.

## Generating Corrective Actions

Use **Machine Dashboard** > **Generate Corrective Actions**.

The system reviews production records and creates action recommendations when machine performance needs attention.

Corrective actions are based on:

1. No Output
2. High Downtime
3. Low Efficiency

The `RCA Config` sheet controls the possible causes and recommended actions.

## Generating Daily Summary

Use **Machine Dashboard** > **Generate Daily Summary**.

This creates a `Daily Summary` sheet with:

- Report date
- Applied filters
- KPI summary
- Top performing machine
- Lowest performing machine

The Daily Summary respects the current Dashboard filters.

## Exporting Reports

Use **Machine Dashboard** > **Export Dashboard PDF** to export the Dashboard sheet.

Use **Machine Dashboard** > **Export Corrective Actions PDF** to export the Corrective Actions sheet.

PDF files are saved to Google Drive with timestamped filenames.

## Troubleshooting

### The Machine Dashboard menu is missing

Reload the spreadsheet. If it still does not appear, open Apps Script and make sure the project is saved.

### The dashboard is blank

Check that `DailyProduction` has data, then run **Refresh Dashboard**.

### Filters are not showing the expected results

Check the selected Date From, Date To, Section, and Machine values. Set Section and Machine back to `All` to remove those filters.

### A date warning appears

Make sure Date From is not later than Date To.

### Corrective Actions is empty

This can happen when no records meet the corrective action rules. It can also happen if DailyProduction has no data.

### PDF export does not work

Google may ask for script authorization the first time. Approve the requested permissions, then try again.
