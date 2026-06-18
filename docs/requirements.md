# Machine Efficiency Dashboard Requirements

## Goal

Build a Google Sheets and Apps Script dashboard for monitoring machine efficiency, section efficiency, machine status, and corrective actions.

## Sheets

The project should use these sheets:

1. DailyProduction
2. Dashboard
3. Corrective Actions
4. RCA Config

## DailyProduction Columns

- Date
- Section
- Machine
- Target Output
- Actual Output
- Downtime Minutes
- Remarks

## Dashboard Features

The Dashboard sheet should include:

- KPI cards
- Overall efficiency
- Average downtime
- Total target output
- Total actual output
- Machine status table
- Daily Machine Efficiency chart
- Daily Section Efficiency chart

## Chart Requirements

Charts must not render from empty or unstable dynamic ranges.

Chart source data should first be written into helper ranges on the Dashboard sheet.

The chart functions should run only after all dashboard data is already written.

## Corrective Actions

If machine efficiency is below the threshold, generate a corrective action row with:

- Date
- Machine
- Target Output
- Actual Output
- Efficiency %
- Issue
- Possible Cause
- Recommended Corrective Action

## RCA Config

Use RCA Config to map issues to possible causes and recommended corrective actions.

## Development Rules

- Keep functions separated by file responsibility.
- Do not put everything in one large script.
- Do not rewrite unrelated files when fixing one issue.
- Use helper ranges for charts.
- Clear old charts before inserting new charts.
- Keep the spreadsheet readable and beginner-friendly.
