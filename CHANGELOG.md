# Changelog

## Phase 7: Daily Summary and PDF Export

- Added Daily Summary reporting based on current Dashboard filters.
- Added Dashboard PDF export.
- Added Corrective Actions PDF export.
- Added reporting menu items to the Machine Dashboard menu.

## Phase 6: Dashboard Filtering

- Added Dashboard filter controls for Date From, Date To, Section, and Machine.
- Added Section and Machine dropdowns based on DailyProduction data.
- Updated dashboard refresh so KPIs, machine table, and charts use filtered production rows.
- Added defensive validation for invalid dates and reversed date ranges.

## Phase 5: Charts and Visualizations

- Added chart helper ranges on the Dashboard sheet.
- Added Machine Efficiency chart data in `P:Q`.
- Added Section Efficiency chart data in `S:T`.
- Added dashboard chart refresh behavior that removes old charts before rebuilding.
- Improved charts with horizontal bar chart formatting.

## Phase 4: Dashboard UI and Formatting

- Reworked the Dashboard sheet into a professional management dashboard.
- Added merged dashboard title styling.
- Added KPI card layout.
- Added conditional coloring for Overall Efficiency.
- Styled the Machine Efficiency Status table.
- Added status color coding for Normal, Warning, and Critical machines.

## Phase 3: Corrective Action Generation

- Added Corrective Actions sheet generation.
- Added RCA Config support.
- Added rule-based issue detection for No Output, High Downtime, and Low Efficiency.
- Added possible cause and recommended action mapping.

## Phase 2: KPI Calculations

- Added dashboard KPI calculations.
- Added total target output, total actual output, overall efficiency, average downtime, total machines, and total sections.
- Added machine-level efficiency aggregation.
- Added machine status rules.

## Phase 1: Sheet Setup and Sample Data

- Added setup for required sheets.
- Added DailyProduction structure.
- Added Corrective Actions and RCA Config sheet setup.
- Added sample production data.
- Added the Machine Dashboard menu foundation.
