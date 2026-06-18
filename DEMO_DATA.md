# Demo Data

## Purpose

The demo data helps another user recreate the project walkthrough without needing real production records. It shows how the dashboard behaves across normal, warning, and critical machine performance.

## Sample Sections

Use these section names:

- Assembly
- Cutting
- Packaging
- Inspection
- Finishing

## Sample Machines

Use these machine names:

- Machine A1
- Machine A2
- Machine C1
- Machine C2
- Machine P1
- Machine P2
- Machine I1
- Machine I2
- Machine F1
- Machine F2

## Sample Production Data

Create these rows in `DailyProduction` after the header row:

| Date | Section | Machine | Target Output | Actual Output | Downtime Minutes | Remarks |
| --- | --- | --- | ---: | ---: | ---: | --- |
| 2026-06-01 | Assembly | Machine A1 | 500 | 470 | 18 | Minor setup delay |
| 2026-06-02 | Assembly | Machine A2 | 520 | 515 | 5 | Normal run |
| 2026-06-03 | Cutting | Machine C1 | 450 | 390 | 42 | Blade change needed |
| 2026-06-04 | Cutting | Machine C2 | 460 | 455 | 8 | Normal run |
| 2026-06-05 | Packaging | Machine P1 | 600 | 580 | 12 | Short material wait |
| 2026-06-06 | Packaging | Machine P2 | 610 | 0 | 90 | Machine unavailable |
| 2026-06-07 | Inspection | Machine I1 | 350 | 340 | 7 | Normal run |
| 2026-06-08 | Inspection | Machine I2 | 360 | 300 | 35 | Sensor alignment issue |
| 2026-06-09 | Finishing | Machine F1 | 420 | 415 | 10 | Normal run |
| 2026-06-10 | Finishing | Machine F2 | 430 | 385 | 28 | Operator workflow check |

## Recreating the Demo

1. Open the Google Sheet.
2. Run `Machine Dashboard` > `Setup Sheets`.
3. Run `Machine Dashboard` > `Load Sample Data`.
4. Run `Machine Dashboard` > `Setup Data Validation`.
5. Run `Machine Dashboard` > `Refresh Dashboard`.
6. Run `Machine Dashboard` > `Generate Corrective Actions`.
7. Run `Machine Dashboard` > `Generate Daily Summary`.
8. Run `Machine Dashboard` > `Generate Historical Trends`.
9. Run `Machine Dashboard` > `Refresh Executive Dashboard`.

After these steps, the workbook is ready for screenshots, portfolio demos, or academic presentation.
