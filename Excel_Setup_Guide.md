# Excel VBA and Power Query Setup Guide

## Important
The repository provides a normal `.xlsx` sample workbook and VBA source module. Excel does not execute VBA from a plain `.xlsx` file. To use macros, save a copy as **Excel Macro-Enabled Workbook (`.xlsm`)**, then import `vba/FinancialMISAutomation.bas`.

## 1. Open the workbook
Open `Financial_MIS_Sample_Data.xlsx` in desktop Microsoft Excel.

## 2. Import the VBA module
1. Save a copy as `Financial_MIS_Reporting_System.xlsm`.
2. Enable the Developer tab if needed: File > Options > Customize Ribbon > Developer.
3. Press `Alt + F11` to open the VBA editor.
4. Select File > Import File and choose `vba/FinancialMISAutomation.bas`.
5. Ensure the workbook has a `Transactions` sheet whose columns match the sample workbook.
6. Run `CreateMISSummarySheet` once, then use `RefreshAllMIS` to refresh connections and `ExportMISSummaryToPDF` to export the summary.

Only enable macros in workbooks and files you trust.

## 3. Configure Power Query
1. In Excel, select Data > Get Data > From File > From Text/CSV.
2. Import `data/transactions.csv`, `data/accounts.csv`, `data/cost_centers.csv`, and `data/budgets.csv`.
3. For reusable transformations, open Power Query Editor, select Home > Advanced Editor, and adapt the matching `.m` file in `power_query/`.
4. Replace `C:\YOUR_PROJECT_PATH\` in each query with the actual full path to your extracted project.
5. Load queries to worksheets or the Data Model.
6. Set relationships using `account_code`, `cost_center_id`, and the month/date fields where appropriate.
7. Use Data > Refresh All to refresh data.

## 4. Suggested Excel MIS layout
Create an `MIS Summary` sheet with KPI cards for posted amount, posted transaction count, pending count, and budget variance. Add monthly revenue/cost/expense charts and a department expense summary using PivotTables or the Data Model.

## 5. Financial control notes
This is an educational sample. Confirm account mapping, sign conventions, cut-off rules, tax treatment, reversals, and budget allocation logic with a finance owner before using any real business data.
