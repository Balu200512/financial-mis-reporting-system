# Financial MIS Reporting System

## 1. Project Overview

The **Financial MIS Reporting System** is a portfolio project that demonstrates how finance teams can organize transaction data, compare actuals with budgets, prepare monthly financial summaries, and automate recurring management information system (MIS) reporting using **Excel VBA, SQL, and Power Query**.

It includes synthetic transaction data, account and cost-center master data, budget data, SQL schema and analytics queries, Power Query M scripts, Excel VBA automation code, and an Excel sample workbook.

The project is designed for learning and demonstration. It is not an audited accounting system and should not be used for statutory reporting without appropriate review and controls.

## 2. Project Objectives

- Organize financial transactions in a structured format.
- Build a relational data model for accounts, cost centers, transactions, and budgets.
- Prepare monthly revenue, cost, and expense summaries.
- Compare actual activity with budget amounts.
- Identify pending and reversed transactions for control review.
- Use Power Query to import, clean, and transform CSV data.
- Use Excel VBA to refresh reports and export the MIS summary to PDF.
- Use SQL to answer finance and management reporting questions.
- Create a repeatable workflow that can be refreshed when source data changes.

## 3. Technologies Used

| Technology | Purpose |
|---|---|
| Microsoft Excel | MIS workbook and summary sheets |
| Excel VBA | Refresh automation and PDF export |
| Power Query (M) | Importing, typing, trimming, and filtering source data |
| SQL | Database schema and financial analysis queries |
| SQLite | Local relational database for the reproducible Python runner |
| Python and Pandas | Optional validation and sample report generation |
| CSV | Portable source datasets |
| GitHub | Version control and portfolio sharing |

## 4. Project Features

- Financial transaction register.
- Chart of accounts and cost-center master tables.
- Monthly budget records.
- Posted, pending, and reversed transaction tracking.
- Monthly financial summary for revenue, cost of sales, operating expenses, and operating result.
- Actual-versus-budget query.
- Department expense analysis.
- Power Query templates for repeatable data preparation.
- VBA macros for refresh, summary-sheet creation, and PDF export.
- Python runner to validate sample data and generate reports.
- Excel workbook containing the sample datasets.

## 5. Sample Dataset

All records are **synthetic sample data** generated for this project. They do not represent a real company, employee, customer, supplier, or financial account.

| Dataset | Records | Purpose |
|---|---:|---|
| `transactions.csv` | 600 | Transaction-level financial activity |
| `accounts.csv` | 10 | Account codes, names, and account types |
| `cost_centers.csv` | 6 | Department, cost-center, and region mapping |
| `budgets.csv` | 720 | Monthly budget amounts by account and cost center |

The sample period is calendar year 2025. Currency is INR. Transaction statuses include Posted, Pending, and Reversed.

### Data files

- `data/transactions.csv`
- `data/accounts.csv`
- `data/cost_centers.csv`
- `data/budgets.csv`

The same source data is provided in `Financial_MIS_Sample_Data.xlsx` across separate worksheets.

## 6. Repository Structure

```text
financial-mis-reporting-system/
│
├── data/
│   ├── transactions.csv
│   ├── accounts.csv
│   ├── cost_centers.csv
│   └── budgets.csv
│
├── sql/
│   ├── schema.sql
│   └── analytics_queries.sql
│
├── vba/
│   └── FinancialMISAutomation.bas
│
├── power_query/
│   ├── Transactions.m
│   ├── Accounts.m
│   └── Budgets.m
│
├── docs/
│   └── Excel_Setup_Guide.md
│
├── reports/
│   └── Generated report outputs
│
├── Financial_MIS_Sample_Data.xlsx
├── mis_reporting.py
├── requirements.txt
├── .gitignore
└── README.md
```

Generated database and report files are created when the Python runner is executed. They are not required to be committed to GitHub.

## 7. Financial Data Model

The project uses four main data tables.

### 7.1 Transactions

Contains transaction-level details such as transaction ID, date, account code, cost center, description, reference number, amount, currency, status, and preparer.

### 7.2 Accounts

Maps each account code to its account name and type. The sample account types are Revenue, Cost of Sales, and Operating Expense.

### 7.3 Cost Centers

Maps cost-center IDs to department and region information. This supports departmental and regional reporting.

### 7.4 Budgets

Stores monthly budget amounts by account and cost center. These records can be compared with actual posted transactions for the same period and reporting dimensions.

### Relationships

- `transactions.account_code` → `accounts.account_code`
- `transactions.cost_center_id` → `cost_centers.cost_center_id`
- `budgets.account_code` → `accounts.account_code`
- `budgets.cost_center_id` → `cost_centers.cost_center_id`

Monthly reporting matches transaction month with the budget month. Use consistent date, account, and cost-center definitions when expanding the project.

## 8. SQL Database and Queries

### Database schema

`sql/schema.sql` defines tables, primary keys, foreign keys, validation checks, and indexes for SQLite.

### Analytics queries

`sql/analytics_queries.sql` includes examples for:

1. Monthly actuals by account type.
2. Monthly revenue, cost of sales, operating expenses, and operating result.
3. Actual-versus-budget analysis by month, department, and account.
4. Pending and reversed transaction controls.
5. Operating expense summary by department.

The SQL is SQLite-oriented. Other database platforms may require small syntax adjustments.

## 9. Power Query Automation

Power Query helps import and prepare source data without manually repeating the same cleanup steps.

Included M scripts:

- `power_query/Transactions.m`
- `power_query/Accounts.m`
- `power_query/Budgets.m`

The scripts demonstrate:

- Importing CSV files.
- Promoting the first row to column headers.
- Assigning data types.
- Trimming text fields.
- Filtering invalid or negative amounts where applicable.

### Important path setting

The M files contain a placeholder path:

`C:\YOUR_PROJECT_PATH\data\...`

Replace `C:\YOUR_PROJECT_PATH\` with the actual full folder path on your computer before refreshing the query. Alternatively, use Excel's Get Data interface to select the CSV file and adapt the transformation steps in Power Query Editor.

Detailed steps are in `docs/Excel_Setup_Guide.md`.

## 10. Excel VBA Automation

The VBA source is provided in:

`vba/FinancialMISAutomation.bas`

It contains three macros:

### `RefreshAllMIS`

- Refreshes workbook connections.
- Waits for supported asynchronous queries to finish.
- Records a refresh timestamp on the MIS Summary sheet.
- Displays a completion or error message.

### `ExportMISSummaryToPDF`

- Exports the `MIS Summary` worksheet to a PDF.
- Saves the PDF beside the workbook.
- Adds a timestamp to the output filename.

### `CreateMISSummarySheet`

- Creates the `MIS Summary` worksheet if it does not exist.
- Adds sample KPI formulas for posted transactions, posted amount, and pending transaction count.

### Macro-enabled workbook requirement

The supplied sample workbook is `.xlsx`, which cannot store executable VBA macros. To use the VBA code, save a copy as `.xlsm` in desktop Excel and import `FinancialMISAutomation.bas` using the VBA editor. See the setup guide before running macros.

Macros should only be enabled for files and code you trust.

## 11. Excel MIS Report Design

The recommended MIS Summary sheet includes:

- Report title and last refresh timestamp.
- Posted transaction count.
- Posted transaction amount.
- Pending transaction count.
- Monthly revenue trend.
- Cost of sales and operating expenses.
- Operating result.
- Actual-versus-budget variance.
- Department expense summary.
- Transaction status control table.

Use PivotTables, PivotCharts, formulas, or the Data Model depending on your reporting preference.

## 12. Key Financial Metrics

### Revenue

Sum of posted transactions mapped to accounts classified as Revenue.

### Cost of Sales

Sum of posted transactions mapped to Cost of Sales accounts.

### Operating Expenses

Sum of posted transactions mapped to Operating Expense accounts.

### Operating Result

For this sample project:

`Operating Result = Revenue - Cost of Sales - Operating Expenses`

### Budget Variance

For the supplied example queries:

`Variance = Actual Amount - Budget Amount`

A positive variance means actual amount exceeds budget; whether that is favorable depends on the account type. Higher revenue may be favorable, while higher expenses may be unfavorable.

### Transaction control

Pending and reversed records are reported separately for review. The sample monthly financial summary uses Posted transactions only.

**Finance note:** This is a simplified demonstration. Real MIS reporting requires approved account mappings, sign conventions, accounting periods, cut-off rules, accruals, tax treatment, reversals, currency conversion, and clear rules for favorable/unfavorable variance.

## 13. Run the Optional Python Validation and Reporting Script

Python is included as a reproducible helper to validate the datasets, create a local SQLite database, and export sample reports. The main business workflow remains Excel VBA, SQL, and Power Query.

### Prerequisites

- Python 3.10 or later.
- Visual Studio Code or another Python editor.
- Microsoft Excel desktop for VBA and Power Query functionality.

### Install dependencies

```bash
python -m venv venv
```

Windows Command Prompt:

```bat
venv\Scripts\activate
```

Install packages:

```bash
python -m pip install -r requirements.txt
```

Run:

```bash
python mis_reporting.py
```

The script validates source columns, checks transaction ID uniqueness, checks non-negative amounts and master-data references, builds `financial_mis.db`, and creates report files in `reports/`.

Expected outputs include:

- `reports/monthly_financial_summary.csv`
- `reports/monthly_financial_summary.md`
- `reports/transaction_control_summary.csv`
- `reports/monthly_financial_trend.png`
- `financial_mis.db`

## 14. Setup Instructions for Excel

1. Extract or clone the repository.
2. Open `Financial_MIS_Sample_Data.xlsx` in desktop Microsoft Excel.
3. Import the CSV files through **Data → Get Data → From File → From Text/CSV**, or use the included M scripts.
4. Update placeholder file paths in the Power Query scripts.
5. Load queries to worksheets or the Data Model.
6. Save a copy of the workbook as `Financial_MIS_Reporting_System.xlsm`.
7. Import `vba/FinancialMISAutomation.bas` using `Alt + F11` → **File → Import File**.
8. Run `CreateMISSummarySheet` once.
9. Configure the required query connections and verify the Transactions worksheet column layout.
10. Run `RefreshAllMIS` to refresh workbook data.
11. Run `ExportMISSummaryToPDF` to save a PDF copy of the MIS Summary.

Refer to `docs/Excel_Setup_Guide.md` for more details.

## 15. Business Questions Answered

This project helps demonstrate how to answer questions such as:

1. What is the monthly revenue trend?
2. What are the monthly cost of sales and operating expenses?
3. What is the operating result for each month?
4. Which departments record the highest operating expenses?
5. How much actual activity is recorded against the budget?
6. Which transactions remain pending?
7. How many transactions were reversed?
8. Which account types contribute to the financial result?
9. How can monthly MIS preparation be refreshed and exported more consistently?

## 16. Skills Demonstrated

- Financial MIS reporting concepts.
- Excel worksheet and report design.
- Excel VBA automation.
- Power Query and M language.
- SQL schema design and reporting queries.
- Data validation and transformation.
- Actual-versus-budget analysis.
- Monthly financial summaries.
- Transaction status controls.
- Python data processing and reporting.
- GitHub repository organization and documentation.

## 17. Limitations

- All sample records are synthetic.
- The included workbook is `.xlsx`; VBA requires a separate `.xlsm` copy in desktop Excel.
- The VBA macros and Power Query scripts must be imported/configured in Excel before use.
- Query paths must be adjusted to the local project folder.
- The SQL schema is SQLite-compatible, and other database systems may need syntax changes.
- This project is not an audited accounting solution or a replacement for a finance system.

## 18. Future Enhancements

- Add a finished `.xlsm` workbook with configured buttons and queries.
- Connect to a live SQL Server or PostgreSQL database.
- Add Power Pivot measures and a dimensional model.
- Add cash flow, balance sheet, and profit-and-loss reporting.
- Add automated email distribution of PDF reports.
- Add variance thresholds and exception alerts.
- Add month-end close checklists and reconciliation controls.
- Add role-based access, audit logs, and automated testing.

## 19. Author

**Balaji Mudugal**

GitHub: [Balu200512](https://github.com/Balu200512)

## 20. License

This project is provided for educational, portfolio, and demonstration purposes. Add a formal open-source license file if you intend to distribute the project under specific reuse terms.

---

⭐ If you find this project useful, consider starring the repository on GitHub.
