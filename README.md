# Financial MIS Reporting System

## 1. Project Overview

The **Financial MIS Reporting System** is a financial reporting project developed using **Excel VBA, SQL, and Power Query**. It helps organize financial transactions, analyze monthly financial performance, compare actual amounts with budgets, and automate recurring MIS reporting activities.

The project includes sample financial datasets, SQL database scripts, Power Query transformation scripts, Excel VBA automation code, and an Excel workbook.

It also includes an optional Python utility to validate the sample data, create a local SQLite database, and generate financial summary reports.

## 2. Project Objectives

- Automate repetitive financial MIS reporting activities.
- Organize transaction data into structured tables.
- Analyze monthly revenue, cost of sales, and operating expenses.
- Compare actual financial activity against budget amounts.
- Track posted, pending, and reversed transactions.
- Clean and transform financial data using Power Query.
- Refresh reports and export summaries using Excel VBA.
- Use SQL queries to generate financial insights.
- Improve reporting consistency and reduce manual work.

## 3. Technologies Used

| Technology | Purpose |
|---|---|
| Microsoft Excel | Financial reporting and MIS summaries |
| Excel VBA | Report refresh automation and PDF export |
| Power Query | Data import, cleaning, and transformation |
| SQL | Financial analysis and database queries |
| SQLite | Local relational database |
| Python | Data validation and report generation |
| Pandas | Data manipulation and analysis |
| Matplotlib | Financial trend visualization |
| CSV | Source data storage |
| GitHub | Version control and project sharing |

## 4. Project Features

- Financial transaction register.
- Account master and cost-center master data.
- Monthly budget dataset.
- Monthly revenue and expense analysis.
- Actual-versus-budget comparison.
- Transaction status monitoring.
- Department-level expense analysis.
- Power Query data-cleaning scripts.
- VBA automation for refreshing data.
- Automated PDF export of the MIS Summary worksheet.
- SQL queries for financial reporting.
- Python-based data validation and report generation.
- Excel workbook containing the sample datasets.

## 5. Sample Dataset

This project uses **synthetic sample data** for educational and portfolio purposes. It does not represent actual company financial information.

| Dataset | Records | Description |
|---|---:|---|
| Transactions | 600 | Financial transaction records |
| Accounts | 10 | Account codes, names, and types |
| Cost Centers | 6 | Department and region information |
| Budgets | 720 | Monthly budget records |

The sample data covers calendar year 2025, and monetary amounts are represented in INR.

### Dataset files

- `data/transactions.csv`
- `data/accounts.csv`
- `data/cost_centers.csv`
- `data/budgets.csv`

The same data is also available in `Financial_MIS_Sample_Data.xlsx` across separate worksheets.

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
│   └── Generated reports
│
├── Financial_MIS_Sample_Data.xlsx
├── mis_reporting.py
├── requirements.txt
├── .gitignore
└── README.md
```

The SQLite database and reports are generated locally when the Python script is executed.

## 7. Financial Data Model

The project contains four main data tables.

### 7.1 Transactions

Stores transaction-level financial activity, including:

- Transaction ID
- Transaction date
- Account code
- Cost-center ID
- Description
- Reference number
- Amount
- Currency
- Transaction status
- Prepared by

### 7.2 Accounts

Contains account codes, account names, and account classifications.

The sample account types are:

- Revenue
- Cost of Sales
- Operating Expense

### 7.3 Cost Centers

Contains cost-center identifiers, department names, and region information.

This data supports departmental and regional reporting.

### 7.4 Budgets

Stores monthly budget amounts by account and cost center.

The budget data can be compared with actual posted transactions for the corresponding reporting period.

### Table relationships

- `transactions.account_code` → `accounts.account_code`
- `transactions.cost_center_id` → `cost_centers.cost_center_id`
- `budgets.account_code` → `accounts.account_code`
- `budgets.cost_center_id` → `cost_centers.cost_center_id`

## 8. SQL Database and Analytics

The `sql` folder contains the database schema and financial reporting queries.

### `schema.sql`

Defines the database tables, primary keys, foreign keys, validation rules, and indexes.

### `analytics_queries.sql`

Contains queries for:

1. Monthly financial actuals by account type.
2. Revenue, cost of sales, and operating expenses.
3. Monthly operating result.
4. Actual-versus-budget analysis.
5. Pending and reversed transaction summaries.
6. Department-wise operating expense analysis.

The supplied SQL is designed for SQLite. Other database systems may require minor syntax adjustments.

## 9. Power Query Automation

Power Query is used to import, clean, and transform the financial datasets.

The project includes:

- `power_query/Transactions.m`
- `power_query/Accounts.m`
- `power_query/Budgets.m`

The scripts demonstrate:

- Importing CSV files.
- Promoting headers.
- Setting appropriate data types.
- Trimming text fields.
- Filtering invalid records.
- Preparing structured data for reporting.

### Configure the file paths

The Power Query scripts contain a placeholder path:

```text
C:\YOUR_PROJECT_PATH\data\
```

Replace `C:\YOUR_PROJECT_PATH\` with the actual path to your extracted project folder before refreshing the queries.

Alternatively, import the CSV files through Excel's **Data → Get Data → From Text/CSV** menu and configure the transformations in Power Query Editor.

## 10. Excel VBA Automation

The VBA source code is available in:

`vba/FinancialMISAutomation.bas`

It contains three main macros.

### 10.1 RefreshAllMIS

- Refreshes workbook data connections.
- Waits for supported asynchronous queries to finish.
- Records the last refresh timestamp.
- Displays a completion or error message.

### 10.2 CreateMISSummarySheet

Creates an MIS Summary worksheet if it does not exist and adds sample KPI formulas for:

- Posted transaction count.
- Posted transaction amount.
- Pending transaction count.

### 10.3 ExportMISSummaryToPDF

Exports the MIS Summary worksheet to PDF format and saves the output alongside the workbook with a timestamped filename.

### Important VBA requirement

The supplied Excel workbook uses the `.xlsx` format, which does not store executable VBA macros.

To use the automation:

1. Open the workbook in desktop Microsoft Excel.
2. Save a copy as `Financial_MIS_Reporting_System.xlsm`.
3. Press `Alt + F11` to open the VBA editor.
4. Select **File → Import File**.
5. Select `vba/FinancialMISAutomation.bas`.
6. Confirm that the worksheet names and column positions match the VBA code.
7. Run `CreateMISSummarySheet` once.
8. Configure the Power Query connections and use `RefreshAllMIS`.
9. Run `ExportMISSummaryToPDF` to export the summary.

Only enable macros from code and files you trust.

## 11. Recommended MIS Dashboard

The Excel MIS Summary worksheet can contain the following sections.

### Key performance indicators

- Posted Transaction Count
- Posted Transaction Amount
- Pending Transaction Count
- Monthly Revenue
- Cost of Sales
- Operating Expenses
- Operating Result
- Actual-versus-Budget Variance

### Recommended charts

- Monthly revenue trend.
- Monthly expense trend.
- Revenue versus expenses.
- Department-wise operating expenses.
- Actual-versus-budget comparison.
- Transaction status distribution.

These visuals can be created using Excel charts, PivotTables, PivotCharts, and Power Query-loaded tables.

## 12. Key Financial Metrics

### Revenue

The total amount of posted transactions classified as Revenue.

### Cost of Sales

The total amount of posted transactions classified as Cost of Sales.

### Operating Expenses

The total amount of posted transactions classified as Operating Expense.

### Operating Result

For this demonstration project:

```text
Operating Result = Revenue - Cost of Sales - Operating Expenses
```

### Budget Variance

The project uses the following variance calculation:

```text
Variance = Actual Amount - Budget Amount
```

A positive variance means actual amounts exceed budget. Whether this is favorable depends on the account type: higher revenue may be favorable, while higher expenses may be unfavorable.

### Transaction Status

Transactions are classified as:

- `Posted`
- `Pending`
- `Reversed`

The sample monthly financial summary includes posted transactions only. Pending and reversed transactions are available for separate control analysis.

**Financial control note:** Real financial reporting requires approved account mappings, accounting-period rules, accruals, cut-off controls, tax treatment, reversal handling, and consistent sign conventions.

## 13. Installation and Setup

### Prerequisites

- Python 3.10 or later for the optional Python utility.
- Microsoft Excel desktop for VBA and Power Query.
- Visual Studio Code or another code editor.
- Git, if you want to manage the project through Git commands.

### Step 1: Download or clone the repository

```bash
git clone https://github.com/Balu200512/financial-mis-reporting-system.git
```

Move into the project directory:

```bash
cd financial-mis-reporting-system
```

Alternatively, download the ZIP from GitHub and extract it.

### Step 2: Create a Python virtual environment

```bash
python -m venv venv
```

For Windows Command Prompt:

```bat
venv\Scripts\activate
```

### Step 3: Install dependencies

```bash
python -m pip install -r requirements.txt
```

### Step 4: Run the Python utility

```bash
python mis_reporting.py
```

The script validates the datasets, creates the SQLite database, and generates summary reports and charts.

## 14. Generated Reports

After successfully running `mis_reporting.py`, the project creates:

```text
financial_mis.db

reports/
├── monthly_financial_summary.csv
├── monthly_financial_summary.md
├── transaction_control_summary.csv
└── monthly_financial_trend.png
```

### Report descriptions

- `monthly_financial_summary.csv` – monthly revenue, cost of sales, expenses, and operating result.
- `monthly_financial_summary.md` – readable monthly summary.
- `transaction_control_summary.csv` – transaction counts and amounts by status.
- `monthly_financial_trend.png` – visualization of monthly revenue and expenses.
- `financial_mis.db` – local SQLite database containing the financial data.

These are generated outputs and do not need to be created manually.

## 15. Business Questions Answered

This project demonstrates how to answer the following financial reporting questions:

1. What is the monthly revenue trend?
2. What are the total cost of sales and operating expenses?
3. What is the operating result for each month?
4. Which departments record the highest operating expenses?
5. How do actual amounts compare with budget amounts?
6. How many transactions are pending?
7. How many transactions have been reversed?
8. Which account types contribute to the financial result?
9. How can monthly reporting be refreshed more consistently?
10. How can an MIS summary be exported as a PDF?

## 16. Skills Demonstrated

- Financial MIS reporting.
- Microsoft Excel.
- Excel VBA automation.
- Power Query and M language.
- SQL database design.
- Financial analysis.
- Budget-versus-actual reporting.
- Data cleaning and validation.
- Monthly financial reporting.
- Excel charting and PivotTables.
- Python and Pandas.
- Data visualization.
- Git and GitHub.

## 17. Limitations

- The sample data is synthetic.
- The provided Excel workbook is not a completed macro-enabled `.xlsm` workbook.
- VBA macros must be imported and configured in desktop Excel.
- Power Query file paths must be updated for the local computer.
- The SQL scripts are SQLite-oriented.
- The generated reports are based on the supplied sample data.
- This is a portfolio and learning project, not an audited accounting system.

## 18. Future Enhancements

Possible improvements include:

- Delivering a configured macro-enabled Excel workbook.
- Connecting directly to SQL Server or PostgreSQL.
- Adding profit-and-loss and cash-flow reports.
- Adding balance-sheet reporting.
- Creating automated email distribution for PDF reports.
- Adding budget threshold alerts.
- Adding reconciliation and month-end closing checklists.
- Implementing audit logs and automated testing.
- Integrating Power Pivot and advanced financial measures.

## 19. Author

**Balaji Mudugal**

GitHub: [Balu200512](https://github.com/Balu200512)

## 20. License

This project is provided for educational, portfolio, and demonstration purposes. Add an open-source license file if you intend to distribute the project under specific reuse terms.

---

⭐ If you find this project useful, consider starring the repository on GitHub.
