\
"""Build the local Financial MIS database and export summary reports."""
from pathlib import Path
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
REPORTS = ROOT / "reports"
DB_PATH = ROOT / "financial_mis.db"

def main():
    REPORTS.mkdir(exist_ok=True)
    accounts = pd.read_csv(DATA / "accounts.csv")
    centers = pd.read_csv(DATA / "cost_centers.csv")
    transactions = pd.read_csv(DATA / "transactions.csv", parse_dates=["transaction_date"])
    budgets = pd.read_csv(DATA / "budgets.csv", parse_dates=["budget_month"])

    required = {
        "transactions": {"transaction_id", "transaction_date", "account_code", "cost_center_id", "amount", "status"},
        "accounts": {"account_code", "account_name", "account_type"},
        "cost_centers": {"cost_center_id", "department", "region"},
        "budgets": {"budget_id", "budget_month", "account_code", "cost_center_id", "budget_amount"},
    }
    frames = {"transactions": transactions, "accounts": accounts, "cost_centers": centers, "budgets": budgets}
    for name, cols in required.items():
        missing = cols - set(frames[name].columns)
        if missing:
            raise ValueError(f"{name} is missing columns: {sorted(missing)}")
    if transactions["transaction_id"].duplicated().any():
        raise ValueError("Duplicate transaction_id values found")
    if (transactions["amount"] < 0).any() or (budgets["budget_amount"] < 0).any():
        raise ValueError("Amounts must be non-negative")
    if not set(transactions["account_code"]).issubset(set(accounts["account_code"])):
        raise ValueError("transactions contains unknown account_code values")
    if not set(transactions["cost_center_id"]).issubset(set(centers["cost_center_id"])):
        raise ValueError("transactions contains unknown cost_center_id values")

    with sqlite3.connect(DB_PATH) as con:
        con.executescript((ROOT / "sql" / "schema.sql").read_text(encoding="utf-8"))
        for name, frame in frames.items():
            frame.to_sql(name, con, if_exists="replace", index=False)
        monthly = pd.read_sql_query("""
            SELECT substr(t.transaction_date,1,7) AS report_month,
                   SUM(CASE WHEN a.account_type='Revenue' THEN t.amount ELSE 0 END) AS revenue,
                   SUM(CASE WHEN a.account_type='Cost of Sales' THEN t.amount ELSE 0 END) AS cost_of_sales,
                   SUM(CASE WHEN a.account_type='Operating Expense' THEN t.amount ELSE 0 END) AS operating_expenses
            FROM transactions t JOIN accounts a ON a.account_code=t.account_code
            WHERE t.status='Posted'
            GROUP BY report_month ORDER BY report_month
        """, con)
    monthly["operating_result"] = monthly["revenue"] - monthly["cost_of_sales"] - monthly["operating_expenses"]
    monthly.to_csv(REPORTS / "monthly_financial_summary.csv", index=False)
    monthly.to_markdown(REPORTS / "monthly_financial_summary.md", index=False)

    status = transactions.groupby("status", as_index=False).agg(
        transaction_count=("transaction_id", "count"), total_amount=("amount", "sum"))
    status.to_csv(REPORTS / "transaction_control_summary.csv", index=False)

    if not monthly.empty:
        ax = monthly.plot(x="report_month", y=["revenue", "cost_of_sales", "operating_expenses"], marker="o", figsize=(10, 5))
        ax.set_title("Monthly Financial MIS - Posted Transactions")
        ax.set_xlabel("Month")
        ax.set_ylabel("Amount (INR)")
        ax.tick_params(axis="x", rotation=45)
        ax.figure.tight_layout()
        ax.figure.savefig(REPORTS / "monthly_financial_trend.png", dpi=150)
        plt.close(ax.figure)

    print("Financial MIS Reporting System completed successfully.")
    print(f"Transactions: {len(transactions):,}")
    print(f"Accounts: {len(accounts):,}")
    print(f"Cost centers: {len(centers):,}")
    print(f"Budgets: {len(budgets):,}")
    print(f"SQLite database: {DB_PATH}")
    print(f"Reports folder: {REPORTS}")

if __name__ == "__main__":
    main()
