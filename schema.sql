-- Financial MIS Reporting System schema (SQLite-compatible)
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS accounts (
    account_code TEXT PRIMARY KEY,
    account_name TEXT NOT NULL,
    account_type TEXT NOT NULL CHECK (account_type IN ('Revenue','Cost of Sales','Operating Expense'))
);

CREATE TABLE IF NOT EXISTS cost_centers (
    cost_center_id TEXT PRIMARY KEY,
    cost_center_name TEXT NOT NULL,
    department TEXT NOT NULL,
    region TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS transactions (
    transaction_id TEXT PRIMARY KEY,
    transaction_date TEXT NOT NULL,
    account_code TEXT NOT NULL,
    cost_center_id TEXT NOT NULL,
    description TEXT,
    reference_no TEXT,
    amount REAL NOT NULL CHECK (amount >= 0),
    currency TEXT NOT NULL DEFAULT 'INR',
    status TEXT NOT NULL CHECK (status IN ('Posted','Pending','Reversed')),
    prepared_by TEXT,
    FOREIGN KEY (account_code) REFERENCES accounts(account_code),
    FOREIGN KEY (cost_center_id) REFERENCES cost_centers(cost_center_id)
);

CREATE TABLE IF NOT EXISTS budgets (
    budget_id TEXT PRIMARY KEY,
    budget_month TEXT NOT NULL,
    account_code TEXT NOT NULL,
    cost_center_id TEXT NOT NULL,
    budget_amount REAL NOT NULL CHECK (budget_amount >= 0),
    currency TEXT NOT NULL DEFAULT 'INR',
    FOREIGN KEY (account_code) REFERENCES accounts(account_code),
    FOREIGN KEY (cost_center_id) REFERENCES cost_centers(cost_center_id)
);

CREATE INDEX IF NOT EXISTS idx_transactions_date ON transactions(transaction_date);
CREATE INDEX IF NOT EXISTS idx_transactions_account ON transactions(account_code);
CREATE INDEX IF NOT EXISTS idx_transactions_cost_center ON transactions(cost_center_id);
