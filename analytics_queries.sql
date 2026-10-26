-- 1. Monthly actuals by account type (posted transactions only)
SELECT substr(t.transaction_date, 1, 7) AS report_month,
       a.account_type,
       SUM(CASE WHEN a.account_type = 'Revenue' THEN t.amount ELSE 0 END) AS revenue,
       SUM(CASE WHEN a.account_type = 'Cost of Sales' THEN t.amount ELSE 0 END) AS cost_of_sales,
       SUM(CASE WHEN a.account_type = 'Operating Expense' THEN t.amount ELSE 0 END) AS operating_expenses
FROM transactions t
JOIN accounts a ON a.account_code = t.account_code
WHERE t.status = 'Posted'
GROUP BY report_month, a.account_type
ORDER BY report_month, a.account_type;

-- 2. Monthly financial summary
SELECT substr(t.transaction_date, 1, 7) AS report_month,
       SUM(CASE WHEN a.account_type = 'Revenue' THEN t.amount ELSE 0 END) AS revenue,
       SUM(CASE WHEN a.account_type = 'Cost of Sales' THEN t.amount ELSE 0 END) AS cost_of_sales,
       SUM(CASE WHEN a.account_type = 'Operating Expense' THEN t.amount ELSE 0 END) AS operating_expenses,
       SUM(CASE WHEN a.account_type = 'Revenue' THEN t.amount ELSE 0 END)
       - SUM(CASE WHEN a.account_type = 'Cost of Sales' THEN t.amount ELSE 0 END)
       - SUM(CASE WHEN a.account_type = 'Operating Expense' THEN t.amount ELSE 0 END) AS operating_result
FROM transactions t
JOIN accounts a ON a.account_code = t.account_code
WHERE t.status = 'Posted'
GROUP BY report_month
ORDER BY report_month;

-- 3. Actual vs budget by month, department and account type
WITH actuals AS (
    SELECT substr(t.transaction_date,1,7) AS report_month,
           t.cost_center_id, t.account_code, SUM(t.amount) AS actual_amount
    FROM transactions t
    WHERE t.status='Posted'
    GROUP BY substr(t.transaction_date,1,7), t.cost_center_id, t.account_code
)
SELECT a.report_month, cc.department, ac.account_type, ac.account_name,
       a.actual_amount, b.budget_amount,
       a.actual_amount - COALESCE(b.budget_amount,0) AS variance
FROM actuals a
JOIN accounts ac ON ac.account_code=a.account_code
JOIN cost_centers cc ON cc.cost_center_id=a.cost_center_id
LEFT JOIN budgets b ON substr(b.budget_month,1,7)=a.report_month
    AND b.cost_center_id=a.cost_center_id AND b.account_code=a.account_code
ORDER BY a.report_month, cc.department, ac.account_type;

-- 4. Pending and reversed transaction control
SELECT status, COUNT(*) AS transaction_count, SUM(amount) AS total_amount
FROM transactions
GROUP BY status
ORDER BY status;

-- 5. Department expense summary
SELECT cc.department, SUM(t.amount) AS operating_expenses
FROM transactions t
JOIN accounts a ON a.account_code=t.account_code
JOIN cost_centers cc ON cc.cost_center_id=t.cost_center_id
WHERE t.status='Posted' AND a.account_type='Operating Expense'
GROUP BY cc.department
ORDER BY operating_expenses DESC;
