-- UK Company Financial Health Review: SQL checks
-- These run against the SQLite database built by src/analysis.py (outputs/financials.db).
-- Table "ratios" holds one row per company and financial year, in GBP millions.

-- name: latest_year_snapshot
-- The most recent financial year for each company, with its key ratios.
SELECT r.company, r.fy_label, r.revenue_m, r.revenue_growth, r.pbt_margin,
       r.current_ratio, r.debt_to_equity, r.interest_cover
FROM ratios r
JOIN (SELECT company, MAX(fy_end) AS last_end FROM ratios GROUP BY company) l
  ON r.company = l.company AND r.fy_end = l.last_end
ORDER BY r.pbt_margin DESC;

-- name: three_year_average
-- Average of each ratio over the three analysed years (base year excluded).
SELECT company,
       ROUND(AVG(revenue_growth), 4) AS avg_revenue_growth,
       ROUND(AVG(pbt_margin), 4)     AS avg_pbt_margin,
       ROUND(AVG(current_ratio), 2)  AS avg_current_ratio,
       ROUND(MIN(current_ratio), 2)  AS lowest_current_ratio
FROM ratios
WHERE in_analysis = 1
GROUP BY company
ORDER BY avg_pbt_margin DESC;

-- name: revenue_declines
-- Every analysed year where revenue fell.
SELECT company, fy_label, revenue_m, ROUND(revenue_growth, 4) AS revenue_growth
FROM ratios
WHERE in_analysis = 1 AND revenue_growth < 0
ORDER BY revenue_growth;

-- name: loss_or_thin_margin_years
-- Years with a loss before tax or a profit before tax margin under 5%.
SELECT company, fy_label, profit_before_tax_m, ROUND(pbt_margin, 4) AS pbt_margin
FROM ratios
WHERE in_analysis = 1 AND pbt_margin < 0.05
ORDER BY pbt_margin;

-- name: missing_inputs
-- Count of ratio results that could not be calculated, by company.
SELECT company,
       SUM(CASE WHEN debt_to_equity IS NULL THEN 1 ELSE 0 END) AS missing_debt_to_equity,
       SUM(CASE WHEN interest_cover IS NULL THEN 1 ELSE 0 END) AS missing_interest_cover
FROM ratios
WHERE in_analysis = 1
GROUP BY company;
