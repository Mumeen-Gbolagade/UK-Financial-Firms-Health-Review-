"""
UK Company Financial Health Review - data import, cleaning and ratio analysis.

Steps
  1. (Optional) pull company profiles and accounts filing lists from the
     Companies House API. Needs your own free API key in the CH_API_KEY
     environment variable. No key is stored in this project.
  2. Load the figures extracted from each company's audited accounts
     (data/raw_financials.csv) and the company register (data/companies.csv).
  3. Check and clean: convert everything to GBP millions, log missing values
     and unusual figures. Nothing is filled in or estimated.
  4. Calculate revenue growth, profit margin, current ratio, debt-to-equity
     and interest coverage.
  5. Load the results into SQLite and run the checks in sql/queries.sql.

Run from the project folder:
    python src/analysis.py
"""

from __future__ import annotations

import json
import os
import sqlite3
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "outputs"
SQL_FILE = ROOT / "sql" / "queries.sql"

MONEY_COLS = [
    "revenue", "net_revenue", "operating_profit", "finance_income",
    "finance_costs", "profit_before_tax", "profit_for_year",
    "current_assets", "current_liabilities", "borrowings",
    "lease_liabilities", "total_equity",
]
UNIT_FACTOR = {"GBP000": 1 / 1000, "GBPm": 1.0}


# ---------------------------------------------------------------- step 1
def fetch_companies_house(numbers: list[str]) -> None:
    """Save each company's profile and accounts filing list as JSON.

    Skipped quietly when CH_API_KEY is not set. The key is read from the
    environment only and is never written to disk.
    """
    key = os.environ.get("CH_API_KEY")
    if not key:
        print("Step 1: CH_API_KEY not set, skipping the Companies House download.")
        return
    import requests  # imported here so the rest runs without it

    folder = DATA / "companies_house"
    folder.mkdir(exist_ok=True)
    base = "https://api.company-information.service.gov.uk"
    for number in numbers:
        profile = requests.get(f"{base}/company/{number}", auth=(key, ""), timeout=30)
        filings = requests.get(
            f"{base}/company/{number}/filing-history",
            params={"category": "accounts", "items_per_page": 20},
            auth=(key, ""), timeout=30,
        )
        for name, resp in (("profile", profile), ("accounts_filings", filings)):
            if resp.ok:
                (folder / f"{number}_{name}.json").write_text(json.dumps(resp.json(), indent=2))
            else:
                print(f"  {number} {name}: HTTP {resp.status_code}")
    print(f"Step 1: Companies House data saved to {folder}")


# ---------------------------------------------------------------- step 2-3
def load_and_clean() -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = pd.read_csv(DATA / "raw_financials.csv", dtype={"company_number": str})
    companies = pd.read_csv(DATA / "companies.csv", dtype={"company_number": str})

    unknown_units = set(raw["units"]) - set(UNIT_FACTOR)
    if unknown_units:
        raise ValueError(f"Unexpected units: {unknown_units}")

    clean = raw.copy()
    factor = clean["units"].map(UNIT_FACTOR)
    for col in MONEY_COLS:
        clean[col + "_m"] = clean[col] * factor
    clean["fy_end"] = pd.to_datetime(clean["fy_end"])
    clean = clean.sort_values(["company", "fy_end"]).reset_index(drop=True)
    clean = clean.merge(companies[["company", "short_name"]], on="company", how="left")

    # Each company has four years; the earliest is only used as the base for growth.
    clean["year_rank"] = clean.groupby("company")["fy_end"].rank(ascending=False)
    clean["in_analysis"] = (clean["year_rank"] <= 3).astype(int)

    log = []
    for _, r in clean.iterrows():
        for col in MONEY_COLS:
            if pd.isna(r[col]):
                log.append((r["company"], r["fy_label"], "Missing value", col,
                            "Not shown in the retrieved statements; left blank, not estimated."))
        if r["profit_before_tax"] < 0:
            log.append((r["company"], r["fy_label"], "Unusual figure", "profit_before_tax",
                        "Loss before tax."))
        if pd.notna(r["operating_profit"]) and r["operating_profit"] != 0:
            gap = (r["profit_before_tax"] - r["operating_profit"]) / abs(r["operating_profit"])
            if abs(gap) > 0.30:
                log.append((r["company"], r["fy_label"], "Unusual figure", "profit_before_tax",
                            f"Profit before tax differs from operating profit by {gap:+.0%} "
                            "(items below operating profit, e.g. finance income or impairment)."))
    quality = pd.DataFrame(log, columns=["company", "fy_label", "type", "field", "detail"])
    return clean, quality


# ---------------------------------------------------------------- step 4
def ratio_or_none(num: pd.Series, den: pd.Series) -> pd.Series:
    """Divide, returning None when either side is missing or the divisor is zero."""
    out = num / den
    out[(den == 0) | den.isna() | num.isna()] = None
    return out


def calculate_ratios(clean: pd.DataFrame) -> pd.DataFrame:
    df = clean.copy()
    prev_rev = df.groupby("company")["revenue_m"].shift(1)
    df["revenue_growth"] = ratio_or_none(df["revenue_m"], prev_rev) - 1
    df["pbt_margin"] = ratio_or_none(df["profit_before_tax_m"], df["revenue_m"])
    df["operating_margin"] = ratio_or_none(df["operating_profit_m"], df["revenue_m"])
    df["current_ratio"] = ratio_or_none(df["current_assets_m"], df["current_liabilities_m"])
    debt = df["borrowings_m"] + df["lease_liabilities_m"]  # NaN if leases missing
    df["total_debt_m"] = debt
    df["debt_to_equity"] = ratio_or_none(debt, df["total_equity_m"])
    ebit = df["profit_before_tax_m"] + df["finance_costs_m"]
    df["interest_cover"] = ratio_or_none(ebit, df["finance_costs_m"])
    return df


def kpi_scorecard(ratios: pd.DataFrame) -> pd.DataFrame:
    """Latest-year KPIs per company with a green / amber / red status.

    Bands come from data/kpi_thresholds.csv. They are reading aids chosen by
    the analyst, not industry rules, and the score is a screen, not a rating.
    """
    bands = pd.read_csv(DATA / "kpi_thresholds.csv")
    df = ratios.sort_values(["company", "fy_end"])
    latest = df.groupby("company").tail(1).set_index("company")
    base_equity = df.groupby("company")["total_equity_m"].first()   # base year, 3 years earlier
    latest["equity_change_3y"] = latest["total_equity_m"] / base_equity - 1
    points = {"Green": 2, "Amber": 1, "Red": 0}
    rows = []
    for company, r in latest.iterrows():
        out = {"company": company, "short_name": r["short_name"], "fy_label": r["fy_label"]}
        got, possible = 0, 0
        for b in bands.itertuples():
            v = r[b.kpi]
            if pd.isna(v):
                status = "No data"
            elif b.better == "higher":
                status = "Green" if v >= b.green_from else "Amber" if v >= b.amber_from else "Red"
            else:
                status = "Green" if v <= b.green_from else "Amber" if v <= b.amber_from else "Red"
            out[b.kpi] = v
            out[b.kpi + "_status"] = status
            if status in points:
                got += points[status]
                possible += 2
        out["kpis_scored"] = possible // 2
        out["screening_score"] = got / possible if possible else None
        rows.append(out)
    return pd.DataFrame(rows).sort_values("screening_score", ascending=False)


# ---------------------------------------------------------------- step 5
def run_sql(ratios: pd.DataFrame) -> dict[str, pd.DataFrame]:
    db_path = OUT / "financials.db"
    if db_path.exists():
        db_path.unlink()
    cols = ["company", "short_name", "fy_label", "fy_end", "in_analysis", "revenue_m",
            "profit_before_tax_m", "revenue_growth", "pbt_margin", "operating_margin",
            "current_ratio", "debt_to_equity", "interest_cover"]
    table = ratios[cols].copy()
    table["fy_end"] = table["fy_end"].dt.strftime("%Y-%m-%d")
    with sqlite3.connect(db_path) as con:
        table.to_sql("ratios", con, index=False)
        results = {}
        for block in SQL_FILE.read_text().split("-- name: ")[1:]:
            name, _, body = block.partition("\n")
            results[name.strip()] = pd.read_sql_query(body, con)
    return results


def main() -> None:
    OUT.mkdir(exist_ok=True)
    companies = pd.read_csv(DATA / "companies.csv", dtype={"company_number": str})
    fetch_companies_house(companies["company_number"].tolist())

    clean, quality = load_and_clean()
    print(f"Step 2-3: {len(clean)} company-years loaded, {len(quality)} data notes logged.")

    ratios = calculate_ratios(clean)
    ratios.to_csv(OUT / "ratios.csv", index=False, float_format="%.4f")
    quality.to_csv(OUT / "data_quality_log.csv", index=False)
    print("Step 4: ratios written to outputs/ratios.csv")
    kpis = kpi_scorecard(ratios)
    kpis.to_csv(OUT / "kpi_scorecard.csv", index=False, float_format="%.4f")
    print("        KPI scorecard written to outputs/kpi_scorecard.csv")
    print(kpis[["short_name", "fy_label", "kpis_scored", "screening_score"]].to_string(index=False))

    results = run_sql(ratios)
    with pd.ExcelWriter(OUT / "sql_results.xlsx") as xw:
        for name, frame in results.items():
            frame.to_excel(xw, sheet_name=name[:31], index=False)
    print("Step 5: SQL checks written to outputs/sql_results.xlsx")
    print("\nLatest year snapshot:")
    with pd.option_context("display.width", 140, "display.float_format", "{:.3f}".format):
        print(results["latest_year_snapshot"].to_string(index=False))


if __name__ == "__main__":
    main()
