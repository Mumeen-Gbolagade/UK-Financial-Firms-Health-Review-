# UK Company Financial Health Review

A portfolio project that follows one sector from raw filed accounts to a short brief for decision-makers. It covers six UK-listed asset and wealth managers and compares their growth, profitability, liquidity and borrowing over three years.

![Dashboard](images/dashboard.png)

---

## Contents

1. [Objective](#1-objective)
2. [Project files](#2-project-files)
3. [Raw data](#3-raw-data)
4. [Data cleaning](#4-data-cleaning)
5. [Data exploration](#5-data-exploration)
6. [Data quality](#6-data-quality)
7. [Analysis](#7-analysis)
8. [Recommendations](#8-recommendations)
9. [Linking the findings to business strategy](#9-linking-the-findings-to-business-strategy)
10. [Stakeholder brief](#10-stakeholder-brief)
11. [Limitations](#11-limitations)
12. [How to view and run the project](#12-how-to-view-and-run-the-project)

---

## 1. Objective

**Main question:** which companies in one UK sector look financially stronger, and what risks should a lender or investor examine before relying on them?

**What the project delivers:**
- A like-for-like comparison of six companies across three financial years.
- Five ratios: revenue growth, profit margin, current ratio, debt-to-equity and interest coverage, with each formula explained in plain English.
- A clear list of the strengths, weaknesses and risks found, each backed by figures.
- An honest account of what the data cannot tell us.

**Why this sector:** asset and wealth managers split their balance sheets into current and non-current items, so all five ratios can be worked out. Banks and insurers do not report this way, which makes a current ratio meaningless for them.

**What this project is not:** it is not investment advice. It does not claim that any company is safe or likely to fail. Its purpose is to show where questions should be asked.

---

## 2. Project files

```
uk-financial-health-review/
├── README.md                     ← this file
├── data/
│   ├── raw_financials.csv        ← figures as printed in the accounts (inputs)
│   ├── companies.csv             ← company register and Companies House links
│   └── sources.csv               ← every source document, link and access date
├── src/
│   ├── analysis.py               ← import, cleaning, quality checks, ratios, SQL
│   ├── build_workbook.py         ← builds the Excel workbook with live formulas
│   ├── make_dashboard.py         ← draws the dashboard image
│   └── build_report.py           ← writes the short report
├── sql/
│   └── queries.sql               ← SQL checks run on the SQLite database
├── outputs/
│   ├── UK_Financial_Health_Review.xlsx   ← main workbook
│   ├── ratios.csv · data_quality_log.csv · sql_results.xlsx · financials.db
├── report/
│   ├── UK_Financial_Health_Review_Report.docx
│   └── UK_Financial_Health_Review_Report.pdf
└── images/
    └── dashboard.png
```

| File | Use it for |
|---|---|
| [outputs/UK_Financial_Health_Review.xlsx](outputs/UK_Financial_Health_Review.xlsx) | The full analysis. Start on the **Dashboard** sheet |
| [report/UK_Financial_Health_Review_Report.pdf](report/UK_Financial_Health_Review_Report.pdf) | The three-page written report ([editable .docx](report/UK_Financial_Health_Review_Report.docx)) |
| [src/analysis.py](src/analysis.py) | Seeing exactly how the data was cleaned and the ratios calculated |
| [sql/queries.sql](sql/queries.sql) | The SQL checks |

**The process from start to finish**

```
Companies House check  →  Raw figures typed in  →  Cleaning (£m, blanks kept)
→  Exploration  →  Quality log  →  Ratios (Excel formulas + Python + SQL)
→  Findings  →  Recommendations  →  Stakeholder brief
```

---

## 3. Raw data

**Companies selected**

| Company | Companies House no. | SIC code | Year end | Years analysed |
|---|---|---|---|---|
| AJ Bell plc | 04503206 | 70100, 82990 | 30 Sep | FY2023–FY2025 |
| Ashmore Group plc | 03675683 | not retrieved | 30 Jun | FY2023–FY2025 |
| Jupiter Fund Management plc | 06150195 | not retrieved | 31 Dec | FY2023–FY2025 |
| Liontrust Asset Management plc | 02954692 | 64999 | 31 Mar | FY2024–FY2026 |
| Polar Capital Holdings plc | 04235369 | 64205 | 31 Mar | FY2024–FY2026 |
| Premier Miton Group plc | 06306664 | 70100 | 30 Sep | FY2023–FY2025 |

**Where the figures came from.** Each company was identified on [Companies House](https://find-and-update.company-information.service.gov.uk/) by its company number. Companies House confirmed the names, year ends and, for most of them, the group accounts filed. The accounts PDFs held there could not be opened in the environment used for this project. The figures were therefore taken from the same audited group statements that each company publishes, either its annual report or its audited results announcement. All sources were accessed on 4–5 November 2026 and are listed in [data/sources.csv](data/sources.csv) and on the workbook's **Sources** sheet.

**What was collected.** Each company has two documents, and each document shows two years, giving four years per company and 24 company-years in total. For each year, 12 figures were recorded: revenue, net revenue, operating profit, finance income, finance costs, profit before tax, profit for the year, current assets, current liabilities, borrowings, lease liabilities and total equity.

**How it was recorded.** The figures were typed in exactly as printed, in the units shown (£000 or £m). Each row carries its source ID and a note, for example *"Leases = current 2,216 + non-current 9,842"*. In the workbook these inputs are shown in blue on the **Raw_Data** sheet.

---

## 4. Data cleaning

| Step | What was done | Why |
|---|---|---|
| Common unit | £000 figures divided by 1,000, so everything is in £ millions | Liontrust, Polar Capital, AJ Bell and Premier Miton report in £000; the others in £m |
| Same order | Rows sorted by company, then by year end | Revenue growth needs consecutive years for the same company |
| Line names matched | "Gross profit" (Liontrust, Premier Miton) and "Net income" (Polar Capital) placed in one *net revenue* column | Each firm labels similar lines differently |
| Debt defined once | Debt = borrowings + lease liabilities, with current and non-current parts added together | Most firms have no bank debt, so leases are the only debt |
| Blanks kept blank | Missing figures left empty. Ratios that need them show "n/a" | Avoids inventing numbers |
| Base year flagged | The earliest of the four years is used only to calculate the first growth figure | Keeps three full years of ratios for every company |

In Excel, the **Clean_GBPm** sheet does all of this with formulas linked to **Raw_Data**. The Python script `src/analysis.py` repeats the same steps.

---

## 5. Data exploration

Before calculating ratios, I looked at the overall shape of the data.

- **Scale varies a lot.** In the latest year, revenue ranged from £63.3m (Premier Miton) to £465.7m (Jupiter). The six firms together earned about £1.39bn.
- **More than half the sector was shrinking.** Of the 18 analysed company-years, 11 showed falling revenue, and the median revenue growth was −1.6%. Only AJ Bell and Polar Capital grew in every year.
- **Margins are widely spread.** Profit before tax margins ranged from a small loss (Liontrust FY2024, −0.3%) to 75.2% (Ashmore FY2025). The median was 25.6%.
- **Liquidity is high across the group.** Every current ratio was above 1.2x. The median was 2.5x, while Ashmore was an outlier at up to 10.7x.
- **Equity moves in different directions.** Equity grew at AJ Bell (£133.4m → £217.5m) but fell at Liontrust (£221.0m → £107.9m) and Ashmore (£966.8m → £790.8m).
- **Profit before tax and operating profit can differ a lot.** At Ashmore, profit before tax was up to 124% above operating profit, because of investment income. At Jupiter in FY2023 it was 89% below, because of a goodwill write-down. This is why both margins are shown.

---

## 6. Data quality

`src/analysis.py` checks every row and writes the results to [outputs/data_quality_log.csv](outputs/data_quality_log.csv) and to the workbook's **Data_Log** sheet.

| Check | Result |
|---|---|
| Company-years loaded | 24 (6 companies × 4 years) |
| Missing values | 18, all left blank and listed |
| Unusual figures flagged | 8 (losses, or profit before tax very different from operating profit) |
| Values estimated or filled in | **None** |
| Excel vs Python cross-check | All 144 ratio cells match |
| Excel formula errors | 0 out of 941 formulas |

**Where the gaps are**
- **Polar Capital:** no operating profit line and no separate lease liabilities, so its operating margin and debt-to-equity are n/a.
- **Ashmore:** no finance cost line, as the company states it has no debt, so interest coverage is n/a. Its FY2026 statements could not be retrieved.
- **AJ Bell:** no net revenue line. Figures come from its audited results announcements, because the full annual report could not be read past the auditor's report.
- **Companies House:** the pages for AJ Bell and Ashmore came from older cached copies, and Jupiter's page could not be opened.

---

## 7. Analysis

**How each ratio is calculated**

| Ratio | Formula in plain words | A higher result may mean | A lower result may mean |
|---|---|---|---|
| Revenue growth | This year's revenue ÷ last year's − 1 | Winning clients, or rising markets | Outflows, fee cuts or falling markets |
| Profit margin | Profit before tax ÷ revenue | Costs under control | High costs or one-off charges |
| Current ratio | Current assets ÷ current liabilities | Bigger short-term cushion | Less room to pay bills due within a year |
| Debt-to-equity | (Borrowings + leases) ÷ equity | More reliance on borrowed money | Funded mainly by shareholders |
| Interest coverage | (Profit before tax + finance costs) ÷ finance costs | Interest easily covered | Earnings struggle to cover interest |

I used profit before tax because it is the one profit measure all six companies report. Operating margin is shown alongside as a check. No ratio should be read on its own. A high margin can come from one-off investment gains, and a current ratio near 1x can simply reflect money passing between investors and funds. The **Formula_Guide** sheet explains each case.

**Latest-year results**

| Company | Year | Revenue £m | Growth | PBT margin | Current ratio | Debt / equity |
|---|---|---|---|---|---|---|
| AJ Bell | FY2025 | 317.8 | +18.0% | 43.4% | 3.65x | 0.06x |
| Polar Capital | FY2026 | 264.3 | +19.0% | 29.1% | 2.53x | n/a |
| Jupiter | FY2025 | 465.7 | +15.7% | 28.3% | 2.46x | 0.04x |
| Ashmore | FY2025 | 144.4 | −23.6% | 75.2% | 7.01x | 0.01x |
| Liontrust | FY2026 | 134.4 | −20.9% | 10.7% | 1.34x | 0.03x |
| Premier Miton | FY2025 | 63.3 | −1.1% | 3.7% | 1.27x | 0.02x |

**What the figures show**
- **AJ Bell (strongest on these measures).** Revenue grew from £218.2m to £317.8m, with growth of 33.2%, 23.5% and 18.0%. The margin rose from 40.2% to 43.4%. Its only debt is £12.1m of leases, against £217.5m of equity.
- **Polar Capital.** Revenue grew in every year (6.7%, 13.9%, 19.0%). The margin stayed between 23% and 29%, and the current ratio was steady at about 2.5x.
- **Jupiter (recovering).** The margin went from 2.3% to 21.9% to 28.3% after a £76.2m goodwill write-down in FY2023. It repaid £50m of notes in April 2025, so debt-to-equity fell from 0.12x to 0.04x.
- **Ashmore (strong balance sheet, shrinking business).** It has the highest current ratio and no debt. However, revenue fell from £193.2m to £144.4m, and £51.1m of finance income made up 47% of its FY2025 profit before tax.
- **Liontrust (weakening).** Revenue fell 45% from FY2023 to FY2026, with a small loss in FY2024, and equity halved.
- **Premier Miton (thin margin).** Revenue fell about 30% from FY2022, and the margin narrowed from 7.9% to 3.7%.

The SQL queries in [sql/queries.sql](sql/queries.sql) confirm these patterns. They list the latest-year snapshot, three-year averages, every year of falling revenue, every year with a margin under 5%, and every ratio that could not be calculated.

---

## 8. Recommendations

These are next steps for anyone relying on the analysis, not advice to buy, sell or lend.

**For a lender**
1. **Look beyond debt ratios.** Most of these firms only owe office leases, so debt-to-equity says little. Ask for covenant terms, guarantees, seed-capital commitments and any defined benefit pension position.
2. **Check what sits in current assets and liabilities.** At Liontrust (1.34x) and Premier Miton (1.27x), ask how much of the balance is money passing between investors and funds before treating the ratio as a warning sign.
3. **Test the capital buffer.** Compare each firm's regulatory capital surplus, published in its ICARA or Pillar 3 disclosures, with a scenario where revenue falls by a further 20%.

**For an equity investor**
1. **Separate fee profit from investment gains.** At Ashmore, look at the 30–40% operating margin rather than the 75% headline margin.
2. **Track assets under management and net flows every quarter.** Revenue in this sector follows these figures closely, and they are published well before the annual accounts.
3. **Check dividend cover where equity is falling.** Liontrust's equity fell from £221.0m to £107.9m in three years.

**For the analysis itself**
1. Add Ashmore's FY2026 figures once the full statements can be retrieved.
2. Add Polar Capital's lease liabilities from the notes to its annual report.
3. Re-run with the Companies House API key option (section 12) to download the filing records directly.

---

## 9. Linking the findings to business strategy

What each company's figures suggest it needs to manage, and what a stakeholder might ask management about:

| Company | Where it stands | Strategic question to raise |
|---|---|---|
| **AJ Bell** | Growing platform with high margins | Part of revenue is interest earned on client cash. How would margins hold up if UK interest rates fall? |
| **Polar Capital** | Steady growth, no borrowings | How dependent is revenue on a few technology-focused funds and on performance fees? |
| **Jupiter** | Recovery under way, debt repaid | Can cost discipline continue now that revenue is growing again, and are fund flows positive? |
| **Ashmore** | Very strong balance sheet, falling fees | How will it rebuild fee income, and how much of its large capital base could be returned to shareholders or invested instead? |
| **Liontrust** | Falling revenue and equity | Is the cost base shrinking as fast as revenue? Is the dividend sustainable from current earnings? |
| **Premier Miton** | Small and thin margin, no debt | Is the business large enough to stand alone, or does it need growth or a partner to reach a sustainable margin? |

**Across the sector:** revenue depends on market levels and on clients moving to cheaper passive funds. Firms with steady net inflows (AJ Bell, Polar Capital) are best placed. Firms losing assets need either lower costs or a new source of fees.

---

## 10. Stakeholder brief

*A one-minute summary for a credit committee, investment committee or board.*

> **Situation.** We compared six UK-listed asset and wealth managers on growth, profitability, liquidity and borrowing over their last three audited years.
>
> **What we found.** None of the six relies on bank debt, and all can cover their short-term bills. The real difference is direction. **AJ Bell** and **Polar Capital** grew revenue every year with healthy margins. **Jupiter** is recovering well after a one-off write-down and has paid off its borrowings. **Ashmore** has the strongest balance sheet, but its fee income keeps falling, and much of its profit comes from investment gains. **Liontrust** and **Premier Miton** have seen revenue fall sharply and margins become thin.
>
> **Why it matters.** In this sector, the risk is not that a firm cannot repay a loan today. It is that revenue keeps shrinking until margins, dividends and capital buffers come under strain.
>
> **What we suggest.** Treat AJ Bell, Polar Capital and Jupiter as the stronger names on current evidence. For Ashmore, Liontrust and Premier Miton, ask for recent fund flows, cost plans and regulatory capital surplus before committing.
>
> **Caveat.** These ratios point to questions to ask. They do not prove any firm is safe or at risk.

The full written version is in the [report](report/UK_Financial_Health_Review_Report.pdf), and the visual version is the **Dashboard** sheet.

---

## 11. Limitations

- **Different year ends.** The companies close their years in March, June, September and December, so the latest years run from June 2025 to March 2026. Market conditions differed across these dates.
- **Source route.** The figures come from each company's published audited statements, not from Companies House PDFs. Every figure is still traceable to a public source.
- **Missing figures.** There are 18 blank values (see section 6), mainly for Polar Capital and Ashmore.
- **Reporting differences.** The firms label their lines differently, and some prior years were re-presented. Every case is listed on the **Data_Log** sheet.
- **Listed versus private companies.** This approach is harder for private companies. Small firms can file abridged accounts with no profit and loss account, so growth and margin cannot be calculated at all.
- **Ratios are not a credit rating.** They leave out fund flows, management quality, regulation and anything that has happened since the accounts were published.

---

## 12. How to view and run the project

**Viewing the workbook.** Open `outputs/UK_Financial_Health_Review.xlsx` in Excel, LibreOffice or Google Sheets. Column headers are frozen, so every sheet scrolls cleanly.

| Sheet | What it shows |
|---|---|
| Dashboard | Latest-year scorecard, three-year trend charts and key points |
| Project_Story | The whole project on one page: objective → data → cleaning → exploration → quality → analysis → recommendations → strategy → stakeholder brief |
| Companies | Company register with Companies House links |
| Raw_Data | Figures as printed (blue = typed-in inputs) |
| Clean_GBPm | Cleaned figures in £m (formulas) |
| Ratios | All ratio formulas for every company-year |
| Formula_Guide | Plain-English explanation of each ratio |
| Data_Log | Missing values, unusual figures and comparability notes |
| SQL_Checks | Results of the SQL queries |
| Sources | Every source document, link and access date |

If you change a blue number on **Raw_Data**, the cleaned figures, ratios and dashboard all update.

**Running the analysis**

```bash
pip install pandas openpyxl matplotlib python-docx
python src/analysis.py        # cleaning, quality checks, ratios, SQLite + SQL
python src/build_workbook.py  # Excel workbook
python src/make_dashboard.py  # dashboard image
python src/build_report.py    # report
```

**Optional: download filing records from Companies House.** Get a free API key from the [Companies House developer hub](https://developer.company-information.service.gov.uk/), then run:

```bash
export CH_API_KEY=your_key_here
python src/analysis.py
```

The key is read from your environment only and is never saved in the project.
