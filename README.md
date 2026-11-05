# UK Company Financial Health Review

A portfolio project that follows one sector from raw filed accounts to a short brief for decision-makers. It covers six UK-listed asset and wealth managers and compares their growth, profitability, liquidity and borrowing over three years.

**Quick links:** [Excel workbook](outputs/UK_Financial_Health_Review.xlsx) · [PDF report](report/UK_Financial_Health_Review_Report.pdf) · [KPI scorecard](outputs/kpi_scorecard.csv) · [Python code](src) · [SQL](sql/queries.sql) · [Screenshots](images)

![Dashboard](images/dashboard.png)

**At a glance (latest audited year)**

| Company | Year | Revenue growth | Profit margin | Current ratio | KPI screening score |
|---|---|---|---|---|---|
| AJ Bell | FY2025 | 🟢 +18.0% | 🟢 43.4% | 🟢 3.65x | **100%** (7 of 7 KPIs) |
| Jupiter | FY2025 | 🟢 +15.7% | 🟢 28.3% | 🟢 2.46x | **100%** (7 of 7) |
| Polar Capital | FY2026 | 🟢 +19.0% | 🟢 29.1% | 🟢 2.53x | **100%** (5 of 7) |
| Ashmore | FY2025 | 🔴 −23.6% | 🟢 75.2% | 🟢 7.01x | **67%** (6 of 7) |
| Liontrust | FY2026 | 🔴 −20.9% | 🟡 10.7% | 🔴 1.34x | **43%** (7 of 7) |
| Premier Miton | FY2025 | 🔴 −1.1% | 🔴 3.7% | 🔴 1.27x | **17%** (6 of 7) |

🟢 green · 🟡 amber · 🔴 red, using the bands in [data/kpi_thresholds.csv](data/kpi_thresholds.csv). The score is a screen, not a rating (see [section 9](#9-kpi-scorecard)).

---

## Contents

1. [Objective](#1-objective)
2. [Screenshots](#2-screenshots)
3. [Project files](#3-project-files)
4. [Raw data](#4-raw-data)
5. [Data cleaning](#5-data-cleaning)
6. [Data exploration](#6-data-exploration)
7. [Data quality](#7-data-quality)
8. [Analysis](#8-analysis)
9. [KPI scorecard](#9-kpi-scorecard)
10. [Recommendations](#10-recommendations)
11. [Linking the findings to business strategy](#11-linking-the-findings-to-business-strategy)
12. [Stakeholder brief](#12-stakeholder-brief)
13. [Limitations](#13-limitations)
14. [How to view and run the project](#14-how-to-view-and-run-the-project)
15. [Publishing this repository on GitHub](#15-publishing-this-repository-on-github)

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

## 2. Screenshots

All images are in the [images](images) folder. They are taken from the Excel workbook itself, apart from the summary dashboard, which is drawn from the same numbers.

**Summary dashboard** ([images/dashboard.png](images/dashboard.png)): latest-year scorecard and three-year trends.

![Summary dashboard](images/dashboard.png)

**Excel: Dashboard sheet** ([images/excel_dashboard.png](images/excel_dashboard.png)): the live dashboard, with brand-coloured scorecard, trend tables and native Excel charts.

![Excel Dashboard sheet](images/excel_dashboard.png)

**Excel: KPI_Scorecard sheet** ([images/excel_kpi_scorecard.png](images/excel_kpi_scorecard.png)): seven KPIs with green / amber / red status and editable bands.

![Excel KPI scorecard sheet](images/excel_kpi_scorecard.png)

**Excel: Ratios sheet** ([images/excel_ratios.png](images/excel_ratios.png)): every ratio for every company-year, all live formulas.

![Excel Ratios sheet](images/excel_ratios.png)

**Excel: Project_Story sheet** ([images/excel_project_story.png](images/excel_project_story.png)): the project from objective to stakeholder brief on one page.

![Excel Project Story sheet](images/excel_project_story.png)

**Excel: Data_Log sheet** ([images/excel_data_log.png](images/excel_data_log.png)): missing values, unusual figures and comparability notes.

![Excel Data Log sheet](images/excel_data_log.png)

---

## 3. Project files

```
UK-Financial-Firms-Health-Review-/
├── README.md                     ← this file
├── data/
│   ├── raw_financials.csv        ← figures as printed in the accounts (inputs)
│   ├── companies.csv             ← company register and Companies House links
│   ├── kpi_thresholds.csv        ← green / amber / red bands for each KPI
│   └── sources.csv               ← every source document, link and access date
├── src/
│   ├── config.py                 ← GitHub address used for links in the workbook and report
│   ├── analysis.py               ← import, cleaning, quality checks, ratios, KPIs, SQL
│   ├── build_workbook.py         ← builds the Excel workbook with live formulas
│   ├── make_dashboard.py         ← draws the summary dashboard image
│   ├── make_screenshots.py       ← exports the Excel sheet screenshots
│   ├── check_links.py            ← checks every link before upload
│   └── build_report.py           ← writes the report (PDF + Word)
├── sql/
│   └── queries.sql               ← SQL checks run on the SQLite database
├── outputs/
│   ├── UK_Financial_Health_Review.xlsx   ← main workbook
│   ├── ratios.csv · kpi_scorecard.csv · data_quality_log.csv
│   └── sql_results.xlsx · financials.db
├── report/
│   ├── UK_Financial_Health_Review_Report.docx
│   └── UK_Financial_Health_Review_Report.pdf
├── images/
│   ├── dashboard.png             ← summary dashboard
│   └── excel_*.png               ← screenshots of the workbook sheets
├── requirements.txt
└── .gitignore
```

| File | Use it for |
|---|---|
| [outputs/UK_Financial_Health_Review.xlsx](outputs/UK_Financial_Health_Review.xlsx) | The full analysis. Start on the **Dashboard** sheet |
| [report/UK_Financial_Health_Review_Report.pdf](report/UK_Financial_Health_Review_Report.pdf) | Three-page written report, then a tools section and an appendix with all SQL and Python code ([editable .docx](report/UK_Financial_Health_Review_Report.docx)) |
| [src/analysis.py](src/analysis.py) | Seeing exactly how the data was cleaned and the ratios calculated |
| [sql/queries.sql](sql/queries.sql) | The SQL checks |
| [outputs/kpi_scorecard.csv](outputs/kpi_scorecard.csv) | KPI values and green / amber / red status for each company |

**The process from start to finish**

```
Companies House check  →  Raw figures typed in  →  Cleaning (£m, blanks kept)
→  Exploration  →  Quality log  →  Ratios (Excel formulas + Python + SQL)
→  Findings  →  Recommendations  →  Stakeholder brief
```

---

## 4. Raw data

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

## 5. Data cleaning

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

## 6. Data exploration

Before calculating ratios, I looked at the overall shape of the data.

- **Scale varies a lot.** In the latest year, revenue ranged from £63.3m (Premier Miton) to £465.7m (Jupiter). The six firms together earned about £1.39bn.
- **More than half the sector was shrinking.** Of the 18 analysed company-years, 11 showed falling revenue, and the median revenue growth was −1.6%. Only AJ Bell and Polar Capital grew in every year.
- **Margins are widely spread.** Profit before tax margins ranged from a small loss (Liontrust FY2024, −0.3%) to 75.2% (Ashmore FY2025). The median was 25.6%.
- **Liquidity is high across the group.** Every current ratio was above 1.2x. The median was 2.5x, while Ashmore was an outlier at up to 10.7x.
- **Equity moves in different directions.** Equity grew at AJ Bell (£133.4m → £217.5m) but fell at Liontrust (£221.0m → £107.9m) and Ashmore (£966.8m → £790.8m).
- **Profit before tax and operating profit can differ a lot.** At Ashmore, profit before tax was up to 124% above operating profit, because of investment income. At Jupiter in FY2023 it was 89% below, because of a goodwill write-down. This is why both margins are shown.

---

## 7. Data quality

`src/analysis.py` checks every row and writes the results to [outputs/data_quality_log.csv](outputs/data_quality_log.csv) and to the workbook's **Data_Log** sheet.

| Check | Result |
|---|---|
| Company-years loaded | 24 (6 companies × 4 years) |
| Missing values | 18, all left blank and listed |
| Unusual figures flagged | 8 (losses, or profit before tax very different from operating profit) |
| Values estimated or filled in | **None** |
| Excel vs Python cross-check | All 144 ratio cells match |
| Excel formula errors | 0 out of 1,043 formulas |

**Where the gaps are**
- **Polar Capital:** no operating profit line and no separate lease liabilities, so its operating margin and debt-to-equity are n/a.
- **Ashmore:** no finance cost line, as the company states it has no debt, so interest coverage is n/a. Its FY2026 statements could not be retrieved.
- **AJ Bell:** no net revenue line. Figures come from its audited results announcements, because the full annual report could not be read past the auditor's report.
- **Companies House:** the pages for AJ Bell and Ashmore came from older cached copies, and Jupiter's page could not be opened.

---

## 8. Analysis

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

## 9. KPI scorecard

The five ratios become seven KPIs: the five ratios themselves, plus operating margin (to separate fee profit from investment gains) and the change in equity over three years (to show whether the capital base is growing or shrinking). Each KPI gets a green, amber or red status from the bands in [data/kpi_thresholds.csv](data/kpi_thresholds.csv). In Excel these bands are yellow input cells on the **KPI_Scorecard** sheet, so you can change them and every status updates.

| KPI | Better when | Green | Amber | Red |
|---|---|---|---|---|
| Revenue growth | higher | 5% or more | 0% to 5% | below 0% |
| Profit margin (PBT) | higher | 25% or more | 10% to 25% | below 10% |
| Operating margin | higher | 25% or more | 10% to 25% | below 10% |
| Current ratio | higher | 2.0x or more | 1.5x to 2.0x | below 1.5x |
| Debt-to-equity | lower | 0.25x or less | 0.25x to 0.5x | above 0.5x |
| Interest coverage | higher | 10x or more | 4x to 10x | below 4x |
| Equity change over 3 years | higher | 0% or more | −10% to 0% | below −10% |

**Screening score** = points ÷ maximum possible points, with green = 2, amber = 1 and red = 0. KPIs with no data are left out rather than counted as red.

| Company | Revenue growth | PBT margin | Operating margin | Current ratio | Debt / equity | Interest cover | Equity change | Score |
|---|---|---|---|---|---|---|---|---|
| AJ Bell | 🟢 18.0% | 🟢 43.4% | 🟢 41.5% | 🟢 3.65x | 🟢 0.06x | 🟢 149x | 🟢 +63.0% | **100%** |
| Jupiter | 🟢 15.7% | 🟢 28.3% | 🟢 27.5% | 🟢 2.46x | 🟢 0.04x | 🟢 40x | 🟢 +7.4% | **100%** |
| Polar Capital | 🟢 19.0% | 🟢 29.1% | ⚪ n/a | 🟢 2.53x | ⚪ n/a | 🟢 389x | 🟢 +5.6% | **100%** |
| Ashmore | 🔴 −23.6% | 🟢 75.2% | 🟢 39.6% | 🟢 7.01x | 🟢 0.01x | ⚪ n/a | 🔴 −18.2% | **67%** |
| Liontrust | 🔴 −20.9% | 🟡 10.7% | 🟡 10.1% | 🔴 1.34x | 🟢 0.03x | 🟢 96x | 🔴 −51.2% | **43%** |
| Premier Miton | 🔴 −1.1% | 🔴 3.7% | 🔴 2.7% | 🔴 1.27x | 🟢 0.02x | ⚪ n/a | 🔴 −11.3% | **17%** |

⚪ = no data. **How to read it:** Polar Capital's 100% rests on only five of the seven KPIs. Ashmore is green on every balance-sheet measure but red on both trend measures. Every firm is green on debt-to-equity, which shows how little borrowing there is in this sector. The bands are my own reading aids, not industry rules, and the score is a quick screen to decide where to look first. It is not a credit rating.

![KPI scorecard](images/excel_kpi_scorecard.png)

---

## 10. Recommendations

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
3. Re-run with the Companies House API key option (section 14) to download the filing records directly.

---

## 11. Linking the findings to business strategy

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

## 12. Stakeholder brief

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

## 13. Limitations

- **Different year ends.** The companies close their years in March, June, September and December, so the latest years run from June 2025 to March 2026. Market conditions differed across these dates.
- **Source route.** The figures come from each company's published audited statements, not from Companies House PDFs. Every figure is still traceable to a public source.
- **Missing figures.** There are 18 blank values (see section 7), mainly for Polar Capital and Ashmore.
- **Reporting differences.** The firms label their lines differently, and some prior years were re-presented. Every case is listed on the **Data_Log** sheet.
- **Listed versus private companies.** This approach is harder for private companies. Small firms can file abridged accounts with no profit and loss account, so growth and margin cannot be calculated at all.
- **Ratios are not a credit rating.** They leave out fund flows, management quality, regulation and anything that has happened since the accounts were published.

---

## 14. How to view and run the project

**Viewing the workbook.** Open `outputs/UK_Financial_Health_Review.xlsx` in Excel, LibreOffice or Google Sheets. The first sheet, **README**, repeats this guide. Column headers are frozen, so every sheet scrolls cleanly.

| Sheet | What it shows |
|---|---|
| README | This guide, inside the workbook |
| Dashboard | Latest-year scorecard, three-year trend charts and key points |
| KPI_Scorecard | Seven KPIs with green / amber / red status, editable bands and a screening score |
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
pip install -r requirements.txt
python src/analysis.py          # cleaning, quality checks, ratios, KPIs, SQLite + SQL
python src/build_workbook.py    # Excel workbook
python src/make_dashboard.py    # summary dashboard image
python src/make_screenshots.py  # Excel sheet screenshots (needs LibreOffice and poppler)
python src/build_report.py      # report, Word and PDF (PDF needs LibreOffice)
python src/check_links.py       # check every link (add --online to test web addresses)
```

**Optional: download filing records from Companies House.** Get a free API key from the [Companies House developer hub](https://developer.company-information.service.gov.uk/), then run:

```bash
export CH_API_KEY=your_key_here
python src/analysis.py
```

The key is read from your environment only and is never saved in the project.

---

## 15. Publishing this repository on GitHub

This project lives at **https://github.com/Mumeen-Gbolagade/UK-Financial-Firms-Health-Review-**.

**How the links work**
- In this README, every link is a relative path such as `report/UK_Financial_Health_Review_Report.pdf`. GitHub resolves these against the repository, so **the folders must be uploaded as folders**. If the files end up loose in the top level of the repository, every link and image breaks.
- In the Excel workbook and the report, file links point to their full GitHub address, for example `https://github.com/Mumeen-Gbolagade/UK-Financial-Firms-Health-Review-/blob/main/outputs/UK_Financial_Health_Review.xlsx`. They keep working after someone downloads a single file. To use another repository address, set `REPO_URL` (see `src/config.py`) and rebuild.
- `python src/check_links.py` checks all of this before you upload.

**Option A: upload in the browser (no Git needed)**
1. Unzip the project on your computer.
2. If the repository already has files loose at the top level, remove them first. The quickest way is **Settings → General → Danger Zone → Delete this repository**, then create a new, empty repository with the same name. Do not tick "Add a README", because this project already has one.
3. On the empty repository page, click **uploading an existing file**. Open the unzipped project folder, select **everything inside it** (the six folders `data`, `images`, `outputs`, `report`, `sql` and `src`, plus `README.md`, `requirements.txt` and `.gitignore`), and drag the selection onto the page. Drag them; don't use "choose your files", because that picker drops the folders. If your computer hides `.gitignore`, you can leave it out.
4. Check that the list on the upload page shows paths such as `images/dashboard.png` and `outputs/UK_Financial_Health_Review.xlsx`, then click **Commit changes**.

**Option B: upload with Git**

```bash
git clone https://github.com/Mumeen-Gbolagade/UK-Financial-Firms-Health-Review-.git
cd UK-Financial-Firms-Health-Review-
git rm -r -q --ignore-unmatch .          # clear any files uploaded loose at the top level
cp -r /path/to/unzipped/project/. .      # copy in the project folder's contents
git add .
git commit -m "UK Company Financial Health Review"
git push origin main
```

**After uploading**, open the repository page. The dashboard image should appear at the top of the README, and the Screenshots section should show all six images. GitHub shows the `.xlsx`, `.docx` and `.db` files with a download button, opens the PDF report in its viewer, and displays the CSV files as tables.
