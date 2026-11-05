"""
Build the Excel workbook: UK_Financial_Health_Review.xlsx

Every ratio in the workbook is a live Excel formula that points back to the
cleaned figures, which in turn point back to the raw extracted figures. Change
a raw number and the ratios and dashboard update.

Run after src/analysis.py:
    python src/build_workbook.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import DataPoint
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, Font as DFont


def chart_title(chart, text: str, size: int = 1100) -> None:
    """Set a smaller, bold chart title in the workbook font."""
    from openpyxl.chart.title import Title
    from openpyxl.chart.text import Text
    from openpyxl.drawing.text import RegularTextRun
    cp = CharacterProperties(sz=size, b=True, solidFill="1F2A44", latin=DFont(typeface="Arial"))
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    chart.title = Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def axis_font(axis, size: int = 900) -> None:
    cp = CharacterProperties(sz=size, solidFill="5B6577", latin=DFont(typeface="Arial"))
    axis.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "outputs"
FILE = OUT / "UK_Financial_Health_Review.xlsx"

# ---------------------------------------------------------------- styling
FONT = "Arial"
INK = "1F2A44"        # dark slate used for headings
MUTED = "5B6577"
LINE = "D5DAE3"
PAPER = "F5F7FA"
WHITE = "FFFFFF"
GOOD_FILL = "E3F2E8"
WARN_FILL = "FDF1DC"
BAD_FILL = "FBE4E4"

F_TITLE = Font(name=FONT, size=18, bold=True, color=WHITE)
F_SUB = Font(name=FONT, size=10, color="DCE3EE")
F_H = Font(name=FONT, size=10, bold=True, color=WHITE)
F_BODY = Font(name=FONT, size=10, color="222222")
F_BOLD = Font(name=FONT, size=10, bold=True, color="222222")
F_INPUT = Font(name=FONT, size=10, color="0000FF")      # hardcoded inputs
F_LINK = Font(name=FONT, size=10, color="008000")       # links to other sheets
F_URL = Font(name=FONT, size=10, color="1155CC", underline="single")
F_SECTION = Font(name=FONT, size=12, bold=True, color=INK)
F_NOTE = Font(name=FONT, size=9, italic=True, color=MUTED)

FILL_INK = PatternFill("solid", fgColor=INK)
FILL_PAPER = PatternFill("solid", fgColor=PAPER)
FILL_WHITE = PatternFill("solid", fgColor=WHITE)
THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

GBP_M = '#,##0.0;(#,##0.0);"-"'
GBP_RAW = '#,##0.0;(#,##0.0);"0"'
PCT = '0.0%;(0.0%);"0.0%"'
TIMES = '0.00"x";(0.00"x");"0.00x"'
COVER = '#,##0.0"x";(#,##0.0"x");"0.0x"'


def style_header(ws, row: int, ncols: int, start_col: int = 1, height: int = 30) -> None:
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = F_H
        cell.fill = FILL_INK
        cell.alignment = CENTER
        cell.border = BOX
    ws.row_dimensions[row].height = height


def title_band(ws, title: str, subtitle: str, ncols: int) -> None:
    for r in (1, 2):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).fill = FILL_INK
    ws["A1"] = title
    ws["A1"].font = F_TITLE
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws["A2"] = subtitle
    ws["A2"].font = F_SUB
    ws["A2"].alignment = Alignment(vertical="top", indent=1)
    ws.row_dimensions[1].height = 34
    ws.row_dimensions[2].height = 22


def set_widths(ws, widths: dict[str, float]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def page_setup(ws, landscape: bool = True) -> None:
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------- data
raw = pd.read_csv(DATA / "raw_financials.csv", dtype={"company_number": str})
raw["fy_end"] = pd.to_datetime(raw["fy_end"])
raw = raw.sort_values(["company", "fy_end"]).reset_index(drop=True)
companies = pd.read_csv(DATA / "companies.csv", dtype={"company_number": str})
sources = pd.read_csv(DATA / "sources.csv")
quality = pd.read_csv(OUT / "data_quality_log.csv")
sql_results = pd.read_excel(OUT / "sql_results.xlsx", sheet_name=None)
BRAND = dict(zip(companies["company"], companies["brand_colour"]))
SHORT = dict(zip(companies["company"], companies["short_name"]))

wb = Workbook()
wb.remove(wb.active)
ws_readme = wb.create_sheet("README")
ws_dash = wb.create_sheet("Dashboard")
ws_kpi = wb.create_sheet("KPI_Scorecard")
ws_story = wb.create_sheet("Project_Story")
ws_co = wb.create_sheet("Companies")
ws_raw = wb.create_sheet("Raw_Data")
ws_cl = wb.create_sheet("Clean_GBPm")
ws_r = wb.create_sheet("Ratios")
ws_g = wb.create_sheet("Formula_Guide")
ws_log = wb.create_sheet("Data_Log")
ws_sql = wb.create_sheet("SQL_Checks")
ws_src = wb.create_sheet("Sources")

FIRST = 5                      # first data row on Raw/Clean/Ratios
LAST = FIRST + len(raw) - 1

# ================================================================ Companies
title_band(ws_co, "Company register",
           "Six UK-listed asset and wealth managers. Company numbers from Companies House. Pages checked 4-5 Nov 2026.", 9)
heads = ["Company", "Short name", "Company number", "Sector", "SIC codes", "Year end",
         "Accounts type", "Latest accounts seen on Companies House", "Companies House link"]
for i, h in enumerate(heads, 1):
    ws_co.cell(row=4, column=i, value=h)
style_header(ws_co, 4, len(heads))
for r, row in enumerate(companies.itertuples(index=False), start=5):
    vals = [row.company, row.short_name, row.company_number, row.sector, row.sic_codes,
            row.year_end, row.accounts_type, row.latest_accounts_seen_on_companies_house,
            row.companies_house_link]
    for c, v in enumerate(vals, 1):
        cell = ws_co.cell(row=r, column=c, value=v)
        cell.font, cell.alignment, cell.border = F_BODY, WRAP, BOX
    ws_co.cell(row=r, column=1).font = Font(name=FONT, size=10, bold=True, color=row.brand_colour)
    ws_co.cell(row=r, column=3).number_format = "@"
    link = ws_co.cell(row=r, column=9)
    link.hyperlink = row.companies_house_link
    link.font = F_URL
    ws_co.cell(row=r, column=1).border = Border(left=Side(style="thick", color=row.brand_colour),
                                               right=THIN, top=THIN, bottom=THIN)
    ws_co.row_dimensions[r].height = 42
note_row = 5 + len(companies) + 1
ws_co.cell(row=note_row, column=1,
           value="Why this sector: asset and wealth managers split their balance sheets into current and "
                 "non-current items, so all five ratios can be calculated. Banks and insurers do not, "
                 "which makes a current ratio meaningless for them.").font = F_NOTE
ws_co.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=9)
ws_co.row_dimensions[note_row].height = 30
ws_co.cell(row=note_row, column=1).alignment = WRAP
set_widths(ws_co, {"A": 30, "B": 14, "C": 15, "D": 30, "E": 14, "F": 13, "G": 22, "H": 34, "I": 58})
ws_co.freeze_panes = "B5"
page_setup(ws_co)

# ================================================================ Raw_Data
raw_heads = ["Company", "Company number", "Financial year", "Year end date", "Units",
             "Revenue", "Net revenue / gross profit", "Operating profit", "Finance income",
             "Finance costs", "Profit before tax", "Profit for the year", "Current assets",
             "Current liabilities", "Borrowings", "Lease liabilities", "Total equity",
             "Source ID", "Notes from the accounts"]
title_band(ws_raw, "Raw data as printed in the accounts",
           "Blue figures are typed in exactly as shown in each source, in the units stated. Blank = not shown; nothing estimated.",
           len(raw_heads))
for i, h in enumerate(raw_heads, 1):
    ws_raw.cell(row=4, column=i, value=h)
style_header(ws_raw, 4, len(raw_heads), height=42)
money_cols = ["revenue", "net_revenue", "operating_profit", "finance_income", "finance_costs",
              "profit_before_tax", "profit_for_year", "current_assets", "current_liabilities",
              "borrowings", "lease_liabilities", "total_equity"]
for r, row in enumerate(raw.itertuples(index=False), start=FIRST):
    vals = [row.company, row.company_number, row.fy_label, row.fy_end.to_pydatetime(), row.units]
    vals += [None if pd.isna(getattr(row, c)) else float(getattr(row, c)) for c in money_cols]
    vals += [row.source_id, row.notes]
    for c, v in enumerate(vals, 1):
        cell = ws_raw.cell(row=r, column=c, value=v)
        cell.border = BOX
        cell.font = F_INPUT if 6 <= c <= 17 else F_BODY
        if 6 <= c <= 17:
            cell.number_format = '#,##0.0;(#,##0.0)'
            if v is None:
                cell.fill = PatternFill("solid", fgColor=WARN_FILL)
    ws_raw.cell(row=r, column=1).font = Font(name=FONT, size=10, bold=True, color=BRAND[row.company])
    ws_raw.cell(row=r, column=2).number_format = "@"
    ws_raw.cell(row=r, column=4).number_format = "dd mmm yyyy"
    ws_raw.cell(row=r, column=19).alignment = WRAP
    ws_raw.row_dimensions[r].height = 40
ws_raw.cell(row=LAST + 2, column=1,
            value="Shaded cells: figure not shown in the retrieved statements. Each source ID is listed on the Sources sheet.").font = F_NOTE
set_widths(ws_raw, {"A": 30, "B": 12, "C": 10, "D": 13, "E": 9, **{get_column_letter(c): 12.5 for c in range(6, 18)},
                    "R": 9, "S": 70})
ws_raw.freeze_panes = "D5"
page_setup(ws_raw)

# ================================================================ Clean_GBPm
clean_heads = ["Company", "Short name", "Financial year", "Year end date", "Source units",
               "Revenue", "Net revenue / gross profit", "Operating profit", "Finance income",
               "Finance costs", "Profit before tax", "Profit for the year", "Current assets",
               "Current liabilities", "Borrowings", "Lease liabilities", "Total equity",
               "Unit factor", "Total debt (borrowings + leases)"]
title_band(ws_cl, "Cleaned data in £ millions",
           "Green = linked to Raw_Data. Figures in £000 are divided by 1,000. Blanks stay blank so no ratio uses a guessed number.",
           len(clean_heads))
for i, h in enumerate(clean_heads, 1):
    ws_cl.cell(row=4, column=i, value=h)
style_header(ws_cl, 4, len(clean_heads), height=42)
for r in range(FIRST, LAST + 1):
    ws_cl.cell(row=r, column=1, value=f"=Raw_Data!A{r}")
    ws_cl.cell(row=r, column=2, value=f'=IFERROR(INDEX(Companies!$B$5:$B$10,MATCH(A{r},Companies!$A$5:$A$10,0)),A{r})')
    ws_cl.cell(row=r, column=3, value=f"=Raw_Data!C{r}")
    ws_cl.cell(row=r, column=4, value=f"=Raw_Data!D{r}")
    ws_cl.cell(row=r, column=5, value=f"=Raw_Data!E{r}")
    for c in range(6, 18):
        col = get_column_letter(c)
        ws_cl.cell(row=r, column=c, value=f'=IF(Raw_Data!{col}{r}="","",Raw_Data!{col}{r}*$R{r})')
        ws_cl.cell(row=r, column=c).number_format = GBP_RAW
    ws_cl.cell(row=r, column=18, value=f'=IF(E{r}="GBP000",0.001,1)')
    ws_cl.cell(row=r, column=19, value=f'=IF(AND(ISNUMBER(O{r}),ISNUMBER(P{r})),O{r}+P{r},"")')
    ws_cl.cell(row=r, column=19).number_format = GBP_RAW
    for c in range(1, 20):
        cell = ws_cl.cell(row=r, column=c)
        cell.border = BOX
        cell.font = F_LINK if c in (1, 3, 4, 5) or 6 <= c <= 17 else F_BODY
    ws_cl.cell(row=r, column=4).number_format = "dd mmm yyyy"
    ws_cl.cell(row=r, column=18).number_format = "0.000"
set_widths(ws_cl, {"A": 30, "B": 14, "C": 10, "D": 13, "E": 9, **{get_column_letter(c): 12.5 for c in range(6, 18)},
                   "R": 9, "S": 15})
ws_cl.freeze_panes = "D5"
page_setup(ws_cl)

# ================================================================ Ratios
r_heads = ["Company", "Short name", "Financial year", "Year end date", "Revenue £m",
           "Profit before tax £m", "Revenue growth", "Profit margin (PBT)",
           "Operating margin", "Current ratio", "Debt-to-equity", "Interest coverage",
           "Later years for this company", "Use in analysis", "Lookup key"]
title_band(ws_r, "Ratio calculations",
           "All cells are formulas on Clean_GBPm. 'n/a' means the accounts do not give the inputs, or the ratio has no meaning (e.g. no interest cost).",
           len(r_heads))
for i, h in enumerate(r_heads, 1):
    ws_r.cell(row=4, column=i, value=h)
style_header(ws_r, 4, len(r_heads), height=42)
CL = "Clean_GBPm!"
for r in range(FIRST, LAST + 1):
    f = {
        1: f"={CL}A{r}",
        2: f"={CL}B{r}",
        3: f"={CL}C{r}",
        4: f"={CL}D{r}",
        5: f"={CL}F{r}",
        6: f"={CL}K{r}",
        7: (f'=IF(AND(A{r}=A{r-1},ISNUMBER(E{r}),ISNUMBER(E{r-1})),IF(E{r-1}<>0,E{r}/E{r-1}-1,"n/a"),"n/a")'
            if r > FIRST else '="n/a"'),
        8: f'=IF(AND(ISNUMBER({CL}K{r}),ISNUMBER({CL}F{r})),IF({CL}F{r}<>0,{CL}K{r}/{CL}F{r},"n/a"),"n/a")',
        9: f'=IF(AND(ISNUMBER({CL}H{r}),ISNUMBER({CL}F{r})),IF({CL}F{r}<>0,{CL}H{r}/{CL}F{r},"n/a"),"n/a")',
        10: f'=IF(AND(ISNUMBER({CL}M{r}),ISNUMBER({CL}N{r})),IF({CL}N{r}<>0,{CL}M{r}/{CL}N{r},"n/a"),"n/a")',
        11: f'=IF(AND(ISNUMBER({CL}S{r}),ISNUMBER({CL}Q{r})),IF({CL}Q{r}>0,{CL}S{r}/{CL}Q{r},"n/a"),"n/a")',
        12: f'=IF(AND(ISNUMBER({CL}J{r}),ISNUMBER({CL}K{r})),IF({CL}J{r}>0,({CL}K{r}+{CL}J{r})/{CL}J{r},"n/a"),"n/a")',
        13: f'=COUNTIFS($A${FIRST}:$A${LAST},A{r},$D${FIRST}:$D${LAST},">"&D{r})',
        14: f'=IF(M{r}<=2,"Analysed","Base year (growth only)")',
        15: f'=A{r}&"|"&M{r}',
    }
    for c, formula in f.items():
        cell = ws_r.cell(row=r, column=c, value=formula)
        cell.border = BOX
        cell.font = F_LINK if c <= 6 else F_BODY
        cell.alignment = RIGHT if 5 <= c <= 12 else Alignment(vertical="center")
    ws_r.cell(row=r, column=4).number_format = "dd mmm yyyy"
    ws_r.cell(row=r, column=5).number_format = GBP_M
    ws_r.cell(row=r, column=6).number_format = GBP_M
    for c in (7, 8, 9):
        ws_r.cell(row=r, column=c).number_format = PCT
    ws_r.cell(row=r, column=10).number_format = TIMES
    ws_r.cell(row=r, column=11).number_format = TIMES
    ws_r.cell(row=r, column=12).number_format = COVER
    ws_r.cell(row=r, column=15).font = Font(name=FONT, size=8, color="9AA3B2")
# light banding by company, brand colour on company name
for r, row in enumerate(raw.itertuples(index=False), start=FIRST):
    ws_r.cell(row=r, column=2).font = Font(name=FONT, size=10, bold=True, color=BRAND[row.company])
    if list(companies["company"]).index(row.company) % 2 == 0:
        for c in range(1, 16):
            ws_r.cell(row=r, column=c).fill = FILL_PAPER
# conditional colours (only numeric cells are affected)
rng = lambda col: f"{col}{FIRST}:{col}{LAST}"
ws_r.conditional_formatting.add(rng("H"), CellIsRule(operator="lessThan", formula=["0.05"], fill=PatternFill("solid", fgColor=BAD_FILL)))
ws_r.conditional_formatting.add(rng("H"), CellIsRule(operator="greaterThanOrEqual", formula=["0.25"], fill=PatternFill("solid", fgColor=GOOD_FILL)))
ws_r.conditional_formatting.add(rng("G"), CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, color="B42318")))
ws_r.conditional_formatting.add(rng("J"), CellIsRule(operator="lessThan", formula=["1.5"], fill=PatternFill("solid", fgColor=WARN_FILL)))
ws_r.cell(row=LAST + 2, column=1,
          value="Shading: green margin = 25% or more; red margin = under 5%; amber current ratio = under 1.5x; red growth = revenue fell. "
                "These bands are reading aids, not pass/fail tests.").font = F_NOTE
ws_r.cell(row=LAST + 3, column=1,
          value="Interest coverage uses profit before tax plus finance costs, divided by finance costs. A negative result means earnings did not cover interest that year.").font = F_NOTE
ws_r.cell(row=7, column=12).comment = None
set_widths(ws_r, {"A": 30, "B": 14, "C": 10, "D": 13, "E": 12, "F": 13, "G": 12, "H": 13,
                  "I": 12, "J": 11, "K": 12, "L": 12, "M": 11, "N": 20, "O": 6})
ws_r.column_dimensions["O"].hidden = False
ws_r.freeze_panes = "D5"
page_setup(ws_r)

# ================================================================ Dashboard
order = (sql_results["latest_year_snapshot"]["company"].tolist())  # sorted by PBT margin
ws = ws_dash
ncol = 12
title_band(ws, "UK Asset & Wealth Managers: Financial Health Review",
           "Latest audited year for each company, figures in £ millions. Built from Companies House-listed group accounts; data accessed 4-5 Nov 2026.",
           ncol)
ws.row_dimensions[3].height = 8

# --- KPI table
ws["A4"] = "Latest-year scorecard (sorted by profit margin)"
ws["A4"].font = F_SECTION
heads = ["Company", "Latest year", "Revenue £m", "Revenue growth", "Profit margin (PBT)",
         "Current ratio", "Debt-to-equity", "Interest coverage", "What stands out"]
for i, h in enumerate(heads, 1):
    ws.cell(row=5, column=i, value=h)
style_header(ws, 5, len(heads), height=34)
ws.merge_cells(start_row=5, start_column=9, end_row=5, end_column=12)

READ = {
    "AJ Bell plc": "Steady growth, high margin, no borrowings. Check how much profit relies on interest earned on client cash.",
    "Ashmore Group plc": "Most liquid balance sheet and no debt, but revenue has fallen three years running. Profit leans on investment income.",
    "Polar Capital Holdings plc": "Revenue up three years in a row, no borrowings. Lease debt not shown separately, so debt-to-equity is incomplete.",
    "Jupiter Fund Management plc": "Recovered from a 2023 goodwill write-down. Repaid its £50m notes in April 2025.",
    "Liontrust Asset Management plc": "Revenue down three years running and equity has halved since FY2023. Lowest liquidity cushion of the profitable group.",
    "Premier Miton Group plc": "No debt, but profit margin is thin and still falling. Revenue has dropped about 30% since FY2022.",
}
R = "Ratios!"
KEYRANGE = f"{R}$O${FIRST}:$O${LAST}"


def lookup(col: str, key_cell: str) -> str:
    return f'=IFERROR(INDEX({R}${col}${FIRST}:${col}${LAST},MATCH({key_cell},{KEYRANGE},0)),"n/a")'


for i, company in enumerate(order):
    r = 6 + i
    ws.cell(row=r, column=1, value=company)
    ws.cell(row=r, column=13, value=f'=A{r}&"|0"')          # helper key (hidden col M)
    ws.cell(row=r, column=2, value=lookup("C", f"$M{r}"))
    ws.cell(row=r, column=3, value=lookup("E", f"$M{r}"))
    ws.cell(row=r, column=4, value=lookup("G", f"$M{r}"))
    ws.cell(row=r, column=5, value=lookup("H", f"$M{r}"))
    ws.cell(row=r, column=6, value=lookup("J", f"$M{r}"))
    ws.cell(row=r, column=7, value=lookup("K", f"$M{r}"))
    ws.cell(row=r, column=8, value=lookup("L", f"$M{r}"))
    ws.cell(row=r, column=9, value=READ[company])
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=12)
    brand = BRAND[company]
    for c in range(1, 13):
        cell = ws.cell(row=r, column=c)
        cell.border = BOX
        cell.font = F_BODY
        cell.alignment = Alignment(vertical="center", horizontal="right" if 3 <= c <= 8 else "left",
                                   wrap_text=(c == 9), indent=1 if c in (1, 9) else 0)
    ws.cell(row=r, column=1).font = Font(name=FONT, size=11, bold=True, color=WHITE)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=brand)
    ws.cell(row=r, column=2).alignment = CENTER
    ws.cell(row=r, column=3).number_format = GBP_M
    ws.cell(row=r, column=4).number_format = PCT
    ws.cell(row=r, column=5).number_format = PCT
    ws.cell(row=r, column=6).number_format = TIMES
    ws.cell(row=r, column=7).number_format = TIMES
    ws.cell(row=r, column=8).number_format = COVER
    ws.row_dimensions[r].height = 36
last_kpi = 6 + len(order) - 1
ws.column_dimensions["M"].hidden = True
kr = lambda col: f"{col}6:{col}{last_kpi}"
ws.conditional_formatting.add(kr("D"), CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, bold=True, color="B42318"), fill=PatternFill("solid", fgColor=BAD_FILL)))
ws.conditional_formatting.add(kr("D"), CellIsRule(operator="greaterThan", formula=["0.1"], font=Font(name=FONT, bold=True, color="1E7B3A"), fill=PatternFill("solid", fgColor=GOOD_FILL)))
ws.conditional_formatting.add(kr("E"), CellIsRule(operator="greaterThanOrEqual", formula=["0.25"], font=Font(name=FONT, bold=True, color="1E7B3A"), fill=PatternFill("solid", fgColor=GOOD_FILL)))
ws.conditional_formatting.add(kr("E"), CellIsRule(operator="lessThan", formula=["0.05"], font=Font(name=FONT, bold=True, color="B42318"), fill=PatternFill("solid", fgColor=BAD_FILL)))
ws.conditional_formatting.add(kr("F"), CellIsRule(operator="lessThan", formula=["1.5"], fill=PatternFill("solid", fgColor=WARN_FILL)))
note_r = last_kpi + 1
ws.cell(row=note_r, column=1,
        value="Latest years differ: Liontrust and Polar Capital to Mar 2026; AJ Bell and Premier Miton to Sep 2025; Ashmore to Jun 2025; Jupiter to Dec 2025. "
              "Interest coverage is n/a where a company reports no interest cost.").font = F_NOTE
ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=12)
ws.row_dimensions[note_r].height = 26
ws.cell(row=note_r, column=1).alignment = WRAP

# --- trend helper tables (feed the charts)
helper_top = note_r + 2
ws.cell(row=helper_top, column=1, value="Three-year trend (chart data)").font = F_SECTION
trend_specs = [("Profit margin (PBT)", "H", PCT), ("Revenue growth", "G", PCT), ("Current ratio", "J", TIMES)]
blocks = {}
row_ptr = helper_top + 1
for label, col, fmt in trend_specs:
    hdr = row_ptr
    for i, h in enumerate([label, "Two years before", "One year before", "Latest year"], 1):
        ws.cell(row=hdr, column=i, value=h)
    style_header(ws, hdr, 4, height=22)
    for i, company in enumerate(order):
        r = hdr + 1 + i
        ws.cell(row=r, column=1, value=SHORT[company])
        ws.cell(row=r, column=1).font = Font(name=FONT, size=10, bold=True, color=BRAND[company])
        for j, rank in enumerate((2, 1, 0)):
            c = ws.cell(row=r, column=2 + j,
                        value=f'=IFERROR(INDEX({R}${col}${FIRST}:${col}${LAST},MATCH("{company}|{rank}",{KEYRANGE},0)),"n/a")')
            c.number_format = fmt
            c.font = F_BODY
            c.alignment = RIGHT
        for cc in range(1, 5):
            ws.cell(row=r, column=cc).border = BOX
    blocks[label] = (hdr, hdr + len(order))
    row_ptr = hdr + len(order) + 2
# years behind each column, for reference
yr_hdr = row_ptr
for i, h in enumerate(["Financial years used", "Two years before", "One year before", "Latest year"], 1):
    ws.cell(row=yr_hdr, column=i, value=h)
style_header(ws, yr_hdr, 4, height=22)
for i, company in enumerate(order):
    r = yr_hdr + 1 + i
    ws.cell(row=r, column=1, value=SHORT[company]).font = Font(name=FONT, size=10, bold=True, color=BRAND[company])
    for j, rank in enumerate((2, 1, 0)):
        c = ws.cell(row=r, column=2 + j,
                    value=f'=IFERROR(INDEX({R}$C${FIRST}:$C${LAST},MATCH("{company}|{rank}",{KEYRANGE},0)),"n/a")')
        c.font, c.alignment = F_BODY, CENTER
    for cc in range(1, 5):
        ws.cell(row=r, column=cc).border = BOX

# --- charts
def line_chart(label: str, title: str, y_fmt: str) -> LineChart:
    hdr, end = blocks[label]
    ch = LineChart()
    chart_title(ch, title)
    ch.height, ch.width = 7.6, 17.5
    ch.y_axis.number_format = y_fmt
    ch.y_axis.majorGridlines = None
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    data = Reference(ws, min_col=1, max_col=4, min_row=hdr + 1, max_row=end)
    ch.add_data(data, from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=4, min_row=hdr, max_row=hdr))
    for s, company in zip(ch.series, order):
        colour = BRAND[company]
        s.graphicalProperties.line.solidFill = colour
        s.graphicalProperties.line.width = 28000
        s.marker = Marker(symbol="circle", size=7)
        s.marker.graphicalProperties.solidFill = colour
        s.marker.graphicalProperties.line.solidFill = colour
        s.smooth = False
    ch.x_axis.tickLblPos = "low"
    axis_font(ch.x_axis)
    axis_font(ch.y_axis)
    ch.legend.position = "b"
    return ch


ws.add_chart(line_chart("Profit margin (PBT)", "Profit before tax margin, last three years", "0%"), "F15")
ws.add_chart(line_chart("Revenue growth", "Revenue growth, last three years", "0%"), "F31")

hdr, end = blocks["Current ratio"]
bar = BarChart()
bar.type = "bar"
chart_title(bar, "Current ratio, latest year (current assets ÷ current liabilities)")
bar.height, bar.width = 7.6, 17.5
axis_font(bar.x_axis)
axis_font(bar.y_axis)
bar.add_data(Reference(ws, min_col=4, min_row=hdr + 1, max_row=end), titles_from_data=False)
bar.set_categories(Reference(ws, min_col=1, min_row=hdr + 1, max_row=end))
bar.legend = None
bar.y_axis.number_format = '0.0"x"'
bar.y_axis.majorGridlines = None
bar.y_axis.delete = False
bar.x_axis.delete = False
bar.x_axis.scaling.orientation = "maxMin"
s = bar.series[0]
for idx, company in enumerate(order):
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = BRAND[company]
    pt.graphicalProperties.line.solidFill = BRAND[company]
    s.dPt.append(pt)
s.dLbls = DataLabelList()
s.dLbls.showVal = True
s.dLbls.showSerName = False
s.dLbls.showCatName = False
s.dLbls.showLegendKey = False
s.dLbls.showPercent = False
s.dLbls.numFmt = '0.0"x"'
bar.gapWidth = 60
ws.add_chart(bar, "F47")

# --- key findings box beneath helper tables
kf_top = yr_hdr + len(order) + 2
ws.cell(row=kf_top, column=1, value="Key points for a lender or investor").font = F_SECTION
points = [
    "AJ Bell combines the strongest growth (revenue +18% to +33% a year) with a rising margin (40% to 43%) and no borrowings.",
    "Ashmore and Polar Capital also carry no borrowings. Ashmore's very high margin depends on investment income, while its fee revenue fell 24% in FY2025.",
    "Jupiter's margin recovered from 2.3% (FY2023, after a £76m goodwill write-down) to 28.3%, and it repaid its only borrowings in 2025.",
    "Liontrust and Premier Miton show the weakest trends: falling revenue, thin margins and current ratios near 1.3x.",
    "None of these ratios proves a firm is safe or at risk. They show where to ask questions next.",
]
for i, p in enumerate(points):
    r = kf_top + 1 + i
    ws.cell(row=r, column=1, value=f"•  {p}")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(row=r, column=1).font = F_BODY
    ws.cell(row=r, column=1).alignment = WRAP
    ws.row_dimensions[r].height = 44

set_widths(ws, {"A": 36, "B": 14, "C": 13, "D": 13, "E": 13, "F": 12, "G": 13, "H": 13,
                "I": 18, "J": 18, "K": 18, "L": 18})
ws.freeze_panes = "A4"
page_setup(ws)
ws.sheet_view.zoomScale = 90

# ================================================================ Formula guide
title_band(ws_g, "How each ratio is worked out",
           "Plain-English formulas, what a higher or lower result may suggest, and why no ratio should be read alone.", 6)
g_heads = ["Ratio", "Formula in plain words", "Excel formula used (row 5 example)",
           "A higher result may indicate", "A lower result may indicate", "Why not to read it on its own"]
for i, h in enumerate(g_heads, 1):
    ws_g.cell(row=4, column=i, value=h)
style_header(ws_g, 4, 6)
guide = [
    ("Revenue growth", "This year's revenue divided by last year's revenue, minus 1.",
     "=E6/E5-1 (same company, consecutive years)",
     "The firm is winning business or benefiting from rising markets.",
     "Clients are leaving, fees are being cut or markets have fallen.",
     "For fund managers, revenue moves with stock and bond markets, so growth can reflect market levels rather than management skill."),
    ("Profit margin (profit before tax)", "Profit before tax divided by revenue. We use profit before tax because Polar Capital does not show an operating profit line, so it is the one measure every company reports.",
     "=Clean_GBPm!K5/Clean_GBPm!F5",
     "More of each £1 of revenue is kept as profit; costs are under control.",
     "Costs are high relative to income, or one-off charges hit the year.",
     "It includes investment income and one-off items. Ashmore's margin is lifted by returns on its own investments; Jupiter's FY2023 margin is cut by a goodwill write-down. Operating margin is shown alongside as a check."),
    ("Current ratio", "Current assets divided by current liabilities.",
     "=Clean_GBPm!M5/Clean_GBPm!N5",
     "More short-term assets than bills due within a year; a bigger cushion.",
     "Less room to meet short-term obligations.",
     "Fund managers often hold money due to and from their funds in receivables and payables. These balances can be large and roughly cancel out, pulling the ratio towards 1x without signalling stress. Too high can also mean idle cash."),
    ("Debt-to-equity", "Borrowings plus lease liabilities, divided by total equity.",
     "=Clean_GBPm!S5/Clean_GBPm!Q5",
     "More reliance on borrowed money, so more fixed obligations.",
     "The firm is funded mostly by shareholders' money.",
     "Most of these firms have no bank debt, so the figure is mainly office leases. It does not capture regulatory capital needs, pension deficits or seed capital commitments."),
    ("Interest coverage", "Profit before tax plus finance costs, divided by finance costs (how many times earnings cover interest).",
     "=(Clean_GBPm!K5+Clean_GBPm!J5)/Clean_GBPm!J5",
     "Interest is easily covered by earnings.",
     "Earnings would struggle to pay interest; a negative figure means a loss.",
     "When interest costs are tiny (lease interest of a few hundred thousand pounds), the ratio becomes very large and adds little. It is n/a where no interest cost is reported."),
]
for r, row in enumerate(guide, start=5):
    for c, v in enumerate(row, 1):
        cell = ws_g.cell(row=r, column=c, value=v)
        if c == 3:
            cell.data_type = "s"   # show the formula as text, not run it
            cell.font = Font(name="Consolas", size=10, color=INK)
        else:
            cell.font = F_BOLD if c == 1 else F_BODY
        cell.alignment = WRAP
        cell.border = BOX
    ws_g.row_dimensions[r].height = 92
ws_g.cell(row=11, column=1,
          value="Operating margin (supplementary): operating profit divided by revenue. Shown on the Ratios sheet to separate the core fee business from investment returns and one-offs.").font = F_NOTE
ws_g.merge_cells("A11:F11")
set_widths(ws_g, {"A": 22, "B": 38, "C": 30, "D": 30, "E": 30, "F": 52})
ws_g.freeze_panes = "B5"
page_setup(ws_g)

# ================================================================ Data log
title_band(ws_log, "Data checks: missing values, unusual figures and comparability",
           "Generated by src/analysis.py, plus analyst notes. Nothing has been estimated or filled in.", 5)
l_heads = ["Company", "Financial year", "Type", "Field", "Detail"]
for i, h in enumerate(l_heads, 1):
    ws_log.cell(row=4, column=i, value=h)
style_header(ws_log, 4, 5)
extra = pd.DataFrame([
    ("All companies", "All", "Comparability", "year end",
     "Year ends differ (March, June, September, December). Latest years range from Jun 2025 to Mar 2026, so market conditions are not identical."),
    ("All companies", "All", "Source", "source document",
     "Companies House accounts PDFs could not be opened in this environment. Figures were taken from the same audited group statements published by each company (annual report or audited results announcement). Each row cites its source."),
    ("Ashmore Group plc", "FY2026", "Missing year", "all fields",
     "FY2026 (year to 30 Jun 2026) results exist but the full statements could not be retrieved, so FY2023-FY2025 are analysed."),
    ("AJ Bell plc", "All", "Source", "source document",
     "Taken from audited results announcements, not the full annual report, which could not be read past the auditor's report."),
    ("Polar Capital Holdings plc", "All", "Missing value", "lease_liabilities",
     "Lease liabilities are within other balance sheet lines, so debt-to-equity is shown as n/a rather than an understated 0."),
    ("Jupiter Fund Management plc", "FY2023", "Unusual figure", "operating_profit",
     "Operating profit (86.0) is before a 76.2 goodwill impairment, which is shown as a separate line."),
    ("Liontrust Asset Management plc", "All", "Comparability", "net_revenue",
     "Net revenue column holds gross profit (revenue less cost of sales), the nearest equivalent line."),
], columns=l_heads[:4] + ["Detail"])
extra.columns = ["company", "fy_label", "type", "field", "detail"]
log_all = pd.concat([extra, quality], ignore_index=True)
for r, row in enumerate(log_all.itertuples(index=False), start=5):
    for c, v in enumerate(row, 1):
        cell = ws_log.cell(row=r, column=c, value=v)
        cell.font, cell.alignment, cell.border = F_BODY, WRAP, BOX
    fill = {"Missing value": WARN_FILL, "Missing year": WARN_FILL, "Unusual figure": BAD_FILL}.get(row.type)
    if fill:
        ws_log.cell(row=r, column=3).fill = PatternFill("solid", fgColor=fill)
    ws_log.row_dimensions[r].height = 30 if len(str(row.detail)) < 110 else 44
set_widths(ws_log, {"A": 30, "B": 12, "C": 16, "D": 20, "E": 100})
ws_log.freeze_panes = "A5"
ws_log.auto_filter.ref = f"A4:E{4 + len(log_all)}"
page_setup(ws_log)

# ================================================================ SQL checks
title_band(ws_sql, "SQL checks (results)",
           "Output of sql/queries.sql run on outputs/financials.db by src/analysis.py. Static values, refreshed when the script runs.", 9)
r = 4
labels = {
    "latest_year_snapshot": "Latest year for each company",
    "three_year_average": "Three-year averages",
    "revenue_declines": "Years where revenue fell",
    "loss_or_thin_margin_years": "Loss or margin below 5%",
    "missing_inputs": "Ratios that could not be calculated",
}
pct_cols = {"revenue_growth", "pbt_margin", "avg_revenue_growth", "avg_pbt_margin"}
x_cols = {"current_ratio", "debt_to_equity", "interest_cover", "avg_current_ratio", "lowest_current_ratio"}
for name, frame in sql_results.items():
    ws_sql.cell(row=r, column=1, value=labels.get(name, name)).font = F_SECTION
    r += 1
    for i, h in enumerate(frame.columns, 1):
        ws_sql.cell(row=r, column=i, value=h.replace("_", " "))
    style_header(ws_sql, r, len(frame.columns), height=28)
    for row in frame.itertuples(index=False):
        r += 1
        for i, v in enumerate(row, 1):
            col = frame.columns[i - 1]
            val = None if pd.isna(v) else v
            cell = ws_sql.cell(row=r, column=i, value=val if val is not None else "n/a")
            cell.font, cell.border = F_BODY, BOX
            if col in pct_cols:
                cell.number_format = PCT
            elif col in x_cols:
                cell.number_format = COVER if col == "interest_cover" else TIMES
            elif col.endswith("_m"):
                cell.number_format = GBP_M
    r += 2
set_widths(ws_sql, {"A": 30, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16})
ws_sql.freeze_panes = "A4"
page_setup(ws_sql)

# ================================================================ Sources
title_band(ws_src, "Sources",
           "Every figure traces to one of these documents. All accessed 4-5 November 2026.", 6)
s_heads = ["Source ID", "Company", "Document", "Years covered", "Link", "Accessed"]
for i, h in enumerate(s_heads, 1):
    ws_src.cell(row=4, column=i, value=h)
style_header(ws_src, 4, 6)
for r, row in enumerate(sources.itertuples(index=False), start=5):
    vals = [row.source_id, row.company, row.document, row.years_covered, row.url, row.accessed]
    for c, v in enumerate(vals, 1):
        cell = ws_src.cell(row=r, column=c, value=v)
        cell.font, cell.alignment, cell.border = F_BODY, WRAP, BOX
    ws_src.cell(row=r, column=5).hyperlink = row.url
    ws_src.cell(row=r, column=5).font = F_URL
    ws_src.row_dimensions[r].height = 44
r0 = 5 + len(sources) + 1
ws_src.cell(row=r0, column=1, value="Companies House pages (company numbers, filing records)").font = F_SECTION
for i, row in enumerate(companies.itertuples(index=False), start=r0 + 1):
    ws_src.cell(row=i, column=2, value=row.company).font = F_BODY
    c = ws_src.cell(row=i, column=5, value=row.companies_house_link)
    c.hyperlink, c.font = row.companies_house_link, F_URL
    ws_src.cell(row=i, column=3, value=row.check_note).font = F_BODY
    ws_src.cell(row=i, column=3).alignment = WRAP
    ws_src.row_dimensions[i].height = 30
set_widths(ws_src, {"A": 10, "B": 30, "C": 48, "D": 16, "E": 80, "F": 12})
ws_src.freeze_panes = "A5"
page_setup(ws_src)

# ================================================================ Project story
title_band(ws_story, "Project story: from raw data to stakeholder brief",
           "How the review was built, step by step, and what it means for the people who will use it.", 3)
STORY = [
    ("1. Objective", [
        ("Main question", "Which UK asset and wealth managers look financially stronger, and what risks should a lender or investor examine?"),
        ("Why this sector", "These firms report current and non-current items, so all five ratios can be calculated. Banks and insurers do not."),
        ("What it is not", "Not investment advice. It does not claim any firm is safe or likely to fail; it shows where to ask questions."),
    ]),
    ("2. Raw data", [
        ("Companies", "AJ Bell (04503206), Ashmore (03675683), Jupiter (06150195), Liontrust (02954692), Polar Capital (04235369), Premier Miton (06306664)."),
        ("Source route", "Identified on Companies House. Figures taken from each company's published audited group statements, because Companies House PDFs could not be opened here. Accessed 4-5 Nov 2026."),
        ("What was collected", "24 company-years (6 x 4), 12 figures each, typed in exactly as printed. See Raw_Data (blue cells) and Sources."),
    ]),
    ("3. Data cleaning", [
        ("Common unit", "£000 figures divided by 1,000, so all figures are in £ millions (Clean_GBPm)."),
        ("Matching line names", "'Gross profit' and 'Net income' grouped as net revenue. Debt = borrowings + lease liabilities."),
        ("No guessing", "Missing figures stay blank; any ratio needing them shows n/a."),
    ]),
    ("4. Exploration", [
        ("Scale", "Latest revenue ranged from £63.3m (Premier Miton) to £465.7m (Jupiter); about £1.39bn in total."),
        ("Direction", "11 of 18 analysed company-years had falling revenue. Median growth was -1.6%. Only AJ Bell and Polar Capital grew every year."),
        ("Spread", "PBT margins ran from -0.3% to 75.2% (median 25.6%). Current ratios were all above 1.2x (median 2.5x)."),
    ]),
    ("5. Data quality", [
        ("Checks", "18 missing values and 8 unusual figures logged (Data_Log). Nothing estimated."),
        ("Accuracy", "All 144 ratio cells match the Python calculation. 0 formula errors."),
        ("Main gaps", "Polar Capital: no operating profit or lease line. Ashmore: no finance cost line and FY2026 not retrieved. AJ Bell: results announcements used."),
    ]),
    ("6. Analysis", [
        ("Stronger", "AJ Bell: growth of 33%, 24%, 18% and a 43.4% margin. Polar Capital: growth every year, 29.1% margin. Jupiter: margin up from 2.3% to 28.3%, £50m notes repaid."),
        ("Mixed", "Ashmore: 7.0x current ratio and no debt, but revenue down to £144.4m, and 47% of profit from investment income."),
        ("Weaker", "Liontrust: revenue down 45% since FY2023 and equity halved. Premier Miton: margin down to 3.7%."),
        ("KPI scorecard", "Seven KPIs rated green / amber / red (KPI_Scorecard sheet). Screening scores: AJ Bell 100%, Jupiter 100%, Polar Capital 100% (5 of 7 KPIs), Ashmore 67%, Liontrust 43%, Premier Miton 17%."),
    ]),
    ("7. Recommendations", [
        ("Lender", "Look past debt ratios to covenants, guarantees and pensions. Ask what makes up current assets and liabilities. Test the regulatory capital surplus against a further 20% revenue fall."),
        ("Investor", "Separate fee profit from investment gains. Track assets under management and net flows each quarter. Check dividend cover where equity is falling."),
        ("Next update", "Add Ashmore FY2026 and Polar Capital lease figures when available."),
    ]),
    ("8. Business strategy", [
        ("AJ Bell", "How would margins hold if UK interest rates fall and interest on client cash drops?"),
        ("Polar Capital", "How dependent is revenue on a few technology funds and on performance fees?"),
        ("Jupiter", "Can cost discipline hold now that revenue is growing again?"),
        ("Ashmore", "How will fee income be rebuilt, and what is the plan for its large capital base?"),
        ("Liontrust", "Is the cost base shrinking as fast as revenue, and is the dividend covered?"),
        ("Premier Miton", "Is the business large enough to stand alone at a sustainable margin?"),
    ]),
    ("9. Stakeholder brief", [
        ("Situation", "Six UK-listed asset and wealth managers compared over their last three audited years."),
        ("Finding", "None relies on bank debt and all can meet short-term bills. The difference is direction: AJ Bell, Polar Capital and Jupiter are growing; Ashmore, Liontrust and Premier Miton are shrinking."),
        ("Why it matters", "The risk is not repaying a loan today. It is revenue falling until margins, dividends and capital buffers come under strain."),
        ("Suggested action", "Treat AJ Bell, Polar Capital and Jupiter as the stronger names on current evidence. Ask the others for recent fund flows, cost plans and capital surplus before committing."),
    ]),
]
row = 4
for section, items in STORY:
    c = ws_story.cell(row=row, column=1, value=section)
    c.font = Font(name=FONT, size=11, bold=True, color=WHITE)
    for col in (1, 2, 3):
        ws_story.cell(row=row, column=col).fill = FILL_INK
    ws_story.row_dimensions[row].height = 22
    row += 1
    for label, text in items:
        a = ws_story.cell(row=row, column=1, value=label)
        b = ws_story.cell(row=row, column=2, value=text)
        ws_story.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        brand_hit = next((co for co in companies["company"] if SHORT[co] == label), None)
        a.font = Font(name=FONT, size=10, bold=True, color=BRAND[brand_hit] if brand_hit else INK)
        b.font = F_BODY
        a.alignment = b.alignment = WRAP
        a.border = b.border = BOX
        ws_story.row_dimensions[row].height = 30 if len(text) < 120 else 44
        row += 1
    row += 1
ws_story.cell(row=row, column=1, value="Full detail: README.md and the report in the report folder.").font = F_NOTE
set_widths(ws_story, {"A": 24, "B": 90, "C": 30})
ws_story.freeze_panes = "A4"
page_setup(ws_story, landscape=False)

# ================================================================ KPI scorecard
bands = pd.read_csv(DATA / "kpi_thresholds.csv")
kpi_order = pd.read_csv(OUT / "kpi_scorecard.csv")["company"].tolist()
ws = ws_kpi
n_k = len(bands)
total_cols = 2 + 2 * n_k + 2
title_band(ws, "KPI scorecard: latest year, green / amber / red",
           "Each KPI is a formula on the Ratios sheet. Change a yellow band below and every status updates. Bands are reading aids, not industry rules.",
           total_cols)

# --- bands (editable inputs)
ws["A4"] = "KPI bands (edit the yellow cells)"
ws["A4"].font = F_SECTION
b_heads = ["KPI", "Better when", "Green from", "Amber from", "Why this band"]
for i, h in enumerate(b_heads, 1):
    ws.cell(row=5, column=i, value=h)
style_header(ws, 5, 4, height=24)
ws.cell(row=5, column=5).font, ws.cell(row=5, column=5).fill = F_H, FILL_INK
ws.merge_cells(start_row=5, start_column=5, end_row=5, end_column=12)
for c in range(5, 13):
    ws.cell(row=5, column=c).fill = FILL_INK
YELLOW = PatternFill("solid", fgColor="FFF6CC")
pct_kpis = {"revenue_growth", "pbt_margin", "operating_margin", "equity_change_3y"}
band_row = {}
for i, b in enumerate(bands.itertuples(), start=6):
    band_row[b.kpi] = i
    fmt = PCT if b.kpi in pct_kpis else '0.00"x"'
    vals = [b.label, b.better, b.green_from, b.amber_from, b.why_this_band]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=i, column=c, value=v)
        cell.border = BOX
        cell.font = F_BODY
        cell.alignment = Alignment(vertical="center", wrap_text=(c == 5))
    for c in (3, 4):
        ws.cell(row=i, column=c).fill = YELLOW
        ws.cell(row=i, column=c).font = F_INPUT
        ws.cell(row=i, column=c).number_format = fmt
        ws.cell(row=i, column=c).alignment = CENTER
    ws.cell(row=i, column=2).alignment = CENTER
    ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=12)
    for c in range(5, 13):
        ws.cell(row=i, column=c).border = BOX
    ws.row_dimensions[i].height = 22
ws.cell(row=6 + n_k, column=1,
        value="'Higher' KPIs: green at or above 'Green from', amber at or above 'Amber from', otherwise red. "
              "'Lower' KPIs (debt-to-equity): green at or below 'Green from', amber at or below 'Amber from'.").font = F_NOTE

# --- scorecard
top = 6 + n_k + 3
ws.cell(row=top - 1, column=1, value="Latest-year KPI scorecard (sorted by screening score)").font = F_SECTION
ws.cell(row=top, column=1, value="Company")
ws.cell(row=top, column=2, value="Latest year")
ws.merge_cells(start_row=top, start_column=1, end_row=top + 1, end_column=1)
ws.merge_cells(start_row=top, start_column=2, end_row=top + 1, end_column=2)
ratio_col = {"revenue_growth": "G", "pbt_margin": "H", "operating_margin": "I",
             "current_ratio": "J", "debt_to_equity": "K", "interest_cover": "L"}
for k, b in enumerate(bands.itertuples()):
    c = 3 + 2 * k
    ws.cell(row=top, column=c, value=b.label)
    ws.merge_cells(start_row=top, start_column=c, end_row=top, end_column=c + 1)
    ws.cell(row=top + 1, column=c, value="Value")
    ws.cell(row=top + 1, column=c + 1, value="Status")
sc = 3 + 2 * n_k
ws.cell(row=top, column=sc, value="KPIs with data")
ws.cell(row=top, column=sc + 1, value="Screening score")
ws.merge_cells(start_row=top, start_column=sc, end_row=top + 1, end_column=sc)
ws.merge_cells(start_row=top, start_column=sc + 1, end_row=top + 1, end_column=sc + 1)
style_header(ws, top, total_cols, height=34)
style_header(ws, top + 1, total_cols, height=18)
KEY = f"{R}$O${FIRST}:$O${LAST}"
status_cells_all = []
for i, company in enumerate(kpi_order):
    r = top + 2 + i
    ws.cell(row=r, column=1, value=company)
    ws.cell(row=r, column=1).font = Font(name=FONT, size=10, bold=True, color=WHITE)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=BRAND[company])
    ws.cell(row=r, column=2, value=f'=IFERROR(INDEX({R}$C${FIRST}:$C${LAST},MATCH($A{r}&"|0",{KEY},0)),"n/a")')
    status_cells = []
    for k, b in enumerate(bands.itertuples()):
        vc, stc = 3 + 2 * k, 4 + 2 * k
        vref = f"{get_column_letter(vc)}{r}"
        if b.kpi == "equity_change_3y":
            f = (f'=IFERROR(INDEX(Clean_GBPm!$Q${FIRST}:$Q${LAST},MATCH($A{r}&"|0",{KEY},0))'
                 f'/INDEX(Clean_GBPm!$Q${FIRST}:$Q${LAST},MATCH($A{r}&"|3",{KEY},0))-1,"n/a")')
        else:
            col = ratio_col[b.kpi]
            f = f'=IFERROR(INDEX({R}${col}${FIRST}:${col}${LAST},MATCH($A{r}&"|0",{KEY},0)),"n/a")'
        ws.cell(row=r, column=vc, value=f)
        br = band_row[b.kpi]
        ws.cell(row=r, column=stc, value=(
            f'=IF(NOT(ISNUMBER({vref})),"No data",IF($B${br}="higher",'
            f'IF({vref}>=$C${br},"Green",IF({vref}>=$D${br},"Amber","Red")),'
            f'IF({vref}<=$C${br},"Green",IF({vref}<=$D${br},"Amber","Red"))))'))
        ws.cell(row=r, column=vc).number_format = PCT if b.kpi in pct_kpis else (COVER if b.kpi == "interest_cover" else TIMES)
        status_cells.append(f"{get_column_letter(stc)}{r}")
    status_cells_all += status_cells
    scored = "+".join(f'({s}<>"No data")' for s in status_cells)
    points = "+".join(f'IF({s}="Green",2,IF({s}="Amber",1,0))' for s in status_cells)
    ws.cell(row=r, column=sc, value=f"={scored}")
    ws.cell(row=r, column=sc + 1, value=f'=IF({get_column_letter(sc)}{r}>0,({points})/(2*{get_column_letter(sc)}{r}),"n/a")')
    ws.cell(row=r, column=sc + 1).number_format = "0%"
    for c in range(1, total_cols + 1):
        cell = ws.cell(row=r, column=c)
        cell.border = BOX
        if c > 1:
            cell.font = F_BODY
            cell.alignment = CENTER
    ws.cell(row=r, column=1).alignment = Alignment(vertical="center", indent=1)
    ws.cell(row=r, column=sc + 1).font = Font(name=FONT, size=11, bold=True, color=INK)
    ws.row_dimensions[r].height = 30
end = top + 1 + len(kpi_order)
for col_i in range(4, 4 + 2 * n_k, 2):
    L = get_column_letter(col_i)
    rng_s = f"{L}{top + 2}:{L}{end}"
    for word, fill, colour in (("Green", GOOD_FILL, "1E7B3A"), ("Amber", WARN_FILL, "B26A00"),
                               ("Red", BAD_FILL, "B42318"), ("No data", "EEF1F6", "5B6577")):
        ws.conditional_formatting.add(rng_s, CellIsRule(operator="equal", formula=[f'"{word}"'],
                                      fill=PatternFill("solid", fgColor=fill), font=Font(name=FONT, bold=True, color=colour)))
score_rng = f"{get_column_letter(sc + 1)}{top + 2}:{get_column_letter(sc + 1)}{end}"
ws.conditional_formatting.add(score_rng, CellIsRule(operator="greaterThanOrEqual", formula=["0.8"], fill=PatternFill("solid", fgColor=GOOD_FILL)))
ws.conditional_formatting.add(score_rng, CellIsRule(operator="lessThan", formula=["0.5"], fill=PatternFill("solid", fgColor=BAD_FILL)))
notes = [
    "Screening score = points ÷ maximum possible points, where green = 2, amber = 1, red = 0. KPIs with no data are left out, not scored as red.",
    "Polar Capital scores 100% on only 5 of 7 KPIs (no operating profit line or separate lease figure), so its score rests on less evidence.",
    "Equity change compares latest total equity with the base year three years earlier. Interest coverage is 'No data' where a company reports no interest cost.",
    "The score is a quick screen to decide where to look first. It is not a credit rating and does not show that any company is safe or at risk.",
]
for j, t in enumerate(notes):
    c = ws.cell(row=end + 2 + j, column=1, value=t)
    c.font = F_NOTE
    ws.merge_cells(start_row=end + 2 + j, start_column=1, end_row=end + 2 + j, end_column=total_cols)
set_widths(ws, {"A": 32, "B": 11, **{get_column_letter(c): 10.5 for c in range(3, total_cols + 1)}})
ws.column_dimensions["E"].width = 10.5
ws.freeze_panes = ws.cell(row=top + 2, column=3)
page_setup(ws)

# ================================================================ README sheet
import re

from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.hyperlink import Hyperlink

SHEET_NAMES = set(wb.sheetnames)
LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")


def plain(text: str) -> str:
    text = LINK_RE.sub(lambda m: m.group(1), text)
    return text.replace("**", "").replace("`", "").replace("*", "").strip()


def first_link(text: str):
    m = LINK_RE.search(text)
    return m.group(2) if m else None


def bold_names(text: str) -> list[str]:
    return re.findall(r"\*\*([^*]+)\*\*", text)


from config import github_link


def resolve(target: str) -> str | None:
    """Turn a README link into a link that works from a downloaded workbook.

    Files and folders in the project point to their page in the GitHub
    repository, so the link still works when the workbook is downloaded on
    its own.
    """
    if target.startswith("http"):
        return target
    if target.startswith("#"):
        return None                       # section anchor, handled separately
    return github_link(target)


def slug(heading: str) -> str:
    s = re.sub(r"[^\w\s-]", "", heading.lower()).strip()
    return re.sub(r"\s+", "-", s)


ws = ws_readme
md = (ROOT / "README.md").read_text(encoding="utf-8").split("\n")
WIDTH_COLS = 8
COL_W = 19
set_widths(ws, {get_column_letter(c): COL_W for c in range(1, WIDTH_COLS + 1)})
CHARS_PER_ROW = int(COL_W * WIDTH_COLS * 0.92)
anchors, pending = {}, []
embedded = set()
row = 1
i = 0
in_code = False


def text_row(r, text, font=F_BODY, fill=None, indent=0, link=None, height=None):
    cell = ws.cell(row=r, column=1, value=text)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=WIDTH_COLS)
    cell.font = font
    cell.alignment = Alignment(wrap_text=True, vertical="top", indent=indent)
    if fill:
        for c in range(1, WIDTH_COLS + 1):
            ws.cell(row=r, column=c).fill = fill
    lines = max(1, -(-len(text) // CHARS_PER_ROW))
    ws.row_dimensions[r].height = height or max(15, 14 * lines + 3)
    if link:
        target = resolve(link)
        if target:
            cell.hyperlink = target
            if font is F_BODY:
                cell.font = Font(name=FONT, size=10, color="1155CC", underline="single")
        elif link.startswith("#"):
            pending.append((cell, link[1:]))
    return cell


while i < len(md):
    line = md[i]
    s = line.strip()
    if s.startswith("```"):
        in_code = not in_code
        if not in_code:
            row += 1          # gap after a code block
        i += 1
        continue
    if in_code:
        text_row(row, line if line else " ", font=Font(name="Consolas", size=9, color=INK),
                 fill=PatternFill("solid", fgColor="EEF1F6"), indent=1, height=15)
        row += 1
        i += 1
        continue
    if not s or s == "---":
        row += 0 if not s else 1
        i += 1
        continue
    if s.startswith("# "):
        title_band(ws, plain(s[2:]), "Project guide. The same text as README.md in the project folder.", WIDTH_COLS)
        row = 4
    elif s.startswith("## "):
        row += 1
        head = plain(s[3:])
        text_row(row, head, font=Font(name=FONT, size=12, bold=True, color=WHITE), fill=FILL_INK, indent=1, height=22)
        anchors[slug(head)] = row
        row += 1
    elif s.startswith("!["):
        img_path = ROOT / first_link(s)
        if img_path.name.startswith("excel_") or img_path in embedded:
            # screenshots of this workbook are not embedded in it; link instead
            text_row(row, f"Image: {first_link(s)}", link=first_link(s))
            row += 1
        elif img_path.exists():
            embedded.add(img_path)
            pic = XLImage(str(img_path))
            scale = (COL_W * WIDTH_COLS * 7.0) / pic.width
            pic.width, pic.height = int(pic.width * scale), int(pic.height * scale)
            ws.add_image(pic, f"A{row}")
            rows_needed = int(pic.height / 20) + 1
            for r in range(row, row + rows_needed):
                ws.row_dimensions[r].height = 15
            row += rows_needed + 1
    elif s.startswith("|"):
        table = []
        while i < len(md) and md[i].strip().startswith("|"):
            cells = [c.strip() for c in md[i].strip().strip("|").split("|")]
            if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                table.append(cells)
            i += 1
        ncols = len(table[0])
        if ncols > WIDTH_COLS:
            spans = [1] * ncols
        else:
            spans = [WIDTH_COLS // ncols] * ncols
            spans[-1] += WIDTH_COLS - sum(spans)
        if ncols == 2:
            spans = [3, 5]
        for t_i, cells in enumerate(table):
            col = 1
            longest = 1
            for c_i, raw_txt in enumerate(cells):
                txt = plain(raw_txt)
                span = spans[c_i]
                cell = ws.cell(row=row, column=col, value=txt)
                if span > 1:
                    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
                for cc in range(col, col + span):
                    ws.cell(row=row, column=cc).border = BOX
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if t_i == 0:
                    cell.font = F_H
                    for cc in range(col, col + span):
                        ws.cell(row=row, column=cc).fill = FILL_INK
                else:
                    cell.font = F_BOLD if c_i == 0 else F_BODY
                    name = txt if txt in SHEET_NAMES else None
                    link = first_link(raw_txt)
                    if name:
                        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{name}'!A1")
                        cell.font = Font(name=FONT, size=10, bold=True, color="1155CC", underline="single")
                    elif link and resolve(link):
                        cell.hyperlink = resolve(link)
                        cell.font = Font(name=FONT, size=10, color="1155CC", underline="single", bold=(c_i == 0))
                    company = next((co for co in companies["company"] if co == txt or SHORT[co] == txt), None)
                    if company and c_i == 0:
                        cell.font = Font(name=FONT, size=10, bold=True, color=BRAND[company])
                longest = max(longest, -(-len(txt) // int(COL_W * span * 1.05)))
                col += span
            ws.row_dimensions[row].height = max(18, 14 * longest + 4)
            row += 1
        row += 1
        continue
    elif s.startswith(">"):
        txt = plain(s.lstrip("> "))
        if txt:
            text_row(row, txt, font=Font(name=FONT, size=10, italic=True, color=INK),
                     fill=PatternFill("solid", fgColor="EEF1F6"), indent=1)
            row += 1
    elif re.match(r"^(- |\d+\. )", s):
        marker = "•" if s.startswith("- ") else s.split(" ", 1)[0]
        body = s[2:] if s.startswith("- ") else s.split(" ", 1)[1]
        text_row(row, f"{marker}  {plain(body)}", indent=1, link=first_link(body))
        row += 1
    else:
        text_row(row, plain(s), font=F_BOLD if s.startswith("**") and s.endswith("**") else F_BODY,
                 link=first_link(s) if s.count("](") == 1 else None)
        row += 1
    i += 1

for cell, target in pending:
    if target in anchors:
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"README!A{anchors[target]}")
        cell.font = Font(name=FONT, size=10, color="1155CC", underline="single")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"
page_setup(ws, landscape=False)

# ---- workbook-wide defaults
for sheet in wb.worksheets:
    sheet.sheet_properties.tabColor = INK if sheet.title == "Dashboard" else "8A94A6"
wb.active = 0
wb.calculation.fullCalcOnLoad = True   # Excel recalculates every formula on open
OUT.mkdir(exist_ok=True)
wb.save(FILE)
print(f"Saved {FILE}")

# ---- optional: one file per sheet, used by src/make_screenshots.py
import sys

if "--screens" in sys.argv:
    shots = OUT / "_screens"
    shots.mkdir(exist_ok=True)
    for target in ("Dashboard", "KPI_Scorecard", "Ratios", "Project_Story", "Data_Log"):
        for sheet in wb.worksheets:
            sheet.sheet_state = "visible" if sheet.title == target else "hidden"
        wb.active = wb.sheetnames.index(target)
        page = wb[target]
        old = page.page_setup.fitToHeight
        page.page_setup.fitToHeight = 1          # whole sheet on one page for the image
        wb.save(shots / f"{target}.xlsx")
        page.page_setup.fitToHeight = old
    print(f"Saved per-sheet copies to {shots}")
