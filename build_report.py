"""
Write the report: report/UK_Financial_Health_Review_Report.docx and .pdf
The main report is about three pages, followed by a tools section and an
appendix holding all project code (SQL and Python).
Run: python src/build_report.py
"""

import subprocess
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "report" / "UK_Financial_Health_Review_Report.docx"
INK = RGBColor(0x1F, 0x2A, 0x44)
MUTED = RGBColor(0x5B, 0x65, 0x77)

doc = Document()
sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
for side in ("left_margin", "right_margin"):
    setattr(sec, side, Cm(2.0))
sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.8)

base = doc.styles["Normal"]
base.font.name, base.font.size = "Arial", Pt(10)
base.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
base.paragraph_format.space_after = Pt(5)
base.paragraph_format.line_spacing = 1.12
for name, size in (("Heading 1", 13), ("Heading 2", 11)):
    st = doc.styles[name]
    st.font.name, st.font.size, st.font.bold = "Arial", Pt(size), True
    st.font.color.rgb = INK
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.paragraph_format.space_before, st.paragraph_format.space_after = Pt(10), Pt(4)


def para(text, bold_lead=None, style=None, size=None, colour=None, italic=False):
    p = doc.add_paragraph(style=style)
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
    r = p.add_run(text)
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    if colour:
        r.font.color.rgb = colour
    return p


def bullet(text, bold_lead=None):
    p = para(text, bold_lead, style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    return p


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


# ---------------------------------------------------------------- title
t = doc.add_paragraph()
r = t.add_run("UK Company Financial Health Review")
r.bold, r.font.size, r.font.color.rgb = True, Pt(18), INK
para("UK-listed asset and wealth managers · data from audited group accounts · accessed 4-5 November 2026",
     size=9.5, colour=MUTED)

doc.add_heading("1. The question", level=1)
para("Which companies in one UK sector look financially stronger, and what should a lender or investor "
     "look into before relying on them? I chose six UK-listed asset and wealth managers: AJ Bell, Ashmore, "
     "Jupiter, Liontrust, Polar Capital and Premier Miton. These firms report their balance sheets in current "
     "and non-current items, so all five ratios can be worked out. Banks and insurers do not report this way, "
     "which is why I did not use them.")

doc.add_heading("2. Data sources", level=1)
para("Each company was identified by its Companies House number (04503206, 03675683, 06150195, 02954692, "
     "04235369 and 06306664). The Companies House pages confirmed the company names, the year ends and, for "
     "most of them, the group accounts filed. The accounts PDFs held at Companies House could not be opened "
     "in this working environment. I therefore took the figures from the same audited group statements that "
     "each company publishes: its annual report, or its audited results announcement where the full report "
     "could not be read. Every figure in the workbook lists its source document and web link.")
para("For each company I used two documents, each showing two years. That gives four years of figures: "
     "three years for the ratios, plus an earlier year as the starting point for revenue growth. Ashmore's "
     "results for the year to June 2026 have been published, but I could not retrieve the full statements. "
     "Its analysis therefore stops at FY2025.")

doc.add_heading("3. Method", level=1)
para("The figures were typed in exactly as printed. Amounts in £000 were then converted to £ millions. "
     "Blank fields were left blank rather than estimated. Every ratio in the workbook is a live Excel "
     "formula, and the same calculations are repeated in Python and checked with SQL queries.")
for lead, text in [
    ("Revenue growth", " = this year's revenue ÷ last year's revenue − 1."),
    ("Profit margin", " = profit before tax ÷ revenue. I used profit before tax because Polar Capital does not "
                      "report an operating profit line, so it is the one profit measure all six share. "
                      "Operating margin is shown alongside as a check."),
    ("Current ratio", " = current assets ÷ current liabilities."),
    ("Debt-to-equity", " = (borrowings + lease liabilities) ÷ total equity."),
    ("Interest coverage", " = (profit before tax + finance costs) ÷ finance costs."),
]:
    bullet(text, lead)

doc.add_heading("4. Findings", level=1)
para("The table shows each company's latest year. The workbook sets out all three years.")
rows = [
    ("Company", "Year", "Revenue £m", "Growth", "PBT margin", "Current ratio", "Debt / equity"),
    ("AJ Bell", "FY2025", "317.8", "+18.0%", "43.4%", "3.65x", "0.06x"),
    ("Polar Capital", "FY2026", "264.3", "+19.0%", "29.1%", "2.53x", "n/a"),
    ("Jupiter", "FY2025", "465.7", "+15.7%", "28.3%", "2.46x", "0.04x"),
    ("Ashmore", "FY2025", "144.4", "−23.6%", "75.2%", "7.01x", "0.01x"),
    ("Liontrust", "FY2026", "134.4", "−20.9%", "10.7%", "1.34x", "0.03x"),
    ("Premier Miton", "FY2025", "63.3", "−1.1%", "3.7%", "1.27x", "0.02x"),
]
brand = {"AJ Bell": "5B2C83", "Polar Capital": "2F80C0", "Jupiter": "00796B", "Ashmore": "14233C",
         "Liontrust": "E07B00", "Premier Miton": "C2185B"}
tbl = doc.add_table(rows=len(rows), cols=len(rows[0]))
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.style = "Table Grid"
for i, row in enumerate(rows):
    for j, val in enumerate(row):
        cell = tbl.cell(i, j)
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(val)
        run.font.size = Pt(9)
        if i == 0:
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            shade(cell, "1F2A44")
        elif j == 0:
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            shade(cell, brand[val])
        if j > 0:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
para("Polar Capital's lease liabilities are not shown separately, so its debt-to-equity is left as n/a. "
     "Interest coverage is in the workbook. It is very high wherever the only interest cost is on office "
     "leases.", size=8.5, colour=MUTED, italic=True)

doc.add_heading("KPI screening", level=2)
para("The five ratios, plus operating margin and the change in equity over three years, were turned into seven "
     "KPIs, each rated green, amber or red against fixed bands (for example, a profit margin of 25% or more is "
     "green, under 10% is red). Green scores 2 points, amber 1 and red 0, and KPIs with no data are left out. "
     "On this screen AJ Bell and Jupiter score 100% on all seven KPIs, and Polar Capital 100% on the five it "
     "reports. Ashmore scores 67%, green on its balance sheet but red on revenue growth and equity change. "
     "Liontrust scores 43% and Premier Miton 17%. The bands are my own reading aids, and the score shows where "
     "to look first; it is not a rating.")

doc.add_heading("Stronger on these measures", level=2)
bullet(" Revenue rose from £218.2m in FY2023 to £317.8m in FY2025, growing 33.2%, 23.5% and 18.0% a year. "
       "The profit margin edged up from 40.2% to 43.4%. The current ratio held at about 3.6x, and its only "
       "debt is £12.1m of leases against £217.5m of equity.", "AJ Bell.")
bullet(" Revenue grew in each of the three years, by 6.7%, 13.9% and 19.0%, reaching £264.3m in FY2026. "
       "The margin stayed between 23% and 29%, and the current ratio was steady at about 2.5x. Its only "
       "interest cost is about £0.2m on leases.", "Polar Capital.")
bullet(" The FY2023 margin was only 2.3% because of a £76.2m goodwill write-down. It then recovered to 21.9% "
       "and 28.3%. Jupiter repaid its £50m subordinated notes in April 2025, so debt-to-equity fell from "
       "0.12x to 0.04x. The current ratio improved from 2.09x to 2.46x.", "Jupiter.")

doc.add_heading("Mixed", level=2)
bullet(" Ashmore has the largest liquidity cushion (current ratio 7.0x to 10.7x) and no debt. However, "
       "revenue fell from £193.2m to £144.4m over the three years, and equity fell from £913.0m to £790.8m. "
       "Its 75.2% margin is flattered by £51.1m of finance income, which is 47% of profit before tax. "
       "Operating margin, at 30% to 40%, gives a better view of the fee business.", "Ashmore.")

doc.add_heading("Weaker on these measures", level=2)
bullet(" Revenue fell in each of the three years, from £243.3m in FY2023 to £134.4m in FY2026, a drop of 45%. "
       "It made a small loss before tax in FY2024. Equity fell from £221.0m to £107.9m. Its current ratio "
       "(1.34x) is the second lowest in the group.", "Liontrust.")
bullet(" Revenue fell from £90.2m in FY2022 to £63.3m in FY2025, about 30% lower. The margin narrowed each "
       "year, from 7.9% to 4.9% and then 3.7%. It has no bank debt, but has the lowest current ratio at "
       "1.27x.", "Premier Miton.")

doc.add_heading("5. Risks a lender or investor should examine", level=1)
bullet(" Fees are charged on assets under management, so a market fall or client outflows reduce revenue "
       "quickly. This shows clearly at Liontrust, Premier Miton and Ashmore. Ask for the latest flows and "
       "assets under management, not just the accounts.", "Market and outflow risk.")
bullet(" Ashmore's profit depends heavily on returns from its own investments, and Jupiter's FY2023 result "
       "was hit by a one-off write-down. Check how much profit comes from the core fee business.",
       "Quality of profit.")
bullet(" For AJ Bell and other platforms, part of revenue is interest earned on client cash. A cut in UK "
       "interest rates could reduce it. Ask what share of revenue this is.", "Interest-rate sensitivity.")
bullet(" Money moving between investors and funds can sit in receivables and payables. These amounts are "
       "large and largely offset each other, which pulls the current ratio towards 1x. Ask for the "
       "breakdown before treating 1.3x as a warning sign.", "What the current ratio contains.")
bullet(" Liontrust's equity halved in three years as dividends and buybacks ran ahead of profit. Check "
       "dividend cover and the regulatory capital surplus reported in each firm's Pillar 3 or ICARA "
       "disclosures.", "Shrinking equity and capital needs.")
bullet(" Most of these firms owe little apart from leases, so debt ratios say little about them. A lender "
       "should look at covenants, guarantees, seed-capital commitments and any defined benefit pension "
       "scheme.", "Hidden obligations.")

doc.add_heading("6. Limits of the data", level=1)
bullet(" The year ends are March, June, September and December, so the latest years run from June 2025 to "
       "March 2026. Market conditions were not the same across these dates.", "Different dates.")
bullet(" The figures come from the audited statements each company published, not from Companies House "
       "PDFs, which could not be read here. AJ Bell's figures come from its audited results "
       "announcements. The Companies House pages for AJ Bell and Ashmore came from older cached copies, "
       "and Jupiter's page could not be opened.", "Source route.")
bullet(" Polar Capital does not report an operating profit line or separate lease liabilities. Ashmore "
       "reports no finance cost line. These cells are marked n/a, not estimated.", "Missing figures.")
bullet(" Each firm labels its lines differently ('net revenue', 'gross profit', 'net income'). Some prior "
       "years were re-presented. Each case is listed in the workbook's Data_Log sheet.",
       "Reporting differences.")
bullet(" These are listed groups, but the same approach applied to private companies is harder. Small "
       "companies can file abridged accounts with no profit and loss account, so growth and margin cannot "
       "be calculated at all.", "Private companies.")

doc.add_heading("7. Conclusion", level=1)
para("On these five measures, AJ Bell, Polar Capital and Jupiter look the strongest. They show growing revenue, "
     "healthy margins, comfortable current ratios and little or no borrowing. Ashmore has a very strong "
     "balance sheet but shrinking fee income. Liontrust and Premier Miton show falling revenue and thin "
     "margins, and their questions are about the trend more than about solvency today. None of this proves "
     "that any firm is safe or likely to fail. The ratios show where a lender or investor should ask "
     "further questions before deciding.")

import csv
import sys

sys.path.insert(0, str(ROOT / "src"))
from config import REPO_URL, github_link  # noqa: E402


def add_link(paragraph, url: str, text: str, size: float = 8.5):
    """Insert a clickable hyperlink run into a paragraph."""
    r_id = paragraph.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                                    is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    for tag, attr in (("w:color", "1155CC"), ("w:u", "single"), ("w:sz", str(int(size * 2)))):
        el = OxmlElement(tag)
        el.set(qn("w:val"), attr)
        rpr.append(el)
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Arial")
    fonts.set(qn("w:hAnsi"), "Arial")
    rpr.append(fonts)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    run.append(t)
    link.append(run)
    paragraph._p.append(link)


doc.add_heading("8. Sources and project files", level=1)
para("All links were checked on 5 November 2026. Figures were accessed on 4-5 November 2026.", size=8.5, colour=MUTED)
with open(ROOT / "data" / "sources.csv", newline="", encoding="utf-8") as fh:
    for src in csv.DictReader(fh):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        lead = p.add_run(f"{src['source_id']}  {src['company']}: {src['document']}. ")
        lead.font.size = Pt(8.5)
        add_link(p, src["url"], "Open source")
with open(ROOT / "data" / "companies.csv", newline="", encoding="utf-8") as fh:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run("Companies House pages: ")
    r.font.size, r.bold = Pt(8.5), True
    for k, co in enumerate(csv.DictReader(fh)):
        if k:
            sep = p.add_run(" · ")
            sep.font.size = Pt(8.5)
        add_link(p, co["companies_house_link"], f"{co['short_name']} ({co['company_number']})")
p = doc.add_paragraph()
r = p.add_run("Project repository (workbook, data, code and dashboard images): ")
r.font.size, r.bold = Pt(8.5), True
add_link(p, REPO_URL, REPO_URL)
p = doc.add_paragraph()
r = p.add_run("Direct files: ")
r.font.size, r.bold = Pt(8.5), True
for k, (label, path) in enumerate([("Excel workbook", "outputs/UK_Financial_Health_Review.xlsx"),
                                   ("Dashboard image", "images/dashboard.png"),
                                   ("KPI scorecard", "outputs/kpi_scorecard.csv"),
                                   ("SQL queries", "sql/queries.sql"),
                                   ("Python code", "src")]):
    if k:
        sep = p.add_run(" · ")
        sep.font.size = Pt(8.5)
    add_link(p, github_link(path), label)

doc.add_heading("9. Tools used", level=1)
tools = [
    ("Tool", "What it was used for"),
    ("Companies House (find-and-update service)", "Confirming company names and numbers, SIC codes, year ends and filed accounts records."),
    ("Company annual reports and audited results", "Source of every financial figure (see the Sources sheet for links and access dates)."),
    ("Python 3 with pandas", "Importing the raw figures, converting units, quality checks, ratio and KPI calculations (src/analysis.py)."),
    ("SQL (SQLite via Python's sqlite3)", "Checks on the results: latest-year snapshot, three-year averages, revenue declines, thin margins, missing inputs (sql/queries.sql)."),
    ("Microsoft Excel format (openpyxl)", "Workbook with live formulas, conditional formatting and native charts. The Dashboard sheet is the BI layer of the project."),
    ("matplotlib and Pillow", "Summary dashboard image and trimmed workbook screenshots for the README (src/make_dashboard.py, src/make_screenshots.py)."),
    ("GitHub", "Hosting the project; links in the workbook and report point to the repository so they work after download (src/config.py, src/check_links.py)."),
    ("python-docx and LibreOffice", "Writing this report and converting it to PDF; LibreOffice also recalculated and checked every workbook formula."),
    ("BI platform", "No separate BI tool (such as Power BI or Tableau) was used. outputs/ratios.csv and outputs/financials.db can be loaded into one directly."),
]
tt = doc.add_table(rows=len(tools), cols=2)
tt.style = "Table Grid"
for i, (a, b) in enumerate(tools):
    for j, val in enumerate((a, b)):
        cell = tt.cell(i, j)
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(val)
        run.font.size = Pt(9)
        if i == 0:
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            shade(cell, "1F2A44")
        elif j == 0:
            run.bold = True
for row in tt.rows:
    row.cells[0].width, row.cells[1].width = Cm(5.5), Cm(11.5)

# ---------------------------------------------------------------- appendix: code
CODE_FILES = [
    ("A1. SQL checks", "sql/queries.sql"),
    ("A2. Python: import, cleaning, quality checks, ratios and SQL", "src/analysis.py"),
    ("A3. Python: Excel workbook and dashboard (BI layer)", "src/build_workbook.py"),
    ("A4. Python: dashboard image", "src/make_dashboard.py"),
    ("A5. Python: workbook screenshots", "src/make_screenshots.py"),
    ("A6. Python: this report", "src/build_report.py"),
    ("A7. Python: shared settings (GitHub address)", "src/config.py"),
    ("A8. Python: link checker", "src/check_links.py"),
]
doc.add_page_break()
doc.add_heading("Appendix: project code", level=1)
para("The complete code used for the analysis, as stored in the project folder. Run the scripts in this order: "
     "analysis.py, build_workbook.py, make_dashboard.py, make_screenshots.py, build_report.py, then check_links.py. The KPI bands are in data/kpi_thresholds.csv. No API keys or passwords are included; "
     "the optional Companies House download reads a key from the CH_API_KEY environment variable.",
     size=9, colour=MUTED)
for title, rel in CODE_FILES:
    doc.add_heading(f"{title}  ({rel})", level=2)
    text = (ROOT / rel).read_text(encoding="utf-8").rstrip("\n")
    for line in text.split("\n"):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_after = pf.space_before = Pt(0)
        pf.line_spacing = 1.0
        pf.left_indent = Cm(0.2)
        run = p.add_run(line if line else " ")
        run.font.name = "Courier New"
        run.element.rPr.rFonts.set(qn("w:eastAsia"), "Courier New")
        run.font.size = Pt(6.5)
        run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)

# page numbers in the footer
footer = doc.sections[0].footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for kind, text in (("begin", None), (None, "PAGE"), ("end", None)):
    run = footer.add_run()
    run.font.size, run.font.color.rgb = Pt(8), MUTED
    if kind:
        fld = OxmlElement("w:fldChar")
        fld.set(qn("w:fldCharType"), kind)
        run._r.append(fld)
    else:
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = text
        run._r.append(instr)
lead = footer.runs[0]._r
pre = footer.add_run("UK Company Financial Health Review · page ")
pre.font.size, pre.font.color.rgb = Pt(8), MUTED
lead.addprevious(pre._r)

OUT.parent.mkdir(exist_ok=True)
doc.save(OUT)
print(f"Saved {OUT}")

# PDF copy (needs LibreOffice; skipped quietly if it is not installed)
try:
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(OUT.parent), str(OUT)],
                   check=True, capture_output=True, timeout=180)
    print(f"Saved {OUT.with_suffix('.pdf')}")
except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
    print("LibreOffice not found: open the .docx and use File > Save as PDF.")
