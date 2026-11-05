"""
Check every link in the project before uploading to GitHub.

  - README.md: relative file links, image paths and #section anchors
  - Excel workbook: links between sheets, and links to files in this repository
  - Word report: links to files in this repository
  - With --online: also requests each external web address

Run from the project folder:
    python src/check_links.py            # offline checks
    python src/check_links.py --online   # also test external URLs (needs requests)
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from config import BRANCH, REPO_URL  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
problems: list[str] = []
external: set[str] = set()


def github_slug(heading: str) -> str:
    s = re.sub(r"[^\w\s-]", "", heading.strip().lower())
    return re.sub(r"\s", "-", s)


def repo_path(url: str) -> Path | None:
    """Map a link to this repository back to a local path."""
    for kind in ("blob", "tree", "raw"):
        prefix = f"{REPO_URL}/{kind}/{BRANCH}/"
        if url.startswith(prefix):
            return ROOT / url[len(prefix):]
    return None


def check_url(url: str, where: str) -> None:
    local = repo_path(url)
    if local is not None:
        if not local.exists():
            problems.append(f"{where}: {url} -> missing file {local.relative_to(ROOT)}")
    elif url.startswith("http"):
        external.add(url)


# ---- README.md
readme = (ROOT / "README.md").read_text(encoding="utf-8")
body = re.sub(r"```.*?```", "", readme, flags=re.S)          # ignore code blocks
anchors = {github_slug(h) for h in re.findall(r"^#{1,6} (.+)$", body, re.M)}
links = re.findall(r"!?\[[^\]]*\]\(([^)\s]+)\)", body)
for target in links:
    if target.startswith("#"):
        if target[1:] not in anchors:
            problems.append(f"README.md: anchor {target} has no matching heading")
    elif target.startswith("http"):
        check_url(target, "README.md")
    elif not (ROOT / target).exists():
        problems.append(f"README.md: {target} does not exist")
print(f"README.md: {len(links)} links checked")

# ---- Excel workbook
wb = load_workbook(ROOT / "outputs" / "UK_Financial_Health_Review.xlsx")
n = 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            h = cell.hyperlink
            if not h:
                continue
            n += 1
            where = f"Workbook {ws.title}!{cell.coordinate}"
            if h.location:
                if h.location.split("!")[0].strip("'") not in wb.sheetnames:
                    problems.append(f"{where}: sheet in {h.location} does not exist")
            elif h.target and h.target.startswith("http"):
                check_url(h.target, where)
            else:
                problems.append(f"{where}: relative link {h.target} breaks once the file is downloaded")
print(f"Workbook: {n} links checked")

# ---- Word report
docx = ROOT / "report" / "UK_Financial_Health_Review_Report.docx"
with zipfile.ZipFile(docx) as z:
    rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
urls = re.findall(r'Target="(http[^"]+)"[^>]*TargetMode="External"', rels)
for u in urls:
    check_url(u.replace("&amp;", "&"), "Report")
print(f"Report: {len(urls)} links checked")

# ---- optional online check
if "--online" in sys.argv:
    import requests

    for url in sorted(external):
        try:
            r = requests.get(url, timeout=20, allow_redirects=True, headers={"User-Agent": "link-check"})
            if r.status_code >= 400:
                problems.append(f"HTTP {r.status_code}: {url}")
        except requests.RequestException as exc:
            problems.append(f"Unreachable: {url} ({exc.__class__.__name__})")
    print(f"Online: {len(external)} external addresses requested")
else:
    print(f"{len(external)} external addresses not requested (add --online to test them)")

if problems:
    print("\nProblems found:")
    for p in problems:
        print("  -", p)
    sys.exit(1)
print("\nAll links OK.")
