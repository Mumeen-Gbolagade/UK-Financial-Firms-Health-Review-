"""
Make the workbook screenshots shown in the README (images/excel_*.png).

Needs LibreOffice (soffice) and poppler (pdftoppm) on the PATH.
Run after src/analysis.py:
    python src/make_screenshots.py
"""

import shutil
import subprocess
from pathlib import Path

from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parents[1]
SHOTS = ROOT / "outputs" / "_screens"
IMAGES = ROOT / "images"
SHEETS = {
    "Dashboard": "excel_dashboard.png",
    "KPI_Scorecard": "excel_kpi_scorecard.png",
    "Ratios": "excel_ratios.png",
    "Project_Story": "excel_project_story.png",
    "Data_Log": "excel_data_log.png",
}


def trim(img: Image.Image, pad: int = 24) -> Image.Image:
    """Cut the blank page margins, keeping a small white border."""
    bg = Image.new(img.mode, img.size, (255, 255, 255))
    box = ImageChops.difference(img, bg).getbbox()
    if not box:
        return img
    l, t, r, b = box
    return img.crop((max(l - pad, 0), max(t - pad, 0), min(r + pad, img.width), min(b + pad, img.height)))


def main() -> None:
    subprocess.run(["python", str(ROOT / "src" / "build_workbook.py"), "--screens"], check=True)
    IMAGES.mkdir(exist_ok=True)
    for sheet, name in SHEETS.items():
        xlsx = SHOTS / f"{sheet}.xlsx"
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(SHOTS), str(xlsx)],
                       check=True, capture_output=True, timeout=180)
        stem = SHOTS / sheet
        subprocess.run(["pdftoppm", "-r", "170", "-png", "-f", "1", "-l", "1", "-singlefile",
                        str(stem.with_suffix(".pdf")), str(stem)], check=True)
        img = trim(Image.open(stem.with_suffix(".png")).convert("RGB"))
        img.save(IMAGES / name, optimize=True)
        print(f"Saved images/{name} ({img.width}x{img.height})")
    shutil.rmtree(SHOTS)


if __name__ == "__main__":
    main()
