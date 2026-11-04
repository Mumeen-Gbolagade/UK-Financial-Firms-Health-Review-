"""
Draw the dashboard image used in the README (images/dashboard.png).
It shows the same numbers as the Dashboard sheet in the workbook.

Run after src/analysis.py:
    python src/make_dashboard.py
"""

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import FancyBboxPatch

ROOT = Path(__file__).resolve().parents[1]
r = pd.read_csv(ROOT / "outputs" / "ratios.csv")
co = pd.read_csv(ROOT / "data" / "companies.csv")
brand = {row.short_name: "#" + row.brand_colour for row in co.itertuples()}
r = r[r["in_analysis"] == 1].copy()
r["pos"] = r.groupby("company")["fy_end"].rank().astype(int)  # 1,2,3 = oldest..latest
latest = r[r["pos"] == 3].sort_values("pbt_margin", ascending=False)
order = latest["short_name"].tolist()

INK, MUTED, GRID = "#1F2A44", "#5B6577", "#E3E7EE"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GRID,
                     "axes.labelcolor": MUTED, "xtick.color": MUTED, "ytick.color": MUTED})

fig = plt.figure(figsize=(16, 10), dpi=150, facecolor="white")
fig.patches.append(FancyBboxPatch((0, 0.915), 1, 0.085, transform=fig.transFigure,
                                  boxstyle="square,pad=0", fc=INK, ec="none"))
fig.text(0.02, 0.962, "UK Asset & Wealth Managers: Financial Health Review", color="white",
         fontsize=20, weight="bold", va="center")
fig.text(0.02, 0.932, "Latest audited year per company · £ millions · group accounts listed on Companies House · data accessed 4-5 Nov 2026",
         color="#DCE3EE", fontsize=10.5, va="center")

# ---- scorecard table
ax_t = fig.add_axes([0.02, 0.56, 0.96, 0.34])
ax_t.axis("off")
cols = ["Company", "Latest year", "Revenue £m", "Revenue growth", "Profit margin (PBT)",
        "Current ratio", "Debt-to-equity", "Interest cover"]
xs = [0.0, 0.20, 0.30, 0.41, 0.53, 0.65, 0.76, 0.87]
ax_t.add_patch(plt.Rectangle((0, 0.86), 1, 0.14, color=INK, transform=ax_t.transAxes))
for x, c in zip(xs, cols):
    ax_t.text(x + (0.01 if x == 0 else 0.05), 0.93, c, color="white", weight="bold",
              fontsize=10.5, va="center", ha="left" if x == 0 else "center", transform=ax_t.transAxes)

def fmt_pct(v):
    return "n/a" if pd.isna(v) else f"{v:+.1%}" if v != 0 else "0.0%"

row_h = 0.86 / len(order)
for i, name in enumerate(order):
    d = latest[latest["short_name"] == name].iloc[0]
    y = 0.86 - (i + 1) * row_h
    if i % 2 == 0:
        ax_t.add_patch(plt.Rectangle((0, y), 1, row_h, color="#F5F7FA", transform=ax_t.transAxes))
    ax_t.add_patch(plt.Rectangle((0, y + 0.01), 0.19, row_h - 0.02, color=brand[name], transform=ax_t.transAxes))
    ax_t.text(0.01, y + row_h / 2, name, color="white", weight="bold", fontsize=11.5, va="center", transform=ax_t.transAxes)
    cells = [d["fy_label"], f"{d['revenue_m']:,.1f}", fmt_pct(d["revenue_growth"]),
             f"{d['pbt_margin']:.1%}", f"{d['current_ratio']:.2f}x",
             "n/a" if pd.isna(d["debt_to_equity"]) else f"{d['debt_to_equity']:.2f}x",
             "n/a" if pd.isna(d["interest_cover"]) else f"{d['interest_cover']:,.0f}x"]
    colours = [INK, INK,
               "#B42318" if d["revenue_growth"] < 0 else "#1E7B3A",
               "#1E7B3A" if d["pbt_margin"] >= 0.25 else ("#B42318" if d["pbt_margin"] < 0.05 else INK),
               "#B26A00" if d["current_ratio"] < 1.5 else INK, INK, INK]
    for x, txt, colr in zip(xs[1:], cells, colours):
        ax_t.text(x + 0.05, y + row_h / 2, txt, color=colr, fontsize=11.5, va="center", ha="center",
                  weight="bold" if colr != INK else "normal", transform=ax_t.transAxes)

# ---- charts
labels = ["Two years before", "One year before", "Latest year"]

def trend(ax, col, title, pct=True):
    for name in order:
        s = r[r["short_name"] == name].sort_values("pos")[col].values
        ax.plot(labels, s, marker="o", lw=2.4, ms=6, color=brand[name], label=name)
    ax.set_title(title, loc="left", fontsize=12.5, weight="bold", color=INK, pad=8)
    ax.grid(axis="y", color=GRID)
    ax.axhline(0, color="#9AA3B2", lw=0.8)
    if pct:
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0%}"))
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)

ax1 = fig.add_axes([0.04, 0.10, 0.28, 0.38])
trend(ax1, "pbt_margin", "Profit before tax margin")
ax2 = fig.add_axes([0.37, 0.10, 0.28, 0.38])
trend(ax2, "revenue_growth", "Revenue growth")
ax3 = fig.add_axes([0.72, 0.10, 0.26, 0.38])
vals = latest.set_index("short_name").loc[order, "current_ratio"]
ax3.barh(order[::-1], vals[::-1], color=[brand[n] for n in order[::-1]], height=0.6)
for y, v in enumerate(vals[::-1]):
    ax3.text(v + 0.12, y, f"{v:.2f}x", va="center", fontsize=10, color=INK)
ax3.set_title("Current ratio, latest year", loc="left", fontsize=12.5, weight="bold", color=INK, pad=8)
ax3.axvline(1, color="#9AA3B2", lw=0.8, ls="--")
ax3.set_xlim(0, vals.max() * 1.18)
for side in ("top", "right"):
    ax3.spines[side].set_visible(False)

handles, lbls = ax1.get_legend_handles_labels()
fig.legend(handles, lbls, loc="lower center", ncol=6, frameon=False, bbox_to_anchor=(0.36, 0.005), fontsize=10)
fig.text(0.98, 0.015, "Ratios point to questions to ask, not verdicts. Sources: see Sources sheet.",
         ha="right", fontsize=9, color=MUTED, style="italic")

out = ROOT / "images" / "dashboard.png"
out.parent.mkdir(exist_ok=True)
fig.savefig(out, facecolor="white")
print(f"Saved {out}")
