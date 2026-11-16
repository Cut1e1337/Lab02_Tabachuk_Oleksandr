"""Presentation-ready charts (large type, one message per chart) built from saved EDA results.

Reads data/processed/auto_mpg_clean.csv and outputs/tables/*.csv, writes outputs/figures/slides/*.png.
Run after notebooks/EDA.ipynb:  python reports/make_slide_figures.py
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "figures" / "slides"
OUT.mkdir(parents=True, exist_ok=True)

df = pd.read_csv(ROOT / "data" / "processed" / "auto_mpg_clean.csv")
corr = pd.read_csv(ROOT / "outputs" / "tables" / "correlations.csv", index_col=0)
L = "l_per_100km"

# palette shared with the slides (build_presentation.py)
INK, MUTED, GRID = "#1F2328", "#6B6F76", "#E6E2D6"
CORAL, TEAL, GREY, OLIVE = "#D1495B", "#2E7D8C", "#B9B6AC", "#8A8B4A"
CARD = "#FFFDF8"
ORIGIN = {"США": TEAL, "Європа": GREY, "Японія": CORAL}

plt.rcParams.update({
    "font.family": "Segoe UI", "font.size": 14, "axes.titlesize": 15,
    "axes.labelsize": 14, "xtick.labelsize": 13, "ytick.labelsize": 13,
    "axes.edgecolor": MUTED, "axes.labelcolor": INK, "xtick.color": MUTED, "ytick.color": MUTED,
    "text.color": INK, "axes.spines.top": False, "axes.spines.right": False,
    "figure.facecolor": CARD, "axes.facecolor": CARD, "savefig.facecolor": CARD,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 1, "axes.axisbelow": True,
})
SIZE = (8.0, 4.55)  # inches; equals the chart slot on the slide (1:1 scale)


def save(fig, name):
    fig.savefig(OUT / f"{name}.png", dpi=220)
    plt.close(fig)


def fmt(x, d=1):
    return f"{x:.{d}f}".replace(".", ",")


# 1. Target distribution (л/100 км)
fig, ax = plt.subplots(figsize=SIZE, layout="constrained")
ax.hist(df[L], bins=np.arange(5, 27, 1), color=TEAL, edgecolor=CARD, linewidth=1.2)
med = df[L].median()
ax.axvline(med, color=CORAL, lw=2.5)
ax.text(med + 0.3, 56, f"медіана {fmt(med)} л", color=CORAL, fontsize=15, fontweight="bold", va="top")
lo, hi = df.loc[df[L].idxmin()], df.loc[df[L].idxmax()]
ax.text(16.6, 50, f"мінімум  {fmt(lo[L])} л — Mazda GLC\nмаксимум  {fmt(hi[L], 0)} л — пікап з V8\nрізниця — у 5 разів",
        fontsize=13, color=INK, va="top", linespacing=1.6)
ax.set(xlabel="Витрата палива, л/100 км", ylabel="Кількість моделей", xlim=(4.5, 27), ylim=(0, 57))
ax.grid(axis="x", visible=False)
save(fig, "s1_target")

# 2. Weight vs consumption
fig, ax = plt.subplots(figsize=SIZE, layout="constrained")
ax.scatter(df["weight_kg"], df[L], s=22, color=TEAL, alpha=0.45, edgecolor="none")
k, b = np.polyfit(df["weight_kg"], df[L], 1)
xs = np.array([df["weight_kg"].min(), df["weight_kg"].max()])
ax.plot(xs, k * xs + b, color=CORAL, lw=3)
r = df["weight_kg"].corr(df[L])
ax.text(760, 23.5, f"r = {fmt(r, 2)}", fontsize=17, fontweight="bold", color=CORAL)
ax.text(760, 21.8, "кожна точка — модель авто,\nлінія — лінійний тренд", fontsize=13, color=MUTED, va="top")
ax.set(xlabel="Маса автомобіля, кг", ylabel="Витрата, л/100 км", ylim=(4, 27))
save(fig, "s2_weight")

# 3. Origin — naive comparison
g = df.groupby("origin").agg(l=(L, "mean"), w=("weight_kg", "mean")).loc[["США", "Європа", "Японія"]]
fig, ax = plt.subplots(figsize=SIZE, layout="constrained")
bars = ax.bar(g.index, g["l"], color=[ORIGIN[o] for o in g.index], width=0.6)
for bar, (o, row) in zip(bars, g.iterrows()):
    ax.text(bar.get_x() + bar.get_width() / 2, row["l"] + 0.25, f"{fmt(row['l'])} л", ha="center",
            fontsize=18, fontweight="bold", color=INK)
    ax.text(bar.get_x() + bar.get_width() / 2, 0.5, f"маса\n{row['w']:.0f} кг", ha="center", va="bottom",
            fontsize=13, color="white", fontweight="bold")
ax.set(ylabel="Середня витрата, л/100 км", ylim=(0, 15))
ax.grid(axis="x", visible=False)
ax.tick_params(axis="x", labelsize=15, colors=INK)
save(fig, "s3_origin")

# 4. Origin within weight groups
grp = df.groupby(["weight_group", "origin"])[L].agg(["mean", "count"]).reset_index()
groups = ["легкі", "середні"]
edges = df.groupby("weight_group")["weight_kg"].agg(["min", "max"])
fig, ax = plt.subplots(figsize=SIZE, layout="constrained")
x = np.arange(len(groups)); wbar = 0.26
for i, o in enumerate(["США", "Європа", "Японія"]):
    vals = [grp[(grp.weight_group == gname) & (grp.origin == o)]["mean"].iloc[0] for gname in groups]
    bs = ax.bar(x + (i - 1) * wbar, vals, wbar * 0.92, color=ORIGIN[o], label=o)
    for bb, v in zip(bs, vals):
        ax.text(bb.get_x() + bb.get_width() / 2, v + 0.2, fmt(v), ha="center", fontsize=14, fontweight="bold")
ax.set_xticks(x, [f"Легкі авто\n{edges.loc['легкі', 'min']:.0f}–{edges.loc['легкі', 'max']:.0f} кг",
                  f"Середні авто\n{edges.loc['середні', 'min']:.0f}–{edges.loc['середні', 'max']:.0f} кг"])
ax.tick_params(axis="x", labelsize=14, colors=INK)
ax.set(ylabel="Середня витрата, л/100 км", ylim=(0, 13))
ax.grid(axis="x", visible=False)
ax.legend(frameon=False, ncol=3, loc="upper left", fontsize=14)
save(fig, "s4_origin_within_weight")

# 5. Trend by year within weight groups
fig, ax = plt.subplots(figsize=SIZE, layout="constrained")
yr = df.groupby("year")[L].mean()
ax.plot(yr.index, yr.values, color=INK, lw=3.2, ls="--", label="усі авто", zorder=3)
for gname, col in [("важкі", TEAL), ("середні", OLIVE), ("легкі", CORAL)]:
    s = df[df.weight_group == gname].groupby("year")[L].mean()
    ax.plot(s.index, s.values, color=col, lw=2.4, marker="o", ms=5)
    ax.text(s.index[-1] + 0.25, s.values[-1], gname, color=col, fontsize=14, fontweight="bold", va="center")
for y_, lbl in [(1973, "нафтова\nкриза"), (1979, "2-га\nкриза")]:
    ax.axvline(y_, color=MUTED, ls=":", lw=1.5)
    ax.text(y_ + 0.15, 5.2, lbl, fontsize=12, color=MUTED, va="bottom")
ax.annotate("усі авто", xy=(1975, yr.loc[1975]), xytext=(1973.7, 13.3), fontsize=14, fontweight="bold",
            color=INK, arrowprops=dict(arrowstyle="-", color=INK, lw=1.2))
ax.set(xlabel="Модельний рік", ylabel="Середня витрата, л/100 км", xlim=(1969.6, 1983.4), ylim=(5, 19))
ax.set_xticks(range(1970, 1983, 2))
save(fig, "s5_trend")

# 6. Multicollinearity heatmap
names = {"weight_kg": "Маса", "displacement_l": "Об'єм", "horsepower": "Потужність",
         "cylinders": "Циліндри", "acceleration": "Розгін", "year": "Рік"}
c = corr.loc[list(names), list(names)].abs()
fig, ax = plt.subplots(figsize=(5.6, 4.55), layout="constrained")
from matplotlib.colors import LinearSegmentedColormap
cmap = LinearSegmentedColormap.from_list("w", [CARD, "#F3D9C4", CORAL])
ax.imshow(c.values, cmap=cmap, vmin=0, vmax=1)
for i in range(len(c)):
    for j in range(len(c)):
        v = c.values[i, j]
        ax.text(j, i, fmt(v, 2), ha="center", va="center", fontsize=13,
                color="white" if v > 0.8 else INK, fontweight="bold" if (i < 4 and j < 4 and i != j) else None)
ax.set_xticks(range(len(c)), names.values(), rotation=30, ha="right"); ax.set_yticks(range(len(c)), names.values())
ax.tick_params(colors=INK, length=0); ax.grid(False)
for sp in ax.spines.values(): sp.set_visible(False)
ax.add_patch(plt.Rectangle((-0.5, -0.5), 4, 4, fill=False, ec=INK, lw=2.5))
ax.set_title("|ρ Спірмена| між факторами", fontsize=14, color=MUTED)
save(fig, "s6_multicollinearity")

print("Saved:", sorted(p.name for p in OUT.glob("*.png")))
