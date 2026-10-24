#!/usr/bin/env python3
"""Make the "Conducting the review" figures for the SLR.

Usage:
    pip install openpyxl matplotlib
    python scripts/plots.py path/to/Final_new_all_394.xlsx

Writes PDF (for LaTeX) and PNG (for GitHub) files to figures/.
"""
import sys
from collections import Counter
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build import ROOT, load  # noqa: E402

OUT = ROOT / "figures"
DEEPFAKE, FAKENEWS = "#2a78d6", "#eb6834"   # categorical slots 1 and 2
SINGLE = "#2a78d6"
INK, MUTED, GRID = "#1f1f1e", "#6b6a64", "#e4e3de"
CATS = [("Deepfake", DEEPFAKE), ("Fake News", FAKENEWS)]

plt.rcParams.update({
    "font.family": "serif", "font.size": 10, "axes.edgecolor": MUTED,
    "axes.labelcolor": INK, "xtick.color": MUTED, "ytick.color": INK,
    "axes.spines.top": False, "axes.spines.right": False,
    "pdf.fonttype": 42, "savefig.bbox": "tight", "savefig.dpi": 300,
})


def save(fig, name):
    OUT.mkdir(exist_ok=True)
    for ext in ("pdf", "png"):
        fig.savefig(OUT / f"{name}.{ext}")
    plt.close(fig)


def stacked_vertical(ax, labels, series):
    """series: list of (name, color, values). Stacked bars with a 2px white gap."""
    bottom = [0] * len(labels)
    for name, color, vals in series:
        ax.bar(labels, vals, bottom=bottom, color=color, width=0.7, label=name,
               edgecolor="white", linewidth=1.5)
        bottom = [b + v for b, v in zip(bottom, vals)]
    for i, t in enumerate(bottom):
        ax.text(i, t + max(bottom) * 0.015, str(t), ha="center", va="bottom", fontsize=9, color=INK)
    ax.yaxis.grid(True, color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)


def stacked_horizontal(ax, labels, series):
    left = [0] * len(labels)
    y = range(len(labels))
    for name, color, vals in series:
        ax.barh(list(y), vals, left=left, color=color, height=0.65, label=name,
                edgecolor="white", linewidth=1.5)
        left = [a + v for a, v in zip(left, vals)]
    for i, t in enumerate(left):
        ax.text(t + max(left) * 0.01, i, str(t), va="center", fontsize=9, color=INK)
    ax.set_yticks(list(y), labels)
    ax.invert_yaxis()
    ax.xaxis.grid(True, color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_xlim(0, max(left) * 1.12)


def by_category(papers, key, order):
    return [(n, c, [sum(1 for p in papers if p["category"] == n and key(p) == o) for o in order])
            for n, c in CATS]


def fig_years(papers):
    years = sorted({p["year"] for p in papers})
    fig, ax = plt.subplots(figsize=(6.5, 3.2))
    stacked_vertical(ax, years, by_category(papers, lambda p: p["year"], years))
    ax.set_xlabel("Publication year")
    ax.set_ylabel("Number of papers")
    ax.legend(frameon=False, loc="upper left")
    save(fig, "fig_papers_per_year")


def fig_venue(papers):
    order = [v for v, _ in Counter(p["venue_type"] for p in papers).most_common()]
    fig, ax = plt.subplots(figsize=(6.5, 1.9))
    stacked_horizontal(ax, order, by_category(papers, lambda p: p["venue_type"], order))
    ax.set_xlabel("Number of papers")
    ax.legend(frameon=False, loc="lower right")
    save(fig, "fig_venue_type")


def fig_publisher(papers):
    order = [v for v, _ in Counter(p["publisher"] for p in papers).most_common()]
    fig, ax = plt.subplots(figsize=(6.5, 2.4))
    stacked_horizontal(ax, order, by_category(papers, lambda p: p["publisher"], order))
    ax.set_xlabel("Number of papers")
    ax.legend(frameon=False, loc="lower right")
    save(fig, "fig_publisher")


def country_of(p):
    c = " ".join(p["country"].split())
    return "Multiple" if " and " in c else (c or "Unspecified")


def fig_countries(papers, top=10):
    counts = Counter(country_of(p) for p in papers)
    order = [c for c, _ in counts.most_common(top)]
    for p in papers:
        p["_country"] = country_of(p) if country_of(p) in order else "Other"
    order.append("Other")
    fig, ax = plt.subplots(figsize=(6.5, 3.6))
    stacked_horizontal(ax, order, by_category(papers, lambda p: p["_country"], order))
    ax.set_xlabel("Number of papers (by first-author country)")
    ax.legend(frameon=False, loc="lower right")
    save(fig, "fig_countries")


def fig_modalities(papers):
    key = lambda p: " + ".join(p["modalities"]) or "Unspecified"
    order = [m for m, _ in Counter(key(p) for p in papers).most_common()]
    fig, ax = plt.subplots(figsize=(6.5, 3.0))
    stacked_horizontal(ax, order, by_category(papers, key, order))
    ax.set_xlabel("Number of papers")
    ax.legend(frameon=False, loc="lower right")
    save(fig, "fig_modalities")


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    papers = load(sys.argv[1])
    fig_years(papers)
    fig_venue(papers)
    fig_publisher(papers)
    fig_countries(papers)
    fig_modalities(papers)
    print(f"Saved figures for {len(papers)} papers to {OUT}")


if __name__ == "__main__":
    main()
