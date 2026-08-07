"""
make_walk_forward_figure.py
Plots the walk-forward results from walk_forward_evaluation.py: test-year
ROC-AUC and F1-macro across the five expanding-window folds, so a reader
can see at a glance whether Random Forest holds up over time or slides.
Reads data/processed/walk_forward_results.csv.

Usage: python src/make_walk_forward_figure.py
"""

import pandas as pd
import matplotlib.pyplot as plt

DATA_DIR = "data/processed"
FIGURES_DIR = "report/figures"


def main():
    d = pd.read_csv(f"{DATA_DIR}/walk_forward_results.csv")
    d = d.sort_values("test_year")

    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Segoe UI", "Arial", "DejaVu Sans"]
    plt.rcParams.update({"figure.dpi": 100, "axes.spines.top": False, "axes.spines.right": False})

    fig, ax1 = plt.subplots(figsize=(6.5, 4.2))
    ax1.plot(d["test_year"], d["test_roc_auc"], marker="o", color="#B22222", label="ROC-AUC")
    ax1.set_ylabel("ROC-AUC")
    ax1.set_ylim(0.9, 1.0)

    ax2 = ax1.twinx()
    ax2.plot(d["test_year"], d["test_f1_macro"], marker="s", color="#1F4E79", label="F1-macro")
    ax2.set_ylabel("F1-macro")
    ax2.set_ylim(0.6, 1.0)
    ax2.spines["top"].set_visible(False)

    ax1.set_xlabel("Test year (trained on every year before it)")
    ax1.set_xticks(d["test_year"])
    ax1.set_title("Random Forest, walk-forward: test-year score as the window moves", fontsize=11, fontweight="bold")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="lower left", frameon=False)

    plt.tight_layout()
    plt.savefig(f"{FIGURES_DIR}/10_walk_forward.png", dpi=220, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved {FIGURES_DIR}/10_walk_forward.png")


if __name__ == "__main__":
    main()
