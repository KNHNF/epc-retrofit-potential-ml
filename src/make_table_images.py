"""
make_table_images.py
Renders Table 1 (model comparison) and Table 2 (Bristol case study) as clean
PNG images, matching the Word tables' own styling (dark header row, green
highlight for the winning model, red highlight for Bristol prediction
errors). Used only in the two-column report variant: three separate real
attempts at floating a genuine Word table across both columns (section
break, DrawingML text box, native w:tblpPr) all failed in real Microsoft
Word, confirmed by direct inspection each time, not assumed. Word's table
model cannot cross the column gutter in a continuous multi-column section
the way a picture's `wp:anchor` can. Rendering the same real data as an
image sidesteps that limitation entirely by reusing the picture-floating
mechanism that already works.

The single-column report keeps real, editable Word tables (add_comparison_
table / add_bristol_table in generate_report.py), no floating is needed
there since a single column already spans the full page width.

Usage: python src/make_table_images.py
"""

import csv
import matplotlib.pyplot as plt

DATA_DIR = "data/processed"
FIGURES_DIR = "report/figures"

HEADER_BG = "#2F2F2F"
WINNER_BG = "#E8F0E3"
ERROR_BG = "#F5DCDC"
RULE_COLOR = "#333333"


def _read_csv(path):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def _apply_booktabs_style(ax, tbl, n_rows):
    """Journal-style rules only (thick top, thin under the header, thick
    bottom), no vertical gridlines and no lines between data rows. Matches
    the reference table style shared for comparison, and the standard
    "booktabs" convention most published papers actually use; the full grid
    this project used before read as a spreadsheet export, not a typeset
    table.

    Cell.visible_edges looks like the obvious tool for this (set 'T' on the
    header row, 'B' on the last row, nothing elsewhere), and it does draw
    the right lines, but it silently makes the WHOLE cell invisible,
    facecolor and text included, the moment a cell's edges aren't the full
    default set. Confirmed by isolating it: a cell with all four edges
    renders correctly; the same cell with only one edge set renders
    nothing at all, not even its dark header fill. Rather than fight that,
    every cell keeps its default (closed) edges but coloured to match its
    own fill so they're invisible, and the three real rules are drawn as
    plain Line2D segments across the whole table instead, independent of
    the cell grid entirely. Requires the table's bbox to be the axes'
    full [0, 0, 1, 1] so each row's y-position is exactly 1/n_rows,
    known in advance rather than queried from cell geometry that isn't
    finalised until a draw event happens."""
    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor(cell.get_facecolor())
        cell.set_linewidth(0.5)
    row_h = 1.0 / n_rows
    for y, lw in [(1.0, 1.6), (1.0 - row_h, 0.9), (0.0, 1.6)]:
        ax.plot([0, 1], [y, y], color=RULE_COLOR, linewidth=lw,
                transform=ax.transAxes, clip_on=False, zorder=10)


def render_comparison_table():
    rows = _read_csv(f"{DATA_DIR}/model_comparison.csv")
    if not rows:
        return
    # Accuracy and recall first (what a marker expects to see), then the
    # three metrics this imbalanced target actually needs. CV columns stay
    # in the CSV but are dropped here too, matching the single-column
    # table, the specific CV-vs-test reversal numbers already live in
    # Section 6's prose with citations.
    headers = ["Model", "Accuracy", "Recall", "Precision", "F1", "ROC-AUC", "PR-AUC"]
    src_keys = ["Model", "Test Accuracy", "Test Recall", "Test Precision",
                "Test F1-macro", "Test ROC-AUC", "Test PR-AUC"]
    # Bolded on the winning row below: accuracy is deliberately excluded,
    # the report's own argument is that accuracy is not the metric doing
    # the real work here, the table shouldn't visually reward it either.
    significant_keys = {"Test Recall", "Test F1-macro", "Test ROC-AUC", "Test PR-AUC"}

    best_model = max(rows, key=lambda r: float(r["Test ROC-AUC"]))["Model"]
    cell_text = [[r[k] for k in src_keys] for r in rows]

    # Bold the actual best value per column, not every significant column on
    # whichever row wins overall: the overall winner is not necessarily best
    # on every individual metric (Random Forest wins ROC-AUC here but has the
    # lowest recall of the four models compared).
    best_per_col = {k: max(float(r[k]) for r in rows) for k in significant_keys}

    n_rows = len(rows) + 1
    fig, ax = plt.subplots(figsize=(9.5, 0.55 * n_rows))
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    tbl = ax.table(cellText=cell_text, colLabels=headers, cellLoc="center",
                   loc="center", bbox=[0, 0, 1, 1])
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)

    for (row, col), cell in tbl.get_celld().items():
        if row == 0:
            cell.set_facecolor(HEADER_BG)
            cell.set_text_props(color="white", fontweight="bold")
        else:
            r = rows[row - 1]
            if r["Model"] == best_model:
                cell.set_facecolor(WINNER_BG)
            key = src_keys[col]
            if key in significant_keys and float(r[key]) == best_per_col[key]:
                cell.set_text_props(fontweight="bold")
    _apply_booktabs_style(ax, tbl, n_rows=n_rows)

    plt.tight_layout()
    plt.savefig(f"{FIGURES_DIR}/table1_model_comparison.png", dpi=220, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved {FIGURES_DIR}/table1_model_comparison.png")


def render_bristol_table():
    rows = _read_csv(f"{DATA_DIR}/bristol_case_study.csv")
    if not rows:
        return
    headers = ["Postcode", "Type", "Rating", "Gap (pts)", "P(headroom)", "Predicted", "Actual"]
    cell_text = []
    mismatches = []
    for r in rows:
        pred = "Yes" if r["PRED_LABEL"] == "1" else "No"
        actual = "Yes" if r["RETROFIT_POTENTIAL"] == "1" else "No"
        mismatches.append(pred != actual)
        cell_text.append([
            r["OUTWARD_POSTCODE"], r["PROPERTY_TYPE"], r["CURRENT_ENERGY_RATING"],
            f"{float(r['EFFICIENCY_GAP']):.0f}", f"{float(r['PRED_PROBA']):.2f}",
            pred, actual,
        ])

    n_rows = len(rows) + 1
    fig, ax = plt.subplots(figsize=(9.5, 0.55 * n_rows))
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    tbl = ax.table(cellText=cell_text, colLabels=headers, cellLoc="center",
                   loc="center", bbox=[0, 0, 1, 1])
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)

    for (row, col), cell in tbl.get_celld().items():
        if row == 0:
            cell.set_facecolor(HEADER_BG)
            cell.set_text_props(color="white", fontweight="bold")
        elif mismatches[row - 1]:
            cell.set_facecolor(ERROR_BG)
    _apply_booktabs_style(ax, tbl, n_rows=n_rows)

    plt.tight_layout()
    plt.savefig(f"{FIGURES_DIR}/table2_bristol.png", dpi=220, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved {FIGURES_DIR}/table2_bristol.png")


def render_cities_table():
    rows = _read_csv(f"{DATA_DIR}/city_case_studies_summary.csv")
    if not rows:
        return
    headers = ["City", "n", "Positive rate", "Recall", "Precision", "F1"]
    cell_text = []
    for r in rows:
        cell_text.append([
            r["city"], f"{int(float(r['n_test_properties'])):,}",
            f"{float(r['positive_rate']):.2f}", f"{float(r['recall']):.2f}",
            f"{float(r['precision']):.2f}", f"{float(r['f1']):.2f}",
        ])

    n_rows = len(rows) + 1
    fig, ax = plt.subplots(figsize=(9.5, 0.55 * n_rows))
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    tbl = ax.table(cellText=cell_text, colLabels=headers, cellLoc="center",
                   loc="center", bbox=[0, 0, 1, 1])
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)

    for (row, col), cell in tbl.get_celld().items():
        if row == 0:
            cell.set_facecolor(HEADER_BG)
            cell.set_text_props(color="white", fontweight="bold")
    _apply_booktabs_style(ax, tbl, n_rows=n_rows)

    plt.tight_layout()
    plt.savefig(f"{FIGURES_DIR}/table3_cities.png", dpi=220, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved {FIGURES_DIR}/table3_cities.png")


if __name__ == "__main__":
    render_comparison_table()
    render_bristol_table()
    render_cities_table()
