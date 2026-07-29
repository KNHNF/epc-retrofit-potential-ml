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


def _read_csv(path):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def render_comparison_table():
    rows = _read_csv(f"{DATA_DIR}/model_comparison.csv")
    if not rows:
        return
    headers = ["Model", "CV F1", "CV AUC", "Test F1", "Test AUC", "Test PR-AUC"]
    src_keys = ["Model", "CV F1-macro", "CV ROC-AUC", "Test F1-macro", "Test ROC-AUC", "Test PR-AUC"]

    best_model = max(rows, key=lambda r: float(r["Test ROC-AUC"]))["Model"]
    cell_text = [[r[k] for k in src_keys] for r in rows]

    fig, ax = plt.subplots(figsize=(9.5, 0.55 * (len(rows) + 1)))
    ax.axis("off")
    tbl = ax.table(cellText=cell_text, colLabels=headers, cellLoc="center", loc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)
    tbl.scale(1, 2.0)

    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor("#888888")
        if row == 0:
            cell.set_facecolor(HEADER_BG)
            cell.set_text_props(color="white", fontweight="bold")
        elif rows[row - 1]["Model"] == best_model:
            cell.set_facecolor(WINNER_BG)
            cell.set_text_props(fontweight="bold")

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

    fig, ax = plt.subplots(figsize=(9.5, 0.55 * (len(rows) + 1)))
    ax.axis("off")
    tbl = ax.table(cellText=cell_text, colLabels=headers, cellLoc="center", loc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)
    tbl.scale(1, 2.0)

    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor("#888888")
        if row == 0:
            cell.set_facecolor(HEADER_BG)
            cell.set_text_props(color="white", fontweight="bold")
        elif mismatches[row - 1]:
            cell.set_facecolor(ERROR_BG)

    plt.tight_layout()
    plt.savefig(f"{FIGURES_DIR}/table2_bristol.png", dpi=220, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved {FIGURES_DIR}/table2_bristol.png")


if __name__ == "__main__":
    render_comparison_table()
    render_bristol_table()
