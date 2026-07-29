"""
regen_correlation_figure.py
Regenerates report/figures/04_correlation_heatmap.png with readable labels
and annotation size. The original (from notebooks/01_EDA.ipynb cell 13) used
raw column names (e.g. CO2_EMISS_CURR_PER_FLOOR_AREA) rotated 90 degrees and
8pt annotations sized for an 11x9in figure shown at full page width; shrunk
to report width (5.5-6.2in) both became too small to read.

Usage: python src/regen_correlation_figure.py
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

DATA_DIR = "data/processed"
FIGURES_DIR = "report/figures"

LABELS = {
    'CURRENT_ENERGY_EFFICIENCY': 'Current efficiency',
    'TOTAL_FLOOR_AREA': 'Floor area',
    'NUMBER_HABITABLE_ROOMS': 'Habitable rooms',
    'CO2_EMISS_CURR_PER_FLOOR_AREA': 'CO2 / floor area',
    'CO2_EMISSIONS_CURRENT': 'CO2 emissions',
    'HEATING_COST_CURRENT': 'Heating cost',
    'ENERGY_CONSUMPTION_CURRENT': 'Energy consumption',
    'NUMBER_HEATED_ROOMS': 'Heated rooms',
    'MULTI_GLAZE_PROPORTION': 'Multi-glaze %',
    'LOW_ENERGY_LIGHTING': 'Low-energy lighting %',
    'RETROFIT_POTENTIAL': 'Retrofit potential',
}


def main():
    df = pd.read_parquet(f"{DATA_DIR}/epc_train_sample_200k.parquet")
    corr_cols = [c for c in LABELS if c in df.columns]
    corr = df[corr_cols].corr()
    corr = corr.rename(index=LABELS, columns=LABELS)
    mask = np.triu(np.ones_like(corr, dtype=bool), k=1)

    fig, ax = plt.subplots(figsize=(9, 7.5))
    sns.heatmap(
        corr, mask=mask, annot=True, fmt='.2f', cmap='RdBu_r',
        vmin=-1, vmax=1, center=0, linewidths=0.5, ax=ax,
        annot_kws={'size': 11}, cbar_kws={'shrink': 0.8},
    )
    ax.set_title('Pearson correlation matrix (numerical features + target)', fontsize=13)
    ax.tick_params(axis='both', labelsize=11)
    plt.setp(ax.get_xticklabels(), rotation=40, ha='right')
    plt.setp(ax.get_yticklabels(), rotation=0)
    plt.tight_layout()
    plt.savefig(f"{FIGURES_DIR}/04_correlation_heatmap.png", bbox_inches='tight', dpi=200)
    plt.close(fig)
    print(f"Saved {FIGURES_DIR}/04_correlation_heatmap.png")


if __name__ == "__main__":
    main()
