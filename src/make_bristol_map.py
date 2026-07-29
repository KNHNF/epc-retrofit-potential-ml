"""
make_bristol_map.py
Builds an area map of Bristol postcode districts (BS1, BS3, BS5, etc.),
coloured by the trained model's predicted high-retrofit-potential rate for
that district's held-out test-set properties, drawn over a real OpenStreetMap
basemap (via contextily) so the city is actually visible, not just floating
circles on a blank background. Reads
data/processed/bristol_district_summary.csv (produced by bristol_case_study.py).

Marker position is still an approximate district centroid (public knowledge
of Bristol's postcode-area geography), not a real boundary polygon, geopandas
and ONS boundary shapefiles are not used here. Marker size encodes property
count (sqrt-scaled, so area is proportional to count, not radius), colour
encodes predicted-positive rate, both from real model output on real
held-out data.

Usage: python src/make_bristol_map.py
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
import contextily as cx
from pyproj import Transformer

DATA_DIR = "data/processed"
FIGURES_DIR = "report/figures"


def main():
    d = pd.read_csv(f"{DATA_DIR}/bristol_district_summary.csv")
    d = d.sort_values("n", ascending=False)

    to_mercator = Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)
    x, y = to_mercator.transform(d["lon"].values, d["lat"].values)
    d = d.assign(x=x, y=y)

    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Segoe UI", "Arial", "DejaVu Sans"]

    fig, ax = plt.subplots(figsize=(7.5, 7.2))

    # Area, not radius, proportional to property count: sqrt-scale so the
    # largest district doesn't visually dominate the plot.
    radius = np.sqrt(d["n"])
    sizes = 260 + 1500 * (radius - radius.min()) / (radius.max() - radius.min())

    # Soft drop shadow first, then the marker on top, gives the dots real
    # depth instead of the flat "plotted point" look.
    ax.scatter(
        d["x"] + 40, d["y"] - 40, s=sizes, c="black", alpha=0.15,
        zorder=2, linewidths=0,
    )
    sc = ax.scatter(
        d["x"], d["y"], s=sizes, c=d["predicted_positive_rate"],
        cmap="Reds", vmin=0, edgecolors="white", linewidths=1.6, zorder=3,
        alpha=0.92,
    )
    for _, row in d.iterrows():
        ax.annotate(
            row["district"], (row["x"], row["y"]),
            ha="center", va="center", fontsize=8.5, fontweight="bold",
            color="white", zorder=4,
            path_effects=[pe.withStroke(linewidth=2.2, foreground="#333333")],
        )

    pad = 0.14 * (d["x"].max() - d["x"].min())
    ax.set_xlim(d["x"].min() - pad, d["x"].max() + pad)
    ax.set_ylim(d["y"].min() - pad, d["y"].max() + pad)
    cx.add_basemap(ax, source=cx.providers.CartoDB.Positron, attribution_size=6)

    cbar = fig.colorbar(sc, ax=ax, shrink=0.75, pad=0.02)
    cbar.set_label("Predicted high-retrofit-potential rate", fontsize=9.5)
    cbar.outline.set_visible(False)
    cbar.ax.tick_params(labelsize=8.5)

    ax.set_title(
        "Bristol postcode districts: predicted retrofit-potential rate",
        fontsize=12, fontweight="bold", pad=32,
    )
    ax.text(
        0.5, 1.045, "Marker area = number of test-set properties (n>=20 per district)",
        transform=ax.transAxes, ha="center", fontsize=9, color="#555555",
    )
    ax.set_xlabel(
        "District centroid, approximate. Basemap: (C) OpenStreetMap contributors, (C) CARTO",
        fontsize=7.5, color="#666666",
    )
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)

    plt.tight_layout()
    plt.savefig(f"{FIGURES_DIR}/07_bristol_district_map.png", dpi=220, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved {FIGURES_DIR}/07_bristol_district_map.png")


if __name__ == "__main__":
    main()
