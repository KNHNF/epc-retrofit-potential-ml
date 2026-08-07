"""
make_bristol_map.py
Builds a choropleth of Bristol postcode districts (BS1, BS3, BS5, etc.),
coloured by the trained model's predicted high-retrofit-potential rate for
that district's held-out test-set properties. District shapes are real
boundary polygons, not point markers on a basemap tile: from
data/external/bs_postcode_districts.geojson, itself the postcode-district
polygons for postcode area BS from Wikipedia's "List of postcode districts
in the United Kingdom" (https://en.wikipedia.org/wiki/List_of_postcode_districts_in_the_United_Kingdom),
retrieved via the uk-postcode-polygons GeoJSON export
(https://github.com/missinglink/uk-postcode-polygons). Reads
data/processed/bristol_district_summary.csv (produced by bristol_case_study.py)
for the rate and property count per district.

No geopandas/shapely dependency: polygons are drawn directly from the
GeoJSON coordinate arrays with matplotlib patches, since no geometric
operations (only filling and centroid-for-label placement) are needed.

Usage: python src/make_bristol_map.py
"""

import json

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize

DATA_DIR = "data/processed"
FIGURES_DIR = "report/figures"
GEOJSON_PATH = "data/external/bs_postcode_districts.geojson"


def polygon_rings(geometry):
    """Yield each polygon's exterior ring (list of [lon, lat] pairs) for a
    GeoJSON Polygon or MultiPolygon geometry, ignoring holes."""
    if geometry["type"] == "Polygon":
        yield geometry["coordinates"][0]
    elif geometry["type"] == "MultiPolygon":
        for poly in geometry["coordinates"]:
            yield poly[0]
    elif geometry["type"] == "GeometryCollection":
        for sub in geometry["geometries"]:
            yield from polygon_rings(sub)


def main():
    d = pd.read_csv(f"{DATA_DIR}/bristol_district_summary.csv")
    d = d.set_index("district")

    with open(GEOJSON_PATH) as f:
        geo = json.load(f)

    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Segoe UI", "Arial", "DejaVu Sans"]

    fig, ax = plt.subplots(figsize=(7.5, 7.2))

    norm = Normalize(vmin=0, vmax=d["predicted_positive_rate"].max())
    cmap = plt.get_cmap("Reds")

    patches = []
    colors = []
    all_x, all_y = [], []
    labelled = set()

    for feature in geo["features"]:
        name = feature["properties"]["name"]
        if name not in d.index:
            continue
        for ring in polygon_rings(feature["geometry"]):
            ring = np.array(ring)
            patches.append(Polygon(ring, closed=True))
            colors.append(cmap(norm(d.loc[name, "predicted_positive_rate"])))
            all_x.extend(ring[:, 0])
            all_y.extend(ring[:, 1])
            if name not in labelled:
                cx, cy = ring[:, 0].mean(), ring[:, 1].mean()
                n = int(d.loc[name, "n"])
                ax.annotate(
                    f"{name}\nn={n}", (cx, cy),
                    ha="center", va="center", fontsize=7.5, fontweight="bold",
                    color="#222222", zorder=4,
                    path_effects=[pe.withStroke(linewidth=2.4, foreground="white")],
                )
                labelled.add(name)

    missing = set(d.index) - labelled
    if missing:
        print(f"Warning: no boundary found for districts {sorted(missing)}, skipped.")

    ax.add_collection(PatchCollection(
        patches, facecolor=colors, edgecolor="white", linewidths=1.0, zorder=2,
    ))

    pad_x = 0.04 * (max(all_x) - min(all_x))
    pad_y = 0.04 * (max(all_y) - min(all_y))
    ax.set_xlim(min(all_x) - pad_x, max(all_x) + pad_x)
    ax.set_ylim(min(all_y) - pad_y, max(all_y) + pad_y)
    ax.set_aspect(1 / np.cos(np.radians(np.mean(all_y))))

    sm = ScalarMappable(norm=norm, cmap=cmap)
    cbar = fig.colorbar(sm, ax=ax, shrink=0.75, pad=0.02)
    cbar.set_label("Predicted high-retrofit-potential rate", fontsize=9.5)
    cbar.outline.set_visible(False)
    cbar.ax.tick_params(labelsize=8.5)

    ax.set_title(
        "Bristol postcode districts: predicted retrofit-potential rate",
        fontsize=12, fontweight="bold", pad=32,
    )
    ax.text(
        0.5, 1.045, "n = number of test-set properties in that district (n>=20)",
        transform=ax.transAxes, ha="center", fontsize=9, color="#555555",
    )
    ax.set_xlabel(
        "District boundaries: Wikipedia, List of postcode districts in the United Kingdom",
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
