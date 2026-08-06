"""
impact_and_fairness.py

Two things the report was missing.

1. Impact in money and carbon. The EPC register already carries the current and
   potential heating cost and CO2 for every certificate, so the saving from
   taking a home to its potential is in the data, not an assumption. The useful
   comparison is what a fixed survey budget captures when it ranks by model
   score against screening on rating band alone.

2. Fairness. All headline metrics are aggregate. This breaks recall down by
   property type and tenure to check the model is not systematically worse for
   one kind of home.

Savings are the register's own current-minus-potential figures. They assume the
full recommended package is installed, which is an upper bound, not a forecast.

Usage: python src/impact_and_fairness.py
Writes data/processed/impact_summary.csv and data/processed/fairness_by_group.csv
"""

import csv
import numpy as np
import pandas as pd
from sklearn.metrics import recall_score, precision_score

DATA_DIR = "data/processed"
SEED = 42
MIN_GROUP = 300          # below this a per-group rate is too noisy to report


def load():
    te = pd.read_parquet(f"{DATA_DIR}/epc_test_sample_50k.parquet")
    te = te.reset_index(drop=True)
    te["pred"] = np.load(f"{DATA_DIR}/rf_test_preds.npy")
    te["proba"] = np.load(f"{DATA_DIR}/rf_test_probs.npy")
    te["cost_saving"] = te.HEATING_COST_CURRENT - te.HEATING_COST_POTENTIAL
    te["co2_saving"] = te.CO2_EMISSIONS_CURRENT - te.CO2_EMISSIONS_POTENTIAL
    # a negative saving means the recommended package costs more to run, which is
    # a data error rather than a real result; drop rather than let it net off
    te = te[(te.cost_saving >= 0) & (te.co2_saving >= 0)].copy()
    return te


def impact(te):
    """What a fixed survey budget captures, ranked by model vs by rating band."""
    rng = np.random.default_rng(SEED)
    dg = te[te.CURRENT_ENERGY_RATING.isin(list("DEFG"))].copy()
    budget = int(0.10 * len(dg))          # survey 10% of the eligible stock

    by_model = dg.nlargest(budget, "proba")
    # rating-band screening has no ordering within the band, so it is a random
    # draw from D-G; average over repeats rather than trusting one draw
    rand_cost, rand_co2, rand_hits = [], [], []
    for _ in range(200):
        s = dg.sample(budget, random_state=int(rng.integers(1e9)))
        rand_cost.append(s.cost_saving.sum())
        rand_co2.append(s.co2_saving.sum())
        rand_hits.append(int(s.RETROFIT_POTENTIAL.sum()))

    rows = [{
        "strategy": "Ranked by model",
        "surveyed": budget,
        "true_positives": int(by_model.RETROFIT_POTENTIAL.sum()),
        "annual_cost_saving_gbp": int(by_model.cost_saving.sum()),
        "annual_co2_saving_t": round(float(by_model.co2_saving.sum()), 1),
    }, {
        "strategy": "Rating band only",
        "surveyed": budget,
        "true_positives": int(np.mean(rand_hits)),
        "annual_cost_saving_gbp": int(np.mean(rand_cost)),
        "annual_co2_saving_t": round(float(np.mean(rand_co2)), 1),
    }]
    for r in rows:
        print(f"  {r['strategy']:18s} {r['true_positives']:>5} hits  "
              f"GBP {r['annual_cost_saving_gbp']:>9,}  {r['annual_co2_saving_t']:>7} t")

    uplift = rows[0]["annual_cost_saving_gbp"] / max(rows[1]["annual_cost_saving_gbp"], 1)
    print(f"  uplift: {uplift:.2f}x on cost captured for the same number of surveys")

    with open(f"{DATA_DIR}/impact_summary.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]) + ["uplift_x"])
        w.writeheader()
        for r in rows:
            r["uplift_x"] = f"{uplift:.2f}"
            w.writerow(r)
    return rows, uplift


def fairness(te):
    """Recall by subgroup. Recall is the metric the report argues matters."""
    out = []
    for col, label in [("PROPERTY_TYPE", "Property type"), ("TENURE", "Tenure")]:
        for val, g in te.groupby(col):
            pos = int(g.RETROFIT_POTENTIAL.sum())
            if len(g) < MIN_GROUP or pos < 30:
                continue
            out.append({
                "dimension": label,
                "group": str(val),
                "n": len(g),
                "positive_rate": f"{g.RETROFIT_POTENTIAL.mean():.3f}",
                "recall": f"{recall_score(g.RETROFIT_POTENTIAL, g.pred):.3f}",
                "precision": f"{precision_score(g.RETROFIT_POTENTIAL, g.pred, zero_division=0):.3f}",
            })
    for r in out:
        print(f"  {r['dimension']:14s} {r['group'][:24]:24s} n={r['n']:>6} "
              f"recall={r['recall']} precision={r['precision']}")
    recalls = [float(r["recall"]) for r in out]
    spread = max(recalls) - min(recalls)
    print(f"  recall spread across groups: {spread:.3f}")

    with open(f"{DATA_DIR}/fairness_by_group.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0]))
        w.writeheader()
        w.writerows(out)
    return out, spread


if __name__ == "__main__":
    te = load()
    print(f"Usable test rows: {len(te):,}\n")
    print("Impact, surveying 10% of D-G stock:")
    impact(te)
    print("\nFairness by subgroup:")
    fairness(te)
