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
from sklearn.metrics import recall_score, precision_score, f1_score

DATA_DIR = "data/processed"
SEED = 42
MIN_GROUP = 300          # below this a per-group rate is too noisy to report


def load():
    # test_meta.csv holds the exact 199,682-row order that X_test.npy / rf_test_preds.npy
    # were built from (saved alongside them in notebook 02), but only the columns the
    # model's fairness slice needs (UPRN, region, age band, property type, tenure). The
    # cost/CO2/rating columns used below live in the full test parquet instead, matched
    # back in by UPRN then reordered to match test_meta's row order, not the parquet's.
    meta = pd.read_csv(f"{DATA_DIR}/test_meta.csv")
    meta["UPRN"] = meta["UPRN"].astype(str)
    full = pd.read_parquet(f"{DATA_DIR}/epc_test_full.parquet")
    full["UPRN"] = full["UPRN"].astype(str)
    extra_cols = ["UPRN", "CURRENT_ENERGY_RATING", "HEATING_COST_CURRENT",
                  "HEATING_COST_POTENTIAL", "CO2_EMISSIONS_CURRENT",
                  "CO2_EMISSIONS_POTENTIAL", "RETROFIT_POTENTIAL"]
    te = meta.merge(full[extra_cols], on="UPRN", how="left")
    assert len(te) == len(meta), "UPRN match dropped or duplicated rows"

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


def flat_threshold(te):
    """Fix the low recall on flats with a group-specific decision threshold.

    The diagnosis in fairness() is that the model is not wrong about flats, it is
    too cautious: precision is high, recall is not. That is a threshold problem,
    not a model problem, so it does not need retraining.

    The threshold has to be chosen somewhere other than where it is scored, or
    the improvement is just fitting the test set. Flats are split in half: pick
    the threshold on the first half, report on the second.
    """
    rng = np.random.default_rng(SEED)
    fl = te[te.PROPERTY_TYPE == "Flat"].copy()
    ho = te[te.PROPERTY_TYPE == "House"]
    target = recall_score(ho.RETROFIT_POTENTIAL, ho.pred)   # match house recall

    idx = rng.permutation(len(fl))
    tune, hold = fl.iloc[idx[: len(fl) // 2]], fl.iloc[idx[len(fl) // 2:]]

    grid = np.arange(0.05, 0.55, 0.01)
    chosen = 0.5
    for t in grid:                       # highest threshold that still hits target
        if recall_score(tune.RETROFIT_POTENTIAL, (tune.proba >= t).astype(int)) >= target:
            chosen = t
    before = (hold.proba >= 0.5).astype(int)
    after = (hold.proba >= chosen).astype(int)

    row = {
        "house_recall": f"{target:.3f}",
        "chosen_threshold": f"{chosen:.2f}",
        "flat_recall_before": f"{recall_score(hold.RETROFIT_POTENTIAL, before):.3f}",
        "flat_recall_after": f"{recall_score(hold.RETROFIT_POTENTIAL, after):.3f}",
        "flat_precision_before": f"{precision_score(hold.RETROFIT_POTENTIAL, before, zero_division=0):.3f}",
        "flat_precision_after": f"{precision_score(hold.RETROFIT_POTENTIAL, after, zero_division=0):.3f}",
        "flat_f1_before": f"{f1_score(hold.RETROFIT_POTENTIAL, before):.3f}",
        "flat_f1_after": f"{f1_score(hold.RETROFIT_POTENTIAL, after):.3f}",
        "n_holdout": len(hold),
    }
    print(f"  target (house recall)      {row['house_recall']}")
    print(f"  threshold chosen on tune   {row['chosen_threshold']}")
    print(f"  flat recall    {row['flat_recall_before']} to {row['flat_recall_after']}")
    print(f"  flat precision {row['flat_precision_before']} to {row['flat_precision_after']}")
    print(f"  flat F1        {row['flat_f1_before']} to {row['flat_f1_after']}")

    with open(f"{DATA_DIR}/flat_threshold.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(row))
        w.writeheader()
        w.writerow(row)
    return row


if __name__ == "__main__":
    te = load()
    print("Usable test rows: {:,}".format(len(te)))
    print()
    print("Impact, surveying 10% of D-G stock:")
    impact(te)
    print()
    print("Fairness by subgroup:")
    fairness(te)
    print()
    print("Flat-specific threshold:")
    flat_threshold(te)
