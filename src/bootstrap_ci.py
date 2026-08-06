"""
bootstrap_ci.py
Bootstrap confidence intervals on test ROC-AUC, and on the paired gap between
the top two models.

The point is the paired gap. Random Forest beats XGBoost by 0.0011 ROC-AUC,
which is small enough to ask whether it is just the luck of one 50,000-row test
sample. Resampling the test set with replacement answers that directly: if the
gap holds across resamples, it is a real ordering, not noise.

Usage: python src/bootstrap_ci.py
Writes data/processed/bootstrap_ci.csv
"""

import csv
import numpy as np
from sklearn.metrics import roc_auc_score

DATA_DIR = "data/processed"
N_RESAMPLES = 2000
SEED = 42

MODELS = {
    "Logistic Regression": "lr",
    "Random Forest": "rf",
    "XGBoost": "xgb",
    "SVM (LinearSVC)": "svm",
}


def main():
    rng = np.random.default_rng(SEED)
    y = np.load(f"{DATA_DIR}/y_test.npy")
    probs = {name: np.load(f"{DATA_DIR}/{key}_test_probs.npy")
             for name, key in MODELS.items()}

    idx = np.arange(len(y))
    draws = [rng.choice(idx, size=len(idx), replace=True) for _ in range(N_RESAMPLES)]
    # same draws for every model so the paired gap below is a like-for-like
    # comparison rather than two independent noise sources
    draws = [s for s in draws if y[s].sum() > 0]

    rows = []
    per_model = {}
    for name in MODELS:
        vals = np.array([roc_auc_score(y[s], probs[name][s]) for s in draws])
        per_model[name] = vals
        lo, hi = np.percentile(vals, [2.5, 97.5])
        rows.append({
            "model": name,
            "test_roc_auc": f"{roc_auc_score(y, probs[name]):.4f}",
            "ci_low": f"{lo:.4f}",
            "ci_high": f"{hi:.4f}",
        })
        print(f"{name:22s} {rows[-1]['test_roc_auc']}  [{lo:.4f}, {hi:.4f}]")

    ranked = sorted(rows, key=lambda r: float(r["test_roc_auc"]), reverse=True)
    top, second = ranked[0]["model"], ranked[1]["model"]
    diff = per_model[top] - per_model[second]
    d_lo, d_hi = np.percentile(diff, [2.5, 97.5])
    share = float((diff > 0).mean())

    print(f"\n{top} minus {second}: {diff.mean():+.4f}  [{d_lo:+.4f}, {d_hi:+.4f}]")
    print(f"ahead in {share:.1%} of resamples")

    with open(f"{DATA_DIR}/bootstrap_ci.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["model", "test_roc_auc", "ci_low", "ci_high"])
        w.writeheader()
        w.writerows(rows)

    with open(f"{DATA_DIR}/bootstrap_gap.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["top_model", "second_model", "gap", "gap_ci_low",
                    "gap_ci_high", "share_ahead", "n_resamples"])
        w.writerow([top, second, f"{diff.mean():.4f}", f"{d_lo:.4f}",
                    f"{d_hi:.4f}", f"{share:.4f}", len(draws)])
    print(f"\nSaved to {DATA_DIR}/bootstrap_ci.csv and bootstrap_gap.csv")


if __name__ == "__main__":
    main()
