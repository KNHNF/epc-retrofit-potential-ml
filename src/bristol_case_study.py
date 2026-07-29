"""
bristol_case_study.py
Real-world application demo: run the trained Random Forest on genuine, held-out
(2025-2026, never trained or tuned on) EPC certificates for Bristol, City of.
Produces data/processed/bristol_case_study.csv, which generate_report.py reads
to build the "Real-World Application" section. No hardcoded example properties,
this is the actual model applied to actual public EPC records.

Usage: python src/bristol_case_study.py
"""

import pickle
import numpy as np
import pandas as pd
from statsmodels.stats.proportion import proportions_ztest
from sklearn.metrics import precision_score, recall_score, f1_score

from preprocessing import extract_wall_type

DATA_DIR = "data/processed"
RAW_DIR = "data/raw"
N_EXAMPLES = 8
SEED = 42

# Approximate centroid (lat, lon) for each Bristol postcode district. These
# are well-known public area locations (city centre, inner suburbs), not
# precise boundary data, used only to place a district-level marker on a
# map, not to draw an actual choropleth. Only districts that actually occur
# in the Bristol test-set data get plotted.
POSTCODE_DISTRICT_CENTROIDS = {
    "BS1":  (51.4536, -2.5951),  # City centre
    "BS2":  (51.4590, -2.5850),  # St Pauls / Kingsdown
    "BS3":  (51.4400, -2.6050),  # Bedminster / Southville
    "BS4":  (51.4370, -2.5650),  # Brislington
    "BS5":  (51.4650, -2.5600),  # Easton / St George
    "BS6":  (51.4680, -2.5900),  # Redland / Cotham
    "BS7":  (51.4820, -2.5850),  # Horfield / Filton
    "BS8":  (51.4600, -2.6200),  # Clifton
    "BS9":  (51.4800, -2.6200),  # Westbury-on-Trym / Henleaze
    "BS10": (51.5000, -2.6000),  # Southmead
    "BS11": (51.4950, -2.6800),  # Avonmouth / Lawrence Weston
    "BS13": (51.4100, -2.6300),  # Hartcliffe / Withywood
    "BS14": (51.4150, -2.5750),  # Whitchurch / Stockwood
    "BS15": (51.4600, -2.5000),  # Kingswood / Hanham
    "BS16": (51.4800, -2.5100),  # Fishponds / Downend
}


def load_bristol_uprns():
    frames = []
    for year in (2025, 2026):
        chunk = pd.read_csv(
            f"{RAW_DIR}/certificates-{year}.csv",
            encoding="latin-1",
            usecols=["uprn", "postcode", "local_authority_label"],
        )
        frames.append(chunk[chunk["local_authority_label"] == "Bristol, City of"])
    bristol = pd.concat(frames, ignore_index=True)
    bristol = bristol.rename(columns={"uprn": "UPRN"})
    bristol = bristol.dropna(subset=["UPRN"])
    bristol["UPRN"] = bristol["UPRN"].astype("int64").astype(str)
    bristol = bristol.drop_duplicates(subset="UPRN", keep="last")
    return bristol[["UPRN", "postcode"]]


def main():
    bristol_uprns = load_bristol_uprns()
    print(f"Bristol certificates found in 2025-2026 raw files: {len(bristol_uprns)}")

    test_full = pd.read_parquet(f"{DATA_DIR}/epc_test_full.parquet")
    test_full["UPRN"] = test_full["UPRN"].astype(str)
    is_bristol = test_full["UPRN"].isin(set(bristol_uprns["UPRN"]))
    bristol = test_full[is_bristol].merge(bristol_uprns, on="UPRN", how="left")
    rest_of_test = test_full[~is_bristol]
    print(f"Bristol properties present in the held-out test set (never trained/tuned on): {len(bristol)}")
    if bristol.empty:
        raise SystemExit("No Bristol properties found in the test set, check local_authority_label match.")

    bristol = extract_wall_type(bristol)

    with open(f"{DATA_DIR}/preprocessor.pkl", "rb") as f:
        preprocessor = pickle.load(f)
    with open(f"{DATA_DIR}/rf_final_model.pkl", "rb") as f:
        rf = pickle.load(f)

    feature_cols = list(preprocessor.feature_names_in_)
    X = bristol[feature_cols]
    X_enc = preprocessor.transform(X)
    proba = rf.predict_proba(X_enc)[:, 1]
    pred = (proba >= 0.5).astype(int)

    bristol = bristol.copy()
    bristol["PRED_PROBA"] = proba
    bristol["PRED_LABEL"] = pred
    bristol["OUTWARD_POSTCODE"] = bristol["postcode"].astype(str).str.split(" ").str[0]

    rng = np.random.RandomState(SEED)

    def sample_group(mask, n):
        pool = bristol[mask]
        if len(pool) == 0:
            return pool
        return pool.sample(n=min(n, len(pool)), random_state=rng)

    correct_pos = sample_group((bristol.RETROFIT_POTENTIAL == 1) & (bristol.PRED_LABEL == 1), 3)
    correct_neg = sample_group((bristol.RETROFIT_POTENTIAL == 0) & (bristol.PRED_LABEL == 0), 3)
    errors = sample_group(bristol.RETROFIT_POTENTIAL != bristol.PRED_LABEL, 2)

    demo = pd.concat([correct_pos, correct_neg, errors]).drop_duplicates(subset="UPRN")
    demo = demo.sample(frac=1.0, random_state=SEED).head(N_EXAMPLES)

    out_cols = [
        "OUTWARD_POSTCODE", "PROPERTY_TYPE", "BUILT_FORM", "CONSTRUCTION_AGE_BAND",
        "CURRENT_ENERGY_RATING", "EFFICIENCY_GAP", "PRED_PROBA", "PRED_LABEL",
        "RETROFIT_POTENTIAL",
    ]
    demo[out_cols].to_csv(f"{DATA_DIR}/bristol_case_study.csv", index=False)

    accuracy = (bristol.PRED_LABEL == bristol.RETROFIT_POTENTIAL).mean()
    positive_rate = bristol.RETROFIT_POTENTIAL.mean()
    majority_baseline_accuracy = max(positive_rate, 1 - positive_rate)
    precision = precision_score(bristol.RETROFIT_POTENTIAL, bristol.PRED_LABEL, zero_division=0)
    recall = recall_score(bristol.RETROFIT_POTENTIAL, bristol.PRED_LABEL, zero_division=0)
    f1 = f1_score(bristol.RETROFIT_POTENTIAL, bristol.PRED_LABEL, zero_division=0)

    n1, x1 = len(bristol), int(bristol.RETROFIT_POTENTIAL.sum())
    n2, x2 = len(rest_of_test), int(rest_of_test.RETROFIT_POTENTIAL.sum())
    rest_rate = x2 / n2
    z_stat, p_value = proportions_ztest([x1, x2], [n1, n2])

    summary = pd.DataFrame([{
        "n_bristol_test_properties": n1,
        "bristol_positive_rate": positive_rate,
        "bristol_accuracy": accuracy,
        "bristol_majority_baseline_accuracy": majority_baseline_accuracy,
        "bristol_precision": precision,
        "bristol_recall": recall,
        "bristol_f1": f1,
        "rest_of_test_positive_rate": rest_rate,
        "n_rest_of_test": n2,
        "positive_rate_gap_pts": (positive_rate - rest_rate) * 100,
        "positive_rate_z": z_stat,
        "positive_rate_pvalue": p_value,
    }])
    summary.to_csv(f"{DATA_DIR}/bristol_case_study_summary.csv", index=False)
    print(demo[out_cols])
    print(summary)

    # District-level aggregation for the Bristol map figure. Only districts
    # that actually occur in the data are kept; centroid is approximate
    # (see POSTCODE_DISTRICT_CENTROIDS), the rate itself is the model's
    # real predicted-positive rate for that district's test-set properties.
    district_rows = []
    for district, group in bristol.groupby("OUTWARD_POSTCODE"):
        if district not in POSTCODE_DISTRICT_CENTROIDS or len(group) < 20:
            continue
        lat, lon = POSTCODE_DISTRICT_CENTROIDS[district]
        district_rows.append({
            "district": district,
            "lat": lat,
            "lon": lon,
            "n": len(group),
            "predicted_positive_rate": group.PRED_LABEL.mean(),
            "actual_positive_rate": group.RETROFIT_POTENTIAL.mean(),
        })
    pd.DataFrame(district_rows).to_csv(f"{DATA_DIR}/bristol_district_summary.csv", index=False)


if __name__ == "__main__":
    main()
