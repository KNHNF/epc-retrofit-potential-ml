"""
city_case_studies.py
Bristol on its own is one city, not proof the model works everywhere.
This runs the same held-out check on two more cities with a different
housing mix, Manchester and Leeds, both northern, both denser and more
terraced than Bristol, using the model already trained on 2020-2024 (no
retraining here, same rf_final_model.pkl and preprocessor.pkl as the
Bristol case study). Produces data/processed/city_case_studies_summary.csv,
which generate_report.py reads for the extra cities table.

Usage: python src/city_case_studies.py
"""

import pickle
import pandas as pd
from statsmodels.stats.proportion import proportions_ztest
from sklearn.metrics import precision_score, recall_score, f1_score

from preprocessing import extract_wall_type

DATA_DIR = "data/processed"
RAW_DIR = "data/raw"

CITIES = ["Manchester", "Leeds"]


def load_city_uprns(local_authority_label):
    frames = []
    for year in (2025, 2026):
        chunk = pd.read_csv(
            f"{RAW_DIR}/certificates-{year}.csv",
            encoding="latin-1",
            usecols=["uprn", "local_authority_label"],
        )
        frames.append(chunk[chunk["local_authority_label"] == local_authority_label])
    city = pd.concat(frames, ignore_index=True)
    city = city.rename(columns={"uprn": "UPRN"})
    city = city.dropna(subset=["UPRN"])
    city["UPRN"] = city["UPRN"].astype("int64").astype(str)
    city = city.drop_duplicates(subset="UPRN", keep="last")
    return city[["UPRN"]]


def evaluate_city(name, preprocessor, model, test_full):
    uprns = load_city_uprns(name)
    is_city = test_full["UPRN"].isin(set(uprns["UPRN"]))
    city = test_full[is_city].copy()
    rest = test_full[~is_city]
    if city.empty:
        print(f"No {name} properties found in the test set, skipping.")
        return None

    city = extract_wall_type(city)
    feature_cols = list(preprocessor.feature_names_in_)
    X = preprocessor.transform(city[feature_cols])
    proba = model.predict_proba(X)[:, 1]
    pred = (proba >= 0.5).astype(int)

    accuracy = (pred == city.RETROFIT_POTENTIAL).mean()
    positive_rate = city.RETROFIT_POTENTIAL.mean()
    majority_baseline = max(positive_rate, 1 - positive_rate)
    precision = precision_score(city.RETROFIT_POTENTIAL, pred, zero_division=0)
    recall = recall_score(city.RETROFIT_POTENTIAL, pred, zero_division=0)
    f1 = f1_score(city.RETROFIT_POTENTIAL, pred, zero_division=0)

    n1, x1 = len(city), int(city.RETROFIT_POTENTIAL.sum())
    n2, x2 = len(rest), int(rest.RETROFIT_POTENTIAL.sum())
    rest_rate = x2 / n2
    _, p_value = proportions_ztest([x1, x2], [n1, n2])

    return {
        "city": name,
        "n_test_properties": n1,
        "positive_rate": positive_rate,
        "accuracy": accuracy,
        "majority_baseline_accuracy": majority_baseline,
        "precision": precision,
        "recall": recall,
        "f1": f1,
        "rest_of_test_positive_rate": rest_rate,
        "positive_rate_gap_pts": (positive_rate - rest_rate) * 100,
        "positive_rate_pvalue": p_value,
    }


def main():
    test_full = pd.read_parquet(f"{DATA_DIR}/epc_test_full.parquet")
    test_full["UPRN"] = test_full["UPRN"].astype(str)

    with open(f"{DATA_DIR}/preprocessor.pkl", "rb") as f:
        preprocessor = pickle.load(f)
    with open(f"{DATA_DIR}/rf_final_model.pkl", "rb") as f:
        model = pickle.load(f)

    rows = [evaluate_city(city, preprocessor, model, test_full) for city in CITIES]
    rows = [r for r in rows if r is not None]

    # Bristol's own numbers already sit in bristol_case_study_summary.csv, folded in
    # here too so the report can build one table across all three cities.
    bristol_path = f"{DATA_DIR}/bristol_case_study_summary.csv"
    try:
        b = pd.read_csv(bristol_path).iloc[0]
        rows.append({
            "city": "Bristol",
            "n_test_properties": int(b["n_bristol_test_properties"]),
            "positive_rate": b["bristol_positive_rate"],
            "accuracy": b["bristol_accuracy"],
            "majority_baseline_accuracy": b["bristol_majority_baseline_accuracy"],
            "precision": b["bristol_precision"],
            "recall": b["bristol_recall"],
            "f1": b["bristol_f1"],
            "rest_of_test_positive_rate": b["rest_of_test_positive_rate"],
            "positive_rate_gap_pts": b["positive_rate_gap_pts"],
            "positive_rate_pvalue": b["positive_rate_pvalue"],
        })
    except FileNotFoundError:
        print("Bristol summary not found, run bristol_case_study.py first for the combined table.")

    out = pd.DataFrame(rows)
    out.to_csv(f"{DATA_DIR}/city_case_studies_summary.csv", index=False)
    print(out.to_string(index=False))


if __name__ == "__main__":
    main()
