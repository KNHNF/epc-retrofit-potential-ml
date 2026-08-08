"""
walk_forward_evaluation.py
Checks whether Random Forest holds up over time rather than just on the
main pipeline's single 2020-2024/2025-2026 split: moves the training
window forward one year at a time, five folds in total. Reuses the
main model's tuned settings by default (TUNE_PER_FOLD=True does a real
search per fold instead, needs real CPU headroom, not this sandbox).
See docs/IMPLEMENTATION_NOTES.md.

Usage: python src/walk_forward_evaluation.py
"""

import pandas as pd
import numpy as np
import pyarrow.dataset as ds
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler, OrdinalEncoder, OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import GridSearchCV, StratifiedKFold
from sklearn.metrics import f1_score, roc_auc_score, average_precision_score, recall_score

from preprocessing import extract_wall_type

DATA_DIR = "data/processed"
TRAIN_FULL = f"{DATA_DIR}/epc_train_full.parquet"  # 2020-2024
TEST_FULL = f"{DATA_DIR}/epc_test_full.parquet"    # 2025-2026

TRAIN_SAMPLE_SIZE = 200_000  # matches the main pipeline's sample size
TEST_SAMPLE_SIZE = 50_000
SEED = 42

# False: reuse the main model's settings (fast, what this ran with by default).
# True: a real 3-fold grid search inside every walk-forward fold, same grid as
# 04_Random_Forest.ipynb. Only turn this on somewhere with real CPU headroom.
TUNE_PER_FOLD = False

NUMERIC_COLS = [
    'CURRENT_ENERGY_EFFICIENCY', 'TOTAL_FLOOR_AREA',
    'NUMBER_HABITABLE_ROOMS', 'NUMBER_HEATED_ROOMS',
    'CO2_EMISS_CURR_PER_FLOOR_AREA', 'CO2_EMISSIONS_CURRENT',
    'ENERGY_CONSUMPTION_CURRENT', 'HEATING_COST_CURRENT',
    'HOT_WATER_COST_CURRENT', 'LIGHTING_COST_CURRENT',
    'EXTENSION_COUNT', 'FIXED_LIGHTING_OUTLETS_COUNT',
    'MULTI_GLAZE_PROPORTION', 'LOW_ENERGY_LIGHTING',
]
ENERGY_RATING_COLS = ['CURRENT_ENERGY_RATING']
ENERGY_RATING_ORDER = [['G', 'F', 'E', 'D', 'C', 'B', 'A']]
EFF_RATING_COLS = [
    'WALLS_ENERGY_EFF', 'ROOF_ENERGY_EFF', 'FLOOR_ENERGY_EFF',
    'WINDOWS_ENERGY_EFF', 'MAINHEAT_ENERGY_EFF', 'HOT_WATER_ENERGY_EFF',
    'LIGHTING_ENERGY_EFF', 'MAINHEATC_ENERGY_EFF',
]
EFF_RATING_VALS = [['N/A', 'Very Poor', 'Poor', 'Average', 'Good', 'Very Good']] * len(EFF_RATING_COLS)
NOMINAL_COLS = ['PROPERTY_TYPE', 'BUILT_FORM', 'TENURE', 'MAINS_GAS_FLAG', 'TRANSACTION_TYPE', 'WALL_TYPE']

READ_COLS = (
    ['LODGEMENT_DATE', 'WALLS_DESCRIPTION', 'RETROFIT_POTENTIAL']
    + NUMERIC_COLS + ENERGY_RATING_COLS + EFF_RATING_COLS
    + [c for c in NOMINAL_COLS if c != 'WALL_TYPE']
)

# Expanding window: train on every year up to and including the fold year,
# test on the year straight after. 2020 alone is too little to call a
# training set on its own, so the first fold starts at 2020-2021.
FOLDS = [
    (list(range(2020, 2022)), 2022),
    (list(range(2020, 2023)), 2023),
    (list(range(2020, 2024)), 2024),
    (list(range(2020, 2025)), 2025),
    (list(range(2020, 2026)), 2026),
]


# Rows per calendar year, counted once up front (LODGEMENT_DATE.dt.year.value_counts()
# on the full files). Loading either full parquet as one pandas frame runs out of memory
# here, so read_year below never does that, it reads in batches and keeps only a sampled
# fraction of each batch, using these counts to work out what that fraction should be.
YEAR_ROW_COUNTS = {
    2020: 1_284_712, 2021: 1_419_776, 2022: 1_533_309,
    2023: 1_489_225, 2024: 1_523_258, 2025: 1_643_058, 2026: 829_638,
}


def read_year(year, n_target, seed):
    path = TRAIN_FULL if year <= 2024 else TEST_FULL
    date_filter = (
        (ds.field('LODGEMENT_DATE') >= f'{year}-01-01')
        & (ds.field('LODGEMENT_DATE') < f'{year + 1}-01-01')
    )
    scanner = ds.dataset(path, format='parquet').scanner(
        filter=date_filter, columns=READ_COLS, batch_size=100_000,
    )
    # Sample a bit over the target fraction from every batch, then trim to the exact
    # count at the end, rather than risk under-shooting on a small last batch.
    frac = min(1.0, 1.15 * n_target / YEAR_ROW_COUNTS[year])
    rng = np.random.RandomState(seed)
    parts = []
    for batch in scanner.to_batches():
        chunk = batch.to_pandas()
        parts.append(chunk.sample(frac=frac, random_state=rng.randint(0, 2**31 - 1)))
    df = pd.concat(parts, ignore_index=True)
    if len(df) > n_target:
        df = df.sample(n=n_target, random_state=seed)
    return df


def load_window(years, total_target, seed):
    per_year = total_target // len(years)
    frames = [read_year(y, per_year, seed) for y in years]
    return pd.concat(frames, ignore_index=True)


def build_preprocessor(cols_present):
    numeric_cols = [c for c in NUMERIC_COLS if c in cols_present]
    energy_r_cols = [c for c in ENERGY_RATING_COLS if c in cols_present]
    eff_rating_cols = [c for c in EFF_RATING_COLS if c in cols_present]
    nominal_cols = [c for c in NOMINAL_COLS if c in cols_present]

    transformers = [
        ('num', Pipeline([
            ('impute', SimpleImputer(strategy='median')),
            ('scale', StandardScaler()),
        ]), numeric_cols),
        ('energy_rating', Pipeline([
            ('impute', SimpleImputer(strategy='most_frequent')),
            ('ord', OrdinalEncoder(categories=ENERGY_RATING_ORDER, handle_unknown='use_encoded_value', unknown_value=-1)),
        ]), energy_r_cols),
        ('eff_rating', Pipeline([
            ('impute', SimpleImputer(strategy='most_frequent')),
            ('ord', OrdinalEncoder(categories=EFF_RATING_VALS[:len(eff_rating_cols)], handle_unknown='use_encoded_value', unknown_value=-1)),
        ]), eff_rating_cols),
        ('nom', Pipeline([
            ('impute', SimpleImputer(strategy='most_frequent')),
            ('ohe', OneHotEncoder(handle_unknown='ignore', sparse_output=False, drop='first')),
        ]), nominal_cols),
    ]
    all_cols = numeric_cols + energy_r_cols + eff_rating_cols + nominal_cols
    return ColumnTransformer(transformers=transformers, remainder='drop'), all_cols


def run_fold(train_years, test_year):
    print(f"Fold: train {train_years[0]}-{train_years[-1]}, test {test_year}")
    train = load_window(train_years, TRAIN_SAMPLE_SIZE, SEED)
    test = read_year(test_year, TEST_SAMPLE_SIZE, SEED)

    train = extract_wall_type(train)
    test = extract_wall_type(test)

    preprocessor, feat_cols = build_preprocessor(train.columns)
    X_train = preprocessor.fit_transform(train[feat_cols])
    X_test = preprocessor.transform(test[feat_cols])
    y_train = train['RETROFIT_POTENTIAL'].values
    y_test = test['RETROFIT_POTENTIAL'].values

    if TUNE_PER_FOLD:
        param_grid = {
            'n_estimators': [200],
            'max_features': ['sqrt', 0.3],
            'min_samples_leaf': [1, 5],
        }
        search = GridSearchCV(
            RandomForestClassifier(class_weight='balanced', random_state=SEED),
            param_grid, cv=StratifiedKFold(n_splits=3, shuffle=True, random_state=SEED),
            scoring='f1_macro', n_jobs=-1,
        )
        search.fit(X_train, y_train)
        rf = RandomForestClassifier(
            n_estimators=300, class_weight='balanced', n_jobs=-1, random_state=SEED,
            max_features=search.best_params_['max_features'],
            min_samples_leaf=search.best_params_['min_samples_leaf'],
        )
        print(f"  tuned: {search.best_params_}")
    else:
        rf = RandomForestClassifier(
            n_estimators=300, max_features=0.3, min_samples_leaf=1,
            class_weight='balanced', n_jobs=-1, random_state=SEED,
        )
    rf.fit(X_train, y_train)
    proba = rf.predict_proba(X_test)[:, 1]
    pred = (proba >= 0.5).astype(int)

    return {
        'train_years': f"{train_years[0]}-{train_years[-1]}",
        'test_year': test_year,
        'n_train': len(train),
        'n_test': len(test),
        'tuned_per_fold': TUNE_PER_FOLD,
        'train_positive_rate': y_train.mean(),
        'test_positive_rate': y_test.mean(),
        'test_roc_auc': roc_auc_score(y_test, proba),
        'test_f1_macro': f1_score(y_test, pred, average='macro'),
        'test_pr_auc': average_precision_score(y_test, proba),
        'test_recall': recall_score(y_test, pred),
    }


def main():
    rows = [run_fold(years, test_year) for years, test_year in FOLDS]
    results = pd.DataFrame(rows)
    results.to_csv(f"{DATA_DIR}/walk_forward_results.csv", index=False)
    print(results.to_string(index=False))


if __name__ == "__main__":
    main()
