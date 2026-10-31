"""
offline_demo.py
Runs the project's method end to end on SYNTHETIC data, with no downloads.

What it shows, using the real functions in src/preprocessing.py:
  1. cleaning and the retrofit-potential target (D-G rated and gap of 20+ points)
  2. a temporal split: certificates lodged before 2025 train, 2025 onwards test
  3. nested cross-validation (outer 3-fold, inner 2-fold grid search) on the training
     years only, with the preprocessing INSIDE the cross-validation so no fold's
     imputation or scaling statistics come from its own validation rows
  4. a final fit on all training years and a prediction on the held-out later years

The data is random numbers shaped like EPC columns. It says nothing about real homes,
and this script deliberately prints no model scores: the point is the structure, which
is checked with assertions. Real results need the real EPC data (see the README).

Usage: python examples/offline_demo.py
"""

import os
import sys

import numpy as np
import pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import GridSearchCV, StratifiedKFold
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))
from preprocessing import (drop_duplicates_by_uprn, engineer_target, extract_wall_type,
                           filter_valid_records)

SPLIT_DATE = '2025-01-01'   # same boundary as src/00_prepare_data.py: 2025-2026 is test
NUMERIC = ['CURRENT_ENERGY_EFFICIENCY', 'TOTAL_FLOOR_AREA', 'NUMBER_HABITABLE_ROOMS']
NOMINAL = ['PROPERTY_TYPE', 'WALL_TYPE']


def band(eff):
    """SAP efficiency score to EPC band letter."""
    edges = [(92, 'A'), (81, 'B'), (69, 'C'), (55, 'D'), (39, 'E'), (21, 'F')]
    return next((letter for lo, letter in edges if eff >= lo), 'G')


def make_synthetic_epc(n=1500, seed=0):
    rng = np.random.default_rng(seed)
    cur = np.clip(rng.normal(60, 13, n), 1, 100).round()
    wall = rng.choice(['Cavity wall, as built', 'Solid brick, as built', 'Timber frame'], n)
    # Potential depends on wall type plus noise, so there is a learnable signal.
    uplift = np.where(np.char.find(wall.astype(str), 'Solid') >= 0, 22, 10) + rng.normal(0, 8, n)
    pot = np.clip(cur + np.clip(uplift, 0, None), 1, 100).round()
    dates = pd.Timestamp('2020-01-01') + pd.to_timedelta(rng.integers(0, 7 * 365, n), unit='D')
    df = pd.DataFrame({
        'UPRN': np.arange(1, n + 1).astype(str),
        'LODGEMENT_DATE': dates.strftime('%Y-%m-%d'),
        'CURRENT_ENERGY_EFFICIENCY': cur,
        'POTENTIAL_ENERGY_EFFICIENCY': pot,
        'CURRENT_ENERGY_RATING': [band(c) for c in cur],
        'POTENTIAL_ENERGY_RATING': [band(p) for p in pot],
        'PROPERTY_TYPE': rng.choice(['House', 'Flat', 'Bungalow'], n),
        'WALLS_DESCRIPTION': wall,
        'TOTAL_FLOOR_AREA': rng.normal(85, 25, n).clip(20, None).round(1),
        'NUMBER_HABITABLE_ROOMS': rng.integers(1, 8, n).astype(float),
    })
    # A few deliberate faults for the cleaning step to catch.
    df.loc[0, 'CURRENT_ENERGY_RATING'] = 'Z'
    df.loc[1, 'TOTAL_FLOOR_AREA'] = -5.0
    df.loc[2, 'TOTAL_FLOOR_AREA'] = np.nan          # missing, must be kept and imputed
    df = pd.concat([df, df.iloc[[10]].assign(LODGEMENT_DATE='2019-12-31')])   # older duplicate
    return df


def build_model():
    pre = ColumnTransformer([
        ('num', Pipeline([('impute', SimpleImputer(strategy='median')),
                          ('scale', StandardScaler())]), NUMERIC),
        ('nom', Pipeline([('impute', SimpleImputer(strategy='most_frequent')),
                          ('ohe', OneHotEncoder(handle_unknown='ignore', drop='first'))]), NOMINAL),
    ])
    return Pipeline([('pre', pre),
                     ('clf', LogisticRegression(class_weight='balanced', max_iter=1000))])


def run_demo(n=1500, seed=0):
    raw = make_synthetic_epc(n, seed)
    df = drop_duplicates_by_uprn(raw)
    df = filter_valid_records(df)
    df = extract_wall_type(df)
    df = engineer_target(df, gap_threshold=20)
    df['LODGEMENT_DATE'] = pd.to_datetime(df['LODGEMENT_DATE'])

    # Temporal holdout: split on date, never at random.
    train = df[df['LODGEMENT_DATE'] < SPLIT_DATE]
    test = df[df['LODGEMENT_DATE'] >= SPLIT_DATE]
    checks = {
        'rows_after_cleaning': len(df),
        'rows_dropped_by_cleaning_and_dedup': len(raw) - len(df),
        'train_rows': len(train),
        'test_rows': len(test),
        'train_ends_before_test_starts': bool(train['LODGEMENT_DATE'].max() < test['LODGEMENT_DATE'].min()),
        'no_uprn_in_both_sets': set(train['UPRN']).isdisjoint(test['UPRN']),
        'missing_floor_area_row_kept': bool(df['TOTAL_FLOOR_AREA'].isna().any()),
    }

    X_train, y_train = train[NUMERIC + NOMINAL], train['RETROFIT_POTENTIAL']
    X_test, y_test = test[NUMERIC + NOMINAL], test['RETROFIT_POTENTIAL']

    # Nested CV on training years only. The inner GridSearchCV refits the whole pipeline
    # (imputer, scaler, encoder, model) on each inner training fold.
    outer = StratifiedKFold(n_splits=3, shuffle=True, random_state=seed)
    inner = StratifiedKFold(n_splits=2, shuffle=True, random_state=seed)
    search = GridSearchCV(build_model(), {'clf__C': [0.1, 1.0]}, cv=inner, scoring='f1_macro')
    outer_scores, chosen_C = [], []
    for tr, va in outer.split(X_train, y_train):
        search.fit(X_train.iloc[tr], y_train.iloc[tr])
        chosen_C.append(search.best_params_['clf__C'])
        outer_scores.append(search.score(X_train.iloc[va], y_train.iloc[va]))
    checks['outer_folds_scored'] = len(outer_scores)

    # Final model: fit on every training row, predict the later period.
    final = build_model().set_params(clf__C=max(set(chosen_C), key=chosen_C.count))
    final.fit(X_train, y_train)
    pred = final.predict(X_test)
    checks['one_prediction_per_test_row'] = bool(len(pred) == len(test))

    assert checks['train_ends_before_test_starts'], 'temporal split leaked'
    assert checks['no_uprn_in_both_sets'], 'a property is in both splits'
    assert checks['one_prediction_per_test_row']
    return checks


def main():
    print('Offline demo on SYNTHETIC data (no model scores are printed on purpose).')
    for key, value in run_demo().items():
        print(f'  {key}: {value}')
    print('Structure checks passed.')


if __name__ == '__main__':
    main()
