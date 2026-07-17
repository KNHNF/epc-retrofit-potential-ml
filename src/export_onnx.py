"""
export_onnx.py
Exports the winning model (preprocessor + Random Forest) to a single ONNX file
that can run entirely client-side in a browser via onnxruntime-web, no backend.

Run this on the SAME machine/environment that trained the models (matching
scikit-learn version). preprocessor.pkl and rf_final_model.pkl were saved with
scikit-learn 1.9.0; loading them with a different version can silently produce
wrong results (sklearn warns about this itself), so don't run this somewhere
the version doesn't match without re-checking.

Why this isn't a straight pickle-to-ONNX conversion, and why there is no
imputer in this export pipeline at all:

skl2onnx's Imputer converter only supports a string sentinel for missing
values on string-typed columns, not NaN, so a direct conversion fails.

But there's a second, more important reason: the ORIGINAL fitted preprocessor
has a real, pre-existing quirk. Several categorical EPC columns
(ROOF_ENERGY_EFF, FLOOR_ENERGY_EFF, MAINS_GAS_FLAG, BUILT_FORM,
TRANSACTION_TYPE) store missing values as Python None, not float NaN.
SimpleImputer(missing_values=np.nan) never catches these, so in the model
that was actually trained and reported on:
  - eff_rating columns: a missing value passes through unimputed and lands on
    OrdinalEncoder's unknown_value=-1 (NOT the mode, despite the design intent).
  - nominal columns: a missing value was seen as a genuine category during
    fit, so literal None already sits inside categories_ for BUILT_FORM,
    MAINS_GAS_FLAG, and TRANSACTION_TYPE.

An imputer that actually imputes (which is what a first attempt at this
export did) is *more correct* than the original, and therefore produces a
DIFFERENT feature matrix and different predictions than the model your
report's numbers are based on. To export faithfully, this script reproduces
that behaviour rather than fixing it: no imputer, and missing categorical
values are represented as the string 'None', with any literal Python None
already inside the original OneHotEncoder's categories_ rewritten to the
string 'None' so the ONNX string tensor can still land on that exact slot.

Install first:
    pip install skl2onnx onnx onnxruntime

Usage:
    python src/export_onnx.py

Outputs:
    report/model/epc_retrofit_rf.onnx   -- the combined preprocessor+model graph
    report/model/feature_schema.json    -- raw column names, types, and valid
                                            category values, for building the
                                            browser form correctly
"""

import os
import json
import pickle

import numpy as np
import pandas as pd

from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OrdinalEncoder, OneHotEncoder

from skl2onnx import convert_sklearn
from skl2onnx.common.data_types import FloatTensorType, StringTensorType

DATA_DIR = "data/processed"
OUT_DIR  = "report/model"
os.makedirs(OUT_DIR, exist_ok=True)

# Must match notebooks/02_Feature_Engineering.ipynb exactly.
NUMERIC_COLS = [
    'CURRENT_ENERGY_EFFICIENCY', 'TOTAL_FLOOR_AREA',
    'NUMBER_HABITABLE_ROOMS', 'NUMBER_HEATED_ROOMS',
    'CO2_EMISS_CURR_PER_FLOOR_AREA', 'CO2_EMISSIONS_CURRENT',
    'ENERGY_CONSUMPTION_CURRENT', 'HEATING_COST_CURRENT',
    'HOT_WATER_COST_CURRENT', 'LIGHTING_COST_CURRENT',
    'EXTENSION_COUNT', 'FIXED_LIGHTING_OUTLETS_COUNT',
    'MULTI_GLAZE_PROPORTION', 'LOW_ENERGY_LIGHTING',
]
ENERGY_RATING_COLS  = ['CURRENT_ENERGY_RATING']
ENERGY_RATING_ORDER = [['G', 'F', 'E', 'D', 'C', 'B', 'A']]

EFF_RATING_COLS = [
    'WALLS_ENERGY_EFF', 'ROOF_ENERGY_EFF', 'FLOOR_ENERGY_EFF',
    'WINDOWS_ENERGY_EFF', 'MAINHEAT_ENERGY_EFF', 'HOT_WATER_ENERGY_EFF',
    'LIGHTING_ENERGY_EFF', 'MAINHEATC_ENERGY_EFF',
]
EFF_RATING_VALS = [['N/A', 'Very Poor', 'Poor', 'Average', 'Good', 'Very Good']] * len(EFF_RATING_COLS)

NOMINAL_COLS = [
    'PROPERTY_TYPE', 'BUILT_FORM', 'TENURE',
    'MAINS_GAS_FLAG', 'TRANSACTION_TYPE', 'WALL_TYPE',
]
STRING_COLS = ENERGY_RATING_COLS + EFF_RATING_COLS + NOMINAL_COLS
ALL_COLS = NUMERIC_COLS + STRING_COLS


def extract_wall_type(df):
    """Same logic as notebook 02 -- WALL_TYPE is derived, not a raw EPC column."""
    df = df.copy()
    if 'WALLS_DESCRIPTION' in df.columns:
        desc = df['WALLS_DESCRIPTION'].str.lower().fillna('')
        df['WALL_TYPE'] = 'other'
        df.loc[desc.str.contains('cavity'), 'WALL_TYPE'] = 'cavity'
        df.loc[desc.str.contains('solid'), 'WALL_TYPE'] = 'solid'
    else:
        df['WALL_TYPE'] = 'unknown'
    return df


def drop_negative_records(df):
    """Same negative-value filter as notebook 02: drop records with physically impossible
    negative energy/emissions/cost values (~0.1%). NaN is kept (missing, imputed later)."""
    non_neg = [
        'CURRENT_ENERGY_EFFICIENCY', 'TOTAL_FLOOR_AREA',
        'NUMBER_HABITABLE_ROOMS', 'NUMBER_HEATED_ROOMS',
        'CO2_EMISS_CURR_PER_FLOOR_AREA', 'CO2_EMISSIONS_CURRENT',
        'ENERGY_CONSUMPTION_CURRENT', 'HEATING_COST_CURRENT',
        'HOT_WATER_COST_CURRENT', 'LIGHTING_COST_CURRENT',
        'EXTENSION_COUNT', 'FIXED_LIGHTING_OUTLETS_COUNT',
        'MULTI_GLAZE_PROPORTION', 'LOW_ENERGY_LIGHTING',
    ]
    present = [c for c in non_neg if c in df.columns]
    if present:
        df = df[~(df[present] < 0).any(axis=1)]
    return df


def stringify_missing(df, cols):
    """Convert real nulls to the literal string 'None', matching how the
    ORIGINAL preprocessor's categories_ already represents missingness for
    columns where its imputer silently failed to catch Python None (see
    module docstring). This is NOT the same as a normal missing-value
    sentinel like '' -- it must match the exact slot the original encoder
    already learned."""
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = df[c].astype(object).where(df[c].notna(), 'None').astype(str)
    return df


def stringify_categories(categories_list):
    """Replace literal Python None inside a fitted OneHotEncoder's
    categories_ with the string 'None', so an ONNX string tensor can still
    address that exact category slot."""
    out = []
    for cats in categories_list:
        out.append([('None' if c is None else c) for c in cats])
    return out


def main():
    print("Loading fitted preprocessor and model...")
    with open(f'{DATA_DIR}/preprocessor.pkl', 'rb') as f:
        original_preprocessor = pickle.load(f)
    with open(f'{DATA_DIR}/rf_final_model.pkl', 'rb') as f:
        rf_model = pickle.load(f)

    original_num_pipeline = None
    original_ohe_categories = None
    for name, trans, cols in original_preprocessor.transformers_:
        if name == 'num':
            original_num_pipeline = trans
        if name == 'nom':
            original_ohe_categories = trans.named_steps['ohe'].categories_

    # No imputer in any of these groups -- see module docstring for why.
    # energy_rating / eff_rating: fixed categories with no 'None' entry, so a
    # missing value (fed in as the string 'None') falls through to
    # unknown_value=-1, exactly matching the original's actual (unimputed)
    # behaviour on these columns.
    export_transformers = [
        ('num', original_num_pipeline, NUMERIC_COLS),
        ('energy_rating', OrdinalEncoder(
            categories=ENERGY_RATING_ORDER,
            handle_unknown='use_encoded_value', unknown_value=-1,
        ), ENERGY_RATING_COLS),
        ('eff_rating', OrdinalEncoder(
            categories=EFF_RATING_VALS,
            handle_unknown='use_encoded_value', unknown_value=-1,
        ), EFF_RATING_COLS),
        # nom: categories pinned to the ORIGINAL fitted encoder's categories_,
        # with any literal None rewritten to the string 'None' so the ONNX
        # string tensor can land on that exact slot instead of falling to
        # "unknown" -- BUILT_FORM, MAINS_GAS_FLAG, and TRANSACTION_TYPE all
        # have None as a real, already-learned category (see module docstring).
        ('nom', OneHotEncoder(
            categories=stringify_categories(original_ohe_categories),
            handle_unknown='ignore', sparse_output=False, drop='first',
        ), NOMINAL_COLS),
    ]

    export_preprocessor = ColumnTransformer(transformers=export_transformers, remainder='drop')

    print("Fitting export preprocessor structure (categories are pinned to the "
          "original's, so this fit only initialises sklearn's internal state, "
          "it does not relearn anything data-driven for the pinned encoders)...")
    train = pd.read_parquet(f'{DATA_DIR}/epc_train_sample_200k.parquet')
    train = drop_negative_records(train)  # drop the ~0.1% negative-value records, as in nb02
    train = extract_wall_type(train)
    train_for_fit = stringify_missing(train, STRING_COLS)
    export_preprocessor.fit(train_for_fit[ALL_COLS])

    # --- Sanity check before ever touching ONNX: does the export preprocessor
    # produce the exact same 51-column output as the original preprocessor on
    # real data, missing values and all? If this doesn't match, stop, don't
    # export a silently wrong model.
    test = pd.read_parquet(f'{DATA_DIR}/epc_test_sample_50k.parquet')
    test = drop_negative_records(test)  # keep schema mins consistent with the cleaned data
    test = extract_wall_type(test)
    sample = test[ALL_COLS].head(500).copy()
    sample_stringified = stringify_missing(sample, STRING_COLS)

    X_original = original_preprocessor.transform(sample)
    X_export   = export_preprocessor.transform(sample_stringified)
    max_feature_diff = np.max(np.abs(X_original - X_export))
    print(f"Max feature-matrix difference (original vs export preprocessor): "
          f"{max_feature_diff:.6f}")
    if max_feature_diff > 1e-6:
        diff = np.abs(X_original - X_export)
        bad_rows, bad_cols = np.where(diff > 1e-6)
        with open(f'{DATA_DIR}/feature_names.pkl', 'rb') as f:
            feature_names = pickle.load(f)
        print(f"STOP: {len(bad_rows)} cells still differ. First few:")
        for r, c in list(zip(bad_rows, bad_cols))[:5]:
            name = feature_names[c] if c < len(feature_names) else f"col_{c}"
            print(f"  row {r}, col {c} ({name}): original={X_original[r,c]}, export={X_export[r,c]}")
        print("Do not proceed with export until this is fixed.")
        return

    print("PASS: export preprocessor reproduces the original feature matrix exactly.")

    export_pipeline = Pipeline([
        ('preprocessor', export_preprocessor),
        ('classifier', rf_model),
    ])

    # --- Confirm the export pipeline's predictions agree with the ORIGINAL
    # trained pipeline (original preprocessor, same RF) on the same 500 rows.
    orig_X = original_preprocessor.transform(sample)
    orig_probs = rf_model.predict_proba(orig_X)[:, 1]
    export_probs_sklearn = export_pipeline.predict_proba(sample_stringified)[:, 1]
    agree_diff = np.max(np.abs(orig_probs - export_probs_sklearn))
    print(f"Max prediction diff vs ORIGINAL trained pipeline on {len(sample)} rows: {agree_diff:.6f}")
    if agree_diff > 1e-6:
        print("STOP: export pipeline predictions do not match the original trained "
              "pipeline. Do not proceed with export until this is fixed.")
        return
    print("PASS: export pipeline predictions match the original trained pipeline exactly.")

    initial_types = []
    for col in NUMERIC_COLS:
        initial_types.append((col, FloatTensorType([None, 1])))
    for col in STRING_COLS:
        initial_types.append((col, StringTensorType([None, 1])))

    print(f"Converting to ONNX ({len(initial_types)} input columns)...")
    onnx_model = convert_sklearn(
        export_pipeline,
        initial_types=initial_types,
        target_opset=15,
        options={id(rf_model): {'zipmap': False}},
    )

    onnx_path = f'{OUT_DIR}/epc_retrofit_rf.onnx'
    with open(onnx_path, 'wb') as f:
        f.write(onnx_model.SerializeToString())
    print(f"Saved {onnx_path} ({os.path.getsize(onnx_path)/1e6:.1f} MB)")

    # --- Validation: real test rows through sklearn (export_pipeline) vs
    # through ONNX Runtime.
    print("\nValidating ONNX output against the export pipeline's own sklearn output...")
    import onnxruntime as rt
    sess = rt.InferenceSession(onnx_path, providers=['CPUExecutionProvider'])
    onnx_inputs = {}
    for col in NUMERIC_COLS:
        onnx_inputs[col] = sample_stringified[[col]].astype(np.float32).values
    for col in STRING_COLS:
        onnx_inputs[col] = sample_stringified[[col]].astype(str).values

    onnx_out = sess.run(None, onnx_inputs)
    onnx_probs = onnx_out[1][:, 1]

    diff = np.abs(export_probs_sklearn - onnx_probs)
    max_diff = diff.max()
    mean_diff = diff.mean()
    sklearn_class = (export_probs_sklearn >= 0.5).astype(int)
    onnx_class = (onnx_probs >= 0.5).astype(int)
    n_flipped = int((sklearn_class != onnx_class).sum())

    print(f"Random Forest: n_estimators={rf_model.n_estimators}, "
          f"max_depth={rf_model.max_depth}")
    print(f"Max probability difference over {len(sample_stringified)} test rows: {max_diff:.6f}")
    print(f"Mean probability difference: {mean_diff:.6f}")
    print(f"Rows where the predicted CLASS (0.5 threshold) flips between "
          f"sklearn and ONNX: {n_flipped} / {len(sample_stringified)}")

    BOUNDARY_BAND = 0.10  # sklearn_prob within 0.5 +/- this counts as genuine boundary noise

    if max_diff < 1e-4:
        print("PASS: ONNX output matches sklearn output.")
    elif n_flipped == 0:
        print("PASS (with caveat): probabilities differ slightly (likely ONNX Runtime's "
              "float32 aggregation over many trees vs sklearn's float64 predict_proba), "
              "but zero classification decisions change at the 0.5 threshold across this "
              "sample. Acceptable for a demo that shows probabilities, but state this "
              "precision caveat wherever the ONNX model's numbers are displayed.")
    else:
        flipped_idx = np.where(sklearn_class != onnx_class)[0]
        all_near_boundary = True
        print(f"{n_flipped} row(s) flip class at the 0.5 threshold:")
        for i in flipped_idx:
            p = export_probs_sklearn[i]
            near = abs(p - 0.5) <= BOUNDARY_BAND
            all_near_boundary = all_near_boundary and near
            print(f"  row {i}: sklearn_prob={p:.6f}, onnx_prob={onnx_probs[i]:.6f}, "
                  f"{'near 0.5 (expected float32 boundary noise)' if near else 'NOT near 0.5 -- unexplained'}")

        if not all_near_boundary:
            print("STOP: at least one flipped row is not near the 0.5 boundary, so this "
                  "is not simple float32 aggregation noise. Do not use this export, "
                  "something else is wrong.")
            return

        print(f"PASS (with caveat): every flipped row is within {BOUNDARY_BAND} of the 0.5 "
              "decision boundary in sklearn too, consistent with float32 (ONNX Runtime) vs "
              "float64 (sklearn) aggregation noise across 300 trees, not a conversion bug. "
              "Proceeding, but: (1) do not pick a demo property whose sklearn probability "
              "sits within this band of 0.5, check schema-time probabilities before choosing "
              "examples, and (2) state this precision caveat on the demo page.")

    # --- Save the raw feature schema for the browser form/demo data: column
    # names, types, and valid category values (with 'None' meaning "missing",
    # not a real option to expose in a form).
    schema = {'numeric': {}, 'categorical': {}}
    for col in NUMERIC_COLS:
        schema['numeric'][col] = {
            'min': float(test[col].min()) if col in test.columns else None,
            'max': float(test[col].max()) if col in test.columns else None,
            'median': float(test[col].median()) if col in test.columns else None,
        }
    schema['categorical'][ENERGY_RATING_COLS[0]] = ENERGY_RATING_ORDER[0]
    for i, col in enumerate(EFF_RATING_COLS):
        schema['categorical'][col] = EFF_RATING_VALS[i]
    for name, trans, cols in export_preprocessor.transformers_:
        if name == 'nom':
            for i, col in enumerate(cols):
                schema['categorical'][col] = [c for c in trans.categories_[i] if c != 'None']

    schema_path = f'{OUT_DIR}/feature_schema.json'
    with open(schema_path, 'w') as f:
        json.dump(schema, f, indent=2)
    print(f"Saved {schema_path}")

    print("\nDone. Next step: build the browser predictor page against these two files.")


if __name__ == "__main__":
    main()
