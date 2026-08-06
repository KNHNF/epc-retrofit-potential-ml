# EPC Retrofit Potential

Predicting which UK residential properties have the most retrofit headroom, from open Energy
Performance Certificate (EPC) data.

MSc Data Science, UWE Bristol. Module: Machine Learning and Predictive Analytics.
Author: Karan Homayounfar (25065219). Coursework deadline: 6 August 2026.

## Problem

The UK has a legally binding net-zero target for 2050. In the certificates used here, 43% of homes
are rated D or below on the EPC energy efficiency scale, and retrofitting all of them at once isn't
financially or logistically possible. The question that matters for policy is which properties will
deliver the most energy savings per pound of public investment.

Most EPC machine learning work predicts the current energy rating label (A-G). That framing is
limited for prioritisation: a property already rated C has little room left to improve regardless of
its label. This project instead predicts retrofit headroom directly: a property is labelled high
potential if it is currently rated D-G and has a gap of 20 or more points between its current and
potential EPC efficiency score.

## Dataset

UK EPC open data, published by the Ministry of Housing, Communities and Local Government
(https://epc.opendatacommunities.org). Certificates from 2020 to 2026 for England and Wales.

Training set: 200,000 records stratified from certificates lodged 2020-2024.
Test set: 50,000 records from certificates lodged 2025-2026.

The split is temporal, not random. This matters: the positive class rate shifts from 21.7% in
training to 10.8% in test, a real distribution shift that a random split would have hidden. Raw
and processed data are gitignored (too large for git, and the raw files contain address data);
regenerate them with `src/00_prepare_data.py`.

## Method

Four classifiers, compared on the same temporal split with the same evaluation protocol:

| Model | Role |
|---|---|
| Logistic Regression | Baseline |
| Random Forest | Main model |
| XGBoost | Gradient boosting comparison |
| SVM (LinearSVC + Platt scaling) | Margin-based comparison |

A standalone Decision Tree was considered and rejected: Breiman et al. (1984) established that
unpruned trees overfit severely, which is the exact weakness Random Forest's ensembling addresses.
See `references/annotated_bibliography.md` for the full justification of every model choice.

Evaluation: nested cross-validation (5-fold outer, 3-fold inner) for hyperparameter selection,
avoiding the optimistic bias of tuning and scoring on the same split (Varma and Simon, 2006), with a
final evaluation on the temporally held-out test set to measure genuine forward-in-time
generalisation, something nested CV alone can't test since its folds all come from the same period.
Metrics: F1-macro and ROC-AUC throughout (not accuracy, given the class imbalance), plus McNemar's
test for pairwise significance and calibration curves for threshold selection.

## Results

| Model | Test ROC-AUC | Test F1-macro | Test PR-AUC |
|---|---|---|---|
| **Random Forest** | **0.9705** | **0.8262** | **0.8051** |
| XGBoost | 0.9694 | 0.8151 | 0.7815 |
| SVM | 0.9629 | 0.8151 | 0.7721 |
| Logistic Regression | 0.9635 | 0.7666 | 0.7729 |

Random Forest wins on every test-set metric, despite XGBoost scoring higher in cross-validation
(0.9873 vs 0.9858 CV ROC-AUC). The model that tuned best on the 2020-2024 training distribution was
not the model that generalised best to the genuinely unseen 2025-2026 period, a finding that only
shows up because of the temporal (not random) test split.

## How to run

```
pip install -r requirements.txt
python src/00_prepare_data.py        # builds data/processed/ from raw EPC CSVs
```

Then run the notebooks in order:

1. `01_EDA.ipynb`, exploratory analysis, class balance, missingness, correlations
2. `02_Feature_Engineering.ipynb`, sklearn preprocessing pipeline, temporal train/test split
3. `03_Baseline_LR.ipynb`, Logistic Regression baseline
4. `04_Random_Forest.ipynb`, Random Forest, main model
5. `04b_XGBoost.ipynb`, XGBoost comparison
6. `05_Comparison_Evaluation.ipynb`, SVM, four-way comparison, McNemar tests, calibration

Then build the figures and the report:

```
python src/bristol_case_study.py     # per-district Bristol predictions
python src/make_bristol_map.py       # Fig. 9, district map
python src/make_table_images.py      # table images for the two-column build
python src/generate_report.py human  # builds the .docx from the saved metrics
pytest tests/
```

The report generator does no modelling. It reads `data/processed/model_comparison.csv` and the
saved `.pkl` metrics and writes the numbers into the prose, so the document cannot state a figure
that disagrees with the run that produced it. Change the data, rebuild, every number updates.

Note (Windows): scikit-learn's `n_jobs=-1` can crash with a joblib `PicklingError` on this platform
when nested inside another parallel call, or when there isn't enough free disk space for its temp
memmap files. `04_Random_Forest.ipynb` runs single-threaded (`n_jobs=1`) for this reason.

## Structure

```
02_Main_Project/
├── data/
│   ├── raw/                   EPC CSVs, gitignored (too large, contains address data)
│   └── processed/             Cleaned/engineered data, model outputs, gitignored
├── notebooks/                 01 through 05, run in order
├── src/
│   ├── 00_prepare_data.py     Raw CSVs to train/test parquet samples
│   ├── data_loader.py
│   ├── preprocessing.py
│   ├── evaluation.py
│   ├── bristol_case_study.py  Per-district Bristol predictions and z-test
│   ├── make_bristol_map.py    Fig. 9, district map on an OpenStreetMap basemap
│   ├── make_table_images.py   Tables as images, for the two-column build
│   ├── two_column_layout.py   Floating figures, endnotes, IEEE-style columns
│   ├── export_onnx.py         Model export
│   └── generate_report.py     Builds the .docx from saved metrics
├── tests/
├── report/
│   ├── figures/
│   ├── REPORT_OVERVIEW.md     Section plan against the marking criteria
│   └── *.docx                 Generated, gitignored
├── references/
│   ├── references.bib
│   └── annotated_bibliography.md
├── requirements.txt
└── .gitignore
```

## References

Full list with UWE Bristol Harvard formatting in `references/references.bib` and
`references/annotated_bibliography.md` (which also explains why each source is cited, not just
what it says).
