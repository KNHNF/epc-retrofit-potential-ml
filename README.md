# EPC Retrofit Potential

Predicting which UK residential properties have the most retrofit headroom, from open Energy
Performance Certificate (EPC) data.

MSc Data Science, UWE Bristol. Module: Machine Learning and Predictive Analytics.
Author: Karan Homayounfar (25065219).

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

An earlier exploratory pass (`01_EDA.ipynb`) used a 200,000-record training sample and a
50,000-record test sample. The final models train on a larger stratified sample instead, just under
1 million training records and about 200,000 test records, assembled on Kaggle rather than locally
since a nested search across four models at that size needs more memory and CPU time than a laptop
comfortably gives.

The split is temporal, not random. This matters: the positive class rate roughly halves from 21.7%
in training to 10.8% in test, a real distribution shift that a random split would have hidden.

### Getting the data

Raw and processed data are gitignored: the raw files are several GB and carry address-level
records. To reproduce from scratch you need to download them yourself.

1. Register (free) at https://epc.opendatacommunities.org and download the domestic England and
   Wales certificates for 2020 through 2026.
2. Put one CSV per year in `data/raw/`, named `certificates-2020.csv` through
   `certificates-2026.csv`. That naming is what `src/00_prepare_data.py` looks for.
3. Run `python src/00_prepare_data.py`. It reads the raw CSVs in chunks (they do not fit in
   memory), applies the cleaning and eligibility filters, and writes the full train and test
   parquet files into `data/processed/`.

`notebooks/02_Feature_Engineering.ipynb` samples straight from those full parquet files, with the
sample size set by the `N_TRAIN`/`N_TEST` constants at the top of the notebook. The random seed is
fixed at 42 throughout, so the samples and splits are reproducible.

### Running the ~1M-record training at this scale

At close to a million rows, the nested cross-validation search for all four models is too slow for
a normal laptop (Random Forest's search alone took several hours on real cloud hardware). The
notebooks were run on Kaggle instead: `N_TRAIN`/`N_TEST` in `02_Feature_Engineering.ipynb` control
the sample size, and each of `03_Baseline_LR.ipynb`, `04_Random_Forest.ipynb`, `04b_XGBoost.ipynb`,
and `05_Comparison_Evaluation.ipynb` just needs its `DATA_DIR` pointed at wherever the previous
notebook's output was uploaded to.

The chosen hyperparameters barely changed between the 200,000-record and 1-million-record samples,
which is reassuring but only one data point. Also worth knowing: the saved preprocessing pipeline
was built with scikit-learn 1.6.1, and loading it with a different installed version can silently
misbehave or fail outright, `requirements.txt` pins the version for this reason.

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

Two further checks go beyond the single train/test split: a walk-forward evaluation
(`src/walk_forward_evaluation.py`) that moves the training window forward one year at a time and
tests each time on the year right after, and a real-world check on held-out properties in three
cities (`src/bristol_case_study.py`, `src/city_case_studies.py`), not just the aggregate test set.

## Results

| Model | Test ROC-AUC | Test F1-macro | Test PR-AUC |
|---|---|---|---|
| **Random Forest** | **0.9721** | **0.8423** | **0.8097** |
| XGBoost | 0.9703 | 0.8122 | 0.7784 |
| SVM | 0.9632 | 0.8137 | 0.7753 |
| Logistic Regression | 0.9639 | 0.7645 | 0.7772 |

Random Forest wins on every test-set metric, despite XGBoost scoring higher in cross-validation
(0.9881 vs 0.9874 CV ROC-AUC). The model that tuned best on the 2020-2024 training distribution was
not the model that generalised best to the genuinely unseen 2025-2026 period, a finding that only
shows up because of the temporal (not random) test split.

The margin over XGBoost is narrow, small enough to ask whether one test sample got lucky.
Resampling the test set 2,000 times, the gap holds in 100% of resamples, so the ordering is real
rather than noise.

### What it is worth

The register carries current and potential heating cost and CO2 per certificate, so the saving is in
the data rather than assumed. Surveying 10% of the eligible test stock (7,131 homes):

| Strategy | Genuine cases found | Annual heating cost | Annual CO2 |
|---|---|---|---|
| Ranked by model | 6,623 | GBP 5.4m | 26,232.8 t |
| Rating band only | 2,127 | GBP 2.9m | 11,457.0 t |

That is 1.85 times the genuine cases found for the same number of surveys. These figures assume the
full recommended package is installed, so they are an upper bound on what the ranking buys, not a
forecast of delivered savings.

### Where it does not work

Aggregate metrics hid a real gap. Recall is 0.84 on houses but 0.36 on flats, so the model missed
most high-potential flats. Precision on flats is *higher* (0.81), which rules out confusion: the
model is not wrong about flats, it is too cautious, having learned their much lower positive rate
than houses.

That is a threshold problem rather than a model problem. Equalising recall across groups with a
per-group threshold is the equality-of-opportunity criterion (Hardt, Price and Srebro, 2016).
Dropping the flat threshold to 0.22 lifts recall to 0.84, level with houses, at a fair precision
cost, F1 rises overall. Checked two ways: tuned on one random half of the flats and scored on the
other, and, more strictly, tuned on 2025's flats and scored on 2026's, the genuinely later period.
Both hold. The second check is what actually confirms the fix survives time passing, not just a
random split.

Worked through step by step in `05_Comparison_Evaluation.ipynb` and `src/impact_and_fairness.py`.

## How to run

```
pip install -r requirements.txt
python src/00_prepare_data.py        # builds data/processed/ from raw EPC CSVs
```

Then run the notebooks in order (on Kaggle or similar for the full ~1M-record sample, see above):

1. `01_EDA.ipynb`, exploratory analysis, class balance, missingness, correlations
2. `02_Feature_Engineering.ipynb`, sklearn preprocessing pipeline, temporal train/test split
3. `03_Baseline_LR.ipynb`, Logistic Regression baseline
4. `04_Random_Forest.ipynb`, Random Forest, main model
5. `04b_XGBoost.ipynb`, XGBoost comparison
6. `05_Comparison_Evaluation.ipynb`, SVM, four-way comparison, McNemar tests, calibration,
   subgroup recall and the flat threshold fix

Then build the figures and the report:

```
python src/bristol_case_study.py       # per-district Bristol predictions
python src/city_case_studies.py        # same check on Manchester and Leeds
python src/impact_and_fairness.py      # cost/CO2 impact, subgroup recall, flat threshold
python src/bootstrap_ci.py             # confidence intervals on the model gap
python src/walk_forward_evaluation.py  # expanding-window check across years
python src/make_bristol_map.py         # Fig. 9, district map
python src/make_walk_forward_figure.py # Fig. 10, walk-forward chart
python src/make_table_images.py        # table images for the two-column build
python src/generate_report.py human    # builds the .docx from the saved metrics
pytest tests/
```

The report generator does no modelling. It reads `data/processed/model_comparison.csv` and the
saved `.pkl`/`.csv` metrics and writes the numbers into the prose, so the document cannot state a
figure that disagrees with the run that produced it. Change the data, rebuild, every number
updates.

Note (Windows): scikit-learn's `n_jobs=-1` can crash with a joblib `PicklingError` on this platform
when nested inside another parallel call, or when there isn't enough free disk space for its temp
memmap files. On Kaggle's Linux containers the same class of problem shows up differently, as a
"No space left on device" error from joblib memmapping large arrays across worker processes; drop
to `n_jobs=1` if that happens. See `docs/IMPLEMENTATION_NOTES.md` for more on decisions like this.

## Structure

```
02_Main_Project/
├── data/
│   ├── raw/                          EPC CSVs, gitignored (too large, contains address data)
│   ├── external/                     Small reference data that is tracked (postcode boundaries)
│   └── processed/                    Cleaned/engineered data, model outputs, gitignored
├── notebooks/                        01 through 05, run in order
├── src/
│   ├── 00_prepare_data.py            Raw CSVs to train/test parquet files
│   ├── data_loader.py
│   ├── preprocessing.py
│   ├── evaluation.py
│   ├── bristol_case_study.py         Per-district Bristol predictions and z-test
│   ├── city_case_studies.py          Same held-out check on Manchester and Leeds
│   ├── impact_and_fairness.py        Cost/CO2 impact, subgroup recall, flat threshold (both checks)
│   ├── bootstrap_ci.py               Confidence intervals on the model gap
│   ├── walk_forward_evaluation.py    Expanding-window check across years
│   ├── make_bristol_map.py           Fig. 9, real district-boundary choropleth
│   ├── make_walk_forward_figure.py   Fig. 10, walk-forward chart
│   ├── make_table_images.py          Tables as images, for the two-column build
│   ├── two_column_layout.py          Floating figures, endnotes, IEEE-style columns, hyperlinks
│   ├── regen_correlation_figure.py   Re-renders Fig. 3 with readable labels
│   └── generate_report.py            Builds the .docx from saved metrics
├── docs/
│   └── IMPLEMENTATION_NOTES.md       The "why" behind non-obvious decisions in src/
├── tests/
├── report/
│   ├── figures/
│   └── *.docx                        Generated, gitignored
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
