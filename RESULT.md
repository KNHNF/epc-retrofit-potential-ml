# RESULT: epc-retrofit-potential-ml review, cloud batch 2

Branch: `cloud/epc-retrofit-potential-ml-2026-11-02`. Nothing merged, no pull request opened, nothing published.
Scope kept to: no report prose, no outcome values (none are added or restated here), nothing merged before 7 Jan.
I read "no src/207" as not editing `src/walk_forward_evaluation.py` (207 lines). I read it but did not change it. If you meant something else, tell me.

## What I read

README.md, requirements.txt, .gitignore, cleanup_data.py, src/00_prepare_data.py, src/data_loader.py, src/preprocessing.py, src/evaluation.py, src/walk_forward_evaluation.py, src/impact_and_fairness.py, both original test files, and every code cell in notebooks 02, 03, 04, 04b and 05.

I did not read: src/generate_report.py (2,071 lines), src/two_column_layout.py, src/make_*.py, src/bristol_case_study.py, src/city_case_studies.py, src/bootstrap_ci.py, notebook 01, docs/ and references/. The report-building code and the case-study scripts are therefore unreviewed.

## What changed

Bug fixes, each with a test that fails on the old code (13 of the new tests fail against the old `src/`, confirmed by stashing it):

1. `src/00_prepare_data.py`, within-year deduplication depended on chunk order. A UPRN was added to the seen set after its first chunk, so a newer certificate in a later chunk of the same file was thrown away and the older one kept. Now the seen set is frozen for the file and updated once at the end, so the newest certificate per UPRN wins.
2. `src/00_prepare_data.py`, a file missing a required column was skipped chunk by chunk with a print, then the script reported "Data preparation complete" on an empty result. It now raises a `ValueError` naming the file and columns.
3. `src/00_prepare_data.py`, with no `certificates-<year>.csv` in `data/raw/` the script printed warnings and finished as if successful. It now raises `FileNotFoundError` pointing at the README.
4. `src/00_prepare_data.py`, `data/processed/` is gitignored, so on a fresh clone writing the parquet files failed. It is now created. (Found by the new end-to-end test, not by reading.)
5. `src/preprocessing.py`, `drop_duplicates_by_uprn` collapsed every row with a missing UPRN into one row. Rows with no UPRN are now kept.
6. `src/preprocessing.py`, missing columns in `filter_valid_records`, `encode_age_band`, `extract_wall_type` and `engineer_target` gave a bare `KeyError`. They now raise a `ValueError` naming the columns.
7. `src/evaluation.py`, `evaluate_model` did not return `y_test`, so `plot_roc_curves` (which reads `r['y_test']`) raised `KeyError` on its output. It is now returned.
8. `src/evaluation.py`, `evaluate_model` reported only weighted F1 although the README says F1-macro throughout. It now also returns `f1_macro` (`f1_weighted` kept).
9. `src/evaluation.py`, a model with no `predict_proba` (for example LinearSVC) silently got no ROC-AUC. It now falls back to `decision_function`. A ROC-AUC of exactly 0.0 is no longer skipped by a truthiness check.
10. `src/evaluation.py`, `plot_feature_importances` silently mislabelled bars (or hit an `IndexError`) when the name list and importances differed in length. It now raises a `ValueError`.
11. `notebooks/05_Comparison_Evaluation.ipynb`, cells 20 and 22 had `\n` inside single-quoted f-strings stored as a real line break, so both cells were a `SyntaxError`. Fixed with a byte-level edit of those two lines (CRLF kept). All code cells in all notebooks now parse. The notebook was not re-run.
12. `requirements.txt`, `python-docx` was imported by `src/generate_report.py` but not listed, so a clean install could not build the report. Added `python-docx>=1.0` (the only new dependency; reason is the missing declaration, not a new feature).

Added:
- `tests/test_prepare_data.py`, `tests/test_data_loader.py`, `tests/test_offline_demo.py`, and extra cases in the two existing test files.
- `examples/offline_demo.py`, a synthetic-data run of the method (see the README check below).
- README: install and offline-example sections, a Limits section, a caution on the temporal shift, a note that the helpers are not what the notebooks run, and the Structure tree fixed (it said `02_Main_Project/`).
- Docstring notes in `src/preprocessing.py` where behaviour is easy to misuse.

Also noticed and left alone: `engineer_target` labels a row with a NaN gap as 0 (the prepare script drops such rows first, so it does not bite the pipeline). Documented in the docstring.

## Tests, real output

Default sandbox environment (Python 3.13.18, scikit-learn 1.9.1, pandas 3.0.5):

```
platform linux -- Python 3.13.18, pytest-9.2.0, pluggy-1.7.0 -- /usr/local/bin/python
... 39 PASSED lines (one per test, listed in tests/) ...
39 passed in 2.78s
```

Fresh virtualenv from `requirements.txt` as pinned (scikit-learn 1.6.1, pandas 3.1.0, numpy 2.5.4, xgboost 3.4.1), then `python-docx` added:

```
39 passed, 32 warnings in 2.57s
```

Before my changes, the original 13 tests passed (13 passed in 3.02s). With only `src/` reverted to the old code, the new suite gave `13 failed, 24 passed`.

The 32 warnings in the pinned environment are scikit-learn 1.6.1 against a current SciPy: `OptimizeWarning: Unknown solver options: iprint` from the lbfgs solver. Harmless noise here, but `requirements.txt` does not pin SciPy, so it will show up on a fresh install. I did not pin SciPy because that is your call.

`python examples/offline_demo.py` runs in about 2 seconds and prints structure checks only. It passed in both environments.

## Temporal holdout and nested CV: implemented as the README says?

Temporal holdout: yes in code. `00_prepare_data.py` assigns 2020 to 2024 files to train and 2025 and 2026 to test, processes years newest first, and drops a UPRN from an older file if a newer file has it. My end-to-end test confirms no UPRN is in both sets, every train date is before every test date, and the later certificate wins. This rests on the assumption that `certificates-YYYY.csv` holds certificates lodged in year YYYY. I could not check that without the data.

Nested CV: yes in notebooks 03, 04, 04b and 05. Outer `StratifiedKFold(5)`, inner `GridSearchCV` with `StratifiedKFold(3)`, scored on F1-macro, final hyperparameters taken as the most frequently chosen across outer folds, final model refit on all training rows and scored once on the temporal test set. The folds are random, not temporal. The README says so ("its folds all come from the same period"), so that is honest. The notebook 02 comment "for temporal CV in model notebooks" is stale: `y_train_years` is saved and never used.

Leakage and selection risks (none fixed, all now noted in the README Limits):
- Preprocessing is fitted on the whole training set before the CV loops (notebook 02, then 03/04/04b/05 split the transformed array). Medians, scaling and encoder categories therefore see outer validation rows. Likely small, but unmeasured. The offline demo shows the in-fold pattern.
- Dedup keeps only the latest certificate per property. Earlier years then contain only properties never re-certified later, so a home with a high-potential 2022 certificate that was retrofitted and re-certified in 2025 appears only as its (low-potential) 2025 certificate in the test set. That could contribute to the positive-rate drop the README calls "a real distribution shift". I have not tested it. The same selection affects the walk-forward folds.
- Rows with a missing UPRN are dropped in the prepare script. The share is unknown to me.
- The test set is reused for McNemar, calibration, subgroup analysis and choosing the flat threshold. The two threshold checks split the flats, but all draw on 2025-2026.
- The final model choice ("Random Forest is the main model") is not shown to be fixed before looking at test results; I cannot tell from the repo.
- Features: I found no potential-score or potential-cost column in the model inputs in notebook 02 or the walk-forward script (only current values). The target depends on `CURRENT_ENERGY_EFFICIENCY` and the rating band, which are both inputs. That is legitimate (known at certification) but makes the task partly a function of two inputs; notebook 05 has a naive baseline on exactly those two columns, which is the right check.

## README check

- Says what it does: yes.
- How to install: was a single `pip install` line with no Python version, no virtualenv, and no mention of the missing `python-docx`. Now fixed and checked in a clean virtualenv.
- One runnable offline example: there was none (everything needed several GB of data or Kaggle). Now `examples/offline_demo.py` plus `pytest tests/`, both checked.
- Honest limits: partly (Windows and Kaggle notes, one-data-point hyperparameter remark) but nothing about reproducibility, reuse of the test set, preprocessing before CV, or the dedup selection effect. Added.
- Claims I could not confirm and left as written: the Results tables, "Random Forest wins on every test-set metric", the cross-validation comparison versus XGBoost, "100% of resamples", the impact table, the recall figures for flats and houses, and the 7.25M and 2.47M row counts in notebook 02. They come from Kaggle runs whose outputs are not in the repo.

## What is missing before it could be called portfolio-ready

1. Reproducibility of the headline numbers: the metric files or a small, publishable results file (not address-level data) are not in the repo, and the notebook outputs were not checked against the README. A reviewer cannot confirm any figure.
2. The tests cover helpers and the prepare script, not the code that produced the results. The notebooks do not import `src/preprocessing.py` or `src/evaluation.py`. The pipeline lives as copies inside notebooks 02 to 05 and `walk_forward_evaluation.py`, which each re-declare the same column lists. Moving that into `src/` and importing it everywhere would remove the drift risk.
3. Notebook 05 has a hard-coded `/kaggle/...` data path, and nothing ties the notebook order together (no Makefile or runner script).
4. No continuous integration. A single GitHub Actions job running `pytest tests/` offline would do, but I did not add one (it is a repo setting change; your call).
5. Dependency hygiene: only `scikit-learn` is pinned, SciPy and others are not, and the pinned scikit-learn plus current SciPy gives the warning above. A lockfile or a tested set of versions is needed.
6. `cleanup_data.py` at the repo root is a one-off script that deletes files. It asks for YES first, but it looks odd in a portfolio repo. Moving it out or deleting it is a decision for you.
7. The dedup selection effect and preprocessing-before-CV questions above should be tested once, on the laptop, and the answer written into the README either way.
8. `src/generate_report.py` is 2,071 lines of document building and the report is an MSc submission. If this is shown publicly, the split between "project" and "coursework" needs thought. I did not review or change it.
9. Model card style content is absent: intended use, who should not rely on it, data licence and attribution line for the EPC open data. The README links the source but I did not check licence wording.
10. Unfixed code smells: `use_label_encoder=False` is passed to `XGBClassifier` in 04b (deprecated; worth checking against the installed version on the laptop), ties in "most frequent hyperparameter" are resolved arbitrarily, and `load_epc_csv` reads the whole file before sampling.

## What I could not check

- Anything needing the real EPC data or the Kaggle outputs: the age band label coverage for 2025-2026 certificates (unmapped labels would silently fall back to the median in `AGE_BAND_ORDER`; notebook 02 does not use age as a feature, so this affects only the `src/preprocessing.py` helper and the fairness slice by age band), file-year versus lodgement-year consistency, the share of dropped rows, and every reported metric.
- That the notebooks run end to end. I only confirmed all code cells parse.
- `src/generate_report.py` and the figure and case-study scripts.
- Windows behaviour (the joblib notes in the README).

## Unverified claims (in the repo, not checked by me)

- All Results, impact and fairness figures in the README.
- The statement that the hyperparameters "barely changed" between the 200,000 and roughly 1 million record runs.
- The joblib and Windows crash claims in code comments and the README (`docs/IMPLEMENTATION_NOTES.md` not read at all).
- The bibliography claims in `references/` (not read).

## What needs the laptop

1. Re-run `00_prepare_data.py` on the raw CSVs after my fixes and compare row counts with the existing parquet files. The within-year dedup fix can change which certificate is kept for properties with several certificates in the same year, so the processed data, and therefore every downstream figure, may shift slightly. Decide whether to regenerate the Kaggle outputs.
2. Run notebook 05 cells 20 and 22 to confirm the syntax fix behaves.
3. Measure the two selection and leakage questions: positive rate by lodgement year without the "latest certificate only" rule, and CV scores with preprocessing inside the folds.
4. Count rows dropped for missing UPRN and check which age band labels appear in 2025 and 2026.
5. Decide on SciPy pinning, CI, `cleanup_data.py`, and whether `src/walk_forward_evaluation.py` should be touched.
