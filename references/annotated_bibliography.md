# Annotated bibliography

What each source is doing in the report, one line each, so you can defend every citation in a viva
without having to re-read the paper. Full BibTeX in `references.bib`. Harvard-formatted versions
are generated automatically into the report by `src/generate_report.py`.

## Methodology justifications (cite these when explaining *why* a choice was made)

**Varma and Simon (2006)**, Justifies nested cross-validation. Shows that using the same CV split
to both tune hyperparameters and estimate performance gives an optimistically biased error estimate.
This is the paper to cite the moment anyone asks "why not just do a normal train/test split with
grid search", the answer is that grid search *on* the test set (or on the same CV folds used to
report the score) leaks information about the test set into model selection.

**Breiman (2001)**, Justifies Random Forest as the main model. Bagging + random feature subsampling
reduces variance without increasing bias, which is the mechanism that makes an ensemble outperform
a single tree.

**Breiman et al. (1984)**, Justifies rejecting a standalone Decision Tree. Establishes that unpruned
trees overfit and are high-variance, which is exactly the weakness Breiman (2001) later fixes with
ensembling. Cite these two together: 1984 states the problem, 2001 states the fix.

**Chen and Guestrin (2016)**, Justifies XGBoost's inclusion and its `scale_pos_weight` mechanism for
class imbalance. Also the citation for why gradient boosting is described as state-of-the-art on
tabular data.

**Cortes and Vapnik (1995)**, Justifies SVM as the margin-based comparison model, the original
support-vector classifier paper.

**Ng and Jordan (2001)**, Justifies Logistic Regression as the baseline over Naive Bayes: discriminative
classifiers reach their asymptotic error with fewer examples than generative ones, and with 200,000
training rows here that asymptotic regime is comfortably reached.

**Strobl et al. (2007)**, Justifies using *permutation* importance for Random Forest rather than the
default impurity-based (MDI) importance. MDI is biased toward high-cardinality/continuous features,
which matters here because the feature set mixes one-hot categoricals with continuous numerics.

**McNemar (1947)**, Justifies the significance test used to compare models. McNemar's test is for
paired binary classifiers evaluated on the *same* test set; the Diebold-Mariano test (used in the
BM-forecasting dissertation project, not this one) is for regression and would be the wrong tool here.

**Dietterich (1998)**, Justifies *why* McNemar's test specifically, not just that it exists. Reviews
five candidate significance tests for comparing two classifiers and shows McNemar's test has
acceptably low Type I error for a single train/test split design, which is exactly this project's
setup (one temporal 2025-2026 hold-out, not repeated resampling). Dietterich's own alternative,
5x2cv, is for designs with multiple resampled splits, not applicable here. Cite this alongside
McNemar (1947) if asked "isn't a 1947 test too old", it shows the choice was checked against modern
methodological literature, not just inherited from a textbook.

**Pedregosa et al. (2011)**, The scikit-learn paper. Cite once, generically, for the software stack.

**Beyer et al. (1999)**, Justifies excluding kNN from the model comparison. Shows that as
dimensionality rises, the distance to the nearest neighbour converges toward the distance to the
farthest one, an effect visible from as few as 10-15 dimensions. The feature set here is 51-dimensional
after encoding, well past that threshold, so Euclidean distance carries little discriminative signal
and a kNN classifier would not be a fair comparison. Cite this rather than asserting "the curse of
dimensionality" as a vague truism, it names the specific mechanism and the dimension count where it
starts to bite.

## Domain / dataset sources

**Hardy and Glew (2019)**, The single most important limitations citation. They estimate the true
EPC database error rate at 36-62% once assessor disagreement is accounted for (27% of records carry
at least one error flag outright). This directly undercuts the `WALL_TYPE` feature engineered from
`WALLS_DESCRIPTION` in notebook 02, since that field is exactly the kind of assessor-entered text
they flag as unreliable. Use this to show you understand a real weakness in your own pipeline rather
than just listing "data quality" as a generic limitation.

**Seyedzadeh et al. (2018)** and **Pasichnyi, Wallin and Kordas (2019)**, Two different existing
approaches, not one. Seyedzadeh reviews ML for predicting the *current* rating label; Pasichnyi builds
building archetypes for city-scale energy modelling, a different problem again. Neither targets
retrofit *headroom* directly, which is the novelty citation: it's what you point to when the report
claims the framing here differs from prior work. Note the author count: Pasichnyi has three authors,
so it's "Pasichnyi, Wallin and Kordas", not "Pasichnyi et al." under UWE Harvard (et al. only applies
at 4+ authors).

**Ministry of Housing, Communities and Local Government (2024)**, The dataset source itself
(epc.opendatacommunities.org).

## What's missing (worth adding if you have time before 6 August)

The EPC_PLAN.md publication-route notes mention **Ali et al. / UCD group** on urban retrofit decisions
(Applied Energy) as a citation to verify and add, that one hasn't been confirmed yet, unlike everything
above which has been checked against the actual published paper (title, journal, volume, pages all
verified). Don't cite the Ali et al. one until it's been checked the same way; a wrong or invented
citation is worse than one fewer reference.
