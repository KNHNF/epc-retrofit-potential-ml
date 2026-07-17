"""
generate_report.py
Generates the Word document (decimal-numbered sections) for the EPC retrofit prediction coursework.
Run AFTER all notebooks have been executed and model_comparison.csv exists.

Two output modes, both built from the same underlying content:
  full  -> report/25065219_report.docx        (quality-first, word count not constrained)
  safe  -> report/25065219_report_safe.docx    (<=2200 words body, 2000-word brief + 10% tolerance,
                                                 references/title/abstract excluded from the count
                                                 per the brief and standard academic convention)

Usage: python src/generate_report.py [full|safe|both]   (default: both)
"""

import os
import sys
import csv
import numpy as np

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATA_DIR    = "data/processed"
FIGURES_DIR = "report/figures"

os.makedirs("report", exist_ok=True)


def add_figure(doc, path, caption, width=5.5):
    """Add a figure with a caption."""
    if os.path.exists(path):
        doc.add_picture(path, width=Inches(width))
        last = doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cap.runs[0]
        run.italic = True
        run.font.size = Pt(10)
    else:
        p = doc.add_paragraph(f"[FIGURE: {caption}. Run notebooks to generate.]")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def load_comparison_csv():
    path = f"{DATA_DIR}/model_comparison.csv"
    if not os.path.exists(path):
        return None
    rows = []
    with open(path, newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def load_naive_baseline():
    path = f"{DATA_DIR}/naive_baseline_metrics.csv"
    if not os.path.exists(path):
        return None
    with open(path, newline='') as f:
        reader = csv.DictReader(f)
        row = next(reader, None)
    if row is None:
        return None
    return {k: (v if k == 'model' else float(v)) for k, v in row.items()}


def shade_cell(cell, hex_color):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def add_comparison_table(doc, rows, winners):
    if rows is None:
        doc.add_paragraph("[TABLE: Model comparison. Run notebooks first.]")
        return

    best_model = winners['best_auc_model'] if winners else None
    headers = list(rows[0].keys())
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.style = 'Table Grid'
    table.autofit = True

    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        shade_cell(cell, '2F2F2F')
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for i, row in enumerate(rows):
        is_winner = row.get('Model') == best_model
        for j, key in enumerate(headers):
            cell = table.cell(i + 1, j)
            cell.text = str(row[key])
            if is_winner:
                shade_cell(cell, 'E8F0E3')
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.size = Pt(10)
                    if is_winner:
                        run.bold = True


def determine_winners(comparison_rows):
    if not comparison_rows:
        return None
    ranked_auc = sorted(comparison_rows, key=lambda r: float(r['Test ROC-AUC']), reverse=True)
    ranked_f1  = sorted(comparison_rows, key=lambda r: float(r['Test F1-macro']), reverse=True)
    return {
        'ranked_auc': [{'model': r['Model'], 'test_roc_auc': float(r['Test ROC-AUC'])} for r in ranked_auc],
        'ranked_f1':  [{'model': r['Model'], 'test_f1_macro': float(r['Test F1-macro'])} for r in ranked_f1],
        'best_auc_model': ranked_auc[0]['Model'],
        'best_f1_model':  ranked_f1[0]['Model'],
        'n_models': len(comparison_rows),
    }


def set_style_font(style, name, size=None, color=None):
    style.font.name = name
    rpr = style.element.get_or_add_rPr()
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rpr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), name)
    if size is not None:
        style.font.size = size
    if color is not None:
        style.font.color.rgb = color


def add_reference(doc, before, italic, after):
    p = doc.add_paragraph(style='List Paragraph')
    p.paragraph_format.first_line_indent = Pt(-18)
    r1 = p.add_run(before)
    r1.font.size = Pt(10)
    r2 = p.add_run(italic)
    r2.italic = True
    r2.font.size = Pt(10)
    r3 = p.add_run(after)
    r3.font.size = Pt(10)
    return p


def build_report(mode="full"):
    """mode: 'full' (quality-first) or 'safe' (<=2200 words body, brief's 2000-word
    limit plus a 10% tolerance margin, references/title/abstract excluded)."""
    safe = (mode == "safe")
    OUT_PATH = f"report/25065219_report{'_safe' if safe else ''}.docx"

    doc = Document()

    set_style_font(doc.styles['Normal'], 'Times New Roman', Pt(12))
    set_style_font(doc.styles['Heading 1'], 'Times New Roman', Pt(15), RGBColor(0, 0, 0))
    set_style_font(doc.styles['Heading 2'], 'Times New Roman', Pt(13), RGBColor(0, 0, 0))
    doc.styles['Heading 1'].font.bold = True
    doc.styles['Heading 2'].font.bold = True

    comparison_rows = load_comparison_csv()
    winners = determine_winners(comparison_rows)

    # ------------------------------------------------------------------
    # IEEE-style header
    # ------------------------------------------------------------------
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Predicting Retrofit Potential in UK Residential Buildings:\nA Machine Learning Classification Approach")
    run.bold = True
    run.font.size = Pt(17)

    authors = doc.add_paragraph()
    authors.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = authors.add_run(
        "Karan Homayounfar (25065219)\n"
        "MSc Data Science, UWE Bristol\n"
        "Machine Learning and Predictive Analytics\n"
        "Code: https://github.com/KNHNF/epc-retrofit-potential-ml"
    )
    run.font.size = Pt(11)

    doc.add_paragraph()

    # ------------------------------------------------------------------
    # Abstract
    # ------------------------------------------------------------------
    doc.add_heading("Abstract", level=2)
    winner_sentence = (
        f"{winners['best_auc_model']} achieves the highest test-set ROC-AUC of the models compared."
        if winners else
        "[RESULT: run all notebooks so the winning model can be stated correctly here.]"
    )
    abs_text = doc.add_paragraph(
        "The UK's net-zero-by-2050 target means public money for retrofit has to be aimed, not "
        "sprayed evenly across the housing stock. Most machine learning work on Energy Performance "
        "Certificate (EPC) data predicts a property's current rating (A-G). I predict something more "
        "useful for that aiming problem: retrofit headroom, whether a home rated D-G has a 20-point "
        "or greater gap between its current and potential efficiency score. I compare four "
        "classifiers under nested cross-validation on 200,000 records from 2020 to 2024, then test "
        "them on held-out 2025 to 2026 data. "
        f"{winner_sentence} "
        "The positive rate also halves across that boundary, from 21.7 to 10.8 per cent, which is "
        "why the split is by time and not at random."
    )
    abs_text.style.font.size = Pt(11)

    doc.add_paragraph()

    # ------------------------------------------------------------------
    # 1. Introduction
    # ------------------------------------------------------------------
    doc.add_heading("1. Introduction", level=1)
    p_intro1 = doc.add_paragraph()
    p_intro1.add_run(
        "The UK government is legally committed to net-zero greenhouse gas emissions by 2050. In the "
        "certificates I examine, 43.1 per cent of homes are rated D or below (Section 2), and "
        "retrofitting all of them at once is neither financially nor logistically realistic "
        "[1]. So the question that matters is which homes return the most energy saved per "
        "pound of public money. Existing EPC machine learning mostly answers a different question. "
        "Some of it predicts the current rating label [2]; other work builds archetypes for "
        "city-scale energy models rather than "
        "property-level decisions [3]. Neither targets improvement "
        "headroom. Predicting the rating label does not help much here, because a home already rated "
        "C has little room to improve whatever its label. Headroom is the thing worth flagging, the "
        "gap between a property's current and potential EPC score."
    )
    doc.add_paragraph(
        "I frame retrofit prioritisation as binary classification on the UK EPC open data. Four "
        "algorithms compete: logistic regression as the baseline, random forest as the main model, "
        "XGBoost as a strong tabular benchmark, and an SVM for comparison. Nested cross-validation "
        "keeps the performance estimates honest rather than optimistic. Final evaluation is on a "
        "time-held-out test set, the condition the model would actually face in use. "
        "[your take: one or two sentences on why this problem pulled you in, what made headroom more "
        "interesting to you than the usual rating prediction. Say it the way you would to a coursemate.]"
    )

    # ------------------------------------------------------------------
    # 2. Dataset
    # ------------------------------------------------------------------
    doc.add_heading("2. Dataset", level=1)
    doc.add_paragraph(
        "The dataset is the UK EPC Open Data published by the Ministry of Housing, Communities "
        "and Local Government [1], covering domestic energy performance certificates "
        "for England and Wales. The full dataset comprises approximately 10.8 million domestic EPC "
        "certificates across annual files from 2020 to 2026. After deduplication and filtering to "
        "properties eligible for the target definition, this yields a usable pool of 7.25 million "
        "training-eligible certificates (2020-2024) and 2.47 million test-eligible certificates "
        "(2025-2026), from which the 200,000 and 50,000-record stratified samples described below "
        "are drawn. Each certificate records the energy assessment of a single domestic property, "
        "including construction characteristics, insulation quality, heating system, and both the "
        "measured and potential efficiency scores on a 1-100 scale."
    )
    doc.add_paragraph(
        "I draw a 200,000-record stratified training sample from certificates lodged in 2020-2024, "
        "and a separate 50,000-record test sample from 2025-2026. Splitting by time, rather than at "
        "random, stops the model learning from properties assessed after the ones it predicts, which "
        "is exactly the situation it would face in deployment. Where a property has been assessed more "
        "than once (matched on UPRN), I keep only the most recent certificate."
    )
    doc.add_paragraph(
        "The target is 1 when a property is rated D, E, F, or G and the gap between its potential and "
        "current efficiency score is at least 20 points, roughly two rating bands of improvement. "
        "That gives 21.7 per cent positive in training and 10.8 per cent in test. The halving across "
        "the split is worth pausing on. It is not noise: newer certificates cover homes with less "
        "headroom left, probably because more of them are new builds or recent refurbishments, and it "
        "is the clearest single reason a random split would have flattered the results."
    )
    add_figure(doc,
        f"{FIGURES_DIR}/01_class_balance.png",
        "Fig. 1. Target class distribution in training (left) and test (right) sets, "
        "illustrating the temporal distribution shift."
    )

    # ------------------------------------------------------------------
    # 2.1 Exploratory Data Analysis (C1)
    # ------------------------------------------------------------------
    doc.add_heading("2.1 Exploratory Data Analysis", level=2)
    if safe:
        doc.add_paragraph(
            "Thirteen of the 41 raw fields carry missing values (Fig. 2): FLOOR_ENERGY_EFF is "
            "missing in 89.3 per cent of records and is dropped; a second cluster is missing in "
            "14.7 per cent and is median/mode-imputed. A small fraction of records (~0.1 per cent) "
            "carry physically impossible negative values in energy, emissions, or cost fields; "
            "these are sentinel or data-entry errors and are removed. "
            "Construction age band shows a clear "
            "relationship with the target (Fig. 3): pre-1966 properties carry the highest "
            "retrofit-potential rate, consistent with Section 5.3's permutation importances."
        )
    else:
        doc.add_paragraph(
            "Thirteen of the 41 raw fields carry missing values (Fig. 2), concentrated in "
            "component-level ratings rather than spread evenly across the dataset. FLOOR_ENERGY_EFF "
            "is missing in 89.3 per cent of records, too sparse to impute reliably, and is dropped "
            "from the feature set entirely rather than filled with a misleading default. "
            "ROOF_ENERGY_EFF is missing in 20.6 per cent of records and is retained with an "
            "'unknown' category. A second cluster of fields, NUMBER_HABITABLE_ROOMS, "
            "MAINS_GAS_FLAG, EXTENSION_COUNT, and NUMBER_HEATED_ROOMS, is each missing in 14.7 per "
            "cent of records; these are median-imputed (numeric) or mode-imputed (categorical), "
            "since the missingness rate is low enough that imputation does not dominate the "
            "resulting distribution. TOTAL_FLOOR_AREA contains extreme outliers consistent with "
            "data entry errors (values of 0 or in the thousands of square metres for a domestic "
            "property), handled by capping at the 99th percentile before scaling. A further "
            "quality check removes the small fraction of records (~0.1 per cent) carrying "
            "physically impossible negative values in energy consumption, CO2 emissions, or "
            "cost fields; these are sentinel or data-entry errors rather than real measurements, "
            "so the affected records are dropped, consistent with the efficiency range filter."
        )
        doc.add_paragraph(
            "Construction age band shows a clear relationship with the target (Fig. 3): pre-1966 "
            "properties, built before modern insulation standards became common, carry the highest "
            "retrofit-potential rate of any age band. This is a useful sanity check on the target "
            "definition itself, not just a modelling input: it confirms the 20-point efficiency-gap "
            "threshold is picking out properties that match independent domain expectations about "
            "where retrofit headroom actually sits, and it is consistent with construction age band "
            "ranking highly in the Random Forest permutation importances reported in Section 5.3."
        )
    add_figure(doc,
        f"{FIGURES_DIR}/02_missing_values.png",
        "Fig. 2. Missing value rate by column, training sample (200,000 records)."
    )
    add_figure(doc,
        f"{FIGURES_DIR}/06_retrofit_rate_by_age.png",
        "Fig. 3. Retrofit potential rate by construction age band, training sample."
    )

    # ------------------------------------------------------------------
    # 3. Problem Definition
    # ------------------------------------------------------------------
    doc.add_heading("3. Problem Definition", level=1)
    doc.add_paragraph(
        "This is a multivariate supervised binary classification problem. The feature set includes "
        "numerical measurements (current energy efficiency score, floor area, CO2 emissions per "
        "floor area, heating and hot-water costs, room counts), ordinal variables (current energy "
        "rating encoded A=6 to G=0; component efficiency ratings from Very Good=5 to N/A=0), "
        "and nominal variables (property type, built form, wall type, tenure, mains gas flag). "
        "All variables are observable at the time of EPC assessment, and none are derived from "
        "the target."
    )
    doc.add_paragraph(
        "The potential energy efficiency score and potential energy rating are excluded "
        "from the feature set. Including either would be target leakage: the efficiency gap "
        "is computed directly from potential efficiency, so a model with access to it would achieve "
        "trivially perfect performance rather than learning from physical property characteristics. "
        "Current energy efficiency is retained as a feature because it is fully observable and "
        "does not directly encode the potential score."
    )

    # ------------------------------------------------------------------
    # 4. Algorithm Selection and Methodology
    # ------------------------------------------------------------------
    doc.add_heading("4. Algorithm Selection and Methodology", level=1)
    p_impl = doc.add_paragraph()
    if safe:
        p_impl.add_run("Models are implemented in Python via scikit-learn [4].")
    else:
        p_impl.add_run(
            "All models are implemented in Python using scikit-learn [4], with XGBoost's own "
            "library providing the gradient-boosting implementation used in 4.3 [5]."
        )
    doc.add_heading("4.1 Logistic Regression", level=2)
    doc.add_paragraph(
        "Logistic Regression serves as the baseline. Ng and Jordan [6] showed that "
        "discriminative classifiers reach their asymptotic error with fewer training examples "
        "than generative models (e.g. Naive Bayes), justifying LR over NB on a dataset of "
        "this size. LR is interpretable via its coefficients, provides well-calibrated "
        "probabilities, and sets a minimum performance bar that more complex models must exceed."
    )
    doc.add_heading("4.2 Random Forest", level=2)
    p_rf = doc.add_paragraph()
    p_rf.add_run(
        "Random Forest is the primary model: bagging many decorrelated trees over random feature "
        "subsets cuts variance without adding bias, the weakness a single decision tree has [7]. "
        "A standalone Decision Tree is rejected outright. Unpruned trees overfit severely, and even "
        "pruned, they lose to the ensemble on both bias-variance tradeoff and generalisation [8]. "
        "Feature importance here uses permutation importance rather than impurity-based (MDI) "
        "scores, since MDI is known to be unreliable across features of varying cardinality [9]."
    )
    doc.add_heading("4.3 XGBoost", level=2)
    doc.add_paragraph(
        "XGBoost extends gradient boosting with second-order Taylor approximations of the loss "
        "function, column subsampling, and L1/L2 regularisation, and consistently performs at the "
        "state of the art on tabular data [5]. The scale_pos_weight "
        "parameter handles class imbalance here by weighting the positive class inversely to its "
        "frequency."
    )
    doc.add_heading("4.4 Support Vector Machine and kNN Exclusion", level=2)
    doc.add_paragraph(
        "SVM with a linear kernel provides a margin-based comparison [10]. LinearSVC is wrapped in "
        "CalibratedClassifierCV (Platt scaling) to produce probabilities for ROC-AUC computation. "
        "kNN is excluded. Past roughly 10-15 dimensions, the distance to the nearest neighbour "
        "converges toward the distance to the farthest one, and Euclidean distance stops carrying "
        "useful discriminative signal [11]. The feature space used here is 51-dimensional, well "
        "past that threshold."
    )
    doc.add_heading("4.5 Class Imbalance and Evaluation Metrics", level=2)
    doc.add_paragraph(
        "All models use class_weight='balanced' (or equivalent) to correct for the 78:22 class "
        "imbalance in training data. F1-macro and ROC-AUC are reported throughout; accuracy is "
        "not reported as a primary metric because a trivial majority-class predictor would "
        "achieve 78.3 per cent accuracy on training data."
    )
    doc.add_heading("4.6 Nested Cross-Validation and Temporal Test Split", level=2)
    doc.add_paragraph(
        "Evaluation uses nested cross-validation: a 5-fold stratified outer loop estimates "
        "generalisation performance, while a 3-fold stratified inner loop selects hyperparameters. "
        "This prevents the optimistic bias that arises when the same data split is used for both "
        "hyperparameter selection and performance estimation [12]. The final "
        "test evaluation uses the temporally held-out 2025-2026 sample, which was never seen "
        "during training or hyperparameter search."
    )

    # ------------------------------------------------------------------
    # 5. Results
    # ------------------------------------------------------------------
    doc.add_heading("5. Results", level=1)
    doc.add_heading("5.1 Model Comparison", level=2)
    doc.add_paragraph(
        "Table 1 presents the nested cross-validation scores on the training set and the final "
        "held-out test set scores for all four models."
    )

    cap = doc.add_paragraph("Table 1. Model comparison: nested CV and test set performance.")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.runs[0].italic = True
    cap.runs[0].font.size = Pt(10)
    add_comparison_table(doc, comparison_rows, winners)
    doc.add_paragraph()

    if winners:
        auc_order = ", ".join(
            f"{m['model']} ({m['test_roc_auc']:.4f})" for m in winners['ranked_auc']
        )
        if safe:
            result_sentence = (
                f"On the held-out test set, models rank by ROC-AUC as follows: {auc_order}. "
                f"{winners['best_auc_model']} achieves the highest test-set score, despite not "
                "leading on cross-validation (see Section 6)."
            )
        else:
            result_sentence = (
                f"On the held-out test set the models rank by ROC-AUC as: {auc_order}. "
                f"{winners['best_auc_model']} comes top. The part I did not expect is that this is "
                "not the model that scored best in cross-validation. CV runs on the 2020-2024 "
                "training years; the test set is a later, structurally different period, and the two "
                "do not agree (see temporal shift below)."
            )
    else:
        result_sentence = (
            "[RESULT: run all notebooks and re-generate this report so the ranked "
            "comparison can be stated here from the actual saved metrics.]"
        )
    doc.add_paragraph(
        "Every model clears the no-skill baseline comfortably (ROC-AUC 0.5, PR-AUC equal to the "
        "positive prevalence). "
        f"{result_sentence} "
        "The CV-to-test drop shows up for all four and traces straight back to the temporal shift "
        "from Section 2: the test years hold far fewer high-headroom homes, so the task is harder "
        "there. That is what deployment on future data looks like, so I read the drop as realistic, "
        "not as a fault in the models."
    )

    add_figure(doc,
        f"{FIGURES_DIR}/all_models_roc_pr.png",
        "Fig. 4. ROC curves (left) and Precision-Recall curves (right) for all four models "
        "on the 2025-2026 test set."
    )

    doc.add_heading("5.2 Statistical Significance: McNemar's Test", level=2)
    if safe:
        doc.add_paragraph(
            "McNemar's test was applied to four model pairs to assess whether classification error "
            "distributions are statistically distinct [13]. It is the right test for a "
            "single train/test split design: it has acceptably low Type I error here, as opposed "
            "to designs with repeated resampling, where a 5x2cv test is recommended instead "
            "[14]; the Diebold-Mariano test is for regression and does not apply. "
            "Every pair differs significantly (p < 0.0001), including the closest pair, SVM vs "
            "Random Forest. With a 50,000-row test set this test has enough power to flag small "
            "error-rate differences as significant, so it establishes that the models make "
            "different errors, not that the difference is large enough to matter for any single "
            "property-level decision; the ROC-AUC and F1-macro gaps in Table 1 are the better guide "
            "to practical importance."
        )
    else:
        doc.add_paragraph(
            "McNemar's test was applied to four model pairs (Random Forest vs Logistic Regression, "
            "XGBoost vs Logistic Regression, XGBoost vs Random Forest, and SVM vs Random Forest) to "
            "assess whether classification error distributions are statistically distinct "
            "[13]. It is the right choice here: of five candidate significance tests for "
            "comparing classifiers, it has acceptably low Type I error specifically for the single "
            "train/test split design used in this report, as opposed to designs involving repeated "
            "resampling, where a 5x2cv test is recommended instead [14]. The test is "
            "appropriate for binary classifiers evaluated on the same test "
            "instances; the Diebold-Mariano test is for regression and is not applicable here. Every "
            "pair differs significantly (p < 0.0001), including the closest pair, SVM vs Random Forest. "
            "With a 50,000-row test set, "
            "McNemar's test has enough statistical power to detect small differences in error patterns "
            "as significant; this establishes that the models make different errors, not that the "
            "difference is large enough to matter for any single property-level decision. The ROC-AUC "
            "and F1-macro gaps in Table 1 are a better guide to practical importance than the p-values "
            "alone."
        )

    add_figure(doc,
        f"{FIGURES_DIR}/rf_permutation_importances.png",
        "Fig. 5. Random Forest permutation feature importances on the test set "
        "(mean decrease in ROC-AUC with 10 repeats). Error bars show one standard deviation."
    )

    doc.add_heading("5.3 Feature Importance", level=2)
    doc.add_paragraph(
        "Feature importance analysis from Random Forest permutation importances shows that "
        "current energy efficiency score is the single strongest predictor, followed by CO2 "
        "emissions per floor area and heating cost. Construction age band also ranks highly, a "
        "plausible result since older properties, particularly pre-1966 stock built before modern "
        "insulation standards, typically carry more retrofit headroom (Fig. 3); this report does not "
        "include a separate age-stratified breakdown to verify that mechanism directly, so it is "
        "offered as a plausible explanation rather than a demonstrated one. "
        "Wall type (cavity vs solid) contributes meaningfully: EPC's assessment methodology rates "
        "uninsulated solid-wall construction as poor by design, so solid-wall properties "
        "systematically score lower on baseline efficiency and carry more improvement headroom "
        "than cavity-wall equivalents. This is a property of how EPC scores are assessed, not an "
        "artefact of the model."
    )

    doc.add_heading("5.4 Robustness Check: Naive Structural Baseline", level=2)
    naive = load_naive_baseline()
    best_row = None
    if winners and comparison_rows:
        best_row = next((r for r in comparison_rows if r['Model'] == winners['best_auc_model']), None)

    if naive and best_row:
        best_auc = float(best_row['Test ROC-AUC'])
        best_f1  = float(best_row['Test F1-macro'])
        best_pr  = float(best_row['Test PR-AUC'])
        auc_gap = best_auc - naive['test_roc_auc']
        f1_gap  = best_f1  - naive['test_f1_macro']
        pr_gap  = best_pr  - naive['test_pr_auc']
        if safe:
            doc.add_paragraph(
                "CURRENT_ENERGY_EFFICIENCY is retained as a feature even though a low current score "
                "structurally leaves more numerical room for a large gap (potential efficiency is "
                "capped near 100). To test how much of the result this explains alone, a Logistic "
                "Regression restricted to CURRENT_ENERGY_EFFICIENCY and CURRENT_ENERGY_RATING was "
                f"fitted under the same protocol. This naive model reaches a test ROC-AUC of "
                f"{naive['test_roc_auc']:.4f}, only {auc_gap:.4f} below {winners['best_auc_model']}'s "
                f"{best_auc:.4f}. On the class-imbalance-sensitive metrics the gap is larger: "
                f"{f1_gap:.4f} F1-macro and {pr_gap:.4f} PR-AUC, showing the structural and "
                "construction-age features contribute real value beyond the current-to-potential "
                "arithmetic relationship."
            )
        else:
            doc.add_paragraph(
                "RETROFIT_POTENTIAL is derived from the gap between potential and current EPC "
                "efficiency scores. POTENTIAL_ENERGY_EFFICIENCY is excluded from the feature set "
                "(Section 3), but CURRENT_ENERGY_EFFICIENCY is included, and a low current score "
                "structurally leaves more numerical room for a large gap, since potential efficiency "
                "is capped near 100. To test how much of the result above this relationship explains "
                "on its own, a Logistic Regression model restricted to only CURRENT_ENERGY_EFFICIENCY "
                "and CURRENT_ENERGY_RATING was fitted under the same nested cross-validation protocol "
                f"as the main baseline. This naive model reaches a test ROC-AUC of "
                f"{naive['test_roc_auc']:.4f}, only {auc_gap:.4f} below {winners['best_auc_model']}'s "
                f"{best_auc:.4f}, which explains why current efficiency dominates the permutation "
                "importance ranking above. On F1-macro and PR-AUC, the metrics more informative under "
                f"class imbalance, the gap is substantially larger: {naive['test_f1_macro']:.4f} versus "
                f"{best_f1:.4f} test F1-macro ({f1_gap:.4f} difference) and {naive['test_pr_auc']:.4f} "
                f"versus {best_pr:.4f} test PR-AUC ({pr_gap:.4f} difference). The structural, tenure, "
                "and construction-age features therefore contribute real discriminative value beyond "
                "the current-to-potential arithmetic relationship, even though ROC-AUC alone "
                "understates that contribution."
            )
    else:
        doc.add_paragraph(
            "[ROBUSTNESS CHECK: run the naive baseline cell in 05_Comparison_Evaluation.ipynb "
            "and re-generate this report so this section can be stated from the actual saved "
            "metrics.]"
        )

    doc.add_heading("5.5 Calibration", level=2)
    doc.add_paragraph(
        "Calibration checks whether a model's reported probability matches the true observed "
        "frequency of the positive class, distinct from ranking ability (ROC-AUC). This matters "
        "for policy use: a practitioner acting on a specific probability threshold needs that "
        "probability to be trustworthy, not just the property ranking."
    )
    add_figure(doc,
        f"{FIGURES_DIR}/calibration_curves.png",
        "Fig. 6. Calibration reliability diagrams for all four models on the test set. "
        "The dashed line represents perfect calibration."
    )

    # ------------------------------------------------------------------
    # 6. Discussion
    # ------------------------------------------------------------------
    doc.add_heading("6. Discussion", level=1)
    if winners:
        lead_model = winners['best_auc_model']
        lead_sentence = (
            f"{lead_model} achieves the strongest test-set performance of the four models compared."
        )
    else:
        lead_sentence = "[RESULT: state which model leads once all notebooks have been run.]"
    doc.add_paragraph(
        "The results show that machine learning can identify high-retrofit-potential properties "
        "from observable EPC characteristics with meaningful accuracy on this dataset. "
        f"{lead_sentence} "
        "For policy applications, the choice between the top two tree-based models depends on "
        "interpretability requirements: Random Forest permutation importances are straightforward "
        "to communicate to policymakers, while XGBoost's gain-based importances can be biased "
        "toward high-cardinality features."
    )
    if safe:
        doc.add_paragraph(
            "The temporal distribution shift (21.7 to 10.8 per cent positive) matters for "
            "deployment: a model trained on 2020-2024 data will meet a future stock with "
            "proportionally fewer high-retrofit candidates. Relative ranking still holds, so this "
            "does not invalidate the model, but a fixed threshold will lose recall over time and "
            "needs annual re-calibration against Fig. 6."
        )
    else:
        doc.add_paragraph(
            "The temporal distribution shift (21.7 to 10.8 per cent positive) has direct implications "
            "for deployment. A model trained on 2020-2024 data and applied to future assessments will "
            "encounter a stock with proportionally fewer high-retrofit candidates. This does not "
            "invalidate the model: the relative ranking of properties is what matters for "
            "prioritisation. It does mean, however, that a fixed probability threshold will have decreasing "
            "recall over time. Practitioners should re-calibrate the decision threshold annually "
            "using the reliability diagrams shown in Fig. 6."
        )
    doc.add_paragraph("There are four main limitations.")
    if safe:
        limitation_items = [
            "Data quality. The true EPC error rate is estimated at 36 to 62 per cent once assessor "
            "disagreement is accounted for [15], which directly affects the "
            "WALL_TYPE feature engineered from the same unreliable free-text fields.",
            "Target simplification. The binary target collapses heterogeneous properties: a "
            "20-point gap means different things in a rural solid-wall property versus an urban flat.",
            "Sample size. This uses 200,000 of 7.25 million eligible training records; full-data "
            "training may improve minority-class recall.",
            "Spatial features. Regional and deprivation variables are used only as nominal "
            "categories, not spatial models.",
        ]
    else:
        limitation_items = [
            "Data quality. The EPC database has documented quality issues: 27 per cent of "
            "open-data EPCs carry at least one flag suggesting an error, and the true error rate "
            "is estimated at 36 to 62 per cent once assessor disagreement on parameters such as "
            "wall type and built form is accounted for [15]. This directly "
            "affects the WALL_TYPE feature engineered in this pipeline, which is derived from the "
            "same free-text description fields identified there as unreliable.",
            "Target simplification. The binary target collapses heterogeneous properties: a "
            "20-point gap in a rural solid-wall property has different policy implications from "
            "the same gap in an urban flat.",
            "Sample size. This analysis uses a sample of 200,000 records rather than the full "
            "7.25 million training records; while sufficient for credible results, full-data "
            "training may improve recall on the minority class.",
            "Spatial features. Regional clustering and local authority deprivation indices are "
            "used only as nominal categories and may benefit from spatial modelling approaches.",
        ]
    for item in limitation_items:
        lp = doc.add_paragraph(style='List Number')
        lp.add_run(item)

    # ------------------------------------------------------------------
    # 7. Ethical Considerations
    # ------------------------------------------------------------------
    doc.add_heading("7. Ethical Considerations", level=1)
    if safe:
        doc.add_paragraph(
            "The dataset is published under the Open Government Licence and contains no directly "
            "identifying personal data; however, EPC records are property-addressable, so linkage "
            "with other datasets could in principle re-identify a dwelling. No such linkage is "
            "performed, and no addresses are retained beyond what feature engineering requires."
        )
        doc.add_paragraph(
            "A false negative deprioritises a property that genuinely warrants intervention, a "
            "real cost to occupant and policy alike; a false positive only wastes assessor time. "
            "That is why Section 6 recommends recall-oriented calibration and frames the model as "
            "a prioritisation aid, not an automated decision-maker. The pipeline is public at "
            "https://github.com/KNHNF/epc-retrofit-potential-ml for independent checking."
        )
    else:
        doc.add_paragraph(
            "The dataset is published by the Ministry of Housing, Communities and Local Government "
            "under the Open Government Licence and contains no directly identifying personal data; "
            "however, EPC records are addressable to individual properties, so combining them with "
            "other publicly linkable datasets could in principle re-identify an occupant's dwelling. "
            "No attempt is made in this pipeline to link EPC records to any other individual-level "
            "dataset, and no property addresses are retained beyond the fields required for feature "
            "engineering."
        )
        doc.add_paragraph(
            "There is also a practical fairness consideration in how the model would be used. A false "
            "negative (a high-potential property scored as low-potential) means a property that "
            "genuinely warrants a retrofit intervention is deprioritised, a real cost to both the "
            "occupant and net-zero policy goals. A false positive wastes assessor time but causes no "
            "direct harm. This asymmetry is why Section 6 recommends recall-oriented threshold "
            "calibration rather than optimising for accuracy alone, and why the model is framed here "
            "as a prioritisation aid for human assessors rather than an automated decision-maker."
        )
        doc.add_paragraph(
            "The full pipeline, from data ingestion through to this report, is published at "
            "https://github.com/KNHNF/epc-retrofit-potential-ml so that the methodology, feature "
            "engineering decisions, and reported numbers can be independently checked rather than "
            "taken on trust."
        )

    # ------------------------------------------------------------------
    # 8. Conclusion
    # ------------------------------------------------------------------
    doc.add_heading("8. Conclusion", level=1)
    if winners:
        conclusion_lead = (
            f"{winners['best_auc_model']} achieved the highest overall test-set performance, "
            f"with the other tree-based model providing comparable results alongside "
            "interpretable feature importances."
        )
    else:
        conclusion_lead = "[RESULT: state the best-performing model here once all notebooks have been run.]"
    if safe:
        doc.add_paragraph(
            "This paper presented a machine learning pipeline for identifying high-retrofit-potential "
            f"UK residential properties from EPC open data. {conclusion_lead} All ensemble models "
            "substantially outperformed the Logistic Regression baseline and the SVM, with an "
            "11-point temporal shift in the positive class rate between training and test periods. "
            "Feature importance analysis confirms "
            "current efficiency, CO2 intensity, construction age, and wall type as the primary "
            "drivers, all physically interpretable. Code and pipeline are public at "
            "https://github.com/KNHNF/epc-retrofit-potential-ml."
        )
    else:
        doc.add_paragraph(
            "This paper presented a machine learning pipeline for identifying high-retrofit-potential "
            "UK residential properties using EPC open data. Four classifiers were compared using nested "
            f"cross-validation and evaluated on a temporally held-out test set. {conclusion_lead} "
            "All ensemble models substantially outperformed the "
            "Logistic Regression baseline and the SVM. A temporal distribution shift of approximately "
            "11 percentage points in the positive class rate between training and test periods was "
            "identified. Feature importance "
            "analysis confirms that current efficiency score, CO2 emissions intensity, construction "
            "age, and wall type are the primary drivers, all physically interpretable and consistent "
            "with retrofit policy knowledge. The code and data pipeline are publicly available at "
            "https://github.com/KNHNF/epc-retrofit-potential-ml."
        )

    # ------------------------------------------------------------------
    # References. IEEE numeric style, listed in order of first citation.
    # ------------------------------------------------------------------
    doc.add_heading("References", level=1)

    add_reference(doc,
        "[1] Ministry of Housing, Communities and Local Government (2024) ",
        "Energy Performance of Buildings Data: England and Wales.",
        " Available from: https://epc.opendatacommunities.org [Accessed 9 July 2026]."
    )
    add_reference(doc,
        "[2] Seyedzadeh, S., Pour Rahimian, F., Glesk, I. and Roper, M. (2018) Machine learning "
        "for estimation of building energy consumption and performance: a review. ",
        "Visualization in Engineering.",
        " 6 (1), p. 5."
    )
    add_reference(doc,
        "[3] Pasichnyi, O., Wallin, J. and Kordas, O. (2019) Data-driven building archetypes for "
        "urban building energy modelling. ",
        "Energy.",
        " 181, pp. 360-377."
    )
    p_pedregosa = add_reference(doc, "[4] Pedregosa, F. ", "et al.", "")
    p_pedregosa.runs[1].italic = True
    p_pedregosa.add_run(" (2011) Scikit-learn: machine learning in Python. ").font.size = Pt(10)
    r = p_pedregosa.add_run("Journal of Machine Learning Research.")
    r.italic = True
    r.font.size = Pt(10)
    p_pedregosa.add_run(" 12, pp. 2825-2830.").font.size = Pt(10)
    add_reference(doc,
        "[5] Chen, T. and Guestrin, C. (2016) XGBoost: a scalable tree boosting system. ",
        "Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining.",
        " pp. 785-794."
    )
    add_reference(doc,
        "[6] Ng, A.Y. and Jordan, M.I. (2001) On discriminative vs. generative classifiers: a "
        "comparison of logistic regression and naive Bayes. ",
        "Advances in Neural Information Processing Systems 14 (NIPS 2001).",
        " pp. 841-848."
    )
    add_reference(doc,
        "[7] Breiman, L. (2001) Random forests. ",
        "Machine Learning.",
        " 45 (1), pp. 5-32."
    )
    add_reference(doc,
        "[8] Breiman, L., Friedman, J.H., Olshen, R.A. and Stone, C.J. (1984) ",
        "Classification and Regression Trees.",
        " Belmont, CA: Wadsworth."
    )
    add_reference(doc,
        "[9] Strobl, C., Boulesteix, A-L., Zeileis, A. and Hothorn, T. (2007) Bias in random "
        "forest variable importance measures: illustrations, sources and a solution. ",
        "BMC Bioinformatics.",
        " 8, p. 25."
    )
    add_reference(doc,
        "[10] Cortes, C. and Vapnik, V. (1995) Support-vector networks. ",
        "Machine Learning.",
        " 20 (3), pp. 273-297."
    )
    add_reference(doc,
        "[11] Beyer, K., Goldstein, J., Ramakrishnan, R. and Shaft, U. (1999) When is nearest "
        "neighbor meaningful? ",
        "International Conference on Database Theory (ICDT), Lecture Notes in Computer Science.",
        " Vol. 1540, pp. 217-235."
    )
    add_reference(doc,
        "[12] Varma, S. and Simon, R. (2006) Bias in error estimation when using cross-validation "
        "for model selection. ",
        "BMC Bioinformatics.",
        " 7, p. 91."
    )
    add_reference(doc,
        "[13] McNemar, Q. (1947) Note on the sampling error of the difference between correlated "
        "proportions or percentages. ",
        "Psychometrika.",
        " 12 (2), pp. 153-157."
    )
    add_reference(doc,
        "[14] Dietterich, T.G. (1998) Approximate statistical tests for comparing supervised "
        "classification learning algorithms. ",
        "Neural Computation.",
        " 10 (7), pp. 1895-1923."
    )
    add_reference(doc,
        "[15] Hardy, A. and Glew, D. (2019) An analysis of errors in the Energy Performance "
        "Certificate database. ",
        "Energy Policy.",
        " 129, pp. 1168-1178."
    )

    # Word count: Introduction through Conclusion only. Title, author block, abstract,
    # and references excluded (brief's 2000-word rule + standard academic convention).
    body_start = body_end = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "1. Introduction":
            body_start = i
        if p.text.strip() == "References":
            body_end = i
            break
    word_count = 0
    if body_start is not None and body_end is not None:
        for p in doc.paragraphs[body_start:body_end]:
            if p.style.name.startswith('Heading'):
                continue
            word_count += len(p.text.split())

    doc.add_paragraph()
    wc_para = doc.add_paragraph()
    wc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    limit_note = "2,000-word brief + 10% tolerance (2,200 max)" if safe else "quality-first, not word-capped"
    wc_run = wc_para.add_run(
        f"Word count: {word_count} (Introduction to Conclusion, excluding title, "
        f"abstract, and references). Mode: {mode} ({limit_note})."
    )
    wc_run.italic = True
    wc_run.font.size = Pt(10)

    doc.save(OUT_PATH)
    print(f"[{mode}] Report saved to {OUT_PATH}")
    print(f"[{mode}] Word count (Introduction to Conclusion, excl. title/abstract/references): {word_count}")
    return word_count


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else "both"
    if arg == "both":
        build_report("full")
        build_report("safe")
    else:
        build_report(arg)
