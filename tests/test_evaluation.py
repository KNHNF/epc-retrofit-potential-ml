"""
test_evaluation.py
Smoke tests for evaluation utility functions.
"""

import matplotlib
matplotlib.use('Agg')
import numpy as np
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from evaluation import (evaluate_model, plot_confusion_matrix, plot_feature_importances,
                        plot_roc_curves)


class DummyModel:
    """Minimal mock classifier for testing evaluation functions."""
    def predict(self, X):
        return np.ones(len(X), dtype=int)

    def predict_proba(self, X):
        return np.column_stack([np.zeros(len(X)), np.ones(len(X))])

    @property
    def feature_importances_(self):
        return np.array([0.5, 0.3, 0.2])


def test_evaluate_model_returns_keys():
    model = DummyModel()
    X = np.zeros((10, 3))
    y = np.array([0, 1, 0, 1, 0, 1, 0, 1, 0, 1])
    result = evaluate_model(model, X, y, model_name='Test')
    assert 'f1_weighted' in result
    assert 'roc_auc' in result


def test_plot_confusion_matrix_runs():
    y_test = np.array([0, 1, 1, 0])
    y_pred = np.array([0, 1, 0, 0])
    fig = plot_confusion_matrix(y_test, y_pred, model_name='Test')
    assert fig is not None


def test_plot_feature_importances_runs():
    model = DummyModel()
    feature_names = ['feature_a', 'feature_b', 'feature_c']
    fig = plot_feature_importances(model, feature_names, top_n=3)
    assert fig is not None


class ScoreOnlyModel:
    """Like LinearSVC: predict and decision_function, but no predict_proba."""
    def predict(self, X):
        return (X[:, 0] > 0).astype(int)

    def decision_function(self, X):
        return X[:, 0]


def test_evaluate_model_reports_macro_f1():
    # README says F1-macro; the helper used to report only the weighted average.
    model = DummyModel()   # predicts all 1
    y = np.array([0, 1, 0, 1, 0, 1, 0, 1, 0, 1])
    result = evaluate_model(model, np.zeros((10, 3)), y, model_name='Test')
    # all-ones prediction on a balanced y: class 1 F1 = 2/3, class 0 F1 = 0, macro = 1/3
    assert result['f1_macro'] == pytest.approx(1 / 3)


def test_evaluate_model_falls_back_to_decision_function():
    X = np.array([[-2.0], [-1.0], [1.0], [2.0]])
    y = np.array([0, 0, 1, 1])
    result = evaluate_model(ScoreOnlyModel(), X, y, model_name='Scores')
    assert result['roc_auc'] == 1.0


def test_evaluate_model_zero_auc_is_still_printed(capsys):
    # roc_auc of exactly 0.0 is falsy; it must not be silently skipped.
    X = np.array([[2.0], [1.0], [-1.0], [-2.0]])
    y = np.array([0, 0, 1, 1])
    evaluate_model(ScoreOnlyModel(), X, y, model_name='Inverted')
    assert 'ROC-AUC: 0.0000' in capsys.readouterr().out


def test_plot_roc_curves_accepts_evaluate_model_output():
    # plot_roc_curves reads r['y_test'], which evaluate_model used not to return (KeyError).
    y = np.array([0, 1, 0, 1, 0, 1, 0, 1, 0, 1])
    results = [evaluate_model(DummyModel(), np.zeros((10, 3)), y, model_name='Dummy')]
    assert plot_roc_curves(results) is not None


def test_plot_feature_importances_length_mismatch_raises():
    with pytest.raises(ValueError, match='feature_names'):
        plot_feature_importances(DummyModel(), ['only_one'], top_n=1)
