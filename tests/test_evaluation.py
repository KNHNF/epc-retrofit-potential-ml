"""
test_evaluation.py
Smoke tests for evaluation utility functions.
"""

import numpy as np
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from evaluation import evaluate_model, plot_confusion_matrix, plot_feature_importances


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
