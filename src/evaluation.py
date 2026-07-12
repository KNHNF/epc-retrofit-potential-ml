"""
evaluation.py
Standardised evaluation functions used across all model notebooks.
"""

import numpy as np
import matplotlib.pyplot as plt
from sklearn.metrics import (
    classification_report, confusion_matrix,
    roc_auc_score, roc_curve, f1_score,
    ConfusionMatrixDisplay
)


def evaluate_model(model, X_test, y_test, model_name: str = "Model") -> dict:
    """
    Run standard evaluation for a fitted classifier.
    Returns a dict of metrics and prints a summary.
    """
    y_pred = model.predict(X_test)
    y_prob = model.predict_proba(X_test)[:, 1] if hasattr(model, 'predict_proba') else None

    f1 = f1_score(y_test, y_pred, average='weighted')
    roc_auc = roc_auc_score(y_test, y_prob) if y_prob is not None else None

    print(f"\n{model_name} results")
    print(classification_report(y_test, y_pred, target_names=['Low potential', 'High potential']))
    if roc_auc:
        print(f"ROC-AUC: {roc_auc:.4f}")

    return {
        'model_name': model_name,
        'f1_weighted': f1,
        'roc_auc': roc_auc,
        'y_pred': y_pred,
        'y_prob': y_prob,
    }


def plot_confusion_matrix(y_test, y_pred, model_name: str = "Model"):
    """Plot a labelled confusion matrix."""
    cm = confusion_matrix(y_test, y_pred)
    disp = ConfusionMatrixDisplay(confusion_matrix=cm,
                                   display_labels=['Low potential', 'High potential'])
    fig, ax = plt.subplots(figsize=(5, 4))
    disp.plot(ax=ax, cmap='Blues', colorbar=False)
    ax.set_title(f'Confusion matrix: {model_name}')
    plt.tight_layout()
    return fig


def plot_roc_curves(results: list, title: str = "ROC curves"):
    """
    Plot ROC curves for multiple models on one chart.
    results: list of dicts returned by evaluate_model, each must have y_prob.
    y_test must be passed separately as it is shared.
    """
    fig, ax = plt.subplots(figsize=(6, 5))
    ax.plot([0, 1], [0, 1], 'k--', label='Random classifier')
    for r in results:
        if r['y_prob'] is not None:
            fpr, tpr, _ = roc_curve(r['y_test'], r['y_prob'])
            ax.plot(fpr, tpr, label=f"{r['model_name']} (AUC={r['roc_auc']:.3f})")
    ax.set_xlabel('False positive rate')
    ax.set_ylabel('True positive rate (recall)')
    ax.set_title(title)
    ax.legend(loc='lower right')
    plt.tight_layout()
    return fig


def plot_feature_importances(model, feature_names: list, top_n: int = 15, title: str = "Feature importances"):
    """Plot top N feature importances from a tree-based model (RF, XGBoost)."""
    importances = model.feature_importances_
    indices = np.argsort(importances)[::-1][:top_n]
    top_features = [feature_names[i] for i in indices]
    top_values = importances[indices]

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(top_features[::-1], top_values[::-1], color='steelblue')
    ax.set_xlabel('Importance')
    ax.set_title(title)
    plt.tight_layout()
    return fig
