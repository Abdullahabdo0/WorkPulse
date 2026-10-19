"""
explain.py
==========
SHAP-based explainability for the Workload (Random Forest) model.
Returns the top N feature contributions for a single employee prediction.
"""

import numpy as np
import pandas as pd
from typing import List, Dict, Any

# Feature labels (human-readable names for display)
FEATURE_NAMES = [
    "avg_daily_hours",
    "overtime_hours_week",
    "absence_days_month",
    "late_arrivals_month",
    "tasks_assigned",
    "vacation_days_last_year",
]

FEATURE_LABELS = {
    "avg_daily_hours":          "High daily working hours",
    "overtime_hours_week":      "Excessive overtime hours",
    "absence_days_month":       "Frequent absences",
    "late_arrivals_month":      "Repeated late arrivals",
    "tasks_assigned":           "High task load",
    "vacation_days_last_year":  "Insufficient vacation days",
}


def get_top_reasons(
    model,
    employee_features: Dict[str, Any],
    sentiment_risk: float,
    top_n: int = 3,
) -> List[str]:
    """
    Generate top reasons for a burnout prediction using SHAP values.

    Parameters
    ----------
    model : fitted sklearn estimator
        The trained Random Forest (or any tree-based) model.
    employee_features : dict
        Numeric feature values (keys = FEATURE_NAMES).
    sentiment_risk : float
        Sentiment risk score [0, 1] — added as context reason if high.
    top_n : int
        Number of top reasons to return.

    Returns
    -------
    List[str]
        Human-readable list of top contributing factors.
    """
    try:
        import shap

        # Build feature array in correct order
        x = np.array([[employee_features[f] for f in FEATURE_NAMES]])
        df_input = pd.DataFrame(x, columns=FEATURE_NAMES)

        # Use TreeExplainer for Random Forest
        explainer = shap.TreeExplainer(model)
        shap_values = explainer.shap_values(df_input)

        # For binary classification, shap_values is a list [class0, class1]
        # We want class 1 (burnout = True)
        if isinstance(shap_values, list):
            sv = shap_values[1][0]
        else:
            sv = shap_values[0]

        # Sort features by absolute SHAP value (descending)
        indices = np.argsort(np.abs(sv))[::-1]

        reasons = []
        for idx in indices:
            if len(reasons) >= top_n:
                break
            feature = FEATURE_NAMES[idx]
            contribution = sv[idx]
            if contribution > 0:  # Only include features that push toward burnout
                reasons.append(FEATURE_LABELS.get(feature, feature))

        # If sentiment is high and we still have room, add it
        if sentiment_risk >= 0.65 and len(reasons) < top_n:
            reasons.append("Negative employee sentiment")

        # Fallback: if no positive contributors found, use top features by value
        if not reasons:
            reasons = _fallback_reasons(employee_features, sentiment_risk, top_n)

        return reasons[:top_n]

    except Exception as e:
        # If SHAP fails for any reason, fall back to rule-based explanation
        print(f"[SHAP] Warning: {e} — using fallback explanation.")
        return _fallback_reasons(employee_features, sentiment_risk, top_n)


def _fallback_reasons(
    features: Dict[str, Any],
    sentiment_risk: float,
    top_n: int = 3,
) -> List[str]:
    """
    Rule-based fallback when SHAP is unavailable.
    Scores each feature against typical thresholds.
    """
    scores = {}

    if features.get("avg_daily_hours", 0) > 9:
        scores["avg_daily_hours"] = features["avg_daily_hours"] - 9

    if features.get("overtime_hours_week", 0) > 10:
        scores["overtime_hours_week"] = features["overtime_hours_week"] - 10

    if features.get("absence_days_month", 0) > 3:
        scores["absence_days_month"] = features["absence_days_month"] - 3

    if features.get("late_arrivals_month", 0) > 5:
        scores["late_arrivals_month"] = features["late_arrivals_month"] - 5

    if features.get("tasks_assigned", 0) > 25:
        scores["tasks_assigned"] = features["tasks_assigned"] - 25

    if features.get("vacation_days_last_year", 25) < 10:
        scores["vacation_days_last_year"] = 10 - features["vacation_days_last_year"]

    # Sort by score descending
    sorted_features = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    reasons = [FEATURE_LABELS.get(f, f) for f, _ in sorted_features[:top_n]]

    # Add sentiment if high
    if sentiment_risk >= 0.65 and len(reasons) < top_n:
        reasons.append("Negative employee sentiment")

    if not reasons:
        reasons = ["Cumulative workload factors", "Work-life imbalance indicators"]

    return reasons[:top_n]
