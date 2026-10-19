"""
risk.py
=======
Risk Engine — combines Workload probability and Sentiment risk into a
single Risk Score and Level.

Formula:
    risk_score = (0.6 × workload_prob + 0.4 × sentiment_risk) × 100

Levels:
    Low    : risk_score <  34
    Medium : risk_score <  67
    High   : risk_score >= 67
"""

from typing import Tuple


# ── Weights ───────────────────────────────────────────────────────────────────
WORKLOAD_WEIGHT  = 0.60
SENTIMENT_WEIGHT = 0.40

# ── Level thresholds ─────────────────────────────────────────────────────────
LOW_MAX    = 34   # 0 – 33  → Low
MEDIUM_MAX = 66   # 34 – 66 → Medium
                  # 67+     → High


def compute_risk(
    workload_prob: float,
    sentiment_risk_score: float
) -> Tuple[int, str]:
    """
    Compute the final burnout Risk Score and Level.

    Parameters
    ----------
    workload_prob : float
        Probability output from the ML workload model [0.0, 1.0].
    sentiment_risk_score : float
        Sentiment risk from the NLP model [0.0, 1.0].

    Returns
    -------
    (risk_score, level) : (int, str)
        risk_score : 0–100
        level      : "Low" | "Medium" | "High"
    """
    raw = WORKLOAD_WEIGHT * workload_prob + SENTIMENT_WEIGHT * sentiment_risk_score
    risk_score = int(round(raw * 100))
    risk_score = max(0, min(100, risk_score))   # clamp to [0, 100]

    if risk_score < LOW_MAX:
        level = "Low"
    elif risk_score <= MEDIUM_MAX:
        level = "Medium"
    else:
        level = "High"

    return risk_score, level


def level_color(level: str) -> str:
    """Return a hex colour code for UI display."""
    return {"Low": "#22c55e", "Medium": "#f59e0b", "High": "#ef4444"}.get(level, "#6b7280")
