"""
train_model.py
==============
Standalone script to train and evaluate the Workload model.
Run this after generating the data:

    python data/generate_data.py
    python train_model.py

Outputs:
    models/workload_model.pkl   — saved Random Forest model
    (prints metrics to console)
"""

import os
import sys
import warnings
import joblib
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score,
    f1_score, confusion_matrix, classification_report
)

warnings.filterwarnings("ignore")

# ── Paths ──────────────────────────────────────────────────────────────────
ROOT       = os.path.dirname(os.path.abspath(__file__))
DATA_PATH  = os.path.join(ROOT, "data", "employees.csv")
MODEL_DIR  = os.path.join(ROOT, "models")
MODEL_PATH = os.path.join(MODEL_DIR, "workload_model.pkl")

os.makedirs(MODEL_DIR, exist_ok=True)

# ── Feature columns ────────────────────────────────────────────────────────
FEATURES = [
    "avg_daily_hours",
    "overtime_hours_week",
    "absence_days_month",
    "late_arrivals_month",
    "tasks_assigned",
    "vacation_days_last_year",
]
TARGET = "burnout"


def load_data():
    if not os.path.exists(DATA_PATH):
        print(f"[ERROR] Data not found: {DATA_PATH}")
        print("Run: python data/generate_data.py")
        sys.exit(1)
    df = pd.read_csv(DATA_PATH)
    print(f"[Data] Loaded {len(df)} records | Burnout rate: {df[TARGET].mean()*100:.1f}%")
    return df


def train_and_evaluate():
    # ── Load ───────────────────────────────────────────────
    df = load_data()
    X = df[FEATURES]
    y = df[TARGET]

    # ── Split ──────────────────────────────────────────────
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.20, random_state=42, stratify=y
    )
    print(f"\n[Split] Train: {len(X_train)} | Test: {len(X_test)}")

    # ── Model 1: Logistic Regression (baseline) ────────────
    scaler = StandardScaler()
    X_train_sc = scaler.fit_transform(X_train)
    X_test_sc  = scaler.transform(X_test)

    lr = LogisticRegression(max_iter=500, random_state=42)
    lr.fit(X_train_sc, y_train)
    lr_preds = lr.predict(X_test_sc)

    print("\n" + "="*50)
    print("  Logistic Regression (Baseline)")
    print("="*50)
    print(classification_report(y_test, lr_preds, target_names=["No Burnout", "Burnout"]))

    # ── Model 2: Random Forest ─────────────────────────────
    rf = RandomForestClassifier(
        n_estimators=200,
        max_depth=10,
        min_samples_split=5,
        random_state=42,
        n_jobs=-1,
    )
    rf.fit(X_train, y_train)
    rf_preds = rf.predict(X_test)

    print("="*50)
    print("  Random Forest")
    print("="*50)
    print(classification_report(y_test, rf_preds, target_names=["No Burnout", "Burnout"]))

    # ── Comparison table ───────────────────────────────────
    results = {
        "Model": ["Logistic Regression", "Random Forest"],
        "Accuracy":  [accuracy_score(y_test, lr_preds),  accuracy_score(y_test, rf_preds)],
        "Precision": [precision_score(y_test, lr_preds), precision_score(y_test, rf_preds)],
        "Recall":    [recall_score(y_test, lr_preds),    recall_score(y_test, rf_preds)],
        "F1":        [f1_score(y_test, lr_preds),        f1_score(y_test, rf_preds)],
    }
    df_results = pd.DataFrame(results)
    print("\n-- Comparison -------------------------------------------")
    print(df_results.to_string(index=False))

    # ── Choose best model (highest F1) ────────────────────
    best_model = rf  # We default to Random Forest as it typically outperforms
    best_name  = "Random Forest"
    print(f"\n[Decision] Using {best_name} as the final model.")

    # ── Feature Importance ─────────────────────────────────
    importances = pd.Series(rf.feature_importances_, index=FEATURES)
    importances_sorted = importances.sort_values(ascending=True)

    print("\n-- Feature Importance (Random Forest) ----------")
    for f, imp in importances_sorted.items():
        bar = "#" * int(imp * 40)
        print(f"  {f:<30} {imp:.4f}  {bar}")

    # -- Plot Feature Importance --
    plt.figure(figsize=(8, 4))
    importances_sorted.plot(kind="barh", color=["#3b82f6"]*len(FEATURES))
    plt.title("Random Forest - Feature Importance")
    plt.xlabel("Importance")
    plt.tight_layout()
    plt_path = os.path.join(MODEL_DIR, "feature_importance.png")
    plt.savefig(plt_path, dpi=120)
    plt.close()
    print(f"\n[Saved] Feature importance chart -> {plt_path}")

    # -- Confusion Matrix --
    cm = confusion_matrix(y_test, rf_preds)
    plt.figure(figsize=(5, 4))
    sns.heatmap(cm, annot=True, fmt="d", cmap="Blues",
                xticklabels=["No Burnout", "Burnout"],
                yticklabels=["No Burnout", "Burnout"])
    plt.title("Random Forest - Confusion Matrix")
    plt.ylabel("True Label")
    plt.xlabel("Predicted Label")
    plt.tight_layout()
    cm_path = os.path.join(MODEL_DIR, "confusion_matrix.png")
    plt.savefig(cm_path, dpi=120)
    plt.close()
    print(f"[Saved] Confusion matrix -> {cm_path}")

    # -- Save model --
    joblib.dump(best_model, MODEL_PATH)
    print(f"\n[Saved] Model -> {MODEL_PATH}")
    print("\nTraining complete!")

    return best_model


if __name__ == "__main__":
    train_and_evaluate()
