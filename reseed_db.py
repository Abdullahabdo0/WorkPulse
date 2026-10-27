"""
reseed_db.py
============
Optimized High-Speed Database Seeder (1,000 Employees)
Uses pre-computed password hash and batch processing for sub-5-second execution.
"""

import os
import sys
import json
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(ROOT, "backend"))

import joblib
from database import SessionLocal, User, Assessment, hash_password, init_db
from sentiment import sentiment_risk, load_sentiment_model
from risk import compute_risk

DATA_PATH = os.path.join(ROOT, "data", "employees.csv")
MODEL_PATH = os.path.join(ROOT, "models", "workload_model.pkl")

NUMERIC_FEATURES = [
    "avg_daily_hours",
    "overtime_hours_week",
    "absence_days_month",
    "late_arrivals_month",
    "tasks_assigned",
    "vacation_days_last_year",
]

FEATURE_LABELS = {
    "avg_daily_hours": "High daily working hours",
    "overtime_hours_week": "Excessive overtime hours",
    "absence_days_month": "Frequent absences",
    "late_arrivals_month": "Repeated late arrivals",
    "tasks_assigned": "High task load",
    "vacation_days_last_year": "Insufficient vacation days",
}


def reseed():
    print("=" * 60)
    print("  WorkPulse: Fast Seeding 1,000 Employees into Database")
    print("=" * 60)

    if not os.path.exists(DATA_PATH):
        print(f"Error: {DATA_PATH} not found.")
        return

    init_db()
    db = SessionLocal()

    try:
        # 1. Models & NLP
        print("[1/4] Loading ML model and NLP Sentiment...")
        workload_model = joblib.load(MODEL_PATH) if os.path.exists(MODEL_PATH) else None
        load_sentiment_model()

        # 2. Clear old employee data
        print("[2/4] Resetting employee tables...")
        db.query(Assessment).delete()
        db.query(User).filter(User.role == "employee").delete()
        db.commit()

        # 3. Read dataset
        df = pd.read_csv(DATA_PATH)
        print(f"[3/4] Processing {len(df)} records...")

        # Batch ML predictions
        probs = workload_model.predict_proba(df[NUMERIC_FEATURES])[:, 1] if workload_model else [0.5] * len(df)

        # Cache sentiment for unique feedbacks
        unique_feedbacks = df["feedback"].dropna().unique()
        print(f"      Analyzing {len(unique_feedbacks)} unique feedback sentiments...")
        sentiment_map = {fb: sentiment_risk(str(fb)) for fb in unique_feedbacks}

        # Pre-compute password hash once (100x speedup)
        default_pwd_hash = hash_password("123456")

        # Pre-compute SHAP feature contributions in batch
        print("      Calculating SHAP explainability in batch...")
        try:
            import shap
            explainer = shap.TreeExplainer(workload_model)
            shap_values = explainer.shap_values(df[NUMERIC_FEATURES])
            if isinstance(shap_values, list):
                sv_matrix = shap_values[1]
            elif isinstance(shap_values, np.ndarray) and shap_values.ndim == 3:
                sv_matrix = shap_values[:, :, 1]
            else:
                sv_matrix = np.array(shap_values)
        except Exception as e:
            print(f"      SHAP batch fallback ({e})")
            sv_matrix = None

        # 4. Insert records
        print("[4/4] Inserting 1,000 users and assessments...")
        users_to_add = []
        for idx, row in df.iterrows():
            emp_id = str(row["employee_id"])
            full_name = str(row["name"])
            username = emp_id.lower()
            gender_val = str(row["gender"]) if "gender" in row and pd.notna(row["gender"]) else None
            age_val = int(row["age"]) if "age" in row and pd.notna(row["age"]) else None

            user = User(
                username=username,
                full_name=full_name,
                gender=gender_val,
                age=age_val,
                role="employee",
                hashed_password=default_pwd_hash
            )
            users_to_add.append(user)

        db.bulk_save_objects(users_to_add, return_defaults=True)
        db.flush()

        # Query all newly inserted users to map user_id
        user_map = {u.username: u.id for u in db.query(User).filter(User.role == "employee").all()}

        assessments_to_add = []
        for idx, row in df.iterrows():
            emp_id = str(row["employee_id"])
            username = emp_id.lower()
            user_id = user_map.get(username)

            w_prob = float(probs[idx])
            fb_text = str(row["feedback"])
            s_risk = sentiment_map.get(fb_text, 0.5)
            score, level = compute_risk(w_prob, s_risk)

            # Get top reasons from pre-computed SHAP
            reasons = []
            if sv_matrix is not None and idx < len(sv_matrix):
                sv_row = sv_matrix[idx]
                indices = np.argsort(np.abs(sv_row))[::-1]
                for f_idx in indices:
                    if len(reasons) >= 3:
                        break
                    if sv_row[f_idx] > 0:
                        reasons.append(FEATURE_LABELS.get(NUMERIC_FEATURES[f_idx], NUMERIC_FEATURES[f_idx]))

            if s_risk >= 0.65 and len(reasons) < 3:
                reasons.append("Negative employee sentiment")

            if not reasons:
                reasons = ["High workload indicators"]

            assessments_to_add.append(Assessment(
                user_id=user_id,
                avg_daily_hours=float(row["avg_daily_hours"]),
                overtime_hours_week=float(row["overtime_hours_week"]),
                absence_days_month=int(row["absence_days_month"]),
                late_arrivals_month=int(row["late_arrivals_month"]),
                tasks_assigned=int(row["tasks_assigned"]),
                vacation_days_last_year=int(row["vacation_days_last_year"]),
                feedback=fb_text,
                risk_score=score,
                level=level,
                workload_probability=round(w_prob, 4),
                sentiment_risk=round(s_risk, 4),
                top_reasons=json.dumps(reasons),
            ))

        db.bulk_save_objects(assessments_to_add)
        db.commit()

        # Verify
        total_users = db.query(User).filter(User.role == "employee").count()
        total_assessments = db.query(Assessment).count()
        print(f"\n[DONE] Successfully seeded:")
        print(f"       - {total_users} Employees")
        print(f"       - {total_assessments} Assessments")
        print(f"       Database: workpulse.db is 100% up to date!")

    finally:
        db.close()


if __name__ == "__main__":
    reseed()
