"""
generate_data.py
================
Generates a synthetic dataset of ~2000 employees with burnout labels.
Columns:
    avg_daily_hours        : float  [6.0, 13.0]
    overtime_hours_week    : float  [0,  25]
    absence_days_month     : int    [0,   8]
    late_arrivals_month    : int    [0,  15]
    tasks_assigned         : int    [5,  40]
    vacation_days_last_year: int    [0,  25]
    feedback               : str    (one of 30 predefined sentences)
    burnout                : int    0 or 1  (target label)

Logic:
    A burnout score is computed from weighted features + noise.
    If score > threshold → burnout = 1, else 0.
    This ensures realistic class distribution (~35-40% burnout).
"""

import numpy as np
import pandas as pd
import random
import os

# ── Reproducibility ─────────────────────────────────────────────────────────
SEED = 42
np.random.seed(SEED)
random.seed(SEED)

N_EMPLOYEES = 1000

# ── Employee names pool (Arabic transliterated + common names) ───────────────
EMPLOYEE_NAMES = [
    "Ahmed Mohamed", "Mohamed Ali", "Omar Hassan", "Youssef Ibrahim",
    "Khaled Mahmoud", "Tariq Saeed", "Bilal Nasser", "Faris Amin",
    "Rami Younis", "Samir Fouad", "Nader Hamid", "Wael Karim",
    "Hassan Zaki", "Ayman Refaat", "Sherif Adel", "Bassam Tawfik",
    "Hossam Gamal", "Emad Saleh", "Ihab Mostafa", "Walid Nabil",
    "Layla Ahmed", "Nour Mohamed", "Sara Hassan", "Fatima Ibrahim",
    "Rania Khaled", "Dina Mahmoud", "Yasmine Saeed", "Mariam Nasser",
    "Hana Amin", "Salma Younis", "Amira Fouad", "Doaa Hamid",
    "Noha Wael", "Eman Karim", "Ghada Zaki", "Mona Refaat",
    "Heba Adel", "Nada Tawfik", "Reem Gamal", "Mai Saleh",
    "Adam Mostafa", "Ziad Nabil", "Karim Adly", "Mazen Fathy",
    "Sami Lotfy", "Tarek Wahid", "Ramzi Sobhy", "Galal Ezz",
    "Hazem Atef", "Shadi Raafat", "Nadine Hany", "Aya Sherif",
    "Rawan Essam", "Lina Ashraf", "Hoda Wafaa", "Rana Magdi",
    "Dalia Shawki", "Omnia Karam", "Engy Samir", "Rasha Ihab",
]

# ── Feedback sentences ───────────────────────────────────────────────────────
NEGATIVE_FEEDBACK = [
    "I feel completely exhausted and overwhelmed with work.",
    "The workload is unbearable and I can't keep up anymore.",
    "I am burned out and feel like I'm running on empty.",
    "There is too much pressure and I feel constantly stressed.",
    "I dread coming to work every single day.",
    "I have no energy left after work, even for my family.",
    "Management never listens to our concerns about the workload.",
    "I feel disconnected and disengaged from my job.",
    "Deadlines are impossible and I'm always behind.",
    "I feel unappreciated despite working extra hours every week.",
]

NEUTRAL_FEEDBACK = [
    "Work is fine, nothing special to report.",
    "Some days are tough but it's manageable overall.",
    "I have an average amount of tasks and I cope with them.",
    "The workload is sometimes challenging but okay.",
    "I feel neutral about my job; it has its ups and downs.",
    "Things are going okay but there is room for improvement.",
    "I handle my tasks but sometimes feel a bit tired.",
    "Work-life balance could be better, but it's acceptable.",
    "I am keeping up with my responsibilities for now.",
    "Nothing extraordinary; the job is routine and steady.",
]

POSITIVE_FEEDBACK = [
    "I love my job and feel very motivated every day.",
    "The work environment is fantastic and very supportive.",
    "I feel valued and recognized for my contributions.",
    "I have a great work-life balance and I'm very satisfied.",
    "My manager is supportive and the team is wonderful.",
    "I feel energized and productive in my current role.",
    "The workload is well-distributed and completely manageable.",
    "I enjoy challenging tasks and feel a sense of achievement.",
    "I am proud of what our team accomplishes together.",
    "Work is going great and I feel happily engaged.",
]

ALL_FEEDBACK = NEGATIVE_FEEDBACK + NEUTRAL_FEEDBACK + POSITIVE_FEEDBACK


def assign_feedback(burnout_score_raw: float) -> str:
    """
    Map raw burnout score to a feedback category with some randomness.
    Higher score → more likely to get negative feedback.
    """
    rand = random.random()
    if burnout_score_raw > 0.65:
        # High burnout → mostly negative, sometimes neutral
        if rand < 0.75:
            return random.choice(NEGATIVE_FEEDBACK)
        else:
            return random.choice(NEUTRAL_FEEDBACK)
    elif burnout_score_raw > 0.35:
        # Medium burnout → mix
        if rand < 0.40:
            return random.choice(NEGATIVE_FEEDBACK)
        elif rand < 0.75:
            return random.choice(NEUTRAL_FEEDBACK)
        else:
            return random.choice(POSITIVE_FEEDBACK)
    else:
        # Low burnout → mostly positive
        if rand < 0.70:
            return random.choice(POSITIVE_FEEDBACK)
        else:
            return random.choice(NEUTRAL_FEEDBACK)


def generate_dataset(n: int = N_EMPLOYEES) -> pd.DataFrame:
    """Generate synthetic employee burnout dataset."""

    # ── Raw feature generation ───────────────────────────────────────────────
    avg_daily_hours = np.random.uniform(6.0, 13.0, n)
    overtime_hours_week = np.random.uniform(0, 25, n)
    absence_days_month = np.random.randint(0, 9, n)           # 0–8
    late_arrivals_month = np.random.randint(0, 16, n)         # 0–15
    tasks_assigned = np.random.randint(5, 41, n)              # 5–40
    vacation_days_last_year = np.random.randint(0, 26, n)     # 0–25

    # ── Burnout score formula (weights based on known risk factors) ──────────
    # Normalize each feature to 0-1 range first
    h_norm  = (avg_daily_hours - 6) / 7          # daily hours weight
    ot_norm = overtime_hours_week / 25            # overtime weight
    ab_norm = absence_days_month / 8             # absence weight
    la_norm = late_arrivals_month / 15           # late arrivals weight
    ta_norm = tasks_assigned / 40               # tasks weight
    vac_norm = vacation_days_last_year / 25      # vacation weight (INVERSE risk)

    # Weighted combination → burnout_score ∈ (0, 1)
    burnout_score_raw = (
        0.30 * h_norm
        + 0.25 * ot_norm
        + 0.15 * ab_norm
        + 0.10 * la_norm
        + 0.15 * ta_norm
        - 0.05 * vac_norm   # More vacation → less burnout risk
    )

    # Add noise to prevent overfitting
    noise = np.random.normal(0, 0.10, n)
    burnout_score_noisy = np.clip(burnout_score_raw + noise, 0, 1)

    # Binary label with threshold
    threshold = 0.50
    burnout_label = (burnout_score_noisy > threshold).astype(int)

    # ── Build feedback column ────────────────────────────────────────────────
    feedback_col = [assign_feedback(s) for s in burnout_score_noisy]

    # ── Build name column (cycle through pool, add suffix to ensure uniqueness)
    names_col = []
    pool_size = len(EMPLOYEE_NAMES)
    for i in range(n):
        base = EMPLOYEE_NAMES[i % pool_size]
        if i >= pool_size:
            suffix = str(i // pool_size + 1)
            parts = base.split(" ")
            base = f"{parts[0]} {parts[1]} {suffix}" if len(parts) >= 2 else f"{base} {suffix}"
        names_col.append(base)

    # ── Assemble DataFrame ───────────────────────────────────────────────────
    df = pd.DataFrame({
        "employee_id":          [f"EMP{str(i+1).zfill(4)}" for i in range(n)],
        "name":                 names_col,
        "avg_daily_hours":      np.round(avg_daily_hours, 2),
        "overtime_hours_week":  np.round(overtime_hours_week, 2),
        "absence_days_month":   absence_days_month,
        "late_arrivals_month":  late_arrivals_month,
        "tasks_assigned":       tasks_assigned,
        "vacation_days_last_year": vacation_days_last_year,
        "feedback":             feedback_col,
        "burnout":              burnout_label,
    })

    return df


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Basic data cleaning pipeline."""
    print(f"[Cleaning] Shape before: {df.shape}")

    # Drop duplicates (on feature columns only, ignore id)
    feature_cols = [
        "avg_daily_hours", "overtime_hours_week", "absence_days_month",
        "late_arrivals_month", "tasks_assigned", "vacation_days_last_year",
        "feedback", "burnout"
    ]
    df = df.drop_duplicates(subset=feature_cols)

    # Drop rows with missing values (shouldn't exist for synthetic data)
    df = df.dropna()

    # Ensure correct dtypes
    df["absence_days_month"]     = df["absence_days_month"].astype(int)
    df["late_arrivals_month"]    = df["late_arrivals_month"].astype(int)
    df["tasks_assigned"]         = df["tasks_assigned"].astype(int)
    df["vacation_days_last_year"]= df["vacation_days_last_year"].astype(int)
    df["burnout"]                = df["burnout"].astype(int)

    # Clip numeric columns to expected ranges
    df["avg_daily_hours"]       = df["avg_daily_hours"].clip(6.0, 13.0)
    df["overtime_hours_week"]   = df["overtime_hours_week"].clip(0, 25)

    print(f"[Cleaning] Shape after:  {df.shape}")
    print(f"[Cleaning] Burnout distribution:\n{df['burnout'].value_counts()}")
    return df


if __name__ == "__main__":
    print("=" * 50)
    print("  Generating Synthetic Employee Burnout Dataset")
    print("=" * 50)

    # Generate
    df = generate_dataset(N_EMPLOYEES)
    print(f"\n[Generated] {len(df)} employee records")

    # Clean
    df = clean_data(df)

    # Save
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "employees.csv")
    df.to_csv(out_path, index=False)
    print(f"\n[Saved] Dataset -> {out_path}")
    print("\nFirst 5 rows:")
    print(df.head().to_string())
    print("\nDone! All files saved.")
