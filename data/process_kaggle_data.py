"""
process_kaggle_data.py
======================
Full Data Cleaning & Transformation Pipeline
=============================================
Source: Kaggle "Are Your Employees Burning Out?" (HackerEarth ML Challenge)
URL: https://www.kaggle.com/datasets/blurredmachine/are-your-employees-burning-out

This script performs:
    1. Load raw Kaggle data (22,750 records with nulls & issues)
    2. Data Cleaning: Drop nulls, fix types, remove outliers
    3. Stratified Sampling: Select 1000 representative employees
    4. Feature Engineering & Mapping: Convert Kaggle columns to WorkPulse schema
    5. NLP Text Fusion: Assign realistic feedback text based on burn rate
    6. Arabic Name Generation: Assign unique employee names
    7. Save final clean employees.csv (1000 employees, 10 columns)

Output:
    data/employees.csv  — Clean dataset ready for the WorkPulse system
"""

import os
import sys
import numpy as np
import pandas as pd
import random

# ── Reproducibility ──────────────────────────────────────────────────────────
SEED = 42
np.random.seed(SEED)
random.seed(SEED)

# ── Paths ────────────────────────────────────────────────────────────────────
ROOT = os.path.dirname(os.path.abspath(__file__))
RAW_PATH = os.path.join(ROOT, "kaggle_raw.csv")
OUT_PATH = os.path.join(ROOT, "employees.csv")

N_EMPLOYEES = 1000

# ── Employee Names Pool (Arabic transliterated, extended for 1000) ───────────
FIRST_NAMES_MALE = [
    "Ahmed", "Mohamed", "Omar", "Youssef", "Khaled", "Tariq", "Bilal", "Faris",
    "Rami", "Samir", "Nader", "Wael", "Hassan", "Ayman", "Sherif", "Bassam",
    "Hossam", "Emad", "Ihab", "Walid", "Adam", "Ziad", "Karim", "Mazen",
    "Sami", "Tarek", "Ramzi", "Galal", "Hazem", "Shadi", "Amr", "Tamer",
    "Mostafa", "Mahmoud", "Ibrahim", "Ali", "Hamza", "Osama", "Adel", "Magdi",
    "Ashraf", "Ehab", "Mohannad", "Seif", "Badr", "Yasser", "Hatem", "Nabil",
    "Fouad", "Gamal", "Fathy", "Sobhy", "Rafik", "Anwar", "Sameh", "Lotfy",
    "Wahid", "Atef", "Raafat", "Hany", "Essam", "Wafaa", "Karam", "Sharif",
]

FIRST_NAMES_FEMALE = [
    "Layla", "Nour", "Sara", "Fatima", "Rania", "Dina", "Yasmine", "Mariam",
    "Hana", "Salma", "Amira", "Doaa", "Noha", "Eman", "Ghada", "Mona",
    "Heba", "Nada", "Reem", "Mai", "Nadine", "Aya", "Rawan", "Lina",
    "Hoda", "Rana", "Dalia", "Omnia", "Engy", "Rasha", "Asmaa", "Samira",
    "Farida", "Habiba", "Nourhan", "Shahd", "Malak", "Jana", "Yara", "Lamia",
    "Nesma", "Abeer", "Sahar", "Maha", "Nawal", "Sawsan", "Soha", "Warda",
    "Nagwa", "Mervat", "Karima", "Aida", "Hanan", "Tahani", "Manal", "Sanaa",
]

LAST_NAMES = [
    "Mohamed", "Ali", "Hassan", "Ibrahim", "Mahmoud", "Saeed", "Nasser", "Amin",
    "Younis", "Fouad", "Hamid", "Karim", "Zaki", "Refaat", "Adel", "Tawfik",
    "Gamal", "Saleh", "Mostafa", "Nabil", "Adly", "Fathy", "Lotfy", "Wahid",
    "Sobhy", "Ezz", "Atef", "Raafat", "Hany", "Sherif", "Essam", "Ashraf",
    "Wafaa", "Magdi", "Shawki", "Karam", "Samir", "Ihab", "Osman", "Khalil",
    "Mansour", "Badawi", "Rashid", "Sultan", "Fahmy", "Darwish", "Farouk", "Morsi",
    "Haroun", "Gaber", "Soliman", "Helmy", "Barakat", "Shehata", "Taha", "Rizk",
]

# ── Feedback Sentences ───────────────────────────────────────────────────────
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
    "The stress at work is affecting my health seriously.",
    "I cannot sleep well because of work anxiety.",
    "Every day feels like a struggle to get through.",
    "I feel trapped in this job with no way out.",
    "My mental health is deteriorating because of work pressure.",
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
    "Some weeks are busier than others but I manage.",
    "The job has its challenges but nothing extreme.",
    "I wish there were fewer meetings but overall it is fine.",
    "My workload fluctuates but stays within reasonable limits.",
    "I feel okay about my current position and duties.",
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
    "I look forward to coming to work every morning.",
    "The company culture is amazing and inclusive.",
    "I feel my career is progressing in the right direction.",
    "My skills are being utilized well and I feel fulfilled.",
    "I am grateful for the opportunities this job provides.",
]


def assign_feedback(burn_rate: float) -> str:
    """
    Map burn rate (0.0 to 1.0) to feedback category with controlled randomness.
    Higher burn rate -> higher probability of negative feedback.
    """
    rand = random.random()
    if burn_rate > 0.60:
        if rand < 0.75:
            return random.choice(NEGATIVE_FEEDBACK)
        else:
            return random.choice(NEUTRAL_FEEDBACK)
    elif burn_rate > 0.35:
        if rand < 0.35:
            return random.choice(NEGATIVE_FEEDBACK)
        elif rand < 0.70:
            return random.choice(NEUTRAL_FEEDBACK)
        else:
            return random.choice(POSITIVE_FEEDBACK)
    else:
        if rand < 0.70:
            return random.choice(POSITIVE_FEEDBACK)
        else:
            return random.choice(NEUTRAL_FEEDBACK)


def generate_unique_names(n: int, genders: pd.Series) -> list:
    """Generate unique employee names based on gender from Kaggle data."""
    names = []
    used = set()
    
    for i in range(n):
        gender = genders.iloc[i] if i < len(genders) else random.choice(["Male", "Female"])
        
        attempts = 0
        while attempts < 100:
            if gender == "Female":
                first = random.choice(FIRST_NAMES_FEMALE)
            else:
                first = random.choice(FIRST_NAMES_MALE)
            last = random.choice(LAST_NAMES)
            full = f"{first} {last}"
            
            if full not in used:
                used.add(full)
                names.append(full)
                break
            attempts += 1
        else:
            # Add numeric suffix for uniqueness if pool exhausted
            suffix = len(used) + 1
            full = f"{first} {last} {suffix}"
            used.add(full)
            names.append(full)
    
    return names


def main():
    print("=" * 70)
    print("  WorkPulse — Kaggle Data Cleaning & Transformation Pipeline")
    print("  Source: HackerEarth 'Are Your Employees Burning Out?' Dataset")
    print("  Target: 1000 Clean Employee Records for WorkPulse System")
    print("=" * 70)

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 1: Load Raw Kaggle Data
    # ═══════════════════════════════════════════════════════════════════════
    print("\n[Step 1] Loading raw Kaggle data...")
    if not os.path.exists(RAW_PATH):
        print(f"  ERROR: Raw data not found at {RAW_PATH}")
        print("  Run: python data/download_kaggle.py first")
        sys.exit(1)
    
    df_raw = pd.read_csv(RAW_PATH)
    print(f"  Loaded: {len(df_raw)} rows x {len(df_raw.columns)} columns")
    print(f"  Columns: {list(df_raw.columns)}")
    print(f"  Missing values BEFORE cleaning:")
    for col in df_raw.columns:
        nulls = df_raw[col].isnull().sum()
        if nulls > 0:
            print(f"    - {col}: {nulls} nulls ({nulls/len(df_raw)*100:.1f}%)")

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 2: Data Cleaning — Handle Missing Values & Outliers
    # ═══════════════════════════════════════════════════════════════════════
    print("\n[Step 2] Cleaning data — dropping nulls in target & critical features...")
    
    # 2a. Drop rows where target (Burn Rate) is null — can't train without labels
    before = len(df_raw)
    df_clean = df_raw.dropna(subset=["Burn Rate"])
    print(f"  Dropped {before - len(df_clean)} rows with null Burn Rate")
    
    # 2b. Drop rows where critical features are null
    before = len(df_clean)
    df_clean = df_clean.dropna(subset=["Resource Allocation", "Mental Fatigue Score"])
    print(f"  Dropped {before - len(df_clean)} rows with null Resource/Fatigue")
    
    # 2c. Remove duplicates on feature columns
    before = len(df_clean)
    df_clean = df_clean.drop_duplicates(
        subset=["Designation", "Resource Allocation", "Mental Fatigue Score", "Burn Rate", "Gender"]
    )
    print(f"  Dropped {before - len(df_clean)} duplicate rows")
    
    # 2d. Clip outliers to valid ranges
    df_clean["Designation"] = df_clean["Designation"].clip(0, 5)
    df_clean["Resource Allocation"] = df_clean["Resource Allocation"].clip(1, 10)
    df_clean["Mental Fatigue Score"] = df_clean["Mental Fatigue Score"].clip(0, 10)
    df_clean["Burn Rate"] = df_clean["Burn Rate"].clip(0, 1)
    
    print(f"  After cleaning: {len(df_clean)} valid rows remain")

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 3: Stratified Sampling — Select 1000 Representative Employees
    # ═══════════════════════════════════════════════════════════════════════
    print(f"\n[Step 3] Stratified sampling of {N_EMPLOYEES} employees...")
    
    # Create binary burnout label for stratification
    df_clean["_burnout_bin"] = (df_clean["Burn Rate"] >= 0.5).astype(int)
    
    # Stratified sample to maintain class distribution
    from sklearn.model_selection import train_test_split
    
    if len(df_clean) > N_EMPLOYEES:
        df_sample, _ = train_test_split(
            df_clean, train_size=N_EMPLOYEES, 
            stratify=df_clean["_burnout_bin"],
            random_state=SEED
        )
    else:
        df_sample = df_clean.head(N_EMPLOYEES)
    
    df_sample = df_sample.reset_index(drop=True)
    burnout_pct = df_sample["_burnout_bin"].mean() * 100
    print(f"  Sampled: {len(df_sample)} employees")
    print(f"  Burnout distribution: {burnout_pct:.1f}% burnout, {100-burnout_pct:.1f}% normal")

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 4: Feature Engineering & Mapping (Kaggle → WorkPulse Schema)
    # ═══════════════════════════════════════════════════════════════════════
    print("\n[Step 4] Transforming Kaggle features to WorkPulse schema...")
    
    # 4a. avg_daily_hours: Derived from Resource Allocation (1-10)
    #     Mapping: RA=1 → ~6hrs (light), RA=10 → ~13hrs (extreme overwork)
    df_sample["avg_daily_hours"] = np.round(
        6.0 + (df_sample["Resource Allocation"] - 1) / 9 * 7.0 
        + np.random.normal(0, 0.3, len(df_sample)),
        2
    ).clip(6.0, 13.0)
    
    # 4b. overtime_hours_week: Correlated with Resource Allocation and Fatigue
    #     Higher resources + higher fatigue → more overtime
    df_sample["overtime_hours_week"] = np.round(
        (df_sample["Resource Allocation"] / 10 * 15) 
        + (df_sample["Mental Fatigue Score"] / 10 * 10) 
        + np.random.normal(0, 2, len(df_sample)),
        2
    ).clip(0, 25)
    
    # 4c. absence_days_month: Correlated with Mental Fatigue (burned out → more absent)
    df_sample["absence_days_month"] = np.round(
        df_sample["Mental Fatigue Score"] / 10 * 8 
        + np.random.normal(0, 1.5, len(df_sample))
    ).clip(0, 8).astype(int)
    
    # 4d. late_arrivals_month: Correlated with Fatigue and Burn Rate
    df_sample["late_arrivals_month"] = np.round(
        df_sample["Burn Rate"] * 15 
        + np.random.normal(0, 2, len(df_sample))
    ).clip(0, 15).astype(int)
    
    # 4e. tasks_assigned: Derived from Designation level (higher role = more tasks)
    df_sample["tasks_assigned"] = np.round(
        5 + (df_sample["Designation"] / 5 * 20) 
        + (df_sample["Resource Allocation"] / 10 * 15) 
        + np.random.normal(0, 3, len(df_sample))
    ).clip(5, 40).astype(int)
    
    # 4f. vacation_days_last_year: Inversely correlated with burn rate
    #     More burned out → less vacation taken (or vice versa)
    df_sample["vacation_days_last_year"] = np.round(
        25 * (1 - df_sample["Burn Rate"]) 
        + np.random.normal(0, 3, len(df_sample))
    ).clip(0, 25).astype(int)
    
    # 4g. burnout: Binary label from Kaggle's continuous Burn Rate
    df_sample["burnout"] = (df_sample["Burn Rate"] >= 0.5).astype(int)

    # 4h. gender & age: Direct demographic preservation & realistic generation
    #     Gender preserved directly from Kaggle dataset ("Male" / "Female")
    df_sample["gender"] = df_sample["Gender"]

    #     Age derived realistically from Designation (seniority level 0-5) + natural variation
    #     Designation 0 (Junior) ~24 yrs, Designation 2 (Mid) ~33 yrs, Designation 5 (Executive) ~50 yrs
    base_age = 23.0 + (df_sample["Designation"] * 5.2) + np.random.normal(0, 2.5, len(df_sample))
    df_sample["age"] = np.round(base_age).clip(22, 60).astype(int)
    
    print("  Feature mapping complete:")
    print(f"    gender:                {df_sample['gender'].value_counts().to_dict()}")
    print(f"    age:                   {df_sample['age'].min()} - {df_sample['age'].max()} yrs (mean: {df_sample['age'].mean():.1f})")
    print(f"    avg_daily_hours:       {df_sample['avg_daily_hours'].min():.1f} - {df_sample['avg_daily_hours'].max():.1f}")
    print(f"    overtime_hours_week:    {df_sample['overtime_hours_week'].min():.1f} - {df_sample['overtime_hours_week'].max():.1f}")
    print(f"    absence_days_month:     {df_sample['absence_days_month'].min()} - {df_sample['absence_days_month'].max()}")
    print(f"    late_arrivals_month:    {df_sample['late_arrivals_month'].min()} - {df_sample['late_arrivals_month'].max()}")
    print(f"    tasks_assigned:         {df_sample['tasks_assigned'].min()} - {df_sample['tasks_assigned'].max()}")
    print(f"    vacation_days_last_year:{df_sample['vacation_days_last_year'].min()} - {df_sample['vacation_days_last_year'].max()}")

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 5: NLP Text Fusion — Assign Feedback Based on Real Burn Rate
    # ═══════════════════════════════════════════════════════════════════════
    print("\n[Step 5] Generating NLP feedback text matched to burn rates...")
    df_sample["feedback"] = df_sample["Burn Rate"].apply(assign_feedback)
    
    # Count feedback categories
    neg_count = sum(1 for f in df_sample["feedback"] if f in NEGATIVE_FEEDBACK)
    pos_count = sum(1 for f in df_sample["feedback"] if f in POSITIVE_FEEDBACK)
    neu_count = len(df_sample) - neg_count - pos_count
    print(f"  Feedback distribution: Negative={neg_count}, Neutral={neu_count}, Positive={pos_count}")

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 6: Generate Unique Names & Employee IDs
    # ═══════════════════════════════════════════════════════════════════════
    print("\n[Step 6] Assigning unique Arabic names and employee IDs...")
    names = generate_unique_names(N_EMPLOYEES, df_sample["Gender"])
    df_sample["employee_id"] = [f"EMP{str(i+1).zfill(4)}" for i in range(N_EMPLOYEES)]
    df_sample["name"] = names
    
    print(f"  Generated {len(set(names))} unique names for {N_EMPLOYEES} employees")

    # ═══════════════════════════════════════════════════════════════════════
    # STEP 7: Final Assembly & Export
    # ═══════════════════════════════════════════════════════════════════════
    print("\n[Step 7] Assembling final dataset...")
    
    # Select only Tawazon schema columns in correct order
    final_columns = [
        "employee_id", "name", "gender", "age", "avg_daily_hours", "overtime_hours_week",
        "absence_days_month", "late_arrivals_month", "tasks_assigned",
        "vacation_days_last_year", "feedback", "burnout"
    ]
    df_final = df_sample[final_columns].copy()
    
    # Final validation
    assert len(df_final) == N_EMPLOYEES, f"Expected {N_EMPLOYEES} rows, got {len(df_final)}"
    assert df_final.isnull().sum().sum() == 0, "Final dataset has null values!"
    
    # Save
    df_final.to_csv(OUT_PATH, index=False)
    print(f"\n  SAVED: {OUT_PATH}")
    print(f"  Shape: {df_final.shape}")
    
    # ═══════════════════════════════════════════════════════════════════════
    # SUMMARY REPORT
    # ═══════════════════════════════════════════════════════════════════════
    print("\n" + "=" * 70)
    print("  PIPELINE COMPLETE — Dataset Description Summary")
    print("=" * 70)
    print(f"  Source:          Kaggle HackerEarth Employee Burnout Dataset")
    print(f"  Original Size:   22,750 records")
    print(f"  After Cleaning:  {len(df_clean)} records (removed nulls & duplicates)")
    print(f"  Final Sample:    {N_EMPLOYEES} employees (stratified sampling)")
    print(f"  Burnout Rate:    {df_final['burnout'].mean()*100:.1f}%")
    print(f"  Columns:         {len(final_columns)}")
    print(f"  Missing Values:  0 (fully clean)")
    print(f"\n  Descriptive Statistics:")
    print(df_final.describe().round(2).to_string())
    print("\n  First 5 rows:")
    print(df_final.head().to_string())
    print("\n  Done!")


if __name__ == "__main__":
    main()
