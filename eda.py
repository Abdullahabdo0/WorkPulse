"""
eda.py
======
Exploratory Data Analysis (EDA) for the burnout dataset.
Generates 6 publication-ready charts and saves them to models/ folder.

Run after generating data:
    python data/generate_data.py
    python eda.py
"""

import os
import warnings
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

warnings.filterwarnings("ignore")
sns.set_theme(style="whitegrid", palette="muted")

# ── Paths ──────────────────────────────────────────────────────────────────
ROOT      = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(ROOT, "data", "employees.csv")
OUT_DIR   = os.path.join(ROOT, "models")
os.makedirs(OUT_DIR, exist_ok=True)

NUMERIC_COLS = [
    "avg_daily_hours", "overtime_hours_week",
    "absence_days_month", "late_arrivals_month",
    "tasks_assigned", "vacation_days_last_year",
]

COLORS = {"0": "#22c55e", "1": "#ef4444"}

# ── Load data ──────────────────────────────────────────────────────────────
df = pd.read_csv(DATA_PATH)
print(f"[EDA] Loaded {len(df)} records")
print(f"[EDA] Burnout rate: {df['burnout'].mean()*100:.1f}%")
print(df.describe())


# ── 1. Class Distribution ──────────────────────────────────────────────────
plt.figure(figsize=(6, 4))
counts = df["burnout"].value_counts()
bars = plt.bar(["No Burnout (0)", "Burnout (1)"], counts.values,
               color=["#22c55e", "#ef4444"], edgecolor="white", linewidth=1.5,
               width=0.5)
for bar, val in zip(bars, counts.values):
    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
             f"{val}\n({val/len(df)*100:.1f}%)", ha="center", fontsize=10, fontweight="bold")
plt.title("1. Class Distribution — Burnout vs. No Burnout", fontsize=13, fontweight="bold")
plt.ylabel("Count")
plt.ylim(0, max(counts.values) * 1.2)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "eda_01_class_distribution.png"), dpi=120)
plt.close()
print("[EDA] Chart 1 saved")


# ── 2. Histograms of numeric features ──────────────────────────────────────
fig, axes = plt.subplots(2, 3, figsize=(14, 8))
axes = axes.flatten()
labels = [
    "Avg Daily Hours", "Overtime Hours/Week",
    "Absence Days/Month", "Late Arrivals/Month",
    "Tasks Assigned", "Vacation Days/Year"
]
for i, (col, label) in enumerate(zip(NUMERIC_COLS, labels)):
    axes[i].hist(df[df["burnout"]==0][col], bins=20, alpha=0.65, color="#22c55e", label="No Burnout")
    axes[i].hist(df[df["burnout"]==1][col], bins=20, alpha=0.65, color="#ef4444", label="Burnout")
    axes[i].set_title(label, fontsize=11, fontweight="bold")
    axes[i].set_xlabel(col)
    axes[i].set_ylabel("Count")
    axes[i].legend(fontsize=8)

fig.suptitle("2. Feature Distributions by Burnout Label", fontsize=14, fontweight="bold", y=1.01)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "eda_02_histograms.png"), dpi=120, bbox_inches="tight")
plt.close()
print("[EDA] Chart 2 saved")


# ── 3. Boxplots ────────────────────────────────────────────────────────────
fig, axes = plt.subplots(2, 3, figsize=(14, 8))
axes = axes.flatten()
for i, (col, label) in enumerate(zip(NUMERIC_COLS, labels)):
    data_0 = df[df["burnout"]==0][col]
    data_1 = df[df["burnout"]==1][col]
    bp = axes[i].boxplot([data_0, data_1], patch_artist=True,
                          tick_labels=["No Burnout", "Burnout"],
                          medianprops=dict(color="black", linewidth=2))
    bp["boxes"][0].set_facecolor("#22c55e")
    bp["boxes"][1].set_facecolor("#ef4444")
    for box in bp["boxes"]:
        box.set_alpha(0.7)
    axes[i].set_title(label, fontsize=11, fontweight="bold")
    axes[i].set_ylabel(col)

fig.suptitle("3. Boxplots — Feature Spread by Burnout Label", fontsize=14, fontweight="bold", y=1.01)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "eda_03_boxplots.png"), dpi=120, bbox_inches="tight")
plt.close()
print("[EDA] Chart 3 saved")


# ── 4. Correlation Heatmap ─────────────────────────────────────────────────
plt.figure(figsize=(9, 7))
corr = df[NUMERIC_COLS + ["burnout"]].corr()
mask = np.triu(np.ones_like(corr, dtype=bool))
sns.heatmap(corr, mask=mask, annot=True, fmt=".2f", cmap="RdYlGn_r",
            center=0, vmin=-1, vmax=1, linewidths=0.5,
            annot_kws={"size": 10})
plt.title("4. Correlation Heatmap (Features vs. Burnout)", fontsize=13, fontweight="bold")
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "eda_04_heatmap.png"), dpi=120)
plt.close()
print("[EDA] Chart 4 saved")


# ── 5. Burnout rate by Overtime bucket ────────────────────────────────────
df["overtime_bucket"] = pd.cut(df["overtime_hours_week"],
                                bins=[0, 5, 10, 15, 20, 25],
                                labels=["0-5", "5-10", "10-15", "15-20", "20-25"])
overtime_burnout = df.groupby("overtime_bucket", observed=True)["burnout"].mean() * 100

plt.figure(figsize=(7, 4))
bars = plt.bar(overtime_burnout.index.astype(str), overtime_burnout.values,
               color=["#22c55e", "#84cc16", "#f59e0b", "#f97316", "#ef4444"],
               edgecolor="white", linewidth=1.5, width=0.6)
for bar, val in zip(bars, overtime_burnout.values):
    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
             f"{val:.1f}%", ha="center", fontsize=10, fontweight="bold")
plt.title("5. Burnout Rate by Weekly Overtime Hours", fontsize=13, fontweight="bold")
plt.xlabel("Overtime Hours / Week")
plt.ylabel("Burnout Rate (%)")
plt.ylim(0, 100)
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "eda_05_overtime_burnout.png"), dpi=120)
plt.close()
print("[EDA] Chart 5 saved")


# ── 6. Scatter: Daily Hours vs Overtime, colored by burnout ───────────────
plt.figure(figsize=(8, 5))
for label, color, name in [(0, "#22c55e", "No Burnout"), (1, "#ef4444", "Burnout")]:
    subset = df[df["burnout"] == label]
    plt.scatter(subset["avg_daily_hours"], subset["overtime_hours_week"],
                alpha=0.4, s=20, color=color, label=name)
plt.xlabel("Avg Daily Hours")
plt.ylabel("Overtime Hours / Week")
plt.title("6. Daily Hours vs. Overtime — Colored by Burnout", fontsize=13, fontweight="bold")
plt.legend()
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, "eda_06_scatter.png"), dpi=120)
plt.close()
print("[EDA] Chart 6 saved")


print("\nEDA complete! All charts saved to models/")
