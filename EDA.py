import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ── Load data ─────────────────────────────
pamap2_df = pd.read_pickle('pamap2_cleaned.pkl')
ppg_df    = pd.read_pickle('ppg_dalia_cleaned.pkl')

X_scaled  = np.load('X_scaled.npy')
hr_raw    = np.load('hr_raw.npy')

feature_names = ['HR Z-Score', 'HR Std', 'RMSSD',
                 'Activity', 'Accel Mean',
                 'Accel Std', 'HR Delta', 'HR Ratio']

# ── BASIC INFO ────────────────────────────
print("PAMAP2 rows:", len(pamap2_df))
print("PPG rows   :", len(ppg_df))
print("Combined samples:", len(X_scaled))

# ============================================================
# 1. HEART RATE DISTRIBUTION (BOTH DATASETS)
# ============================================================
plt.figure(figsize=(10,4))

plt.subplot(1,2,1)
plt.hist(pamap2_df['heart_rate'], bins=30)
plt.title("PAMAP2 HR Distribution")
plt.xlabel("BPM")

plt.subplot(1,2,2)
plt.hist(ppg_df['heart_rate'], bins=30)
plt.title("PPG HR Distribution")
plt.xlabel("BPM")

plt.tight_layout()
plt.show()


# ============================================================
# 2. HR VS ACTIVITY (BOTH DATASETS)
# ============================================================
plt.figure(figsize=(10,4))

plt.subplot(1,2,1)
pamap2_df.boxplot(column='heart_rate', by='activity')
plt.title("PAMAP2 HR vs Activity")
plt.suptitle("")

plt.subplot(1,2,2)
ppg_df.boxplot(column='heart_rate', by='activity')
plt.title("PPG HR vs Activity")
plt.suptitle("")

plt.tight_layout()
plt.show()


# ============================================================
# 3. ACTIVITY DISTRIBUTION (BOTH DATASETS)
# ============================================================
plt.figure(figsize=(10,4))

plt.subplot(1,2,1)
pamap2_df['activity'].value_counts().plot(kind='bar')
plt.title("PAMAP2 Activity Count")

plt.subplot(1,2,2)
ppg_df['activity'].value_counts().plot(kind='bar')
plt.title("PPG Activity Count")

plt.tight_layout()
plt.show()


# ============================================================
# 4. 8 FEATURES (COMBINED DATASET - MAIN PART)
# ============================================================
plt.figure(figsize=(12,5))

for i in range(8):
    plt.subplot(2,4,i+1)
    plt.hist(X_scaled[:, i], bins=30)
    plt.title(feature_names[i], fontsize=9)

plt.suptitle("Feature Distribution (Combined Dataset)")
plt.tight_layout()
plt.show()


# ============================================================
# 5. RAW HR vs Z-SCORE (COMBINED DATASET)
# ============================================================
plt.figure()

plt.hist(hr_raw, bins=30, alpha=0.6, label='Raw HR')
plt.hist(X_scaled[:, 0], bins=30, alpha=0.6, label='Z-score')

plt.title("Raw HR vs Z-score (Combined)")
plt.xlabel("Value")
plt.ylabel("Count")
plt.legend()

plt.show()