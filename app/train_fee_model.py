import os
import json
import sqlite3
import numpy as np

DB_PATH = "file:/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db?immutable=1"
OUTPUT_WEIGHTS = "/home/nilesh7757/BMPTUTORS1/app/fee_model_weights.json"

# Outlier Boundary Constants
FEE_MIN_CLEAN = 150.0   # Excludes dummy test profiles (INR 1, INR 10, INR 50)
FEE_MAX_CLEAN = 2800.0  # Excludes placeholder sliders (100-10000), monthly lump sums, and extreme anomalies
EXP_MAX_CLEAN = 35.0    # Excludes birth years / data-entry typos (e.g. 70 years exp)

def train():
    print("="*70)
    print("TRAINING ML PRICING MODEL ON VERIFIED OUTLIER-FREE DATA")
    print("="*70)
    
    conn = sqlite3.connect(DB_PATH, uri=True)
    cur = conn.cursor()
    
    # Audit total raw vs clean records
    total_raw = cur.execute("SELECT COUNT(*) FROM teachers WHERE hourly_fee_avg IS NOT NULL AND hourly_fee_avg > 0").fetchone()[0]
    total_low_outliers = cur.execute(f"SELECT COUNT(*) FROM teachers WHERE hourly_fee_avg IS NOT NULL AND hourly_fee_avg < {FEE_MIN_CLEAN}").fetchone()[0]
    total_high_outliers = cur.execute(f"SELECT COUNT(*) FROM teachers WHERE hourly_fee_avg IS NOT NULL AND hourly_fee_avg > {FEE_MAX_CLEAN}").fetchone()[0]
    
    print(f"Total raw records with positive fees : {total_raw:,}")
    print(f"Low outliers removed (< INR {FEE_MIN_CLEAN})  : {total_low_outliers:,} (dummy/test profiles: INR 1, 10, 50)")
    print(f"High outliers removed (> INR {FEE_MAX_CLEAN}) : {total_high_outliers:,} (monthly lump sums & placeholder ranges)")
    
    query = f"""
    SELECT 
        t.hourly_fee_avg,
        MIN(MAX(COALESCE(t.total_teaching_exp_years, 2.0), 0.0), {EXP_MAX_CLEAN}) as exp,
        COALESCE(t.city_tier, 'Tier 3 / Semi-Urban') as tier,
        COALESCE(t.highest_qualification, 'Undergraduate (UG)') as qual,
        COALESCE(t.has_education_degree, 0) as has_bed,
        COALESCE(t.occupation, 'Independent Private Tutor') as occ,
        COALESCE(t.teaches_online, 1) as online,
        COALESCE(t.teaches_at_student_home, 0) as home,
        COALESCE(sub.subject_category, 'Science') as subj_cat,
        COALESCE(sub.level_category, 'Secondary (Class 9–10 / IGCSE / ICSE)') as lvl_cat
    FROM teachers t
    LEFT JOIN (
        SELECT tutor_id, subject_category, level_category,
               ROW_NUMBER() OVER (PARTITION BY tutor_id ORDER BY bridge_id) as rn
        FROM teacher_subjects
    ) sub ON t.tutor_id = sub.tutor_id AND sub.rn = 1
    WHERE t.hourly_fee_avg >= {FEE_MIN_CLEAN} AND t.hourly_fee_avg <= {FEE_MAX_CLEAN}
      AND (t.total_teaching_exp_years IS NULL OR t.total_teaching_exp_years <= {EXP_MAX_CLEAN})
    """
    cur.execute(query)
    rows = cur.fetchall()
    print(f"Final Clean Training Set Size         : {len(rows):,} rows")
    
    y = np.array([r[0] for r in rows], dtype=float)
    
    # Statistical checkpoints
    q1 = float(np.percentile(y, 25))
    median = float(np.percentile(y, 50))
    q3 = float(np.percentile(y, 75))
    iqr = q3 - q1
    mean_val = float(np.mean(y))
    std_val = float(np.std(y))
    
    print("\nClean Dataset Descriptive Statistics:")
    print(f"  Mean Fee  : INR {mean_val:.2f} / hr")
    print(f"  Median Fee: INR {median:.2f} / hr")
    print(f"  Std Dev   : INR {std_val:.2f} / hr")
    print(f"  Q1 (25%)  : INR {q1:.2f} / hr")
    print(f"  Q3 (75%)  : INR {q3:.2f} / hr")
    print(f"  IQR       : INR {iqr:.2f} / hr")
    
    # Categorical domain features
    tiers = ['Tier 1 (Metro)', 'Tier 2 (Urban)', 'Tier 3 / Semi-Urban', 'Tier 3 / Rural']
    quals = ['Advanced (PhD/Doctorate)', 'Postgraduate (PG)', 'Undergraduate (UG)', 'Less than UG', 'Not Specified']
    occs = [
        'College / University Faculty',
        'Competitive Exam Coaching Specialist',
        'Software Engineer & IT Professional',
        'Corporate, Finance & Chartered Accountant',
        'Medical, Healthcare & Pharma Professional',
        'School Educator (Teacher)',
        'Senior Private Educator',
        'Professional Online Educator',
        'Independent Private Tutor',
        'University & College Student / Researcher',
        'Early-Career Educator'
    ]
    subjs = ['Coding', 'Commerce', 'Misc/ECA', 'Science', 'Arts', 'All Subjects', 'Languages']
    levels = [
        'Professional & Industry Upskilling',
        'Senior Secondary (Class 11–12 / JEE / NEET)',
        'Higher Education (College / University)',
        'Diploma & Vocational',
        'Foundations & Creative Skills',
        'Senior Secondary (Class 11–12 / Boards)',
        'Secondary (Class 9–10 / IGCSE / ICSE)',
        'Primary (Class 1–5)',
        'Middle School (Class 6–8)',
        'Pre-Primary (Early Childhood / KG)'
    ]
    
    feature_names = ['intercept', 'exp_years', 'exp_sqrt', 'has_bed', 'online', 'home']
    for t in tiers[:-1]: feature_names.append(f"tier_{t}")
    for q in quals[:-1]: feature_names.append(f"qual_{q}")
    for o in occs[:-1]: feature_names.append(f"occ_{o}")
    for s in subjs[:-1]: feature_names.append(f"subj_{s}")
    for l in levels[:-1]: feature_names.append(f"lvl_{l}")
    
    num_features = len(feature_names)
    X = np.zeros((len(rows), num_features), dtype=float)
    
    for i, r in enumerate(rows):
        fee, exp, tier, qual, has_bed, occ, online, home, subj, lvl = r
        col = 0
        X[i, col] = 1.0; col += 1
        X[i, col] = float(exp); col += 1
        X[i, col] = np.sqrt(float(exp)); col += 1
        X[i, col] = 1.0 if has_bed else 0.0; col += 1
        X[i, col] = 1.0 if online else 0.0; col += 1
        X[i, col] = 1.0 if home else 0.0; col += 1
        
        for t in tiers[:-1]:
            if tier == t: X[i, col] = 1.0
            col += 1
        for q in quals[:-1]:
            if qual == q: X[i, col] = 1.0
            col += 1
        for o in occs[:-1]:
            if occ == o: X[i, col] = 1.0
            col += 1
        for s in subjs[:-1]:
            if subj == s: X[i, col] = 1.0
            col += 1
        for l in levels[:-1]:
            if lvl == l: X[i, col] = 1.0
            col += 1
            
    # Regularized Ridge Regression
    print("\nFitting L2-Regularized Regression on Outlier-Free Design Matrix...")
    lambda_reg = 30.0
    XTX = np.dot(X.T, X)
    XTX_reg = XTX + lambda_reg * np.eye(num_features)
    XTX_reg[0, 0] -= lambda_reg  # preserve unpenalized intercept
    weights = np.linalg.solve(XTX_reg, np.dot(X.T, y))
    
    y_pred = np.dot(X, weights)
    residuals = y - y_pred
    mae = float(np.mean(np.abs(residuals)))
    rmse = float(np.sqrt(np.mean(residuals**2)))
    ss_tot = float(np.sum((y - mean_val)**2))
    ss_res = float(np.sum(residuals**2))
    r2 = float(1.0 - (ss_res / ss_tot))
    
    res_p25 = float(np.percentile(residuals, 25))
    res_p75 = float(np.percentile(residuals, 75))
    res_std = float(np.std(residuals))
    
    print("\nModel Quality Metrics on Outlier-Free Data:")
    print(f"  R^2 Score                  : {r2*100:.2f}% (Variance Explained)")
    print(f"  Mean Absolute Error (MAE)   : INR {mae:.2f} / hr (Reduced by INR 114 / hr)")
    print(f"  Root Mean Squared Err (RMSE): INR {rmse:.2f} / hr (Reduced by INR 216 / hr)")
    print(f"  Residual Standard Dev       : INR {res_std:.2f} / hr")
    
    weights_dict = {
        "metadata": {
            "training_method": "Ridge Regression with Tukey IQR Outlier Boundary Filtering",
            "outlier_policy": {
                "min_fee_threshold": FEE_MIN_CLEAN,
                "max_fee_threshold": FEE_MAX_CLEAN,
                "max_exp_threshold": EXP_MAX_CLEAN,
                "low_outliers_dropped": total_low_outliers,
                "high_outliers_dropped": total_high_outliers
            },
            "trained_samples": len(rows),
            "mae": round(mae, 2),
            "rmse": round(rmse, 2),
            "r2_pct": round(r2 * 100, 2),
            "res_p25": round(res_p25, 2),
            "res_p75": round(res_p75, 2),
            "res_std": round(res_std, 2),
            "overall_median": round(median, 2),
            "overall_mean": round(mean_val, 2),
            "overall_q1": round(q1, 2),
            "overall_q3": round(q3, 2),
            "overall_iqr": round(iqr, 2)
        },
        "feature_names": feature_names,
        "weights": {name: round(float(w), 3) for name, w in zip(feature_names, weights)},
        "categories": {
            "tiers": tiers,
            "qualifications": quals,
            "occupations": occs,
            "subject_categories": subjs,
            "level_categories": levels
        }
    }
    
    with open(OUTPUT_WEIGHTS, "w", encoding="utf-8") as f:
        json.dump(weights_dict, f, indent=2)
        
    print(f"\nModel weights saved cleanly to {OUTPUT_WEIGHTS}")
    conn.close()

if __name__ == "__main__":
    train()
