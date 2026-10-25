import os
import json
import sqlite3
import numpy as np

DB_PATH = "file:/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db?immutable=1"
OUTPUT_WEIGHTS = "/home/nilesh7757/BMPTUTORS1/app/fee_model_weights.json"

def train():
    print("Connecting to database and fetching training rows...")
    conn = sqlite3.connect(DB_PATH, uri=True)
    cur = conn.cursor()
    
    query = """
    SELECT 
        t.hourly_fee_avg,
        COALESCE(t.total_teaching_exp_years, 2.0) as exp,
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
    WHERE t.hourly_fee_avg >= 50 AND t.hourly_fee_avg <= 5000
    """
    cur.execute(query)
    rows = cur.fetchall()
    print(f"Total training records: {len(rows):,}")
    
    y = np.array([r[0] for r in rows], dtype=float)
    
    # Define categorical mappings
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
    
    # Feature names
    feature_names = ['intercept', 'exp_years', 'exp_sqrt', 'has_bed', 'online', 'home']
    for t in tiers[:-1]:
        feature_names.append(f"tier_{t}")
    for q in quals[:-1]:
        feature_names.append(f"qual_{q}")
    for o in occs[:-1]:
        feature_names.append(f"occ_{o}")
    for s in subjs[:-1]:
        feature_names.append(f"subj_{s}")
    for l in levels[:-1]:
        feature_names.append(f"lvl_{l}")
        
    num_features = len(feature_names)
    print(f"Number of engineered features: {num_features}")
    
    X = np.zeros((len(rows), num_features), dtype=float)
    
    for i, r in enumerate(rows):
        fee, exp, tier, qual, has_bed, occ, online, home, subj, lvl = r
        
        # Base numericals
        col = 0
        X[i, col] = 1.0  # intercept
        col += 1
        X[i, col] = min(max(float(exp), 0.0), 35.0)
        col += 1
        X[i, col] = np.sqrt(min(max(float(exp), 0.0), 35.0))
        col += 1
        X[i, col] = 1.0 if has_bed else 0.0
        col += 1
        X[i, col] = 1.0 if online else 0.0
        col += 1
        X[i, col] = 1.0 if home else 0.0
        col += 1
        
        # Tier
        for t in tiers[:-1]:
            if tier == t:
                X[i, col] = 1.0
            col += 1
            
        # Qual
        for q in quals[:-1]:
            if qual == q:
                X[i, col] = 1.0
            col += 1
            
        # Occ
        for o in occs[:-1]:
            if occ == o:
                X[i, col] = 1.0
            col += 1
            
        # Subj
        for s in subjs[:-1]:
            if subj == s:
                X[i, col] = 1.0
            col += 1
            
        # Level
        for l in levels[:-1]:
            if lvl == l:
                X[i, col] = 1.0
            col += 1
            
    # Ridge regression: w = (X^T X + lambda I)^(-1) X^T y
    print("Fitting Ridge Regression Model (lambda=25.0)...")
    lambda_reg = 25.0
    XTX = np.dot(X.T, X)
    XTX_reg = XTX + lambda_reg * np.eye(num_features)
    XTX_reg[0, 0] -= lambda_reg  # don't penalize intercept
    XTy = np.dot(X.T, y)
    weights = np.linalg.solve(XTX_reg, XTy)
    
    y_pred = np.dot(X, weights)
    residuals = y - y_pred
    mae = np.mean(np.abs(residuals))
    rmse = np.sqrt(np.mean(residuals**2))
    ss_tot = np.sum((y - np.mean(y))**2)
    ss_res = np.sum(residuals**2)
    r2 = 1.0 - (ss_res / ss_tot)
    
    print(f"Model Trained Successfully:")
    print(f"  R^2 Score: {r2*100:.2f}%")
    print(f"  Mean Absolute Error (MAE): INR {mae:.2f} / hr")
    print(f"  Root Mean Squared Error (RMSE): INR {rmse:.2f} / hr")
    print(f"  Residual Std Dev: INR {np.std(residuals):.2f} / hr")
    
    # Residual quantiles for empirical prediction intervals (e.g. 25% and 75%)
    res_p25 = float(np.percentile(residuals, 25))
    res_p75 = float(np.percentile(residuals, 75))
    res_std = float(np.std(residuals))
    
    # Save model weights and configuration
    weights_dict = {
        "metadata": {
            "trained_samples": len(rows),
            "mae": round(float(mae), 2),
            "rmse": round(float(rmse), 2),
            "r2_pct": round(float(r2 * 100), 2),
            "res_p25": round(res_p25, 2),
            "res_p75": round(res_p75, 2),
            "res_std": round(res_std, 2),
            "overall_median": round(float(np.median(y)), 2),
            "overall_mean": round(float(np.mean(y)), 2)
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
        
    print(f"Saved model weights to {OUTPUT_WEIGHTS}")
    
    # Print key feature weights
    print("\nTop Positive Feature Drivers:")
    sorted_w = sorted(zip(feature_names, weights), key=lambda x: x[1], reverse=True)
    for name, w in sorted_w[:12]:
        print(f"  {name:35s}: +INR {w:.2f}")

if __name__ == "__main__":
    train()
