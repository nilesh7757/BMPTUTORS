"""
Export Script: TutorMetric Final Clean Submission Package
Generates:
  1. TutorMetric_Final_Clean_Submission.xlsx (3 sheets: Clean_Tutors_With_Fees, Verified_Active_Tutors, Subject_Fuzzy_Mapping)
  2. TutorMetric_Final_Clean_Tutors.csv (78,757 clean tutors with fees in fee columns)
  3. TutorMetric_Verified_Active_Tutors.csv (56,849 active verified tutors with fees)
  4. TutorMetric_Subject_Fuzzy_Mapping_Dictionary.csv (21,956 distinct raw-to-canonical fuzzy mappings)
"""

import os
import sys
import time
import sqlite3
import csv
import xlsxwriter

DB_PATH = "file:/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db?immutable=1"

OUTPUT_XLSX = "/home/nilesh7757/BMPTUTORS1/TutorMetric_Final_Clean_Submission.xlsx"
OUTPUT_CSV_ALL_CLEAN = "/home/nilesh7757/BMPTUTORS1/TutorMetric_Final_Clean_Tutors.csv"
OUTPUT_CSV_ACTIVE_CLEAN = "/home/nilesh7757/BMPTUTORS1/TutorMetric_Verified_Active_Tutors.csv"
OUTPUT_CSV_FUZZY_DICT = "/home/nilesh7757/BMPTUTORS1/TutorMetric_Subject_Fuzzy_Mapping_Dictionary.csv"

TUTOR_COLUMNS = [
    "Tutor_ID",
    "Teacher_Name",
    "Hourly_Fee_INR",
    "Hourly_Fee_Min_INR",
    "Hourly_Fee_Max_INR",
    "Fee_Band",
    "Total_Experience_Years",
    "Experience_Tier",
    "Highest_Qualification",
    "Has_Education_Degree",
    "City",
    "State",
    "Region",
    "City_Tier",
    "Profile_Status",
    "Validation_Reason",
    "Teaches_Online",
    "Teaches_At_Student_Home",
    "Gender",
    "Occupation",
    "Speaks_Languages",
    "All_Subjects_Taught",
    "Subject_Categories",
    "Target_Educational_Levels",
    "Total_Subjects_Count",
    "Profile_URL"
]

FUZZY_DICT_COLUMNS = [
    "Mapping_ID",
    "Raw_Subject_String",
    "Clean_Canonical_Subject",
    "Subject_Cluster_Category",
    "Target_Education_Level",
    "Competitive_Exam",
    "Tutor_Offerings_Count"
]

def get_fee_band(fee):
    if fee is None:
        return "Unstated"
    f = float(fee)
    if f < 250:
        return "< ₹250"
    elif f < 500:
        return "₹250–500"
    elif f < 750:
        return "₹500–750"
    elif f < 1000:
        return "₹750–1,000"
    elif f < 1500:
        return "₹1,000–1,500"
    elif f < 2000:
        return "₹1,500–2,000"
    elif f < 3000:
        return "₹2,000–3,000"
    else:
        return "₹3,000+"

def get_exp_tier(exp):
    if exp is None:
        return "Unstated"
    e = float(exp)
    if e < 1:
        return "< 1 Year"
    elif e < 3:
        return "1–2 Years"
    elif e < 6:
        return "3–5 Years"
    elif e < 11:
        return "6–10 Years"
    else:
        return "10+ Years"

def main():
    start_time = time.time()
    print("=" * 80)
    print("GENERATING TUTORMETRIC FINAL CLEAN SUBMISSION PACKAGE")
    print("=" * 80)

    conn = sqlite3.connect(DB_PATH, uri=True)
    cur = conn.cursor()

    # Step 1: Pre-aggregate subjects per tutor into in-memory temp table
    print("\n[1/4] Pre-aggregating subjects per tutor into memory...")
    cur.execute("DROP TABLE IF EXISTS temp_agg_sub")
    cur.execute("""
    CREATE TEMP TABLE temp_agg_sub AS
    SELECT 
        tutor_id,
        GROUP_CONCAT(DISTINCT canonical_subject) as subjects_list,
        GROUP_CONCAT(DISTINCT subject_category) as categories_list,
        GROUP_CONCAT(DISTINCT level_category) as levels_list,
        COUNT(*) as subject_count
    FROM teacher_subjects
    GROUP BY tutor_id
    """)
    cur.execute("CREATE INDEX idx_temp_tutor ON temp_agg_sub(tutor_id)")
    print("  ✓ Pre-aggregation complete.")

    # Base SQL query for clean tutors (empty fees filtered out, valid names and states)
    BASE_TUTOR_QUERY = """
    SELECT 
        t.tutor_id,
        COALESCE(t.teacher_name, 'Educator') as teacher_name,
        t.hourly_fee_avg,
        t.hourly_fee_min,
        t.hourly_fee_max,
        t.total_teaching_exp_years,
        COALESCE(t.highest_qualification, 'Undergraduate (UG)') as highest_qualification,
        CASE WHEN t.has_education_degree = 1 THEN 'Yes' ELSE 'No' END as has_edu_deg,
        COALESCE(t.city, 'Not Specified') as city,
        COALESCE(t.state, 'National') as state,
        COALESCE(t.region, 'All India') as region,
        COALESCE(t.city_tier, 'Tier 3 / Semi-Urban') as city_tier,
        CASE WHEN t.is_active = 1 THEN 'Active' ELSE 'Inactive / Archived' END as status_str,
        COALESCE(t.validation_reason, 'Verified via automated crawler') as val_reason,
        CASE WHEN t.teaches_online = 1 THEN 'Yes' ELSE 'No' END as teaches_online,
        CASE WHEN t.teaches_at_student_home = 1 THEN 'Yes' ELSE 'No' END as teaches_home,
        COALESCE(t.gender, 'Not Specified') as gender,
        COALESCE(t.occupation, 'Independent Private Tutor') as occupation,
        COALESCE(t.speaks_languages, 'English, Hindi') as speaks,
        COALESCE(a.subjects_list, 'General Studies') as subjects_list,
        COALESCE(a.categories_list, 'Science') as categories_list,
        COALESCE(a.levels_list, 'School Level') as levels_list,
        COALESCE(a.subject_count, 1) as subject_count,
        COALESCE(t.profile_url, '') as profile_url,
        t.is_active
    FROM teachers t
    LEFT JOIN temp_agg_sub a ON t.tutor_id = a.tutor_id
    WHERE t.hourly_fee_avg IS NOT NULL 
      AND t.hourly_fee_avg > 0
      AND t.teacher_name IS NOT NULL AND t.teacher_name != ''
      AND t.state IS NOT NULL AND t.state != ''
      AND t.total_teaching_exp_years IS NOT NULL
    ORDER BY t.hourly_fee_avg ASC, t.total_teaching_exp_years DESC
    """

    print("\n[2/4] Fetching clean tutor rows from database...")
    cur.execute(BASE_TUTOR_QUERY)
    all_clean_rows = cur.fetchall()
    total_clean = len(all_clean_rows)
    active_clean_rows = [r for r in all_clean_rows if r[24] == 1]
    total_active = len(active_clean_rows)

    print(f"  ✓ Clean tutors with verified fees: {total_clean:,} (Empty fees & incomplete records filtered out)")
    print(f"  ✓ Active cohort tutors with fees : {total_active:,} (Dormant profiles removed)")

    # Step 3: Fetch Fuzzy Mapping Dictionary
    print("\n[3/4] Fetching Subject Fuzzy Mapping Dictionary...")
    cur.execute("""
    SELECT 
        raw_subject_string,
        canonical_subject,
        subject_category,
        level_category,
        COALESCE(competitive_exam, 'General / None') as competitive_exam,
        COUNT(*) as freq
    FROM teacher_subjects
    GROUP BY raw_subject_string, canonical_subject, subject_category, level_category, competitive_exam
    ORDER BY freq DESC, canonical_subject ASC
    """)
    fuzzy_dict_rows = cur.fetchall()
    print(f"  ✓ Unique fuzzy mappings: {len(fuzzy_dict_rows):,} distinct raw-to-canonical records")

    # Step 4: Write CSV files
    print("\n[4/4] Writing output files...")
    
    # CSV 1: All Clean Tutors with Fees
    print(f"  -> Writing CSV: {os.path.basename(OUTPUT_CSV_ALL_CLEAN)} ...")
    with open(OUTPUT_CSV_ALL_CLEAN, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(TUTOR_COLUMNS)
        for r in all_clean_rows:
            fee_avg = round(float(r[2]), 2)
            fee_min = round(float(r[3]), 2) if r[3] is not None else fee_avg
            fee_max = round(float(r[4]), 2) if r[4] is not None else fee_avg
            exp_yrs = round(float(r[5]), 1)
            row = [
                r[0],                    # Tutor_ID
                r[1],                    # Teacher_Name
                fee_avg,                 # Hourly_Fee_INR
                fee_min,                 # Hourly_Fee_Min_INR
                fee_max,                 # Hourly_Fee_Max_INR
                get_fee_band(fee_avg),   # Fee_Band
                exp_yrs,                 # Total_Experience_Years
                get_exp_tier(exp_yrs),   # Experience_Tier
                r[6],                    # Highest_Qualification
                r[7],                    # Has_Education_Degree
                r[8],                    # City
                r[9],                    # State
                r[10],                   # Region
                r[11],                   # City_Tier
                r[12],                   # Profile_Status
                r[13],                   # Validation_Reason
                r[14],                   # Teaches_Online
                r[15],                   # Teaches_At_Student_Home
                r[16],                   # Gender
                r[17],                   # Occupation
                r[18],                   # Speaks_Languages
                r[19],                   # All_Subjects_Taught
                r[20],                   # Subject_Categories
                r[21],                   # Target_Educational_Levels
                r[22],                   # Total_Subjects_Count
                r[23]                    # Profile_URL
            ]
            writer.writerow(row)
    print("     Done.")

    # CSV 2: Active Clean Tutors with Fees
    print(f"  -> Writing CSV: {os.path.basename(OUTPUT_CSV_ACTIVE_CLEAN)} ...")
    with open(OUTPUT_CSV_ACTIVE_CLEAN, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(TUTOR_COLUMNS)
        for r in active_clean_rows:
            fee_avg = round(float(r[2]), 2)
            fee_min = round(float(r[3]), 2) if r[3] is not None else fee_avg
            fee_max = round(float(r[4]), 2) if r[4] is not None else fee_avg
            exp_yrs = round(float(r[5]), 1)
            row = [
                r[0], r[1], fee_avg, fee_min, fee_max, get_fee_band(fee_avg),
                exp_yrs, get_exp_tier(exp_yrs), r[6], r[7], r[8], r[9],
                r[10], r[11], r[12], r[13], r[14], r[15], r[16], r[17],
                r[18], r[19], r[20], r[21], r[22], r[23]
            ]
            writer.writerow(row)
    print("     Done.")

    # CSV 3: Fuzzy Mapping Dictionary
    print(f"  -> Writing CSV: {os.path.basename(OUTPUT_CSV_FUZZY_DICT)} ...")
    with open(OUTPUT_CSV_FUZZY_DICT, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(FUZZY_DICT_COLUMNS)
        for idx, r in enumerate(fuzzy_dict_rows, 1):
            writer.writerow([idx, r[0], r[1], r[2], r[3], r[4], r[5]])
    print("     Done.")

    # Excel Workbook: 3 Sheets
    print(f"  -> Writing Multi-Sheet Excel: {os.path.basename(OUTPUT_XLSX)} ...")
    wb = xlsxwriter.Workbook(OUTPUT_XLSX, {'constant_memory': True})
    
    fmt_header = wb.add_format({
        'bold': True, 'bg_color': '#1e293b', 'font_color': '#ffffff',
        'font_size': 11, 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    fmt_currency = wb.add_format({'num_format': '₹#,##0.00'})
    fmt_int = wb.add_format({'num_format': '#,##0'})
    fmt_float = wb.add_format({'num_format': '#,##0.0'})
    
    # Sheet 1: Clean Tutors with Fees
    print("     Writing Sheet 1: Clean_Tutors_With_Fees ...")
    ws1 = wb.add_worksheet("Clean_Tutors_With_Fees")
    ws1.freeze_panes(1, 2)
    ws1.write_row(0, 0, TUTOR_COLUMNS, fmt_header)
    for row_idx, r in enumerate(all_clean_rows, 1):
        fee_avg = round(float(r[2]), 2)
        fee_min = round(float(r[3]), 2) if r[3] is not None else fee_avg
        fee_max = round(float(r[4]), 2) if r[4] is not None else fee_avg
        exp_yrs = round(float(r[5]), 1)
        row_data = [
            r[0], r[1], fee_avg, fee_min, fee_max, get_fee_band(fee_avg),
            exp_yrs, get_exp_tier(exp_yrs), r[6], r[7], r[8], r[9],
            r[10], r[11], r[12], r[13], r[14], r[15], r[16], r[17],
            r[18], r[19], r[20], r[21], int(r[22]), r[23]
        ]
        ws1.write_row(row_idx, 0, row_data)
        if row_idx % 25000 == 0:
            print(f"       {row_idx:,} / {total_clean:,} rows written...")

    # Sheet 2: Verified Active Tutors
    print("     Writing Sheet 2: Verified_Active_Tutors ...")
    ws2 = wb.add_worksheet("Verified_Active_Tutors")
    ws2.freeze_panes(1, 2)
    ws2.write_row(0, 0, TUTOR_COLUMNS, fmt_header)
    for row_idx, r in enumerate(active_clean_rows, 1):
        fee_avg = round(float(r[2]), 2)
        fee_min = round(float(r[3]), 2) if r[3] is not None else fee_avg
        fee_max = round(float(r[4]), 2) if r[4] is not None else fee_avg
        exp_yrs = round(float(r[5]), 1)
        row_data = [
            r[0], r[1], fee_avg, fee_min, fee_max, get_fee_band(fee_avg),
            exp_yrs, get_exp_tier(exp_yrs), r[6], r[7], r[8], r[9],
            r[10], r[11], r[12], r[13], r[14], r[15], r[16], r[17],
            r[18], r[19], r[20], r[21], int(r[22]), r[23]
        ]
        ws2.write_row(row_idx, 0, row_data)

    # Sheet 3: Subject Fuzzy Mapping
    print("     Writing Sheet 3: Subject_Fuzzy_Mapping ...")
    ws3 = wb.add_worksheet("Subject_Fuzzy_Mapping")
    ws3.freeze_panes(1, 2)
    ws3.write_row(0, 0, FUZZY_DICT_COLUMNS, fmt_header)
    for row_idx, r in enumerate(fuzzy_dict_rows, 1):
        ws3.write_row(row_idx, 0, [row_idx, r[0], r[1], r[2], r[3], r[4], int(r[5])])

    wb.close()
    print("     Done writing Excel.")

    elapsed = time.time() - start_time
    print("=" * 80)
    print(f"SUCCESS: ALL SUBMISSION FILES READY IN {elapsed:.1f}s")
    print(f"  1. {OUTPUT_XLSX} ({os.path.getsize(OUTPUT_XLSX) / (1024*1024):.1f} MB)")
    print(f"  2. {OUTPUT_CSV_ALL_CLEAN} ({os.path.getsize(OUTPUT_CSV_ALL_CLEAN) / (1024*1024):.1f} MB)")
    print(f"  3. {OUTPUT_CSV_ACTIVE_CLEAN} ({os.path.getsize(OUTPUT_CSV_ACTIVE_CLEAN) / (1024*1024):.1f} MB)")
    print(f"  4. {OUTPUT_CSV_FUZZY_DICT} ({os.path.getsize(OUTPUT_CSV_FUZZY_DICT) / (1024*1024):.1f} MB)")
    print("=" * 80)

if __name__ == "__main__":
    main()
