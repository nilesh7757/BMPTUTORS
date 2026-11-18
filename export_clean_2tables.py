#!/usr/bin/env python3
"""
TutorMetric - Clean 2-Table Database Exporter
============================================
Creates a pristine, 2-tab Excel workbook and clean CSV files with:
- NO raw/scraping artifacts (source_tab, raw_subject_string, fee_raw, location_raw, etc. removed)
- NO long text bio/hookline clutter
- Dedicated clean numeric fees and fee bands
- Clean relational structure:
    Tab 1: Tutors (107,673 rows)
    Tab 2: Subjects (403,221 rows) linked by Tutor_ID
"""

import os
import sys
import time
import sqlite3
import csv
import xlsxwriter

DB_PATH = "cleaned_tutors_data.db"
OUT_EXCEL = "TutorMetric_Clean_Database_2Tables.xlsx"
OUT_TUTORS_CSV = "TutorMetric_Clean_Tutors.csv"
OUT_SUBJECTS_CSV = "TutorMetric_Clean_Subjects.csv"

def get_fee_band(fee):
    if fee is None or fee <= 0:
        return "Unstated / Negotiable"
    elif fee < 250:
        return "< ₹250"
    elif fee < 500:
        return "₹250–₹500"
    elif fee < 1000:
        return "₹500–₹1,000"
    elif fee < 2000:
        return "₹1,000–₹2,000"
    else:
        return "₹2,000+"

def export_clean_tables():
    start_time = time.time()
    print("=" * 80)
    print("EXPORTING CLEAN 2-TABLE DATABASE (NO RAW ARTIFACTS / NO UNNECESSARY CLUTTER)")
    print("=" * 80)

    if not os.path.exists(DB_PATH):
        print(f"Error: Database {DB_PATH} not found.")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Define Clean Columns for Tab 1: Tutors
    tutor_headers = [
        "Tutor_ID",
        "Teacher_Name",
        "Hourly_Fee_INR",
        "Hourly_Fee_Min_INR",
        "Hourly_Fee_Max_INR",
        "Fee_Band",
        "Total_Experience_Years",
        "Online_Experience_Years",
        "Experience_Tier",
        "Highest_Qualification",
        "Has_Education_Degree",
        "City",
        "State",
        "Region",
        "City_Tier",
        "Gender",
        "Occupation",
        "Teaches_Online",
        "Teaches_At_Student_Home",
        "Speaks_Languages",
        "Profile_Status",
        "Validation_Reason",
        "Profile_URL"
    ]

    # Define Clean Columns for Tab 2: Subjects
    subject_headers = [
        "Bridge_ID",
        "Tutor_ID",
        "Subject_Index",
        "Canonical_Subject",
        "Subject_Category",
        "Target_Education_Level",
        "Competitive_Exam"
    ]

    # --- 1. Export Clean Tutors CSV ---
    print(f"\n[1/3] Generating Clean Tutors CSV ({OUT_TUTORS_CSV})...")
    c.execute("""
        SELECT tutor_id, teacher_name, hourly_fee_avg, hourly_fee_min, hourly_fee_max,
               total_teaching_exp_years, online_teaching_exp_years, seniority_tier, highest_qualification,
               has_education_degree, city, state, region, city_tier, gender, occupation,
               teaches_online, teaches_at_student_home, speaks_languages, is_active, validation_reason, profile_url
        FROM teachers
        ORDER BY CAST(SUBSTR(tutor_id, 2) AS INTEGER);
    """)

    tutor_rows = []
    with open(OUT_TUTORS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(tutor_headers)
        for r in c:
            fee = r["hourly_fee_avg"]
            fee_band = get_fee_band(fee)
            clean_row = [
                r["tutor_id"],
                r["teacher_name"],
                fee if fee is not None else "",
                r["hourly_fee_min"] if r["hourly_fee_min"] is not None else "",
                r["hourly_fee_max"] if r["hourly_fee_max"] is not None else "",
                fee_band,
                r["total_teaching_exp_years"],
                r["online_teaching_exp_years"],
                r["seniority_tier"],
                r["highest_qualification"],
                "Yes" if r["has_education_degree"] == 1 else "No",
                r["city"],
                r["state"],
                r["region"],
                r["city_tier"],
                r["gender"] or "Unstated",
                r["occupation"],
                "Yes" if r["teaches_online"] == 1 else "No",
                "Yes" if r["teaches_at_student_home"] == 1 else "No",
                r["speaks_languages"],
                "Active" if r["is_active"] == 1 else "Inactive",
                r["validation_reason"],
                r["profile_url"]
            ]
            writer.writerow(clean_row)
            tutor_rows.append(clean_row)
    print(f"      ✓ Exported {len(tutor_rows):,} clean tutor rows.")

    # --- 2. Export Clean Subjects CSV ---
    print(f"\n[2/3] Generating Clean Subjects CSV ({OUT_SUBJECTS_CSV})...")
    c.execute("""
        SELECT bridge_id, tutor_id, subject_index, canonical_subject, subject_category, level_category, competitive_exam
        FROM teacher_subjects
        ORDER BY bridge_id;
    """)

    subject_rows = []
    with open(OUT_SUBJECTS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(subject_headers)
        for r in c:
            clean_row = [
                r["bridge_id"],
                r["tutor_id"],
                r["subject_index"],
                r["canonical_subject"],
                r["subject_category"],
                r["level_category"],
                r["competitive_exam"] or "None"
            ]
            writer.writerow(clean_row)
            subject_rows.append(clean_row)
    print(f"      ✓ Exported {len(subject_rows):,} clean subject offering rows.")

    # --- 3. Export Clean Multi-Sheet Excel (2 Tabs) ---
    print(f"\n[3/3] Generating Clean 2-Tab Excel ({OUT_EXCEL})...")
    wb = xlsxwriter.Workbook(OUT_EXCEL, {'constant_memory': True, 'strings_to_urls': False})

    header_fmt = wb.add_format({
        "bold": True,
        "bg_color": "#1E293B",
        "font_color": "#FFFFFF",
        "font_name": "Calibri",
        "font_size": 11,
        "border": 1
    })

    # Tab 1: Tutors
    print("      Writing Tab 1: Tutors (107,673 rows)...")
    ws_tutors = wb.add_worksheet("Tutors")
    ws_tutors.set_tab_color("#2563EB")
    for col_idx, h in enumerate(tutor_headers):
        ws_tutors.write(0, col_idx, h, header_fmt)

    for row_idx, row in enumerate(tutor_rows, start=1):
        for col_idx, val in enumerate(row):
            if val != "" and val is not None:
                ws_tutors.write(row_idx, col_idx, val)

    # Free memory
    del tutor_rows

    # Tab 2: Subjects
    print("      Writing Tab 2: Subjects (403,221 rows)...")
    ws_subjs = wb.add_worksheet("Subjects")
    ws_subjs.set_tab_color("#059669")
    for col_idx, h in enumerate(subject_headers):
        ws_subjs.write(0, col_idx, h, header_fmt)

    for row_idx, row in enumerate(subject_rows, start=1):
        for col_idx, val in enumerate(row):
            if val != "" and val is not None:
                ws_subjs.write(row_idx, col_idx, val)

    # Free memory
    del subject_rows

    print("      Saving workbook...")
    wb.close()
    conn.close()

    elapsed = round(time.time() - start_time, 1)
    print("\n" + "=" * 80)
    print(f"SUCCESS: CLEAN 2-TABLE DATABASE READY IN {elapsed}s")
    for p in [OUT_EXCEL, OUT_TUTORS_CSV, OUT_SUBJECTS_CSV]:
        size_mb = os.path.getsize(p) / (1024 * 1024)
        print(f"  • {p} ({size_mb:.1f} MB)")
    print("=" * 80)

if __name__ == "__main__":
    export_clean_tables()
