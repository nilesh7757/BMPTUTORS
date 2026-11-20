#!/usr/bin/env python3
"""
TutorMetric - 2-Table Database Exporter (All Tutors with Price)
==============================================================
Rules strictly applied:
1. Filter out ONLY rows where the price is completely empty / 0 (keeps all 78,757 tutors with stated fees).
2. Keep all rows even if an individual cell is empty / dash (no aggressive row-dropping).
3. Drop columns that are entirely empty across the entire dataset (e.g. monthly_fee_min/max/avg).
4. Strictly original database columns only (NO engineered extra columns like fee_band).
5. Clean 2-table relational structure:
   - Tab 1: Tutors (78,757 rows, 30 columns)
   - Tab 2: Subjects (306,035 rows, 7 columns)

Generates:
  1. TutorMetric_Clean_Database_2Tables.xlsx
  2. TutorMetric_Clean_Tutors.csv
  3. TutorMetric_Clean_Subjects.csv
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

def export():
    start_time = time.time()
    print("=" * 80)
    print("STARTING EXPORT: ALL TUTORS WITH STATED PRICE (RETAIN DATA EVEN IF CELL IS EMPTY)")
    print("=" * 80)

    if not os.path.exists(DB_PATH):
        print(f"Error: Database {DB_PATH} not found.")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Exact original columns for Table 1: Tutors
    # (Completely empty columns like monthly_fee_* and raw scraping artifacts excluded)
    tutor_cols = [
        "tutor_id",
        "teacher_name",
        "hourly_fee_avg",
        "hourly_fee_min",
        "hourly_fee_max",
        "fee_unit",
        "total_teaching_exp_years",
        "online_teaching_exp_years",
        "seniority_tier",
        "experience_role",
        "highest_qualification",
        "has_education_degree",
        "city",
        "state",
        "region",
        "city_tier",
        "gender",
        "occupation",
        "works_as",
        "teaches_online",
        "teaches_at_student_home",
        "can_travel",
        "homework_help",
        "speaks_languages",
        "num_languages",
        "registered_date",
        "last_login_date",
        "is_active",
        "validation_reason",
        "profile_url"
    ]

    # Exact original columns for Table 2: Subjects
    subject_cols = [
        "bridge_id",
        "tutor_id",
        "subject_index",
        "canonical_subject",
        "subject_category",
        "level_category",
        "competitive_exam"
    ]

    # 1. Fetch all tutors where price is present (>0)
    print("\n[1/4] Extracting all tutors with price (keeping rows even if any cell is empty)...")
    query_tutors = f"""
        SELECT {', '.join(tutor_cols)}
        FROM teachers
        WHERE hourly_fee_avg IS NOT NULL AND hourly_fee_avg > 0
        ORDER BY CAST(SUBSTR(tutor_id, 2) AS INTEGER);
    """
    c.execute(query_tutors)

    tutor_rows = []
    retained_tutor_ids = set()

    for r in c:
        row_vals = []
        for col in tutor_cols:
            val = r[col]
            if isinstance(val, str):
                val = val.strip()
            row_vals.append(val if val is not None else "")
        tutor_rows.append(row_vals)
        retained_tutor_ids.add(r["tutor_id"])

    total_tutors = len(tutor_rows)
    print(f"      ✓ Extracted {total_tutors:,} tutors with prices (0 dropped for partial empty cells).")

    # 2. Fetch subject offerings for all retained tutors
    print("\n[2/4] Extracting subject offerings for retained tutors...")
    query_subjects = f"""
        SELECT {', '.join(subject_cols)}
        FROM teacher_subjects
        ORDER BY bridge_id;
    """
    c.execute(query_subjects)

    subject_rows = []
    for r in c:
        tid = r["tutor_id"]
        if tid in retained_tutor_ids:
            row_vals = []
            for col in subject_cols:
                val = r[col]
                if isinstance(val, str):
                    val = val.strip()
                row_vals.append(val if val is not None else "")
            subject_rows.append(row_vals)

    total_subjects = len(subject_rows)
    print(f"      ✓ Extracted {total_subjects:,} subject offerings.")

    # 3. Write CSVs
    print(f"\n[3/4] Writing CSV files...")
    with open(OUT_TUTORS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(tutor_cols)
        writer.writerows(tutor_rows)
    print(f"      ✓ {OUT_TUTORS_CSV} written ({total_tutors:,} rows, {len(tutor_cols)} cols).")

    with open(OUT_SUBJECTS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(subject_cols)
        writer.writerows(subject_rows)
    print(f"      ✓ {OUT_SUBJECTS_CSV} written ({total_subjects:,} rows, {len(subject_cols)} cols).")

    # 4. Write Excel Workbook (2 Tabs)
    print(f"\n[4/4] Writing Multi-Sheet Excel ({OUT_EXCEL})...")
    wb = xlsxwriter.Workbook(OUT_EXCEL, {'constant_memory': True, 'strings_to_urls': False})

    header_fmt = wb.add_format({
        "bold": True,
        "bg_color": "#1E293B",
        "font_color": "#FFFFFF",
        "font_name": "Calibri",
        "font_size": 11,
        "border": 1
    })

    # Sheet 1: Tutors
    print(f"      Writing Tab 1: Tutors ({total_tutors:,} rows)...")
    ws_tutors = wb.add_worksheet("Tutors")
    ws_tutors.set_tab_color("#2563EB")
    for col_idx, h in enumerate(tutor_cols):
        ws_tutors.write(0, col_idx, h, header_fmt)

    for row_idx, row in enumerate(tutor_rows, start=1):
        for col_idx, val in enumerate(row):
            if val != "" and val is not None:
                ws_tutors.write(row_idx, col_idx, val)

    del tutor_rows

    # Sheet 2: Subjects
    print(f"      Writing Tab 2: Subjects ({total_subjects:,} rows)...")
    ws_subjs = wb.add_worksheet("Subjects")
    ws_subjs.set_tab_color("#059669")
    for col_idx, h in enumerate(subject_cols):
        ws_subjs.write(0, col_idx, h, header_fmt)

    for row_idx, row in enumerate(subject_rows, start=1):
        for col_idx, val in enumerate(row):
            if val != "" and val is not None:
                ws_subjs.write(row_idx, col_idx, val)

    del subject_rows

    wb.close()
    conn.close()

    elapsed = round(time.time() - start_time, 1)
    print("\n" + "=" * 80)
    print(f"SUCCESS: EXPORT COMPLETE IN {elapsed}s")
    for p in [OUT_EXCEL, OUT_TUTORS_CSV, OUT_SUBJECTS_CSV]:
        size_mb = os.path.getsize(p) / (1024 * 1024)
        print(f"  • {p} ({size_mb:.1f} MB)")
    print("=" * 80)

if __name__ == "__main__":
    export()
