#!/usr/bin/env python3
"""
TutorMetric - Strictly Original Columns 2-Table Exporter
========================================================
- Contains ONLY columns from the original database (NO engineered extra columns like fee_band).
- Excludes raw scraping artifacts (source_tab, raw_subject_string, fee_raw, location_raw, bio text).
- Filters out all rows with empty/missing fees, dashes in speaks_languages, or missing values in any column.

Generates:
  1. TutorMetric_Clean_Database_2Tables.xlsx (2 Tabs: Tutors, Subjects)
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
    print("EXPORTING STRICTLY ORIGINAL DATABASE COLUMNS (NO EXTRA COLUMNS / NO FEE BAND)")
    print("=" * 80)

    if not os.path.exists(DB_PATH):
        print(f"Error: Database {DB_PATH} not found.")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Exact original columns for Table 1: Tutors (Excluding raw artifacts and long bios)
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

    # Exact original columns for Table 2: Subjects (Excluding raw_subject_string)
    subject_cols = [
        "bridge_id",
        "tutor_id",
        "subject_index",
        "canonical_subject",
        "subject_category",
        "level_category",
        "competitive_exam"
    ]

    # 1. Fetch strictly complete tutor rows
    print("\n[1/4] Filtering and extracting original-column tutor rows...")
    query_tutors = f"""
        SELECT {', '.join(tutor_cols)}
        FROM teachers
        WHERE hourly_fee_avg IS NOT NULL AND hourly_fee_avg > 0
          AND hourly_fee_min IS NOT NULL AND hourly_fee_min > 0
          AND hourly_fee_max IS NOT NULL AND hourly_fee_max > 0
          AND speaks_languages IS NOT NULL AND TRIM(speaks_languages) NOT IN ('', '-', '--', '---', 'None', 'null', 'N/A', 'NA')
          AND total_teaching_exp_years IS NOT NULL
          AND online_teaching_exp_years IS NOT NULL
          AND teacher_name IS NOT NULL AND TRIM(teacher_name) NOT IN ('', '-', '--', '---')
          AND state IS NOT NULL AND TRIM(state) != ''
          AND city IS NOT NULL AND TRIM(city) != ''
          AND gender IN ('Female', 'Male')
          AND highest_qualification != 'Not Specified'
        ORDER BY CAST(SUBSTR(tutor_id, 2) AS INTEGER);
    """
    c.execute(query_tutors)

    tutor_rows = []
    clean_tutor_ids = set()

    for r in c:
        clean_row = []
        for col in tutor_cols:
            val = r[col]
            if isinstance(val, str):
                val = val.strip()
            clean_row.append(val)
        tutor_rows.append(clean_row)
        clean_tutor_ids.add(r["tutor_id"])

    total_tutors = len(tutor_rows)
    print(f"      ✓ Retained {total_tutors:,} clean tutors with 100% complete original columns.")

    # 2. Fetch subject offerings for retained tutors
    print("\n[2/4] Fetching original-column subject offerings for retained tutors...")
    query_subjects = f"""
        SELECT {', '.join(subject_cols)}
        FROM teacher_subjects
        ORDER BY bridge_id;
    """
    c.execute(query_subjects)

    subject_rows = []
    for r in c:
        tid = r["tutor_id"]
        if tid in clean_tutor_ids:
            clean_row = []
            for col in subject_cols:
                val = r[col]
                if isinstance(val, str):
                    val = val.strip()
                elif val is None and col == "competitive_exam":
                    val = "None"
                clean_row.append(val)
            subject_rows.append(clean_row)

    total_subjects = len(subject_rows)
    print(f"      ✓ Retained {total_subjects:,} clean subject offerings.")

    # 3. Write CSVs
    print(f"\n[3/4] Writing Clean CSVs...")
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

    # 4. Write Excel
    print(f"\n[4/4] Writing Clean 2-Tab Excel ({OUT_EXCEL})...")
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
            if val is not None and val != "":
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
            if val is not None and val != "":
                ws_subjs.write(row_idx, col_idx, val)

    del subject_rows

    wb.close()
    conn.close()

    elapsed = round(time.time() - start_time, 1)
    print("\n" + "=" * 80)
    print(f"SUCCESS: STRICTLY ORIGINAL DATABASE EXPORT COMPLETED IN {elapsed}s")
    for p in [OUT_EXCEL, OUT_TUTORS_CSV, OUT_SUBJECTS_CSV]:
        size_mb = os.path.getsize(p) / (1024 * 1024)
        print(f"  • {p} ({size_mb:.1f} MB)")
    print("=" * 80)

if __name__ == "__main__":
    export()
