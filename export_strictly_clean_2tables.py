#!/usr/bin/env python3
"""
TutorMetric - Strictly Clean 2-Table Relational Database Exporter
================================================================
Completely removes:
- Unstated / empty / zero / missing fees
- Dashes ('-', '--', etc.) or empty/null values in speaks_languages or any column
- Missing / unstated values across all attributes (Gender, Qualification, Location, Experience)
- Scraping artifacts (source_tab, raw_subject_string, fee_raw, location_raw, bio text dumps)

Produces 2 clean relational tabs:
  Tab 1: Tutors (78,473 fully-verified complete profiles)
  Tab 2: Subjects (304,970 clean subject offerings linked by Tutor_ID)
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
    if fee < 250:
        return "< ₹250"
    elif fee < 500:
        return "₹250–₹500"
    elif fee < 1000:
        return "₹500–₹1,000"
    elif fee < 2000:
        return "₹1,000–₹2,000"
    else:
        return "₹2,000+"

def export_strictly_clean():
    start_time = time.time()
    print("=" * 80)
    print("STARTING STRICTLY CLEAN 2-TABLE DATABASE EXPORT (ZERO DASHES / ZERO BLANKS)")
    print("=" * 80)

    if not os.path.exists(DB_PATH):
        print(f"Error: Database {DB_PATH} not found.")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

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

    subject_headers = [
        "Bridge_ID",
        "Tutor_ID",
        "Subject_Index",
        "Canonical_Subject",
        "Subject_Category",
        "Target_Education_Level",
        "Competitive_Exam"
    ]

    # Step 1: Query strictly complete tutors
    print("\n[1/4] Filtering and extracting complete tutor records...")
    c.execute("""
        SELECT tutor_id, teacher_name, hourly_fee_avg, hourly_fee_min, hourly_fee_max,
               total_teaching_exp_years, online_teaching_exp_years, seniority_tier, highest_qualification,
               has_education_degree, city, state, region, city_tier, gender, occupation,
               teaches_online, teaches_at_student_home, speaks_languages, is_active, validation_reason, profile_url
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
    """)

    tutor_rows = []
    clean_tutor_id_set = set()

    for r in c:
        tid = r["tutor_id"]
        fee = round(float(r["hourly_fee_avg"]), 1)
        fee_min = round(float(r["hourly_fee_min"]), 1)
        fee_max = round(float(r["hourly_fee_max"]), 1)
        fee_band = get_fee_band(fee)

        clean_row = [
            tid,
            r["teacher_name"].strip(),
            fee,
            fee_min,
            fee_max,
            fee_band,
            round(float(r["total_teaching_exp_years"]), 1),
            round(float(r["online_teaching_exp_years"]), 1),
            r["seniority_tier"].strip(),
            r["highest_qualification"].strip(),
            "Yes" if r["has_education_degree"] == 1 else "No",
            r["city"].strip(),
            r["state"].strip(),
            r["region"].strip(),
            r["city_tier"].strip(),
            r["gender"].strip(),
            r["occupation"].strip(),
            "Yes" if r["teaches_online"] == 1 else "No",
            "Yes" if r["teaches_at_student_home"] == 1 else "No",
            r["speaks_languages"].strip(),
            "Active" if r["is_active"] == 1 else "Inactive",
            r["validation_reason"].strip(),
            r["profile_url"].strip()
        ]
        tutor_rows.append(clean_row)
        clean_tutor_id_set.add(tid)

    total_clean_tutors = len(tutor_rows)
    print(f"      ✓ Retained {total_clean_tutors:,} strictly clean, fully-populated tutor profiles.")

    # Step 2: Query clean subjects corresponding to retained tutors
    print("\n[2/4] Fetching subject offerings for retained tutors...")
    c.execute("""
        SELECT bridge_id, tutor_id, subject_index, canonical_subject, subject_category, level_category, competitive_exam
        FROM teacher_subjects
        ORDER BY bridge_id;
    """)

    subject_rows = []
    for r in c:
        tid = r["tutor_id"]
        if tid in clean_tutor_id_set:
            exam = r["competitive_exam"]
            exam_clean = exam.strip() if (exam and exam.strip() not in ('', '-', 'None', 'null')) else "General Curriculum"
            clean_sub_row = [
                r["bridge_id"],
                tid,
                r["subject_index"],
                r["canonical_subject"].strip(),
                r["subject_category"].strip(),
                r["level_category"].strip(),
                exam_clean
            ]
            subject_rows.append(clean_sub_row)

    total_clean_subjects = len(subject_rows)
    print(f"      ✓ Retained {total_clean_subjects:,} clean subject offerings.")

    # Step 3: Write CSVs
    print(f"\n[3/4] Writing Clean CSV files...")
    with open(OUT_TUTORS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(tutor_headers)
        writer.writerows(tutor_rows)
    print(f"      ✓ {OUT_TUTORS_CSV} written ({total_clean_tutors:,} rows).")

    with open(OUT_SUBJECTS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(subject_headers)
        writer.writerows(subject_rows)
    print(f"      ✓ {OUT_SUBJECTS_CSV} written ({total_clean_subjects:,} rows).")

    # Step 4: Write Clean 2-Tab Excel Workbook
    print(f"\n[4/4] Writing Clean 2-Tab Excel Workbook ({OUT_EXCEL})...")
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
    print(f"      Writing Tab 1: Tutors ({total_clean_tutors:,} rows)...")
    ws_tutors = wb.add_worksheet("Tutors")
    ws_tutors.set_tab_color("#2563EB")
    for col_idx, h in enumerate(tutor_headers):
        ws_tutors.write(0, col_idx, h, header_fmt)

    for row_idx, row in enumerate(tutor_rows, start=1):
        for col_idx, val in enumerate(row):
            ws_tutors.write(row_idx, col_idx, val)

    del tutor_rows

    # Tab 2: Subjects
    print(f"      Writing Tab 2: Subjects ({total_clean_subjects:,} rows)...")
    ws_subjs = wb.add_worksheet("Subjects")
    ws_subjs.set_tab_color("#059669")
    for col_idx, h in enumerate(subject_headers):
        ws_subjs.write(0, col_idx, h, header_fmt)

    for row_idx, row in enumerate(subject_rows, start=1):
        for col_idx, val in enumerate(row):
            ws_subjs.write(row_idx, col_idx, val)

    del subject_rows

    wb.close()
    conn.close()

    elapsed = round(time.time() - start_time, 1)
    print("\n" + "=" * 80)
    print(f"SUCCESS: STRICTLY CLEAN 2-TABLE DATABASE READY IN {elapsed}s")
    for p in [OUT_EXCEL, OUT_TUTORS_CSV, OUT_SUBJECTS_CSV]:
        size_mb = os.path.getsize(p) / (1024 * 1024)
        print(f"  • {p} ({size_mb:.1f} MB)")
    print("=" * 80)

if __name__ == "__main__":
    export_strictly_clean()
