#!/usr/bin/env python3
"""
TutorMetric - Whole Relational Database Exporter
================================================
Exports the ENTIRE 2-table relational database from `cleaned_tutors_data.db`:
  Table 1: Tutors Table (107,673 educators)
  Table 2: Subjects Table (403,221 subject offerings connected by Tutor_ID)

Generates:
  1. TutorMetric_Full_Database_2Tables.xlsx (Multi-sheet Excel workbook)
  2. TutorMetric_All_Tutors_107k.csv (Full Tutors Table CSV)
  3. TutorMetric_All_Subjects_403k.csv (Full Subjects Bridge Table CSV)
"""

import os
import sys
import time
import sqlite3
import csv
import xlsxwriter

DB_PATH = "cleaned_tutors_data.db"
OUT_EXCEL = "TutorMetric_Full_Database_2Tables.xlsx"
OUT_TUTORS_CSV = "TutorMetric_All_Tutors_107k.csv"
OUT_SUBJECTS_CSV = "TutorMetric_All_Subjects_403k.csv"

def export_all():
    start_time = time.time()
    print("=" * 80)
    print("STARTING EXPORT OF ENTIRE 2-TABLE RELATIONAL DATABASE")
    print("=" * 80)

    if not os.path.exists(DB_PATH):
        print(f"Error: Database {DB_PATH} not found.")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Column mappings and headers for Table 1: Tutors
    tutor_cols = [
        ("tutor_id", "Tutor_ID"),
        ("teacher_name", "Teacher_Name"),
        ("profile_url", "Profile_URL"),
        ("is_active", "Is_Active"),
        ("validation_reason", "Validation_Reason"),
        ("gender", "Gender"),
        ("city", "City"),
        ("city_tier", "City_Tier"),
        ("state", "State"),
        ("region", "Region"),
        ("location_raw", "Location_Raw"),
        ("highest_qualification", "Highest_Qualification"),
        ("has_education_degree", "Has_Education_Degree"),
        ("experience_role", "Experience_Role"),
        ("seniority_tier", "Seniority_Tier"),
        ("total_teaching_exp_years", "Total_Teaching_Exp_Years"),
        ("online_teaching_exp_years", "Online_Teaching_Exp_Years"),
        ("occupation", "Occupation"),
        ("works_as", "Works_As"),
        ("speaks_languages", "Speaks_Languages"),
        ("num_languages", "Num_Languages"),
        ("teaches_online", "Teaches_Online"),
        ("teaches_at_student_home", "Teaches_At_Student_Home"),
        ("can_travel", "Can_Travel"),
        ("homework_help", "Homework_Help"),
        ("hourly_fee_avg", "Hourly_Fee_Avg_INR"),
        ("hourly_fee_min", "Hourly_Fee_Min_INR"),
        ("hourly_fee_max", "Hourly_Fee_Max_INR"),
        ("fee_unit", "Fee_Unit"),
        ("monthly_fee_avg", "Monthly_Fee_Avg_INR"),
        ("monthly_fee_min", "Monthly_Fee_Min_INR"),
        ("monthly_fee_max", "Monthly_Fee_Max_INR"),
        ("fee_raw", "Fee_Raw"),
        ("registered_date", "Registered_Date"),
        ("last_login_date", "Last_Login_Date"),
        ("source_tab", "Source_Tab"),
        ("teacher_hookline", "Teacher_Hookline"),
        ("description", "Description")
    ]

    # Column mappings and headers for Table 2: Subjects
    subject_cols = [
        ("bridge_id", "Bridge_ID"),
        ("tutor_id", "Tutor_ID"),
        ("subject_index", "Subject_Index"),
        ("canonical_subject", "Canonical_Subject"),
        ("subject_category", "Subject_Category"),
        ("level_category", "Target_Education_Level"),
        ("competitive_exam", "Competitive_Exam"),
        ("raw_subject_string", "Raw_Subject_String")
    ]

    tutor_db_fields = [col[0] for col in tutor_cols]
    tutor_headers = [col[1] for col in tutor_cols]

    subject_db_fields = [col[0] for col in subject_cols]
    subject_headers = [col[1] for col in subject_cols]

    # 1. Write Tutors CSV
    print(f"\n[1/3] Exporting Tutors Table to CSV ({OUT_TUTORS_CSV})...")
    c.execute(f"SELECT {', '.join(tutor_db_fields)} FROM teachers ORDER BY CAST(SUBSTR(tutor_id, 2) AS INTEGER);")
    tutor_count = 0
    with open(OUT_TUTORS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(tutor_headers)
        for row in c:
            writer.writerow(list(row))
            tutor_count += 1
    print(f"      ✓ Exported {tutor_count:,} tutors to {OUT_TUTORS_CSV}")

    # 2. Write Subjects CSV
    print(f"\n[2/3] Exporting Subjects Table to CSV ({OUT_SUBJECTS_CSV})...")
    c.execute(f"SELECT {', '.join(subject_db_fields)} FROM teacher_subjects ORDER BY bridge_id;")
    subject_count = 0
    with open(OUT_SUBJECTS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(subject_headers)
        for row in c:
            writer.writerow(list(row))
            subject_count += 1
    print(f"      ✓ Exported {subject_count:,} subject offerings to {OUT_SUBJECTS_CSV}")

    # 3. Write Combined Multi-Sheet Excel Workbook
    print(f"\n[3/3] Exporting Combined Multi-Sheet Excel ({OUT_EXCEL})...")
    wb = xlsxwriter.Workbook(OUT_EXCEL, {'constant_memory': True, 'strings_to_urls': False})

    header_fmt = wb.add_format({
        "bold": True,
        "bg_color": "#1E293B",
        "font_color": "#FFFFFF",
        "font_name": "Calibri",
        "font_size": 11,
        "border": 1
    })

    # Sheet 1: Tutors_Table
    print("      Writing Sheet 1: Tutors_Table (107,673 rows)...")
    ws_tutors = wb.add_worksheet("Tutors_Table")
    ws_tutors.set_tab_color("#2563EB")
    for col_idx, h in enumerate(tutor_headers):
        ws_tutors.write(0, col_idx, h, header_fmt)

    c.execute(f"SELECT {', '.join(tutor_db_fields)} FROM teachers ORDER BY CAST(SUBSTR(tutor_id, 2) AS INTEGER);")
    for row_idx, row in enumerate(c, start=1):
        for col_idx, val in enumerate(row):
            if val is not None:
                ws_tutors.write(row_idx, col_idx, val)

    # Sheet 2: Subjects_Table
    print("      Writing Sheet 2: Subjects_Table (403,221 rows)...")
    ws_subjs = wb.add_worksheet("Subjects_Table")
    ws_subjs.set_tab_color("#059669")
    for col_idx, h in enumerate(subject_headers):
        ws_subjs.write(0, col_idx, h, header_fmt)

    c.execute(f"SELECT {', '.join(subject_db_fields)} FROM teacher_subjects ORDER BY bridge_id;")
    for row_idx, row in enumerate(c, start=1):
        for col_idx, val in enumerate(row):
            if val is not None:
                ws_subjs.write(row_idx, col_idx, val)

    print("      Closing workbook and saving to disk...")
    wb.close()
    conn.close()

    elapsed = round(time.time() - start_time, 1)
    print("\n" + "=" * 80)
    print(f"SUCCESS: WHOLE 2-TABLE DATABASE EXPORTED IN {elapsed}s")
    for p in [OUT_EXCEL, OUT_TUTORS_CSV, OUT_SUBJECTS_CSV]:
        size_mb = os.path.getsize(p) / (1024 * 1024)
        print(f"  • {p} ({size_mb:.1f} MB)")
    print("=" * 80)

if __name__ == "__main__":
    export_all()
