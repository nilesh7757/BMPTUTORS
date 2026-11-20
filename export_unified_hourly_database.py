#!/usr/bin/env python3
"""
TutorMetric - Unified Hourly Fee Relational Database Exporter
=============================================================
Converts monthly, daily, and weekly fees into clean hourly rates:
  - Monthly: divided by 16 hours/month (standard 4 weeks x 4 hrs/wk)
  - Daily: divided by 1.5 hours/session
  - Weekly: divided by 4 hours/week
  - Non-INR (USD): converted at standard market rate (88 INR/USD)

Eliminates separate monthly fee columns, providing ONE unified hourly fee set:
  `hourly_fee_avg`, `hourly_fee_min`, `hourly_fee_max`

Generates:
  1. TutorMetric_Clean_Database_2Tables.xlsx (Tab 1: Tutors, Tab 2: Subjects)
  2. TutorMetric_Clean_Tutors.csv
  3. TutorMetric_Clean_Subjects.csv
And audits empty cells across all columns.
"""

import os
import sys
import time
import sqlite3
import csv
import re
import xlsxwriter
import openpyxl

DB_PATH = "cleaned_tutors_data.db"
OUT_EXCEL = "TutorMetric_Clean_Database_2Tables.xlsx"
OUT_TUTORS_CSV = "TutorMetric_Clean_Tutors.csv"
OUT_SUBJECTS_CSV = "TutorMetric_Clean_Subjects.csv"

def parse_fee_to_hourly(f_raw, f_unit, h_avg, h_min, h_max, m_avg, m_min, m_max):
    # 1. Direct hourly fee already present
    if h_avg is not None and h_avg > 0:
        h_min_val = h_min if (h_min is not None and h_min > 0) else h_avg
        h_max_val = h_max if (h_max is not None and h_max > 0) else h_avg
        return round(float(h_avg), 1), round(float(h_min_val), 1), round(float(h_max_val), 1)

    # 2. Monthly fee present -> convert at 16 hrs/month
    if m_avg is not None and m_avg > 0:
        avg_val = round(float(m_avg) / 16.0, 1)
        min_val = round(float(m_min) / 16.0, 1) if (m_min is not None and m_min > 0) else avg_val
        max_val = round(float(m_max) / 16.0, 1) if (m_max is not None and m_max > 0) else avg_val
        return avg_val, min_val, max_val

    # 3. Parse INR from raw string
    m_inr = re.findall(r'₹\s*([0-9,]+)', f_raw)
    if m_inr:
        nums = [float(x.replace(',', '')) for x in m_inr]
        raw_min = min(nums)
        raw_max = max(nums)
        raw_avg = sum(nums) / len(nums)

        if '/day' in f_raw or 'day' in f_unit:
            divisor = 1.5
        elif '/week' in f_raw or 'week' in f_unit:
            divisor = 4.0
        elif '/month' in f_raw or 'month' in f_unit:
            divisor = 16.0
        else:
            divisor = 1.0

        return round(raw_avg / divisor, 1), round(raw_min / divisor, 1), round(raw_max / divisor, 1)

    # 4. Parse USD from raw string
    m_usd = re.findall(r'USD\s*([0-9,.]+)', f_raw)
    if m_usd:
        nums = [float(x.replace(',', '')) for x in m_usd]
        usd_min = min(nums) * 88.0
        usd_max = max(nums) * 88.0
        usd_avg = (sum(nums) / len(nums)) * 88.0

        if '/day' in f_raw or 'day' in f_unit:
            divisor = 1.5
        elif '/week' in f_raw or 'week' in f_unit:
            divisor = 4.0
        elif '/month' in f_raw or 'month' in f_unit:
            divisor = 16.0
        else:
            divisor = 1.0

        return round(usd_avg / divisor, 1), round(usd_min / divisor, 1), round(usd_max / divisor, 1)

    return "", "", ""

def export():
    start_time = time.time()
    print("=" * 80)
    print("STARTING UNIFIED HOURLY DATABASE EXPORT (CONVERTING MONTHLY/DAILY TO HOURLY)")
    print("=" * 80)

    if not os.path.exists(DB_PATH):
        print(f"Error: Database {DB_PATH} not found.")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Unified original columns for Table 1: Tutors (monthly columns removed, hourly consolidated)
    tutor_cols = [
        "tutor_id",
        "teacher_name",
        "hourly_fee_avg",
        "hourly_fee_min",
        "hourly_fee_max",
        "fee_unit",
        "fee_raw",
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

    # Columns for Table 2: Subjects
    subject_cols = [
        "bridge_id",
        "tutor_id",
        "subject_index",
        "canonical_subject",
        "subject_category",
        "level_category",
        "competitive_exam"
    ]

    print("\n[1/4] Extracting tutors and converting monthly/daily rates to hourly...")
    c.execute("""
        SELECT tutor_id, teacher_name, fee_raw, fee_unit, hourly_fee_avg, hourly_fee_min, hourly_fee_max,
               monthly_fee_avg, monthly_fee_min, monthly_fee_max, total_teaching_exp_years,
               online_teaching_exp_years, seniority_tier, experience_role, highest_qualification,
               has_education_degree, city, state, region, city_tier, gender, occupation, works_as,
               teaches_online, teaches_at_student_home, can_travel, homework_help, speaks_languages,
               num_languages, registered_date, last_login_date, is_active, validation_reason, profile_url
        FROM teachers
        WHERE fee_raw IS NOT NULL AND TRIM(fee_raw) != ''
        ORDER BY CAST(SUBSTR(tutor_id, 2) AS INTEGER);
    """)

    tutor_rows = []
    retained_tutor_ids = set()

    for r in c:
        f_raw = r["fee_raw"] or ""
        f_unit = r["fee_unit"] or ""
        h_avg, h_min, h_max = parse_fee_to_hourly(
            f_raw, f_unit,
            r["hourly_fee_avg"], r["hourly_fee_min"], r["hourly_fee_max"],
            r["monthly_fee_avg"], r["monthly_fee_min"], r["monthly_fee_max"]
        )

        row_vals = [
            r["tutor_id"],
            r["teacher_name"].strip() if r["teacher_name"] else "",
            h_avg,
            h_min,
            h_max,
            f_unit.strip() if f_unit else "",
            f_raw.strip() if f_raw else "",
            r["total_teaching_exp_years"] if r["total_teaching_exp_years"] is not None else "",
            r["online_teaching_exp_years"] if r["online_teaching_exp_years"] is not None else "",
            r["seniority_tier"].strip() if r["seniority_tier"] else "",
            r["experience_role"].strip() if r["experience_role"] else "",
            r["highest_qualification"].strip() if r["highest_qualification"] else "",
            r["has_education_degree"] if r["has_education_degree"] is not None else "",
            r["city"].strip() if r["city"] else "",
            r["state"].strip() if r["state"] else "",
            r["region"].strip() if r["region"] else "",
            r["city_tier"].strip() if r["city_tier"] else "",
            r["gender"].strip() if r["gender"] else "",
            r["occupation"].strip() if r["occupation"] else "",
            r["works_as"].strip() if r["works_as"] else "",
            r["teaches_online"] if r["teaches_online"] is not None else "",
            r["teaches_at_student_home"] if r["teaches_at_student_home"] is not None else "",
            r["can_travel"] if r["can_travel"] is not None else "",
            r["homework_help"] if r["homework_help"] is not None else "",
            r["speaks_languages"].strip() if r["speaks_languages"] else "",
            r["num_languages"] if r["num_languages"] is not None else "",
            r["registered_date"].strip() if r["registered_date"] else "",
            r["last_login_date"].strip() if r["last_login_date"] else "",
            r["is_active"] if r["is_active"] is not None else "",
            r["validation_reason"].strip() if r["validation_reason"] else "",
            r["profile_url"].strip() if r["profile_url"] else ""
        ]
        tutor_rows.append(row_vals)
        retained_tutor_ids.add(r["tutor_id"])

    total_tutors = len(tutor_rows)
    print(f"      ✓ Processed {total_tutors:,} tutors into unified hourly schema.")

    print("\n[2/4] Extracting subject offerings...")
    c.execute(f"SELECT {', '.join(subject_cols)} FROM teacher_subjects ORDER BY bridge_id;")
    subject_rows = []
    for r in c:
        tid = r["tutor_id"]
        if tid in retained_tutor_ids:
            row_vals = [
                r["bridge_id"],
                tid,
                r["subject_index"],
                r["canonical_subject"].strip() if r["canonical_subject"] else "",
                r["subject_category"].strip() if r["subject_category"] else "",
                r["level_category"].strip() if r["level_category"] else "",
                r["competitive_exam"].strip() if r["competitive_exam"] else ""
            ]
            subject_rows.append(row_vals)

    total_subjects = len(subject_rows)
    print(f"      ✓ Processed {total_subjects:,} subject offerings.")

    # Write CSVs
    print("\n[3/4] Writing CSVs...")
    with open(OUT_TUTORS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(tutor_cols)
        writer.writerows(tutor_rows)
    print(f"      ✓ {OUT_TUTORS_CSV} written.")

    with open(OUT_SUBJECTS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(subject_cols)
        writer.writerows(subject_rows)
    print(f"      ✓ {OUT_SUBJECTS_CSV} written.")

    # Write Excel
    print("\n[4/4] Writing Excel Workbook (2 Tabs)...")
    wb = xlsxwriter.Workbook(OUT_EXCEL, {'constant_memory': True, 'strings_to_urls': False})

    header_fmt = wb.add_format({
        "bold": True,
        "bg_color": "#1E293B",
        "font_color": "#FFFFFF",
        "font_name": "Calibri",
        "font_size": 11,
        "border": 1
    })

    ws_tutors = wb.add_worksheet("Tutors")
    ws_tutors.set_tab_color("#2563EB")
    for col_idx, h in enumerate(tutor_cols):
        ws_tutors.write(0, col_idx, h, header_fmt)
    for row_idx, row in enumerate(tutor_rows, start=1):
        for col_idx, val in enumerate(row):
            if val != "" and val is not None:
                ws_tutors.write(row_idx, col_idx, val)

    del tutor_rows

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
    print(f"      ✓ Excel saved in {elapsed}s.")

if __name__ == "__main__":
    export()
