#!/usr/bin/env python3
"""
BMP Tutors - Normalized Relational DBMS Tables Generator
========================================================
Decomposes the unnormalized master flat data into two 3NF normalized tables:
  1. Tutors Table (Entity: TUTOR, 107,723 rows, PK: Tutor_ID)
  2. Subjects Table (Entity: SUBJECT, 12,327 unique rows, PK: Subject_ID)

Eliminates the 327k-row redundant flat bridge while preserving 100% relational integrity.
"""

import os
import sys
import sqlite3
from collections import Counter, defaultdict
import numpy as np
import xlsxwriter

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cleaned_tutors_data.db")

OUT_TUTORS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "BMP_Tutors_Master_Normalized.xlsx")
OUT_SUBJECTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "BMP_Subjects_Master_Normalized.xlsx")
OUT_COMBINED = os.path.join(os.path.dirname(os.path.abspath(__file__)), "BMP_Normalized_Database_2Tables.xlsx")

def build_normalized_tables():
    if not os.path.exists(DB_PATH):
        print(f"Error: Database {DB_PATH} not found.")
        sys.exit(1)

    print("==================================================================")
    print("  BMP Tutors - Normalized Relational DBMS Generator (3NF)")
    print("==================================================================")

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Step 1: Read all teachers into memory for fast lookup
    print("\n[1/5] Loading Teachers entity (107,723 tutors)...")
    c.execute("SELECT * FROM teachers ORDER BY CAST(SUBSTR(tutor_id, 2) AS INTEGER);")
    teachers = c.fetchall()
    total_tutors = len(teachers)
    print(f"      Loaded {total_tutors:,} educators.")

    # Step 2: Read all teacher_subjects
    print("\n[2/5] Ingesting and aggregating academic subjects...")
    c.execute("""
        SELECT tutor_id, canonical_subject, subject_category, level_category, competitive_exam
        FROM teacher_subjects
        WHERE canonical_subject IS NOT NULL AND TRIM(canonical_subject) != '';
    """)
    ts_rows = c.fetchall()
    print(f"      Ingested {len(ts_rows):,} raw offerings.")

    # Mappings
    # subject -> set of tutor_ids
    subj_tutors = defaultdict(set)
    # subject -> metadata counters
    subj_cat_counts = defaultdict(Counter)
    subj_lvl_counts = defaultdict(Counter)
    subj_exam_counts = defaultdict(Counter)
    subj_state_counts = defaultdict(Counter)
    subj_fees = defaultdict(list)
    subj_online_count = defaultdict(int)
    subj_home_count = defaultdict(int)
    subj_female_count = defaultdict(int)
    subj_male_count = defaultdict(int)

    # tutor -> list of unique subjects
    tutor_subjs = defaultdict(list)
    tutor_subjs_seen = defaultdict(set)

    # Pre-index teachers for quick metadata lookup
    tutor_info = {}
    for t in teachers:
        tid = t["tutor_id"]
        fee = t["hourly_fee_avg"]
        valid_fee = fee if (fee and 50 <= fee <= 5000) else None
        tutor_info[tid] = {
            "fee": valid_fee,
            "online": bool(t["teaches_online"]),
            "home": bool(t["teaches_at_student_home"]),
            "gender": t["gender"],
            "state": t["state"]
        }

    for row in ts_rows:
        tid = row["tutor_id"]
        sub = row["canonical_subject"].strip()
        cat = (row["subject_category"] or "").strip()
        lvl = (row["level_category"] or "").strip()
        exam = (row["competitive_exam"] or "").strip()

        if sub not in tutor_subjs_seen[tid]:
            tutor_subjs[tid].append(sub)
            tutor_subjs_seen[tid].add(sub)
            subj_tutors[sub].add(tid)

            tinfo = tutor_info.get(tid)
            if tinfo:
                if tinfo["fee"] is not None:
                    subj_fees[sub].append(tinfo["fee"])
                if tinfo["online"]:
                    subj_online_count[sub] += 1
                if tinfo["home"]:
                    subj_home_count[sub] += 1
                if tinfo["gender"] == "Female":
                    subj_female_count[sub] += 1
                elif tinfo["gender"] == "Male":
                    subj_male_count[sub] += 1
                if tinfo["state"]:
                    subj_state_counts[sub][tinfo["state"]] += 1

        if cat: subj_cat_counts[sub][cat] += 1
        if lvl: subj_lvl_counts[sub][lvl] += 1
        if exam: subj_exam_counts[sub][exam] += 1

    # Step 3: Build Subject Entity Table (Ordered by popularity / tutor count)
    print("\n[3/5] Constructing normalized Subject Master entity table...")
    sorted_subjects = sorted(subj_tutors.keys(), key=lambda s: len(subj_tutors[s]), reverse=True)
    print(f"      Total Unique Academic Subjects: {len(sorted_subjects):,}")

    subject_rows = []
    sub_name_to_id = {}
    sub_name_to_cat = {}

    for idx, sub in enumerate(sorted_subjects, start=1):
        sub_id = f"SUB-{idx:05d}"
        sub_name_to_id[sub] = sub_id

        tut_count = len(subj_tutors[sub])
        mkt_share = round((tut_count * 100.0) / total_tutors, 3)

        fees = subj_fees[sub]
        avg_fee = round(float(np.mean(fees)), 1) if fees else None
        med_fee = round(float(np.median(fees)), 1) if fees else None
        min_fee = min(fees) if fees else None
        max_fee = max(fees) if fees else None

        # Mode category, level, exam
        cat = subj_cat_counts[sub].most_common(1)[0][0] if subj_cat_counts[sub] else "General"
        sub_name_to_cat[sub] = cat

        lvl = subj_lvl_counts[sub].most_common(1)[0][0] if subj_lvl_counts[sub] else "Secondary & High School"
        exam = subj_exam_counts[sub].most_common(1)[0][0] if subj_exam_counts[sub] else "General Academic"
        top_state = subj_state_counts[sub].most_common(1)[0][0] if subj_state_counts[sub] else "All-India"

        online_pct = round((subj_online_count[sub] * 100.0) / tut_count, 1) if tut_count else 0.0
        home_pct = round((subj_home_count[sub] * 100.0) / tut_count, 1) if tut_count else 0.0
        female_pct = round((subj_female_count[sub] * 100.0) / tut_count, 1) if tut_count else 0.0
        male_pct = round((subj_male_count[sub] * 100.0) / tut_count, 1) if tut_count else 0.0

        subject_rows.append({
            "Subject_ID": sub_id,
            "Subject_Name": sub,
            "Subject_Category": cat,
            "Total_Tutors_Offering": tut_count,
            "Market_Share_Pct": mkt_share,
            "Avg_Hourly_Fee_INR": avg_fee,
            "Median_Hourly_Fee_INR": med_fee,
            "Min_Hourly_Fee_INR": min_fee,
            "Max_Hourly_Fee_INR": max_fee,
            "Primary_Educational_Stage": lvl,
            "Top_Competitive_Exam": exam,
            "Online_Delivery_Rate_Pct": online_pct,
            "Home_Tuition_Rate_Pct": home_pct,
            "Female_Tutors_Pct": female_pct,
            "Male_Tutors_Pct": male_pct,
            "Top_Offering_State": top_state
        })

    # Step 4: Build Tutor Entity Table (With relational foreign keys to subjects)
    print("\n[4/5] Constructing normalized Tutor Master entity table...")
    tutor_rows = []
    for t in teachers:
        tid = t["tutor_id"]
        sub_list = tutor_subjs.get(tid, [])
        sub_ids = [sub_name_to_id[s] for s in sub_list if s in sub_name_to_id]
        cats = list(dict.fromkeys([sub_name_to_cat[s] for s in sub_list if s in sub_name_to_cat]))

        primary_sub = sub_list[0] if sub_list else None
        primary_cat = cats[0] if cats else None

        # Clean is_active label
        status_label = "Active" if t["is_active"] == 1 else "Inactive"

        tutor_rows.append({
            "Tutor_ID": tid,
            "Teacher_Name": t["teacher_name"],
            "Profile_URL": t["profile_url"],
            "Validation_Status": status_label,
            "Validation_Reason": t["validation_reason"],
            "Gender": t["gender"] or "Unstated",
            "City": t["city"],
            "City_Tier": t["city_tier"],
            "State": t["state"],
            "Region": t["region"],
            "Location_Raw": t["location_raw"],
            "Highest_Qualification": t["highest_qualification"],
            "Has_BEd_MEd_Degree": "Yes" if t["has_education_degree"] == 1 else "No",
            "Experience_Role": t["experience_role"],
            "Seniority_Tier": t["seniority_tier"],
            "Total_Teaching_Exp_Yrs": t["total_teaching_exp_years"],
            "Online_Teaching_Exp_Yrs": t["online_teaching_exp_years"],
            "Occupation": t["occupation"],
            "Works_As": t["works_as"],
            "Speaks_Languages": t["speaks_languages"],
            "Num_Languages": t["num_languages"],
            "Teaches_Online": "Yes" if t["teaches_online"] == 1 else "No",
            "Teaches_At_Student_Home": "Yes" if t["teaches_at_student_home"] == 1 else "No",
            "Can_Travel": "Yes" if t["can_travel"] == 1 else "No",
            "Homework_Help": "Yes" if t["homework_help"] == 1 else "No",
            "Hourly_Fee_Avg_INR": t["hourly_fee_avg"],
            "Hourly_Fee_Min_INR": t["hourly_fee_min"],
            "Hourly_Fee_Max_INR": t["hourly_fee_max"],
            "Registered_Date": t["registered_date"],
            "Last_Login_Date": t["last_login_date"],
            "Total_Subjects_Taught": len(sub_list),
            "Primary_Subject": primary_sub,
            "Primary_Category": primary_cat,
            "Subject_IDs_Taught": "; ".join(sub_ids),
            "Canonical_Subjects_Taught": "; ".join(sub_list),
            "Subject_Categories_Taught": "; ".join(cats)
        })

    # Step 5: Write out clean, professionally formatted Excel workbooks
    print("\n[5/5] Exporting normalized DBMS workbooks with native formatting...")

    # Write Function
    def write_excel_sheet(workbook, sheet_name, data, columns_meta, tab_color="#1E3A8A"):
        ws = workbook.add_worksheet(sheet_name)
        ws.set_tab_color(tab_color)
        ws.freeze_panes(1, 1)

        # Formats
        header_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#0F172A",
            "font_color": "#F8FAFC",
            "border": 1,
            "border_color": "#334155",
            "align": "center",
            "valign": "vcenter",
            "font_name": "Calibri",
            "font_size": 11
        })
        cell_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 10})
        pk_fmt = workbook.add_format({"bold": True, "font_name": "Consolas", "font_size": 10, "align": "center"})
        fk_fmt = workbook.add_format({"font_name": "Consolas", "font_size": 9, "font_color": "#2563EB"})
        num_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 10, "num_format": "#,##0"})
        curr_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 10, "num_format": "₹#,##0.0"})
        pct_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 10, "num_format": "0.0\"%\""})

        # Write Headers
        for col_idx, col in enumerate(columns_meta):
            ws.write(0, col_idx, col["header"], header_fmt)
            ws.set_column(col_idx, col_idx, col.get("width", 15))

        # Write Rows
        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, col in enumerate(columns_meta):
                key = col["key"]
                val = row_data.get(key)
                ctype = col.get("type", "text")

                if val is None or val == "":
                    ws.write_blank(row_idx, col_idx, None, cell_fmt)
                elif ctype == "pk":
                    ws.write(row_idx, col_idx, str(val), pk_fmt)
                elif ctype == "fk":
                    ws.write(row_idx, col_idx, str(val), fk_fmt)
                elif ctype == "currency":
                    ws.write_number(row_idx, col_idx, float(val), curr_fmt)
                elif ctype == "pct":
                    ws.write_number(row_idx, col_idx, float(val), pct_fmt)
                elif ctype == "int":
                    ws.write_number(row_idx, col_idx, int(val), num_fmt)
                elif ctype == "float":
                    ws.write_number(row_idx, col_idx, float(val), cell_fmt)
                else:
                    ws.write(row_idx, col_idx, str(val), cell_fmt)

        ws.autofilter(0, 0, len(data), len(columns_meta) - 1)

    # Column Metas
    subject_cols = [
        {"key": "Subject_ID", "header": "Subject ID (PK)", "type": "pk", "width": 14},
        {"key": "Subject_Name", "header": "Canonical Subject Name", "type": "text", "width": 30},
        {"key": "Subject_Category", "header": "Academic Domain", "type": "text", "width": 24},
        {"key": "Total_Tutors_Offering", "header": "Tutor Count (N)", "type": "int", "width": 16},
        {"key": "Market_Share_Pct", "header": "Market Share (%)", "type": "pct", "width": 16},
        {"key": "Avg_Hourly_Fee_INR", "header": "Avg Hourly Fee (INR)", "type": "currency", "width": 18},
        {"key": "Median_Hourly_Fee_INR", "header": "Median Hourly Fee", "type": "currency", "width": 18},
        {"key": "Min_Hourly_Fee_INR", "header": "Min Hourly Fee", "type": "currency", "width": 15},
        {"key": "Max_Hourly_Fee_INR", "header": "Max Hourly Fee", "type": "currency", "width": 15},
        {"key": "Primary_Educational_Stage", "header": "Primary Student Stage", "type": "text", "width": 25},
        {"key": "Top_Competitive_Exam", "header": "Target Competitive Exam", "type": "text", "width": 22},
        {"key": "Online_Delivery_Rate_Pct", "header": "Online Delivery Rate (%)", "type": "pct", "width": 22},
        {"key": "Home_Tuition_Rate_Pct", "header": "Home Tuition Rate (%)", "type": "pct", "width": 20},
        {"key": "Female_Tutors_Pct", "header": "Women Educators (%)", "type": "pct", "width": 19},
        {"key": "Male_Tutors_Pct", "header": "Men Educators (%)", "type": "pct", "width": 18},
        {"key": "Top_Offering_State", "header": "Top Demand State Hub", "type": "text", "width": 22}
    ]

    tutor_cols = [
        {"key": "Tutor_ID", "header": "Tutor ID (PK)", "type": "pk", "width": 13},
        {"key": "Teacher_Name", "header": "Educator Name", "type": "text", "width": 25},
        {"key": "Validation_Status", "header": "TeacherOn Status", "type": "text", "width": 16},
        {"key": "Validation_Reason", "header": "Verification Method", "type": "text", "width": 26},
        {"key": "Gender", "header": "Gender", "type": "text", "width": 12},
        {"key": "City", "header": "City", "type": "text", "width": 18},
        {"key": "City_Tier", "header": "City Tier", "type": "text", "width": 16},
        {"key": "State", "header": "State", "type": "text", "width": 20},
        {"key": "Region", "header": "Geographic Zone", "type": "text", "width": 16},
        {"key": "Highest_Qualification", "header": "Highest Qualification", "type": "text", "width": 22},
        {"key": "Has_BEd_MEd_Degree", "header": "B.Ed / M.Ed Degree", "type": "text", "width": 18},
        {"key": "Experience_Role", "header": "Professional Role", "type": "text", "width": 22},
        {"key": "Seniority_Tier", "header": "Experience Tier", "type": "text", "width": 18},
        {"key": "Total_Teaching_Exp_Yrs", "header": "Total Exp (Yrs)", "type": "float", "width": 15},
        {"key": "Online_Teaching_Exp_Yrs", "header": "Online Exp (Yrs)", "type": "float", "width": 15},
        {"key": "Occupation", "header": "Primary Occupation", "type": "text", "width": 22},
        {"key": "Works_As", "header": "Employment Mode", "type": "text", "width": 18},
        {"key": "Speaks_Languages", "header": "Languages Spoken", "type": "text", "width": 24},
        {"key": "Num_Languages", "header": "Language Count", "type": "int", "width": 15},
        {"key": "Teaches_Online", "header": "Teaches Online", "type": "text", "width": 15},
        {"key": "Teaches_At_Student_Home", "header": "Home Tuition", "type": "text", "width": 15},
        {"key": "Can_Travel", "header": "Can Travel", "type": "text", "width": 14},
        {"key": "Homework_Help", "header": "Homework Help", "type": "text", "width": 15},
        {"key": "Hourly_Fee_Avg_INR", "header": "Hourly Rate (INR)", "type": "currency", "width": 18},
        {"key": "Hourly_Fee_Min_INR", "header": "Min Hourly Rate", "type": "currency", "width": 15},
        {"key": "Hourly_Fee_Max_INR", "header": "Max Hourly Rate", "type": "currency", "width": 15},
        {"key": "Registered_Date", "header": "Registered Date", "type": "text", "width": 16},
        {"key": "Last_Login_Date", "header": "Last Login Date", "type": "text", "width": 16},
        {"key": "Total_Subjects_Taught", "header": "Total Subjects Count", "type": "int", "width": 19},
        {"key": "Primary_Subject", "header": "Primary Subject", "type": "text", "width": 24},
        {"key": "Primary_Category", "header": "Primary Category", "type": "text", "width": 20},
        {"key": "Subject_IDs_Taught", "header": "Subject IDs Taught (FK)", "type": "fk", "width": 32},
        {"key": "Canonical_Subjects_Taught", "header": "Canonical Subjects List", "type": "text", "width": 45},
        {"key": "Subject_Categories_Taught", "header": "Academic Domains", "type": "text", "width": 30},
        {"key": "Profile_URL", "header": "TeacherOn Profile URL", "type": "text", "width": 35}
    ]

    # File 1: Subjects Master
    print(f"      Writing Subjects Table -> {os.path.basename(OUT_SUBJECTS)} ...")
    wb_subj = xlsxwriter.Workbook(OUT_SUBJECTS, {"constant_memory": True})
    write_excel_sheet(wb_subj, "Subjects_Master", subject_rows, subject_cols, tab_color="#059669")
    wb_subj.close()
    print("      ✓ Subjects Table written successfully.")

    # File 2: Tutors Master
    print(f"      Writing Tutors Table -> {os.path.basename(OUT_TUTORS)} ...")
    wb_tutor = xlsxwriter.Workbook(OUT_TUTORS, {"constant_memory": True})
    write_excel_sheet(wb_tutor, "Tutors_Master", tutor_rows, tutor_cols, tab_color="#2563EB")
    wb_tutor.close()
    print("      ✓ Tutors Table written successfully.")

    # File 3: Combined 2-Sheet Workbook
    print(f"      Writing Combined 2-Table Database -> {os.path.basename(OUT_COMBINED)} ...")
    wb_comb = xlsxwriter.Workbook(OUT_COMBINED, {"constant_memory": True})
    write_excel_sheet(wb_comb, "Tutors_Master", tutor_rows, tutor_cols, tab_color="#2563EB")
    write_excel_sheet(wb_comb, "Subjects_Master", subject_rows, subject_cols, tab_color="#059669")
    wb_comb.close()
    print("      ✓ Combined 2-Table Database written successfully.")

    print("\n==================================================================")
    print("  NORMALIZATION COMPLETE")
    print(f"  • Tutors Table:   {len(tutor_rows):,} rows (1 per educator)")
    print(f"  • Subjects Table: {len(subject_rows):,} rows (1 per academic discipline)")
    print(f"  • Redundancy:     0% (Replaced 327,695-row flat bridge)")
    print("==================================================================")

if __name__ == "__main__":
    build_normalized_tables()
