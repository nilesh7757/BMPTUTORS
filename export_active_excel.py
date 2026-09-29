import sqlite3
import xlsxwriter
import time
import os

DB_PATH = "/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db"
OUTPUT_EXCEL = "/home/nilesh7757/BMPTUTORS1/BMP_Tutors_Verified_Active_Master.xlsx"

def generate_active_master_excel():
    start_time = time.time()
    print(f"Generating 100% URL-Verified Active Tutors Master Excel file: {OUTPUT_EXCEL}")
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Step 1: Precompute aggregated subjects per active tutor
    print("Pre-aggregating bridge data for active tutors...")
    t0 = time.time()
    cur.execute("DROP TABLE IF EXISTS _temp_active_tutor_subjects_agg")
    cur.execute("""
        CREATE TEMP TABLE _temp_active_tutor_subjects_agg AS
        SELECT 
            ts.tutor_id,
            GROUP_CONCAT(DISTINCT ts.canonical_subject) as subjects_list,
            GROUP_CONCAT(DISTINCT ts.subject_category) as categories_list,
            GROUP_CONCAT(DISTINCT ts.level_category) as levels_list,
            GROUP_CONCAT(DISTINCT ts.competitive_exam) as exams_list,
            COUNT(*) as subject_count
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        WHERE t.is_active = 1
        GROUP BY ts.tutor_id
    """)
    cur.execute("CREATE INDEX idx_tmp_active_agg_tid ON _temp_active_tutor_subjects_agg(tutor_id)")
    print(f"Active bridge pre-aggregation complete in {time.time()-t0:.2f}s")

    # Step 2: Initialize xlsxwriter
    if os.path.exists(OUTPUT_EXCEL):
        os.remove(OUTPUT_EXCEL)
        
    wb = xlsxwriter.Workbook(OUTPUT_EXCEL, {'constant_memory': True, 'strings_to_urls': False, 'default_date_format': 'yyyy-mm-dd'})
    
    # Formats
    header_fmt = wb.add_format({
        'bold': True,
        'bg_color': '#1E3A8A',
        'font_color': '#FFFFFF',
        'font_size': 11,
        'border': 1,
        'border_color': '#CBD5E1'
    })
    
    bridge_header_fmt = wb.add_format({
        'bold': True,
        'bg_color': '#047857',
        'font_color': '#FFFFFF',
        'font_size': 11,
        'border': 1,
        'border_color': '#CBD5E1'
    })

    dict_header_fmt = wb.add_format({
        'bold': True,
        'bg_color': '#475569',
        'font_color': '#FFFFFF',
        'font_size': 11,
        'border': 1,
        'border_color': '#CBD5E1'
    })

    # =========================================================================
    # SHEET 1: Verified Active Tutors (76,705 rows)
    # =========================================================================
    print("Writing Sheet 1: Tutors_Verified_Active...")
    ws1 = wb.add_worksheet('Tutors_Verified_Active')
    
    headers1 = [
        "Tutor ID",
        "Source Tab",
        "Is Active (Validation)",
        "Validation Reason",
        "Profile URL (TeacherOn Verified)",
        "Teacher Name",
        "Teacher Hookline",
        "Description (Bio)",
        "Original Location",
        "Clean Location",
        "Clean City",
        "City Tier",
        "Clean State",
        "Region",
        "Clean Highest Qualification",
        "Has Education Degree (B.Ed/M.Ed)",
        "Original Education",
        "Tutor Seniority Tier",
        "Clean Experience Role",
        "Occupation",
        "Total Teaching Exp (Yrs)",
        "Online Teaching Exp (Yrs)",
        "Original Experience",
        "Hourly Fee Avg (INR)",
        "Hourly Fee Min (INR)",
        "Hourly Fee Max (INR)",
        "Fee Unit",
        "Original Fee",
        "Teaches Online",
        "Teaches At Student Home",
        "Can Travel",
        "Homework Help",
        "Gender",
        "Works As",
        "Speaks Languages",
        "Number of Languages",
        "Registered Date",
        "Last Login Date",
        "All Subjects Taught (Cleaned)",
        "Subject Categories",
        "Target Educational Levels (100% Classified)",
        "Competitive Exams Coached",
        "Total Subjects Count"
    ]
    ws1.write_row(0, 0, headers1, header_fmt)

    query1 = """
    SELECT 
        t.tutor_id,
        t.source_tab,
        'Active (Verified)' as is_active_str,
        t.validation_reason as val_reason,
        t.profile_url,
        COALESCE(t.teacher_name, '') as teacher_name,
        COALESCE(t.teacher_hookline, '') as hookline,
        COALESCE(t.description, '') as description,
        COALESCE(t.location_raw, '') as raw_loc,
        COALESCE(t.city, '') as city,
        COALESCE(t.city_tier, '') as city_tier,
        COALESCE(t.state, '') as state,
        COALESCE(t.region, '') as region,
        COALESCE(t.highest_qualification, '') as qual,
        CASE WHEN t.has_education_degree = 1 THEN 'Yes' ELSE 'No' END as has_edu_deg,
        COALESCE(t.education_raw, '') as raw_edu,
        COALESCE(t.seniority_tier, '') as seniority,
        COALESCE(t.experience_role, '') as exp_role,
        COALESCE(t.occupation, '') as occupation,
        t.total_teaching_exp_years,
        t.online_teaching_exp_years,
        COALESCE(t.experience_raw, '') as raw_exp,
        t.hourly_fee_avg,
        t.hourly_fee_min,
        t.hourly_fee_max,
        COALESCE(t.fee_unit, '') as fee_unit,
        COALESCE(t.fee_raw, '') as raw_fee,
        CASE WHEN t.teaches_online = 1 THEN 'Yes' ELSE 'No' END as teaches_online,
        CASE WHEN t.teaches_at_student_home = 1 THEN 'Yes' ELSE 'No' END as teaches_home,
        CASE WHEN t.can_travel = 1 THEN 'Yes' ELSE 'No' END as can_travel,
        CASE WHEN t.homework_help = 1 THEN 'Yes' ELSE 'No' END as hw_help,
        COALESCE(t.gender, 'Unspecified') as gender,
        COALESCE(t.works_as, '') as works_as,
        COALESCE(t.speaks_languages, '') as speaks,
        t.num_languages,
        COALESCE(t.registered_date, '') as registered,
        COALESCE(t.last_login_date, '') as last_login,
        COALESCE(a.subjects_list, '') as subjects_list,
        COALESCE(a.categories_list, '') as categories_list,
        COALESCE(a.levels_list, '') as levels_list,
        COALESCE(a.exams_list, '') as exams_list,
        COALESCE(a.subject_count, 0) as subject_count
    FROM teachers t
    LEFT JOIN _temp_active_tutor_subjects_agg a ON t.tutor_id = a.tutor_id
    WHERE t.is_active = 1
    ORDER BY t.rowid ASC
    """

    cur.execute(query1)
    
    row_idx = 1
    t1 = time.time()
    while True:
        rows = cur.fetchmany(5000)
        if not rows:
            break
        for r in rows:
            clean_loc = f"{r['city']}, {r['state']}" if r['city'] and r['state'] else (r['state'] or r['city'] or '')
            row_data = [
                r['tutor_id'],
                r['source_tab'],
                r['is_active_str'],
                r['val_reason'],
                r['profile_url'],
                r['teacher_name'],
                r['hookline'],
                r['description'],
                r['raw_loc'],
                clean_loc,
                r['city'],
                r['city_tier'],
                r['state'],
                r['region'],
                r['qual'],
                r['has_edu_deg'],
                r['raw_edu'],
                r['seniority'],
                r['exp_role'],
                r['occupation'],
                r['total_teaching_exp_years'],
                r['online_teaching_exp_years'],
                r['raw_exp'],
                r['hourly_fee_avg'],
                r['hourly_fee_min'],
                r['hourly_fee_max'],
                r['fee_unit'],
                r['raw_fee'],
                r['teaches_online'],
                r['teaches_home'],
                r['can_travel'],
                r['hw_help'],
                r['gender'],
                r['works_as'],
                r['speaks'],
                r['num_languages'],
                r['registered'],
                r['last_login'],
                r['subjects_list'],
                r['categories_list'],
                r['levels_list'],
                r['exams_list'],
                r['subject_count']
            ]
            ws1.write_row(row_idx, 0, row_data)
            row_idx += 1
            
        if row_idx % 25000 == 1:
            print(f"  Written {row_idx-1:,} active tutors... ({time.time()-t1:.1f}s)")
            
    print(f"Sheet 1 complete ({row_idx-1:,} active rows) in {time.time()-t1:.2f}s")

    # =========================================================================
    # SHEET 2: Teacher_Subjects_Bridge (Only for Active Tutors)
    # =========================================================================
    print("Writing Sheet 2: Teacher_Subjects_Bridge...")
    ws2 = wb.add_worksheet('Teacher_Subjects_Bridge')
    
    headers2 = [
        "Bridge ID",
        "Tutor ID",
        "Subject Index",
        "Canonical Subject (Cleaned)",
        "Subject Category",
        "Target Educational Level (100% Classified)",
        "Competitive Exam Coached",
        "Original Raw Subject String"
    ]
    ws2.write_row(0, 0, headers2, bridge_header_fmt)

    query2 = """
    SELECT 
        ts.bridge_id,
        ts.tutor_id,
        ts.subject_index,
        ts.canonical_subject,
        ts.subject_category,
        ts.level_category,
        COALESCE(ts.competitive_exam, '') as competitive_exam,
        ts.raw_subject_string
    FROM teacher_subjects ts
    JOIN teachers t ON ts.tutor_id = t.tutor_id
    WHERE t.is_active = 1
    ORDER BY ts.bridge_id ASC
    """
    cur.execute(query2)
    
    row_idx2 = 1
    t2 = time.time()
    while True:
        rows = cur.fetchmany(10000)
        if not rows:
            break
        for r in rows:
            row_data = [
                r['bridge_id'],
                r['tutor_id'],
                r['subject_index'],
                r['canonical_subject'],
                r['subject_category'],
                r['level_category'],
                r['competitive_exam'],
                r['raw_subject_string']
            ]
            ws2.write_row(row_idx2, 0, row_data)
            row_idx2 += 1
            
        if row_idx2 % 100000 == 1:
            print(f"  Written {row_idx2-1:,} bridge links... ({time.time()-t2:.1f}s)")
            
    print(f"Sheet 2 complete ({row_idx2-1:,} rows) in {time.time()-t2:.2f}s")

    # =========================================================================
    # SHEET 3: Data Dictionary
    # =========================================================================
    print("Writing Sheet 3: Data_Dictionary...")
    ws3 = wb.add_worksheet('Data_Dictionary')
    headers3 = ["Sheet Name", "Column Name", "Data Type", "URL Verification / Logic", "Sample / Description"]
    ws3.write_row(0, 0, headers3, dict_header_fmt)

    dict_rows = [
        ("Tutors_Verified_Active", "Tutor ID", "String", "Unique Identifier", "T1, T2, ..."),
        ("Tutors_Verified_Active", "Profile URL", "URL String", "100% Verified Live URL", "https://www.teacheron.com/tutor/chBV - Live TeacherOn profile"),
        ("Tutors_Verified_Active", "Is Active (Validation)", "String", "All 100% Active", "Active (Verified) - Crawler verified page exists and educator is live"),
        ("Tutors_Verified_Active", "Validation Reason", "String", "Crawler Verification Result", "OK (Verified active profile)"),
        ("Tutors_Verified_Active", "Teacher Name", "String", "Cleaned Name", "Real verified tutor name"),
        ("Tutors_Verified_Active", "Teacher Hookline", "String", "Headline / Tagline", "Educator profile headline"),
        ("Tutors_Verified_Active", "Description (Bio)", "String", "Full Profile Bio", "Tutor detailed self-introduction"),
        ("Tutors_Verified_Active", "Clean City / State / Region", "String", "Standardized Geography", "All 37 Indian States and 594 Districts mapped"),
        ("Tutors_Verified_Active", "City Tier", "String", "Urban Classification", "Tier 1, Tier 2, Tier 3"),
        ("Tutors_Verified_Active", "Clean Highest Qualification", "String", "Standardized Degrees", "PhD, Postgraduate (PG), Undergraduate (UG)"),
        ("Tutors_Verified_Active", "Occupation", "String", "Professional Domain", "19 standardized domains (Software Engineer, College Faculty, etc.)"),
        ("Tutors_Verified_Active", "Total Teaching Exp (Yrs)", "Float", "Standardized Numerical Exp", "2.0, 5.0, 10.0"),
        ("Tutors_Verified_Active", "Hourly Fee Avg (INR)", "Float", "Normalized Hourly Rate", "Market trimmed mean: ₹848/hr, Median: ₹650/hr"),
        ("Tutors_Verified_Active", "Teaches Online / Home", "String", "Delivery Modality", "Yes / No"),
        ("Tutors_Verified_Active", "All Subjects Taught (Cleaned)", "String", "Deduplicated Subjects", "Comma-separated list of canonical subjects"),
        ("Tutors_Verified_Active", "Total Subjects Count", "Integer", "Competency Count", "Number of distinct verified subjects taught"),
        ("Teacher_Subjects_Bridge", "Bridge ID", "Integer", "Primary Key", "1, 2, 3..."),
        ("Teacher_Subjects_Bridge", "Canonical Subject (Cleaned)", "String", "Canonical Taxonomy", "Mathematics, Physics, Computer Science & Programming, etc."),
        ("Teacher_Subjects_Bridge", "Subject Category", "String", "High-level Subject Domain", "Science, Coding, Languages, Commerce, Arts, Humanities"),
        ("Teacher_Subjects_Bridge", "Target Educational Level", "String", "100% Classified Level", "Senior Secondary, Secondary, Middle, Primary, College")
    ]
    for idx3, drow in enumerate(dict_rows, 1):
        ws3.write_row(idx3, 0, drow)

    print("Closing workbook and finalizing compression...")
    wb.close()
    conn.close()
    print(f"100% Verified Active Master Excel file successfully generated: {OUTPUT_EXCEL}")
    print(f"Total time elapsed: {time.time()-start_time:.2f}s")

if __name__ == "__main__":
    generate_active_master_excel()
