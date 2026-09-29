import sqlite3
import xlsxwriter
import time
import os

DB_PATH = "/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db"
OUTPUT_EXCEL = "/home/nilesh7757/BMPTUTORS1/BMP_Tutors_Cleaned_Master.xlsx"

def generate_master_excel():
    start_time = time.time()
    print(f"Generating Complete 3-Sheet Master Excel file: {OUTPUT_EXCEL}")
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Step 1: Precompute aggregated subjects per tutor into a temp table for high performance
    print("Pre-aggregating bridge data per tutor...")
    t0 = time.time()
    cur.execute("DROP TABLE IF EXISTS _temp_tutor_subjects_agg")
    cur.execute("""
        CREATE TEMP TABLE _temp_tutor_subjects_agg AS
        SELECT 
            tutor_id,
            GROUP_CONCAT(DISTINCT canonical_subject) as subjects_list,
            GROUP_CONCAT(DISTINCT subject_category) as categories_list,
            GROUP_CONCAT(DISTINCT level_category) as levels_list,
            GROUP_CONCAT(DISTINCT competitive_exam) as exams_list,
            COUNT(*) as subject_count
        FROM teacher_subjects
        GROUP BY tutor_id
    """)
    cur.execute("CREATE INDEX idx_tmp_agg_tid ON _temp_tutor_subjects_agg(tutor_id)")
    print(f"Bridge pre-aggregation complete in {time.time()-t0:.2f}s")

    # Step 2: Initialize xlsxwriter with constant_memory for zero-memory streaming
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

    # =========================================================================
    # SHEET 1: All Tutors Master Cleaned (107,723 rows)
    # =========================================================================
    print("Writing Sheet 1: Tutors_Master_Cleaned...")
    ws1 = wb.add_worksheet('Tutors_Master_Cleaned')
    
    headers1 = [
        "Tutor ID",
        "Source Tab",
        "Is Active (Validation)",
        "Validation Reason",
        "Profile URL",
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
        CASE WHEN t.is_active = 1 THEN 'Yes (Active)' WHEN t.is_active = 0 THEN 'No (Inactive)' ELSE 'Unknown' END as is_active_str,
        COALESCE(t.validation_reason, '') as val_reason,
        COALESCE(t.profile_url, '') as profile_url,
        COALESCE(t.teacher_name, 'Anonymous') as teacher_name,
        COALESCE(t.teacher_hookline, '') as hookline,
        COALESCE(t.description, '') as description,
        COALESCE(t.location_raw, '') as raw_loc,
        COALESCE(t.city, '') as city,
        COALESCE(t.city_tier, '') as city_tier,
        COALESCE(t.state, '') as state,
        COALESCE(t.region, '') as region,
        COALESCE(t.highest_qualification, 'Not Specified') as qual,
        CASE WHEN t.has_education_degree = 1 THEN 'Yes' ELSE 'No' END as has_edu_deg,
        COALESCE(t.education_raw, '') as raw_edu,
        COALESCE(t.seniority_tier, '') as seniority,
        COALESCE(t.experience_role, 'Tutor') as exp_role,
        COALESCE(t.occupation, 'Independent Private Tutor') as occupation,
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
        COALESCE(t.gender, 'Not Specified') as gender,
        COALESCE(t.works_as, 'Individual teacher') as works_as,
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
    LEFT JOIN _temp_tutor_subjects_agg a ON t.tutor_id = a.tutor_id
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
            print(f"  Written {row_idx-1:,} tutors... ({time.time()-t1:.1f}s)")
            
    print(f"Sheet 1 complete ({row_idx-1:,} rows) in {time.time()-t1:.2f}s")

    # =========================================================================
    # SHEET 2: Teacher_Subjects_Bridge (403,353 rows)
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
        bridge_id,
        tutor_id,
        subject_index,
        canonical_subject,
        subject_category,
        level_category,
        COALESCE(competitive_exam, '') as competitive_exam,
        raw_subject_string
    FROM teacher_subjects
    ORDER BY bridge_id ASC
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
    # SHEET 3: Data_Dictionary & Summary
    # =========================================================================
    print("Writing Sheet 3: Data_Dictionary...")
    ws3 = wb.add_worksheet('Data_Dictionary')
    dict_header_fmt = wb.add_format({'bold': True, 'bg_color': '#475569', 'font_color': '#FFFFFF', 'font_size': 11})
    
    ws3.write_row(0, 0, ["Dataset Summary Metric", "Value", "Notes / Description"], dict_header_fmt)
    
    summary_metrics = [
        ("Total Unique Tutors", "107,723", "De-duplicated across scraped (76,706) and not scraped (31,017)"),
        ("Total Bridge Records", "403,353", "100% classified subject-level relationships"),
        ("Validation Active Tutors", "76,706 (71.2%)", "Verified active profile URLs"),
        ("Validation Inactive Tutors", "31,017 (28.8%)", "URLs that returned inactive or 404"),
        ("Average Hourly Fee", "INR 848.4 / hr", "Trimmed mean between INR 50 and INR 5,000 / hr"),
        ("Pedagogical Seniority Tiers", "4 Standard Tiers", "Master Educator (10+ Yrs), Experienced Practitioner (2-5 Yrs), Senior Specialist (5-10 Yrs), Early-Career (<2 Yrs)"),
        ("Target Educational Levels", "11 Canonical Levels", "100% concrete educational stages (0 'Skill Levels' / 0 'Not Specified')"),
        ("City Tiers", "Tier 1 (Metro), Tier 2 (Urban), Tier 3 (Semi-Urban / Rural)", "Standardized against 48 major urban agglomerations in India"),
        ("Occupational Classification", "100% Comprehensive Coverage", "19 Professional Domains: Software & IT, Medicine, Finance & CA, Engineering, Research Scholars, School Teachers, Faculty, etc."),
        ("Generated At", time.strftime("%Y-%m-%d %H:%M:%S"), "Exported directly from relational SQLite clean star schema")
    ]
    
    for r_i, (m, v, n) in enumerate(summary_metrics, start=1):
        ws3.write_row(r_i, 0, [m, v, n])
        
    print("Closing workbook and finalizing compression...")
    wb.close()
    conn.close()
    
    file_size_mb = os.path.getsize(OUTPUT_EXCEL) / (1024 * 1024)
    print(f"Master Excel file successfully generated: {OUTPUT_EXCEL} ({file_size_mb:.2f} MB)")
    print(f"Total time elapsed: {time.time()-start_time:.2f}s")

if __name__ == "__main__":
    generate_master_excel()
