import os
import sys
import time
import sqlite3

# Ensure xlsxwriter is loaded
sys.path.append('/home/nilesh7757/.cache/uv/archive-v0/b2_g7sKwC5mPN6Zp/lib/python3.13/site-packages')
import xlsxwriter

DB_PATH = "file:/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db?immutable=1"
TWO_SHEET_MASTER = "/home/nilesh7757/BMPTUTORS1/BMP_Tutors_Cleaned_Master.xlsx"
TEACHERS_STANDALONE = "/home/nilesh7757/BMPTUTORS1/BMP_Teachers_Master.xlsx"
SUBJECTS_STANDALONE = "/home/nilesh7757/BMPTUTORS1/BMP_Subjects_Bridge.xlsx"

HEADERS_TEACHERS = [
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

QUERY_TEACHERS = """
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

HEADERS_SUBJECTS = [
    "Bridge ID",
    "Tutor ID",
    "Subject Index",
    "Canonical Subject (Cleaned)",
    "Subject Category",
    "Target Educational Level (100% Classified)",
    "Competitive Exam Coached",
    "Original Raw Subject String"
]

QUERY_SUBJECTS = """
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

def write_teachers_worksheet(ws, cur, header_fmt):
    ws.write_row(0, 0, HEADERS_TEACHERS, header_fmt)
    cur.execute(QUERY_TEACHERS)
    
    row_idx = 1
    t0 = time.time()
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
            ws.write_row(row_idx, 0, row_data)
            row_idx += 1
            
        if row_idx % 30000 == 1:
            print(f"    Written {row_idx-1:,} teachers... ({time.time()-t0:.1f}s)")
            
    print(f"  Teachers sheet written successfully: {row_idx-1:,} rows ({time.time()-t0:.1f}s)")
    return row_idx - 1

def write_subjects_worksheet(ws, cur, header_fmt):
    ws.write_row(0, 0, HEADERS_SUBJECTS, header_fmt)
    cur.execute(QUERY_SUBJECTS)
    
    row_idx = 1
    t0 = time.time()
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
            ws.write_row(row_idx, 0, row_data)
            row_idx += 1
            
        if row_idx % 100000 == 1:
            print(f"    Written {row_idx-1:,} subject links... ({time.time()-t0:.1f}s)")
            
    print(f"  Subjects sheet written successfully: {row_idx-1:,} rows ({time.time()-t0:.1f}s)")
    return row_idx - 1

def main():
    print(f"Connecting to database {DB_PATH}...")
    conn = sqlite3.connect(DB_PATH, uri=True)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    
    print("Pre-aggregating bridge data per tutor into temporary in-memory table...")
    t_start = time.time()
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
    print(f"Bridge pre-aggregation complete in {time.time()-t_start:.2f}s")
    
    # -------------------------------------------------------------
    # 1. Master 2-Sheet Excel: Exactly Sheet 1: Teachers, Sheet 2: Subjects
    # -------------------------------------------------------------
    print(f"\n[1/3] Generating Master 2-Sheet Excel: {TWO_SHEET_MASTER}...")
    if os.path.exists(TWO_SHEET_MASTER):
        os.remove(TWO_SHEET_MASTER)
    wb_master = xlsxwriter.Workbook(TWO_SHEET_MASTER, {'constant_memory': True, 'strings_to_urls': False})
    fmt_blue = wb_master.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': '#FFFFFF', 'font_size': 11, 'border': 1, 'border_color': '#CBD5E1'})
    fmt_green = wb_master.add_format({'bold': True, 'bg_color': '#047857', 'font_color': '#FFFFFF', 'font_size': 11, 'border': 1, 'border_color': '#CBD5E1'})
    
    ws1 = wb_master.add_worksheet('Teachers')
    write_teachers_worksheet(ws1, cur, fmt_blue)
    
    ws2 = wb_master.add_worksheet('Subjects')
    write_subjects_worksheet(ws2, cur, fmt_green)
    
    wb_master.close()
    size_mb = os.path.getsize(TWO_SHEET_MASTER) / (1024 * 1024)
    print(f"Master 2-Sheet Excel generated: {TWO_SHEET_MASTER} ({size_mb:.2f} MB)")
    
    # -------------------------------------------------------------
    # 2. Standalone Teachers Excel
    # -------------------------------------------------------------
    print(f"\n[2/3] Generating Standalone Teachers Excel: {TEACHERS_STANDALONE}...")
    if os.path.exists(TEACHERS_STANDALONE):
        os.remove(TEACHERS_STANDALONE)
    wb_t = xlsxwriter.Workbook(TEACHERS_STANDALONE, {'constant_memory': True, 'strings_to_urls': False})
    fmt_blue_t = wb_t.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': '#FFFFFF', 'font_size': 11, 'border': 1, 'border_color': '#CBD5E1'})
    ws_t = wb_t.add_worksheet('Teachers')
    write_teachers_worksheet(ws_t, cur, fmt_blue_t)
    wb_t.close()
    size_mb_t = os.path.getsize(TEACHERS_STANDALONE) / (1024 * 1024)
    print(f"Standalone Teachers Excel generated: {TEACHERS_STANDALONE} ({size_mb_t:.2f} MB)")
    
    # -------------------------------------------------------------
    # 3. Standalone Subjects Excel
    # -------------------------------------------------------------
    print(f"\n[3/3] Generating Standalone Subjects Excel: {SUBJECTS_STANDALONE}...")
    if os.path.exists(SUBJECTS_STANDALONE):
        os.remove(SUBJECTS_STANDALONE)
    wb_s = xlsxwriter.Workbook(SUBJECTS_STANDALONE, {'constant_memory': True, 'strings_to_urls': False})
    fmt_green_s = wb_s.add_format({'bold': True, 'bg_color': '#047857', 'font_color': '#FFFFFF', 'font_size': 11, 'border': 1, 'border_color': '#CBD5E1'})
    ws_s = wb_s.add_worksheet('Subjects')
    write_subjects_worksheet(ws_s, cur, fmt_green_s)
    wb_s.close()
    size_mb_s = os.path.getsize(SUBJECTS_STANDALONE) / (1024 * 1024)
    print(f"Standalone Subjects Excel generated: {SUBJECTS_STANDALONE} ({size_mb_s:.2f} MB)")
    
    conn.close()
    print("\n" + "="*60)
    print("ALL 3 WORKBOOKS COMPLETED SUCCESSFULLY (NO 3RD SHEET)")
    print("="*60)

if __name__ == "__main__":
    main()
