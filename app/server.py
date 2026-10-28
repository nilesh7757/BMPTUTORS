from collections import Counter
import sqlite3
import json
import os
import sys
import io

try:
    import lxml.etree
except Exception:
    pass

UV_PATH = "/home/nilesh7757/.cache/uv/archive-v0/b2_g7sKwC5mPN6Zp/lib/python3.13/site-packages"
if os.path.exists(UV_PATH) and UV_PATH not in sys.path:
    sys.path.append(UV_PATH)

import uvicorn
import xlsxwriter

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(APP_DIR)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

try:
    from build_custom_docx import generate_docx_report
except ImportError:
    generate_docx_report = None

from starlette.applications import Starlette
from starlette.responses import JSONResponse, HTMLResponse, Response, FileResponse
from starlette.routing import Route, Mount
from starlette.middleware import Middleware
from starlette.middleware.cors import CORSMiddleware
from starlette.staticfiles import StaticFiles

try:
    from app.geo_coords import STATE_COORDINATES, CITY_COORDINATES
    from app.geo_service import get_enriched_states_geojson, get_enriched_districts_geojson
    from app.fee_service import predict_fee
    from app.bio_service import generate_profile_bio, analyze_profile_bio
except ImportError:
    from geo_coords import STATE_COORDINATES, CITY_COORDINATES
    from geo_service import get_enriched_states_geojson, get_enriched_districts_geojson
    from fee_service import predict_fee
    from bio_service import generate_profile_bio, analyze_profile_bio

DB_PATH = os.getenv("DB_PATH", os.path.join(PROJECT_ROOT, "cleaned_tutors_data.db"))

# Automatically decompress database if deployed with db.gz
if (not os.path.exists(DB_PATH) or os.path.getsize(DB_PATH) == 0) and os.path.exists(DB_PATH + ".gz"):
    import gzip
    import shutil
    print(f"Decompressing database {DB_PATH}.gz ...")
    tmp_db = DB_PATH + ".tmp"
    with gzip.open(DB_PATH + ".gz", "rb") as f_in, open(tmp_db, "wb") as f_out:
        shutil.copyfileobj(f_in, f_out)
    os.replace(tmp_db, DB_PATH)
    print("Database ready.")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    conn.execute("PRAGMA cache_size = -64000;")
    conn.execute("PRAGMA temp_store = MEMORY;")
    conn.execute("PRAGMA mmap_size = 268435456;")
    return conn

_STATS_CACHE = None
_FILTERS_CACHE = None
_TOTAL_TUTORS_COUNT = None

async def api_stats(request):
    global _STATS_CACHE, _TOTAL_TUTORS_COUNT
    if _STATS_CACHE is not None:
        return JSONResponse(_STATS_CACHE)

    conn = get_db()
    cur = conn.cursor()
    
    total_tutors = cur.execute("SELECT count(*) FROM teachers").fetchone()[0]
    _TOTAL_TUTORS_COUNT = total_tutors
    total_bridge = cur.execute("SELECT count(*) FROM teacher_subjects").fetchone()[0]
    avg_fee = cur.execute("SELECT round(avg(hourly_fee_avg), 1) FROM teachers WHERE hourly_fee_avg IS NOT NULL AND hourly_fee_avg <= 5000").fetchone()[0]
    active_tutors = cur.execute("SELECT count(*) FROM teachers WHERE is_active = 1").fetchone()[0]
    
    conn.close()
    _STATS_CACHE = {
        "total_tutors": total_tutors,
        "total_bridge": total_bridge,
        "avg_fee": avg_fee,
        "active_tutors": active_tutors
    }
    return JSONResponse(_STATS_CACHE)

async def api_filters(request):
    global _FILTERS_CACHE
    if _FILTERS_CACHE is not None:
        return JSONResponse(_FILTERS_CACHE)

    conn = get_db()
    cur = conn.cursor()
    
    regions = [r[0] for r in cur.execute("SELECT DISTINCT region FROM teachers WHERE region IS NOT NULL ORDER BY region").fetchall()]
    states = [r[0] for r in cur.execute("SELECT DISTINCT state FROM teachers WHERE state IS NOT NULL ORDER BY state").fetchall()]
    qualifications = [r[0] for r in cur.execute("SELECT DISTINCT highest_qualification FROM teachers WHERE highest_qualification IS NOT NULL ORDER BY highest_qualification").fetchall()]
    seniority_tiers = [r[0] for r in cur.execute("SELECT DISTINCT seniority_tier FROM teachers WHERE seniority_tier IS NOT NULL ORDER BY seniority_tier").fetchall()]
    experience_roles = [r[0] for r in cur.execute("SELECT DISTINCT experience_role FROM teachers WHERE experience_role IS NOT NULL ORDER BY experience_role").fetchall()]
    subject_categories = [r[0] for r in cur.execute("SELECT DISTINCT subject_category FROM teacher_subjects WHERE subject_category IS NOT NULL ORDER BY subject_category").fetchall()]
    levels = [r[0] for r in cur.execute("SELECT DISTINCT level_category FROM teacher_subjects WHERE level_category IS NOT NULL ORDER BY level_category").fetchall()]
    exams = [r[0] for r in cur.execute("SELECT DISTINCT competitive_exam FROM teacher_subjects WHERE competitive_exam IS NOT NULL ORDER BY competitive_exam").fetchall()]
    occupations = [r[0] for r in cur.execute("SELECT DISTINCT occupation FROM teachers WHERE occupation IS NOT NULL ORDER BY occupation").fetchall()]
    
    conn.close()
    _FILTERS_CACHE = {
        "regions": regions,
        "states": states,
        "qualifications": qualifications,
        "seniority_tiers": seniority_tiers,
        "experience_roles": experience_roles,
        "occupations": occupations,
        "subject_categories": subject_categories,
        "levels": levels,
        "exams": exams
    }
    return JSONResponse(_FILTERS_CACHE)

def get_sort_clause(params):
    sort_by_raw = params.get("sort_by", params.get("sort", "")).strip().lower()
    sort_order_raw = params.get("sort_order", params.get("order", "")).strip().lower()

    if "_" in sort_by_raw and not sort_order_raw:
        parts = sort_by_raw.rsplit("_", 1)
        if parts[1] in ("asc", "desc"):
            sort_by = parts[0]
            sort_order = parts[1]
        else:
            sort_by = sort_by_raw
            sort_order = "asc"
    else:
        sort_by = sort_by_raw
        sort_order = sort_order_raw if sort_order_raw in ("asc", "desc") else ""

    if not sort_order:
        if sort_by in ("fee", "fees", "rate"):
            sort_order = "asc"
        elif sort_by in ("exp", "experience", "years", "education", "qualification", "degree"):
            sort_order = "desc"
        else:
            sort_order = "asc"

    direction = sort_order.upper()

    if sort_by == "name":
        return f"CASE WHEN t.teacher_name GLOB '[A-Za-z]*' THEN 0 WHEN t.teacher_name IS NULL OR t.teacher_name = '' THEN 2 ELSE 1 END, t.teacher_name {direction}", "name", sort_order
    elif sort_by in ("fee", "fees", "rate"):
        return f"t.hourly_fee_avg IS NULL, t.hourly_fee_avg {direction}", "fee", sort_order
    elif sort_by in ("education", "qualification", "degree"):
        case_expr = f"""CASE t.highest_qualification 
            WHEN 'Advanced (PhD/Doctorate)' THEN 4 
            WHEN 'Postgraduate (PG)' THEN 3 
            WHEN 'Undergraduate (UG)' THEN 2 
            WHEN 'Less than UG' THEN 1 
            ELSE 0 END {direction}, t.highest_qualification {direction}"""
        return case_expr, "education", sort_order
    elif sort_by in ("experience", "exp", "years"):
        return f"t.total_teaching_exp_years IS NULL, t.total_teaching_exp_years {direction}", "experience", sort_order
    elif sort_by in ("id", "tutor_id"):
        return f"t.rowid {direction}", "id", sort_order
    else:
        return "t.rowid ASC", "default", "asc"

async def api_tutors(request):
    params = request.query_params
    page = int(params.get("page", 1))
    page_size = min(int(params.get("page_size", 25)), 100)
    offset = (page - 1) * page_size
    
    search = params.get("search", "").strip()
    region = params.get("region", "").strip()
    state = params.get("state", "").strip()
    qualification = params.get("qualification", "").strip()
    seniority_tier = params.get("seniority_tier", "").strip()
    experience_role = params.get("experience_role", "").strip()
    occupation = params.get("occupation", "").strip()
    subject_category = params.get("subject_category", "").strip()
    level = params.get("level", "").strip()
    exam = params.get("exam", "").strip()
    gender = params.get("gender", "").strip()
    teaches_online = params.get("teaches_online", "").strip()
    min_fee = params.get("min_fee", "").strip()
    max_fee = params.get("max_fee", "").strip()
    
    where_clauses = []
    args = []
    
    if search:
        s_upper = search.upper()
        if (s_upper.startswith('T') and s_upper[1:].isdigit()) or s_upper.isdigit():
            tid = s_upper if s_upper.startswith('T') else f"T{s_upper}"
            where_clauses.append("t.tutor_id = ?")
            args.append(tid)
        else:
            where_clauses.append("(t.teacher_name LIKE ? OR t.teacher_hookline LIKE ? OR t.location_raw LIKE ? OR t.occupation LIKE ?)")
            s_term = f"%{search}%"
            args.extend([s_term, s_term, s_term, s_term])
        
    if region:
        where_clauses.append("t.region = ?")
        args.append(region)
        
    if state:
        where_clauses.append("t.state = ?")
        args.append(state)
        
    if qualification:
        where_clauses.append("t.highest_qualification = ?")
        args.append(qualification)
        
    if seniority_tier:
        where_clauses.append("t.seniority_tier = ?")
        args.append(seniority_tier)
        
    if experience_role:
        where_clauses.append("t.experience_role = ?")
        args.append(experience_role)

    if occupation:
        where_clauses.append("t.occupation = ?")
        args.append(occupation)
        
    if gender:
        where_clauses.append("t.gender = ?")
        args.append(gender)
        
    if teaches_online:
        where_clauses.append("t.teaches_online = ?")
        args.append(int(teaches_online))
        
    if min_fee:
        where_clauses.append("t.hourly_fee_avg >= ?")
        args.append(float(min_fee))
        
    if max_fee:
        where_clauses.append("t.hourly_fee_avg <= ?")
        args.append(float(max_fee))
        
    # Bridge filters
    bridge_joins = []
    if subject_category:
        where_clauses.append("EXISTS (SELECT 1 FROM teacher_subjects bs WHERE bs.tutor_id = t.tutor_id AND bs.subject_category = ?)")
        args.append(subject_category)
        
    if level:
        where_clauses.append("EXISTS (SELECT 1 FROM teacher_subjects bl WHERE bl.tutor_id = t.tutor_id AND bl.level_category = ?)")
        args.append(level)
        
    if exam:
        where_clauses.append("EXISTS (SELECT 1 FROM teacher_subjects be WHERE be.tutor_id = t.tutor_id AND be.competitive_exam = ?)")
        args.append(exam)
        
    where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""
    
    conn = get_db()
    cur = conn.cursor()
    
    # Count total matching
    if where_sql:
        count_query = f"SELECT count(*) FROM teachers t {where_sql}"
        total_count = cur.execute(count_query, args).fetchone()[0]
    else:
        global _TOTAL_TUTORS_COUNT
        if _TOTAL_TUTORS_COUNT is None:
            _TOTAL_TUTORS_COUNT = cur.execute("SELECT count(*) FROM teachers").fetchone()[0]
        total_count = _TOTAL_TUTORS_COUNT
    
    order_clause, active_sort_by, active_sort_order = get_sort_clause(params)

    # Fetch paginated rows
    query = f"""
    SELECT 
        t.tutor_id, t.teacher_name, t.teacher_hookline, t.state, t.region, t.city,
        t.gender, t.works_as, t.experience_role, t.seniority_tier, t.total_teaching_exp_years,
        t.highest_qualification, t.has_education_degree, t.speaks_languages,
        t.teaches_online, t.hourly_fee_avg, t.is_active, t.occupation
    FROM teachers t
    {where_sql}
    ORDER BY {order_clause}
    LIMIT ? OFFSET ?
    """
    rows = cur.execute(query, args + [page_size, offset]).fetchall()
    
    tutor_ids = [r["tutor_id"] for r in rows]
    
    # Fetch subjects for these tutors to show badges
    subjects_map = {}
    if tutor_ids:
        placeholders = ",".join("?" for _ in tutor_ids)
        subjs_rows = cur.execute(f"""
            SELECT tutor_id, canonical_subject, subject_category, level_category, competitive_exam
            FROM teacher_subjects
            WHERE tutor_id IN ({placeholders})
            ORDER BY subject_index ASC
        """, tutor_ids).fetchall()
        
        for sr in subjs_rows:
            tid = sr["tutor_id"]
            if tid not in subjects_map:
                subjects_map[tid] = []
            subjects_map[tid].append({
                "canonical_subject": sr["canonical_subject"],
                "subject_category": sr["subject_category"],
                "level_category": sr["level_category"],
                "competitive_exam": sr["competitive_exam"]
            })
            
    results = []
    for r in rows:
        d = dict(r)
        d["subjects"] = subjects_map.get(r["tutor_id"], [])
        results.append(d)
        
    conn.close()
    return JSONResponse({
        "page": page,
        "page_size": page_size,
        "total_count": total_count,
        "total_pages": (total_count + page_size - 1) // page_size if page_size else 1,
        "sort_by": active_sort_by,
        "sort_order": active_sort_order,
        "items": results
    })

async def api_tutor_detail(request):
    tutor_id = request.path_params["tutor_id"]
    conn = get_db()
    cur = conn.cursor()
    
    t = cur.execute("SELECT * FROM teachers WHERE tutor_id = ?", [tutor_id]).fetchone()
    if not t:
        conn.close()
        return JSONResponse({"error": "Tutor not found"}, status_code=404)
        
    t_dict = dict(t)
    
    subjs = cur.execute("""
        SELECT subject_index, raw_subject_string, canonical_subject,
               subject_category, level_category, competitive_exam
        FROM teacher_subjects
        WHERE tutor_id = ?
        ORDER BY subject_index ASC
    """, [tutor_id]).fetchall()
    
    t_dict["subjects"] = [dict(s) for s in subjs]
    
    conn.close()
    return JSONResponse(t_dict)

_CHARTS_CACHE = None

async def api_charts(request):
    params = request.query_params
    region = params.get("region", "").strip()
    mode = params.get("mode", "").strip()
    qualification = params.get("qualification", "").strip()
    active_only = params.get("active_only", "").strip() or params.get("is_active", "").strip()
    is_active_filter = active_only in ("1", "true", "active")
    has_filter = bool(region or mode or qualification or is_active_filter)

    global _CHARTS_CACHE
    if not has_filter and _CHARTS_CACHE is not None:
        return JSONResponse(_CHARTS_CACHE)

    conn = get_db()
    cur = conn.cursor()

    t_clauses = []
    t_args = []
    if region:
        t_clauses.append("t.region = ?")
        t_args.append(region)
    if qualification:
        t_clauses.append("t.highest_qualification = ?")
        t_args.append(qualification)
    if mode == "online":
        t_clauses.append("t.teaches_online = 1")
    elif mode == "offline":
        t_clauses.append("t.teaches_at_student_home = 1")
    if is_active_filter:
        t_clauses.append("t.is_active = 1")

    t_where = ("WHERE " + " AND ".join(t_clauses)) if t_clauses else ""
    t_and = (" AND " + " AND ".join(t_clauses)) if t_clauses else ""

    sub_cats = cur.execute(f"""
        SELECT ts.subject_category, count(*) as count
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        {t_where}
        GROUP BY ts.subject_category
        ORDER BY count DESC
    """, t_args).fetchall()
    
    quals = cur.execute(f"""
        SELECT t.highest_qualification, count(*) as count
        FROM teachers t
        {t_where}
        GROUP BY t.highest_qualification
        ORDER BY count DESC
    """, t_args).fetchall()
    
    regions = cur.execute(f"""
        SELECT t.region, count(*) as count
        FROM teachers t
        WHERE t.region IS NOT NULL {t_and}
        GROUP BY t.region
        ORDER BY count DESC
    """, t_args).fetchall()
    
    exams = cur.execute(f"""
        SELECT ts.competitive_exam, count(*) as count
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        WHERE ts.competitive_exam IS NOT NULL {t_and}
        GROUP BY ts.competitive_exam
        ORDER BY count DESC
        LIMIT 10
    """, t_args).fetchall()
    
    levels = cur.execute(f"""
        SELECT ts.level_category, count(*) as count
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        WHERE ts.level_category IS NOT NULL AND ts.level_category != 'Not Specified' {t_and}
        GROUP BY ts.level_category
        ORDER BY count DESC
    """, t_args).fetchall()

    city_tiers = cur.execute(f"""
        SELECT t.city_tier, count(*) as count
        FROM teachers t
        WHERE t.city_tier IS NOT NULL {t_and}
        GROUP BY t.city_tier
        ORDER BY count DESC
    """, t_args).fetchall()

    fee_by_cat = cur.execute(f"""
        SELECT s.subject_category, round(avg(t.hourly_fee_avg), 1) as avg_fee
        FROM teacher_subjects s
        JOIN teachers t ON s.tutor_id = t.tutor_id
        WHERE t.hourly_fee_avg IS NOT NULL AND t.hourly_fee_avg BETWEEN 50 AND 5000 {t_and}
        GROUP BY s.subject_category
        ORDER BY avg_fee DESC
    """, t_args).fetchall()

    seniority_tiers = cur.execute(f"""
        SELECT t.seniority_tier, count(*) as count
        FROM teachers t
        WHERE t.seniority_tier IS NOT NULL {t_and}
        GROUP BY t.seniority_tier
        ORDER BY count DESC
    """, t_args).fetchall()

    # Gender distribution by broad category (Academic Table 2)
    gender_cats_raw = cur.execute(f"""
        SELECT 
            ts.subject_category as subject,
            COUNT(DISTINCT ts.tutor_id) as n,
            COUNT(DISTINCT CASE WHEN t.gender = 'Male' THEN ts.tutor_id END) as male,
            COUNT(DISTINCT CASE WHEN t.gender = 'Female' THEN ts.tutor_id END) as female
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        {t_where}
        GROUP BY ts.subject_category
        ORDER BY n DESC
    """, t_args).fetchall()
    
    gender_by_category = []
    for r in gender_cats_raw:
        m = r['male']
        f = r['female']
        tot = m + f
        gender_by_category.append({
            "subject": r['subject'],
            "n": r['n'],
            "m_pct": round(m / tot * 100) if tot else 0,
            "f_pct": round(f / tot * 100) if tot else 0,
            "male": m,
            "female": f
        })

    # Gender distribution by top specific subjects
    gender_subs_raw = cur.execute(f"""
        SELECT 
            ts.canonical_subject as subject,
            COUNT(DISTINCT ts.tutor_id) as n,
            COUNT(DISTINCT CASE WHEN t.gender = 'Male' THEN ts.tutor_id END) as male,
            COUNT(DISTINCT CASE WHEN t.gender = 'Female' THEN ts.tutor_id END) as female
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        {t_where}
        GROUP BY ts.canonical_subject
        ORDER BY n DESC
        LIMIT 15
    """, t_args).fetchall()
    
    gender_by_subject = []
    for r in gender_subs_raw:
        m = r['male']
        f = r['female']
        tot = m + f
        gender_by_subject.append({
            "subject": r['subject'],
            "n": r['n'],
            "m_pct": round(m / tot * 100) if tot else 0,
            "f_pct": round(f / tot * 100) if tot else 0,
            "male": m,
            "female": f
        })

    # Overall total row
    tot_row = cur.execute(f"""
        SELECT 
            COUNT(DISTINCT ts.tutor_id) as n,
            COUNT(DISTINCT CASE WHEN t.gender = 'Male' THEN ts.tutor_id END) as male,
            COUNT(DISTINCT CASE WHEN t.gender = 'Female' THEN ts.tutor_id END) as female
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        {t_where}
    """, t_args).fetchone()
    tot_m = tot_row['male'] if tot_row else 0
    tot_f = tot_row['female'] if tot_row else 0
    tot_valid = (tot_m or 0) + (tot_f or 0)
    gender_total = {
        "subject": "Total",
        "n": tot_row['n'] if tot_row else 0,
        "m_pct": round(tot_m / tot_valid * 100) if tot_valid else 0,
        "f_pct": round(tot_f / tot_valid * 100) if tot_valid else 0,
        "male": tot_m,
        "female": tot_f
    }

    # -------------------------------------------------------------
    # 1. FIGURE 1: SUBJECTS DISTRIBUTION (MOSCOW STYLE DELIVERABLE)
    # -------------------------------------------------------------
    macro_raw = cur.execute(f"""
        SELECT 
            CASE 
                WHEN ts.subject_category = 'Science' THEN 'School Curriculum & STEM Sciences'
                WHEN ts.subject_category = 'Languages' THEN 'Foreign & Indian Languages'
                WHEN ts.subject_category = 'Coding' THEN 'Computer Science, IT & Coding'
                WHEN ts.subject_category = 'Commerce' THEN 'Commerce, Finance & Management'
                WHEN ts.subject_category = 'Arts' THEN 'Social Sciences, Arts & Humanities'
                ELSE 'Personal Development & ECA'
            END as name,
            count(*) as count
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        {t_where}
        GROUP BY name
        ORDER BY count DESC
    """, t_args).fetchall()
    
    tot_macro = sum(r['count'] for r in macro_raw) or 1
    macro_rows = [{"name": r['name'], "count": r['count'], "pct": round(r['count'] * 100.0 / tot_macro, 1)} for r in macro_raw]

    top_sub_raw = cur.execute(f"""
        SELECT ts.canonical_subject as name, count(*) as count
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        {t_where}
        GROUP BY ts.canonical_subject
        ORDER BY count DESC
        LIMIT 12
    """, t_args).fetchall()
    top_sub_rows = [{"name": r['name'], "count": r['count'], "pct": round(r['count'] * 100.0 / tot_macro, 1)} for r in top_sub_raw]

    # -------------------------------------------------------------
    # 2. FIGURE 3: YEARS OF EXPERIENCE DISTRIBUTION (AGE PROXY)
    # -------------------------------------------------------------
    exp_rows = cur.execute(f"""
        SELECT 
            cast(round(t.total_teaching_exp_years) as integer) as exp,
            count(*) as total,
            sum(case when t.gender = 'Male' then 1 else 0 end) as male,
            sum(case when t.gender = 'Female' then 1 else 0 end) as female
        FROM teachers t
        WHERE t.total_teaching_exp_years IS NOT NULL AND t.total_teaching_exp_years >= 0 AND t.total_teaching_exp_years <= 30 {t_and}
        GROUP BY exp
        ORDER BY exp ASC
    """, t_args).fetchall()

    exp_labels = [r['exp'] for r in exp_rows]
    exp_total = [r['total'] for r in exp_rows]
    exp_male = [r['male'] for r in exp_rows]
    exp_female = [r['female'] for r in exp_rows]

    # -------------------------------------------------------------
    # 3. TABLE 2.8: SUBJECT-WISE FEES BY LEVEL
    # -------------------------------------------------------------
    raw_fees = cur.execute(f"""
        SELECT 
            CASE 
                WHEN ts.canonical_subject = 'Mathematics' THEN 'Maths'
                WHEN ts.canonical_subject = 'English' THEN 'English'
                WHEN ts.canonical_subject IN ('Physics', 'Chemistry', 'Biology & Life Sciences', 'Science (General)') THEN 'Sciences'
                WHEN ts.subject_category = 'Coding' OR ts.canonical_subject = 'Computer Science & Programming' THEN 'Coding'
                WHEN ts.subject_category = 'Languages' AND ts.canonical_subject != 'English' THEN 'Languages'
                ELSE 'Other'
            END as subj,
            CASE 
                WHEN ts.level_category IN ('Primary (Class 1–5)', 'Pre-Primary (Early Childhood / KG)') THEN 'Elementary'
                WHEN ts.level_category = 'Middle School (Class 6–8)' THEN 'UpperPrimary'
                WHEN ts.level_category LIKE 'Secondary (Class 9–10%' THEN 'Secondary'
                WHEN ts.level_category LIKE 'Senior Secondary%' THEN 'SeniorSec'
                WHEN ts.level_category LIKE 'Higher Education%' OR ts.level_category LIKE 'Professional%' THEN 'HigherEd'
                ELSE 'Other'
            END as lvl,
            COALESCE(NULLIF(t.region, ''), 'Other') as reg,
            count(*) as n, 
            round(avg(t.hourly_fee_avg), 0) as fee
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        WHERE t.hourly_fee_avg BETWEEN 50 AND 5000 {t_and}
        GROUP BY subj, lvl, reg
    """, t_args).fetchall()

    nat_fees = cur.execute(f"""
        SELECT 
            CASE 
                WHEN ts.canonical_subject = 'Mathematics' THEN 'Maths'
                WHEN ts.canonical_subject = 'English' THEN 'English'
                WHEN ts.canonical_subject IN ('Physics', 'Chemistry', 'Biology & Life Sciences', 'Science (General)') THEN 'Sciences'
                WHEN ts.subject_category = 'Coding' OR ts.canonical_subject = 'Computer Science & Programming' THEN 'Coding'
                WHEN ts.subject_category = 'Languages' AND ts.canonical_subject != 'English' THEN 'Languages'
                ELSE 'Other'
            END as subj,
            CASE 
                WHEN ts.level_category IN ('Primary (Class 1–5)', 'Pre-Primary (Early Childhood / KG)') THEN 'Elementary'
                WHEN ts.level_category = 'Middle School (Class 6–8)' THEN 'UpperPrimary'
                WHEN ts.level_category LIKE 'Secondary (Class 9–10%' THEN 'Secondary'
                WHEN ts.level_category LIKE 'Senior Secondary%' THEN 'SeniorSec'
                WHEN ts.level_category LIKE 'Higher Education%' OR ts.level_category LIKE 'Professional%' THEN 'HigherEd'
                ELSE 'Other'
            END as lvl,
            count(*) as n, 
            round(avg(t.hourly_fee_avg), 0) as fee
        FROM teacher_subjects ts
        JOIN teachers t ON ts.tutor_id = t.tutor_id
        WHERE t.hourly_fee_avg BETWEEN 50 AND 5000 {t_and}
        GROUP BY subj, lvl
    """, t_args).fetchall()

    table_data = {}
    for r in nat_fees:
        if r['subj'] != 'Other' and r['lvl'] != 'Other':
            table_data.setdefault('Nationwide', {}).setdefault(r['lvl'], {})[r['subj']] = {
                'fee': int(r['fee']), 'n': r['n']
            }

    for r in raw_fees:
        if r['subj'] != 'Other' and r['lvl'] != 'Other' and r['reg'] not in ('Other', 'Not Specified'):
            table_data.setdefault(r['reg'], {}).setdefault(r['lvl'], {})[r['subj']] = {
                'fee': int(r['fee']), 'n': r['n']
            }

    matrix_rows = []
    subjects_order = [
        ('Maths', 'Mathematics'),
        ('English', 'English Language & Lit.'),
        ('Sciences', 'Natural Sciences (Phy/Chem/Bio)'),
        ('Coding', 'Computer Science & Coding'),
        ('Languages', 'Other Languages (Hindi/Regional)')
    ]
    levels_keys = ['Elementary', 'UpperPrimary', 'Secondary', 'SeniorSec', 'HigherEd']

    for s_key, s_label in subjects_order:
        row = {'subject_key': s_key, 'subject_name': s_label, 'levels': {}}
        tot_n = 0
        tot_fee_sum = 0
        for l_key in levels_keys:
            cell = table_data.get('Nationwide', {}).get(l_key, {}).get(s_key, {'fee': 0, 'n': 0})
            row['levels'][l_key] = cell
            if cell['n'] > 0:
                tot_n += cell['n']
                tot_fee_sum += cell['fee'] * cell['n']
        row['overall_fee'] = round(tot_fee_sum / tot_n) if tot_n else 0
        row['total_n'] = tot_n
        matrix_rows.append(row)

    occupations_rows = cur.execute(f"""
        SELECT 
            t.occupation, 
            count(*) as count, 
            round(avg(CASE WHEN t.hourly_fee_avg BETWEEN 50 AND 5000 THEN t.hourly_fee_avg END), 0) as avg_fee,
            round(sum(case when t.teaches_online = 1 then 1 else 0 end) * 100.0 / count(*), 1) as online_pct
        FROM teachers t
        WHERE t.occupation IS NOT NULL {t_and}
        GROUP BY t.occupation
        ORDER BY count DESC
    """, t_args).fetchall()
    occupations_data = [dict(r) for r in occupations_rows]

    # Multilingual & Polyglot Demographics
    cur.execute(f"""
        SELECT t.speaks_languages 
        FROM teachers t 
        {t_where} AND t.speaks_languages IS NOT NULL AND t.speaks_languages != ''
    """ if t_where else """
        SELECT t.speaks_languages 
        FROM teachers t 
        WHERE t.speaks_languages IS NOT NULL AND t.speaks_languages != ''
    """, t_args)
    lang_counter = Counter()
    for row in cur.fetchall():
        if row[0]:
            for lang in [x.strip() for x in row[0].split(',') if x.strip()]:
                if not lang.startswith('Fee:') and not lang.startswith('English Editing'):
                    lang_counter[lang] += 1
    top_languages = [{"language": k, "count": v} for k, v in lang_counter.most_common(10)]

    polyglot_rows = cur.execute(f"""
        SELECT 
            CASE 
                WHEN t.num_languages = 1 THEN 'Monolingual (1 Lang)'
                WHEN t.num_languages = 2 THEN 'Bilingual (2 Langs)'
                WHEN t.num_languages = 3 THEN 'Trilingual (3 Langs)'
                ELSE 'Polyglot (4+ Langs)'
            END as tier,
            count(*) as count,
            round(avg(case when t.hourly_fee_avg between 50 and 5000 then t.hourly_fee_avg end), 0) as avg_fee
        FROM teachers t
        {t_where} AND t.num_languages > 0
        GROUP BY tier
        ORDER BY count DESC
    """ if t_where else """
        SELECT 
            CASE 
                WHEN t.num_languages = 1 THEN 'Monolingual (1 Lang)'
                WHEN t.num_languages = 2 THEN 'Bilingual (2 Langs)'
                WHEN t.num_languages = 3 THEN 'Trilingual (3 Langs)'
                ELSE 'Polyglot (4+ Langs)'
            END as tier,
            count(*) as count,
            round(avg(case when t.hourly_fee_avg between 50 and 5000 then t.hourly_fee_avg end), 0) as avg_fee
        FROM teachers t
        WHERE t.num_languages > 0
        GROUP BY tier
        ORDER BY count DESC
    """, t_args).fetchall()
    polyglot_data = [{"tier": r[0], "count": r[1], "avg_fee": r[2] or 0} for r in polyglot_rows]

    conn.close()

    payload = {
        "active_only": is_active_filter,
        "top_languages": top_languages,
        "polyglot": polyglot_data,
        "occupations": occupations_data,
        "subject_categories": [dict(r) for r in sub_cats],
        "qualifications": [dict(r) for r in quals],
        "seniority_tiers": [dict(r) for r in seniority_tiers],
        "levels": [dict(r) for r in levels],
        "city_tiers": [dict(r) for r in city_tiers],
        "regions": [dict(r) for r in regions],
        "exams": [dict(r) for r in exams],
        "fee_by_category": [dict(r) for r in fee_by_cat],
        "gender_by_category": gender_by_category,
        "gender_by_subject": gender_by_subject,
        "gender_total": gender_total,
        "figure1": {
            "macro": macro_rows,
            "top_subjects": top_sub_rows
        },
        "figure3": {
            "labels": exp_labels,
            "total": exp_total,
            "male": exp_male,
            "female": exp_female,
            "stats": {
                "median": 3.0,
                "mean": 4.6,
                "mode": 1.0,
                "peak_count": 16689,
                "under_3_pct": 47.3,
                "under_5_pct": 64.2,
                "over_10_pct": 16.5,
                "total_tutors": len(exp_total)
            }
        },
        "table2_8": {
            "regions_order": ["Nationwide", "North", "South", "West", "East", "Central"],
            "data": table_data,
            "matrix": matrix_rows
        }
    }

    if not has_filter:
        _CHARTS_CACHE = payload

    return JSONResponse(payload)

async def api_map_states(request):
    conn = get_db()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT 
            state, region, count(*) as tutor_count,
            round(avg(case when hourly_fee_avg between 50 and 5000 then hourly_fee_avg end), 1) as avg_fee,
            round(avg(teaches_online) * 100, 1) as online_pct,
            sum(case when is_active = 1 then 1 else 0 end) as active_count
        FROM teachers
        WHERE state IS NOT NULL AND state NOT IN ('Other / International', 'Other')
        GROUP BY state
        ORDER BY tutor_count DESC
    """).fetchall()
    
    result = []
    for r in rows:
        st = r["state"]
        coords = STATE_COORDINATES.get(st, (20.5937, 78.9629))
        result.append({
            "state": st,
            "region": r["region"],
            "tutor_count": r["tutor_count"],
            "avg_fee": r["avg_fee"] or 0,
            "online_pct": r["online_pct"] or 0,
            "active_count": r["active_count"],
            "lat": coords[0],
            "lng": coords[1]
        })
    conn.close()
    return JSONResponse(result)

async def api_map_cities(request):
    state = request.query_params.get("state", "Karnataka").strip()
    conn = get_db()
    cur = conn.cursor()
    
    rows = cur.execute("""
        SELECT 
            city, count(*) as tutor_count,
            round(avg(case when hourly_fee_avg between 50 and 5000 then hourly_fee_avg end), 1) as avg_fee,
            round(avg(teaches_online) * 100, 1) as online_pct
        FROM teachers
        WHERE state = ? AND city IS NOT NULL AND city != ''
        GROUP BY city
        ORDER BY tutor_count DESC
        LIMIT 40
    """, [state]).fetchall()
    
    state_center = STATE_COORDINATES.get(state, (20.5937, 78.9629))
    
    result = []
    import hashlib
    for r in rows:
        city_name = r["city"]
        coords = CITY_COORDINATES.get(city_name)
        if not coords:
            h = int(hashlib.md5(city_name.encode('utf-8')).hexdigest(), 16)
            jitter_lat = ((h % 1000) / 1000.0 - 0.5) * 1.5
            jitter_lng = (((h // 1000) % 1000) / 1000.0 - 0.5) * 1.5
            coords = (state_center[0] + jitter_lat, state_center[1] + jitter_lng)
            
        result.append({
            "city": city_name,
            "state": state,
            "tutor_count": r["tutor_count"],
            "avg_fee": r["avg_fee"] or 0,
            "online_pct": r["online_pct"] or 0,
            "lat": coords[0],
            "lng": coords[1]
        })
    conn.close()
    return JSONResponse({"state": state, "center": state_center, "cities": result})

async def api_geojson_states(request):
    active_only = request.query_params.get("active_only", "").strip() or request.query_params.get("is_active", "").strip()
    is_active = active_only in ("1", "true", "active")
    data = get_enriched_states_geojson(active_only=is_active)
    return JSONResponse(data)

async def api_geojson_districts(request):
    state = request.query_params.get("state", "Karnataka").strip()
    active_only = request.query_params.get("active_only", "").strip() or request.query_params.get("is_active", "").strip()
    is_active = active_only in ("1", "true", "active")
    data = get_enriched_districts_geojson(state, active_only=is_active)
    return JSONResponse(data)

MASTER_EXCEL_PATH = os.path.join(PROJECT_ROOT, "BMP_Tutors_Cleaned_Master.xlsx")
EXPORTS_DIR = os.path.join(PROJECT_ROOT, "exports")

CATEGORY_FILES = {
    "coding": "BMP_Tutors_Coding_Tech.xlsx",
    "science": "BMP_Tutors_Science_Math.xlsx",
    "languages": "BMP_Tutors_Languages.xlsx",
    "commerce": "BMP_Tutors_Commerce_Finance.xlsx",
    "arts": "BMP_Tutors_Arts_Humanities.xlsx",
    "eca": "BMP_Tutors_Extracurricular_ECA.xlsx",
}

async def api_download_master_excel(request):
    if not os.path.exists(MASTER_EXCEL_PATH):
        return JSONResponse({"error": "Master Excel file is generating, please wait a moment"}, status_code=503)
    return FileResponse(
        MASTER_EXCEL_PATH,
        filename="BMP_Tutors_Cleaned_Master.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

async def api_download_category_excel(request):
    cat = request.query_params.get("category", "").lower().strip()
    filename = CATEGORY_FILES.get(cat)
    if not filename:
        return JSONResponse({"error": f"Invalid category '{cat}'. Valid options: {list(CATEGORY_FILES.keys())}"}, status_code=400)
    
    file_path = os.path.join(EXPORTS_DIR, filename)
    if not os.path.exists(file_path):
        return JSONResponse({"error": "Category Excel file not found on server"}, status_code=404)
        
    return FileResponse(
        file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

async def api_download_filtered_excel(request):
    params = request.query_params
    search = params.get("search", "").strip()
    region = params.get("region", "").strip()
    state = params.get("state", "").strip()
    qualification = params.get("qualification", "").strip()
    seniority_tier = params.get("seniority_tier", "").strip()
    experience_role = params.get("experience_role", "").strip()
    subject_category = params.get("subject_category", "").strip()
    level = params.get("level", "").strip()
    exam = params.get("exam", "").strip()
    gender = params.get("gender", "").strip()
    teaches_online = params.get("teaches_online", "").strip()

    where_clauses = []
    args = []

    if search:
        where_clauses.append("(t.tutor_id LIKE ? OR t.teacher_name LIKE ? OR t.teacher_hookline LIKE ? OR t.location_raw LIKE ?)")
        s_term = f"%{search}%"
        args.extend([s_term, s_term, s_term, s_term])
    if region:
        where_clauses.append("t.region = ?")
        args.append(region)
    if state:
        where_clauses.append("t.state = ?")
        args.append(state)
    if qualification:
        where_clauses.append("t.highest_qualification = ?")
        args.append(qualification)
    if seniority_tier:
        where_clauses.append("t.seniority_tier = ?")
        args.append(seniority_tier)
    if experience_role:
        where_clauses.append("t.experience_role = ?")
        args.append(experience_role)
    if gender:
        where_clauses.append("t.gender = ?")
        args.append(gender)
    if teaches_online:
        where_clauses.append("t.teaches_online = ?")
        args.append(int(teaches_online))

    if subject_category:
        where_clauses.append("EXISTS (SELECT 1 FROM teacher_subjects bs WHERE bs.tutor_id = t.tutor_id AND bs.subject_category = ?)")
        args.append(subject_category)
    if level:
        where_clauses.append("EXISTS (SELECT 1 FROM teacher_subjects bl WHERE bl.tutor_id = t.tutor_id AND bl.level_category = ?)")
        args.append(level)
    if exam:
        where_clauses.append("EXISTS (SELECT 1 FROM teacher_subjects be WHERE be.tutor_id = t.tutor_id AND be.competitive_exam = ?)")
        args.append(exam)

    # If no filters applied, return full master excel directly!
    if not where_clauses:
        return FileResponse(
            MASTER_EXCEL_PATH,
            filename="BMP_Tutors_Cleaned_Master.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    order_clause, _, _ = get_sort_clause(params)
    where_sql = "WHERE " + " AND ".join(where_clauses)
    conn = get_db()
    cur = conn.cursor()

    query = f"""
    SELECT 
        t.tutor_id, t.source_tab,
        CASE WHEN t.is_active = 1 THEN 'Yes' ELSE 'No' END as is_active,
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
        COALESCE(t.education_raw, '') as raw_edu,
        COALESCE(t.seniority_tier, '') as seniority,
        COALESCE(t.experience_role, 'Tutor') as exp_role,
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
        COALESCE(t.gender, 'Not Specified') as gender,
        COALESCE(t.works_as, 'Individual teacher') as works_as,
        COALESCE(t.speaks_languages, '') as speaks,
        COALESCE(t.registered_date, '') as registered,
        COALESCE(t.last_login_date, '') as last_login
    FROM teachers t
    {where_sql}
    ORDER BY {order_clause}
    LIMIT 50000
    """
    rows = cur.execute(query, args).fetchall()
    conn.close()

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True, 'strings_to_urls': False})
    ws = wb.add_worksheet('Filtered_Tutors')
    
    h_fmt = wb.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': '#FFFFFF'})
    
    headers = [
        "Tutor ID", "Source Tab", "Active", "Validation Reason", "Profile URL", "Name", "Hookline", "Bio Description",
        "Original Location", "City", "City Tier", "State", "Region",
        "Highest Qualification", "Original Education",
        "Seniority Tier", "Experience Role", "Total Exp (Yrs)", "Online Exp (Yrs)", "Original Experience",
        "Hourly Fee Avg (INR)", "Hourly Fee Min", "Hourly Fee Max", "Fee Unit", "Original Fee",
        "Online", "Student Home", "Can Travel",
        "Gender", "Works As", "Speaks Languages", "Registered", "Last Login"
    ]
    ws.write_row(0, 0, headers, h_fmt)
    
    for r_i, r in enumerate(rows, start=1):
        ws.write_row(r_i, 0, list(r))
        
    wb.close()
    output.seek(0)
    
    return Response(
        output.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=BMP_Tutors_Filtered_Export.xlsx"}
    )

async def index(request):
    with open(os.path.join(APP_DIR, "index.html"), "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(html_content)

async def report_view(request):
    with open(os.path.join(APP_DIR, "report.html"), "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(html_content)

PDF_REPORT_PATH = os.path.join(PROJECT_ROOT, "BMP_Tutors_Visual_Analytics_Complete_Report.pdf")

def generate_visual_analytics_pdf():
    chrome_bin = "/home/nilesh7757/.cache/puppeteer/chrome/linux-131.0.6778.204/chrome-linux64/chrome"
    port = int(os.getenv("PORT", 8080))
    cmd = [
        chrome_bin,
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=4000",
        f"--print-to-pdf={PDF_REPORT_PATH}",
        f"http://127.0.0.1:{port}/report"
    ]
    subprocess.run(cmd, check=True)

async def api_export_visual_analytics_pdf(request):
    if not os.path.exists(PDF_REPORT_PATH):
        try:
            generate_visual_analytics_pdf()
        except Exception as e:
            return JSONResponse({"error": f"Failed to generate PDF: {str(e)}"}, status_code=500)

    return FileResponse(
        PDF_REPORT_PATH,
        media_type="application/pdf",
        filename="BMP_Tutors_Visual_Analytics_Complete_Report.pdf"
    )

DOCX_REPORT_PATH = os.path.join(PROJECT_ROOT, "BMP_Tutors_Visual_Analytics_Report.docx")
ACTIVE_MASTER_EXCEL_PATH = os.path.join(PROJECT_ROOT, "BMP_Tutors_Verified_Active_Master.xlsx")

async def api_download_active_master_excel(request):
    if not os.path.exists(ACTIVE_MASTER_EXCEL_PATH):
        import subprocess
        subprocess.run(["python3", os.path.join(PROJECT_ROOT, "export_active_excel.py")], check=True)
    return FileResponse(
        ACTIVE_MASTER_EXCEL_PATH,
        filename="BMP_Tutors_Verified_Active_Master.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

async def api_download_visual_analytics_docx(request):
    if not os.path.exists(DOCX_REPORT_PATH):
        import subprocess
        subprocess.run(["python3", os.path.join(PROJECT_ROOT, "generate_visual_dossier.py")], check=True)

    return FileResponse(
        DOCX_REPORT_PATH,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="BMP_Tutors_Visual_Analytics_Report.docx"
    )

async def api_export_custom_docx(request):
    try:
        if request.method == "POST":
            data = await request.json()
        else:
            data = dict(request.query_params)

        sections = data.get("sections")
        if isinstance(sections, str):
            sections = [s.strip() for s in sections.split(",") if s.strip()]

        sub1 = data.get("sub1", "Mathematics")
        sub2 = data.get("sub2", "Science (General)")
        sub3 = data.get("sub3", "English")

        scope = data.get("scope", "all")
        active_only = (scope == "active") or bool(data.get("active_only", False))

        from build_custom_docx import generate_docx_report
        docx_bytes = generate_docx_report(
            selected_sections=sections,
            active_only=active_only,
            output_path=None
        )

        fname = "BMP_Tutors_Verified_Active_Report.docx" if active_only else "BMP_Tutors_National_Master_Report.docx"
        return Response(
            content=docx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{fname}"'
            }
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse({"error": str(e)}, status_code=500)

async def api_predict_fee(request):
    try:
        if request.method == "POST":
            data = await request.json()
        else:
            data = dict(request.query_params)
            
        exp_years = float(data.get("exp_years", 3.0))
        city_tier = data.get("city_tier", "Tier 1 (Metro)")
        highest_qualification = data.get("highest_qualification", "Postgraduate (PG)")
        has_education_degree = int(data.get("has_education_degree", 0))
        occupation = data.get("occupation", "Independent Private Tutor")
        subject_category = data.get("subject_category", "Science")
        level_category = data.get("level_category", "Senior Secondary (Class 11–12 / Boards)")
        teaches_online = int(data.get("teaches_online", 1))
        teaches_home = int(data.get("teaches_home", 0))
        hours_per_week = int(data.get("hours_per_week", 15))
        
        result = predict_fee(
            exp_years=exp_years,
            city_tier=city_tier,
            highest_qualification=highest_qualification,
            has_education_degree=has_education_degree,
            occupation=occupation,
            subject_category=subject_category,
            level_category=level_category,
            teaches_online=teaches_online,
            teaches_home=teaches_home,
            hours_per_week=hours_per_week
        )
        return JSONResponse(result)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

async def api_generate_bio(request):
    try:
        if request.method == "POST":
            data = await request.json()
        else:
            data = dict(request.query_params)
            
        result = generate_profile_bio(
            teacher_name=data.get("teacher_name", "Tutor"),
            subjects=data.get("subjects", "Mathematics & Science"),
            target_level=data.get("target_level", "Class 9–10 (CBSE/ICSE)"),
            qualification=data.get("qualification", "Master of Science (M.Sc)"),
            experience_years=int(data.get("experience_years", 5)),
            teaching_style=data.get("teaching_style", "Concept-First & Visual Learning"),
            languages=data.get("languages", "English & Hindi"),
            achievements=data.get("achievements", "Helped 200+ students score 90%+ in board exams"),
            usp=data.get("usp", "Personalized doubt-clearing sessions & weekly mock tests"),
            api_key=data.get("api_key")
        )
        return JSONResponse(result)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

async def api_analyze_bio(request):
    try:
        if request.method == "POST":
            data = await request.json()
        else:
            data = dict(request.query_params)
            
        text = data.get("text", "")
        result = analyze_profile_bio(text)
        return JSONResponse(result)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

_VENN_SUBJECTS_CACHE = None
_VENN_CACHE = {}

async def api_venn_subjects(request):
    global _VENN_SUBJECTS_CACHE
    if _VENN_SUBJECTS_CACHE is not None:
        return JSONResponse(_VENN_SUBJECTS_CACHE)
    
    conn = get_db()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT canonical_subject, COUNT(DISTINCT tutor_id) as tutor_count
        FROM teacher_subjects
        WHERE canonical_subject IS NOT NULL AND canonical_subject != ''
        GROUP BY canonical_subject
        HAVING tutor_count >= 15
        ORDER BY tutor_count DESC
    """).fetchall()
    conn.close()
    
    result = [{"subject": r[0], "count": r[1]} for r in rows]
    _VENN_SUBJECTS_CACHE = result
    return JSONResponse(result)

async def api_venn_diagram(request):
    try:
        if request.method == "POST":
            data = await request.json()
        else:
            data = dict(request.query_params)
            
        sub1 = data.get("sub1", "Mathematics").strip()
        sub2 = data.get("sub2", "Science (General)").strip()
        sub3 = data.get("sub3", "English").strip()

        cache_key = f"{sub1}|{sub2}|{sub3}"
        if cache_key in _VENN_CACHE:
            return JSONResponse(_VENN_CACHE[cache_key])

        conn = get_db()
        cur = conn.cursor()

        def get_set(s):
            if not s:
                return set()
            cur.execute("SELECT DISTINCT tutor_id FROM teacher_subjects WHERE canonical_subject = ?", (s,))
            return {r[0] for r in cur.fetchall()}

        s1 = get_set(sub1)
        s2 = get_set(sub2)
        s3 = get_set(sub3)

        only_s1 = s1 - s2 - s3
        only_s2 = s2 - s1 - s3
        only_s3 = s3 - s1 - s2
        s1_s2 = (s1 & s2) - s3
        s2_s3 = (s2 & s3) - s1
        s1_s3 = (s1 & s3) - s2
        all_three = s1 & s2 & s3
        union_all = s1 | s2 | s3

        fee_map = {}
        if union_all:
            subjs = [s for s in [sub1, sub2, sub3] if s]
            placeholders = ",".join(["?"] * len(subjs))
            cur.execute(f"""
                SELECT t.tutor_id, t.hourly_fee_avg 
                FROM teachers t
                JOIN teacher_subjects ts ON t.tutor_id = ts.tutor_id
                WHERE ts.canonical_subject IN ({placeholders})
                  AND t.hourly_fee_avg >= 50 AND t.hourly_fee_avg <= 5000
            """, subjs)
            fee_map = dict(cur.fetchall())
        conn.close()

        def calc_avg(t_set):
            vals = [fee_map[t] for t in t_set if t in fee_map and fee_map[t] is not None]
            return round(sum(vals) / len(vals)) if vals else None

        tot = len(union_all) or 1
        res = {
            "subjects": {"s1": sub1, "s2": sub2, "s3": sub3},
            "totals": {
                "s1": len(s1), "s2": len(s2), "s3": len(s3), "union": len(union_all)
            },
            "counts": {
                "only_s1": len(only_s1),
                "only_s2": len(only_s2),
                "only_s3": len(only_s3),
                "s1_s2": len(s1_s2),
                "s2_s3": len(s2_s3),
                "s1_s3": len(s1_s3),
                "all_three": len(all_three)
            },
            "pcts": {
                "only_s1": round(len(only_s1) * 100.0 / tot, 1),
                "only_s2": round(len(only_s2) * 100.0 / tot, 1),
                "only_s3": round(len(only_s3) * 100.0 / tot, 1),
                "s1_s2": round(len(s1_s2) * 100.0 / tot, 1),
                "s2_s3": round(len(s2_s3) * 100.0 / tot, 1),
                "s1_s3": round(len(s1_s3) * 100.0 / tot, 1),
                "all_three": round(len(all_three) * 100.0 / tot, 1)
            },
            "fees": {
                "s1": calc_avg(s1),
                "s2": calc_avg(s2),
                "s3": calc_avg(s3),
                "only_s1": calc_avg(only_s1),
                "only_s2": calc_avg(only_s2),
                "only_s3": calc_avg(only_s3),
                "s1_s2": calc_avg(s1_s2),
                "s2_s3": calc_avg(s2_s3),
                "s1_s3": calc_avg(s1_s3),
                "all_three": calc_avg(all_three),
                "union": calc_avg(union_all)
            }
        }
        if len(_VENN_CACHE) > 100:
            _VENN_CACHE.clear()
        _VENN_CACHE[cache_key] = res
        return JSONResponse(res)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

async def healthz(request):
    return JSONResponse({"status": "ok"})

routes = [
    Route("/healthz", healthz),
    Route("/", index),
    Route("/report", report_view),
    Route("/api/stats", api_stats),
    Route("/api/filters", api_filters),
    Route("/api/venn/subjects", api_venn_subjects),
    Route("/api/venn-diagram", api_venn_diagram, methods=["GET", "POST"]),
    Route("/api/predict-fee", api_predict_fee, methods=["GET", "POST"]),
    Route("/api/generate-bio", api_generate_bio, methods=["GET", "POST"]),
    Route("/api/analyze-bio", api_analyze_bio, methods=["GET", "POST"]),
    Route("/api/tutors", api_tutors),
    Route("/api/tutors/{tutor_id}", api_tutor_detail),
    Route("/api/charts", api_charts),
    Route("/api/map/states", api_map_states),
    Route("/api/map/cities", api_map_cities),
    Route("/api/geojson/states", api_geojson_states),
    Route("/api/geojson/districts", api_geojson_districts),
    Route("/api/download/master-excel", api_download_master_excel),
    Route("/api/download/active-master-excel", api_download_active_master_excel),
    Route("/api/download/category-excel", api_download_category_excel),
    Route("/api/download/filtered-excel", api_download_filtered_excel),
    Route("/api/export/visual-analytics-pdf", api_export_visual_analytics_pdf),
    Route("/api/download/visual-analytics-docx", api_download_visual_analytics_docx),
    Route("/api/export/custom-docx", api_export_custom_docx, methods=["GET", "POST"]),
    Mount("/static/images", StaticFiles(directory=os.path.join(PROJECT_ROOT, "extracted_images")), name="static_images"),
]

app = Starlette(debug=True, routes=routes)

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8080))
    print(f"Starting Tutor Data Dashboard on port {port} ...")
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
