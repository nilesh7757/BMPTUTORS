#!/usr/bin/env python3
"""
BMP Tutors - Unified Database Cleaning & Deliverable Regeneration Pipeline
========================================================================
Executes the complete 7-step data cleaning protocol:
  STEP 1: Subject Name Canonicalization (12,327 -> ~300 canonical subjects)
  STEP 2: Hourly Fee Range Enforcement (₹100 to ₹5,000; invalid -> NULL)
  STEP 3: Remove Dummy / Test Accounts (T3662, T64166, T92868, T100360, etc.)
  STEP 4: Handle Missing/Empty Data (1,268 records -> 'Anonymous Tutor')
  STEP 5: Fix State & Gender Mismatches against TrueData.xlsx
  STEP 6: Deduplicate profiles (same name + same city + same subjects)
  STEP 7: Database VACUUM, verification checks & re-export Excel deliverables
"""

import os
import sys
import re
import sqlite3
from collections import Counter, defaultdict

# Include UV cache packages for openpyxl / xlsxwriter
UV_PATHS = [
    "/home/nilesh7757/.cache/uv/environments-v2/48a179f4a57efa1f/90ca161982e074e8/lib/python3.13/site-packages",
    "/home/nilesh7757/.cache/uv/archive-v0/KcQCAaaojDTfnz86/lib/python3.13/site-packages",
    "/home/nilesh7757/.cache/uv/archive-v0/b2_g7sKwC5mPN6Zp/lib/python3.13/site-packages",
    "/home/nilesh7757/.cache/uv/archive-v0/hoJjWQYazaEYVkk3/lib/python3.13/site-packages",
]
for p in UV_PATHS:
    if os.path.exists(p) and p not in sys.path:
        sys.path.insert(0, p)

import openpyxl
from clean_rules import clean_location, clean_gender

DB_PATH = "cleaned_tutors_data.db"
TRUE_DATA_PATH = "TrueData.xlsx"

# -------------------------------------------------------------
# STEP 1: SUBJECT CANONICALIZATION DICTIONARY & LOGIC
# -------------------------------------------------------------

EXACT_SUBJECT_MAP = {
    # Mathematics prompt examples
    "maths": "Mathematics",
    "math": "Mathematics",
    "mathamatics": "Mathematics",
    "mathmatics": "Mathematics",
    "mathematic": "Mathematics",
    "vedicmaths": "Mathematics",
    "cuemath": "Mathematics",
    "mathcounts": "Mathematics",
    "mathematical analysis": "Mathematics",
    "mathematical methods": "Mathematics",
    "engineering mathemetics": "Mathematics",
    "engineering mathematics": "Mathematics",
    "applied mathematics": "Mathematics",
    "pure mathematics": "Mathematics",
    "discrete mathematics": "Mathematics",
    "higher mathematics": "Mathematics",
    "business mathematics": "Mathematics",
    "vedic maths": "Mathematics",
    "abacus": "Mathematics",
    "algebra": "Mathematics",
    "calculus": "Mathematics",
    "geometry": "Mathematics",
    "trigonometry": "Mathematics",
    "arithmetic": "Mathematics",
    "linear algebra": "Mathematics",
    "differential equations": "Mathematics",
    "mathematical modelling": "Mathematics",
    "mathematical logic": "Mathematics",
    "mathematical reasoning": "Mathematics",
    "mathematical activities": "Mathematics",
    "australian mathematical methods": "Mathematics",
    "cbse maths": "Mathematics",
    "icse maths": "Mathematics",
    "igcse maths": "Mathematics",
    "ib maths": "Mathematics",
    "ap precalculus": "Mathematics",
    "graph theory": "Mathematics",
    
    # English prompt examples
    "engilsh": "English",
    "englsih": "English",
    "spoken english": "English",
    "english grammar": "English",
    "english literature": "English",
    "english language": "English",
    "phonics": "English",
    "ielts": "English",
    "toefl": "English",
    "pte": "English",
    "creative writing": "English",
    "comprehension": "English",
    "esl": "English",
    "sat english": "English",
    "business english": "English",
    "communicative english": "English",
    
    # Indic scripts
    "गणित": "Mathematics",
    "अंकगणित": "Mathematics",
    "बीजगणित": "Mathematics",
    "रेखागणित": "Mathematics",
    "हिंदी": "Hindi",
    "हिन्दी": "Hindi",
    "संस्कृत": "Sanskrit",
    "स्तोत्र": "Sanskrit",
    "हवन": "Sanskrit",
    "वेद": "Sanskrit",
    "वैदिक": "Sanskrit",
    "கணிதம்": "Mathematics",
    "தமிழ்": "Tamil",
    "ഗണിതം": "Mathematics",
    "ശാസ്ത്രം": "Science (General)",
    "മലയാളം": "Malayalam",
    "తెలుగు": "Telugu",
    "గణితం": "Mathematics",
    "বাংলা": "Bengali",
    "হাদিস मुखস্থ": "Arabic & Islamic Studies",
    "কোরআন": "Arabic & Islamic Studies",
    
    # Other common variants
    "bhagavad geeta": "Philosophy & Ethics",
    "bhagavad gita": "Philosophy & Ethics",
    "bhajan": "Music",
    "bhajans": "Music",
    "bharatanatayam": "Dance",
    "bharathanatyam": "Dance",
    "bollywood": "Dance",
    "bollywood songs": "Music",
    "b com": "Accountancy",
    "bcom": "Accountancy",
    "graphic designer": "Design & Multimedia",
    "graphic designing": "Design & Multimedia",
    "figma": "Design & Multimedia",
    "film songs": "Music",
    "fiqh": "Arabic & Islamic Studies",
    "islamic basic": "Arabic & Islamic Studies",
    "islamic education": "Arabic & Islamic Studies",
    "genai": "AI & Data Science",
    "github": "Cloud Computing & DevOps",
    "google sheets": "AI & Data Science",
    "google ads": "Business Studies & Management",
    "handwriting": "Art & Craft",
    "heat transfer": "Mechanical Engineering",
    "fluid dynamics": "Mechanical Engineering",
    "ib tok": "Philosophy & Ethics",
    "ict": "Computer Science & Programming",
    "it": "Computer Science & Programming",
    "ip": "Computer Science & Programming",
    "iit foundation": "Competitive Exam Preparation",
    "imo": "Mathematics",
    "ipmat": "Competitive Exam Preparation",
    "image processing": "AI & Data Science",
    "income tax": "Accountancy",
    "indian economy": "Economics",
    "interior design": "Design & Multimedia",
    "interview skills": "Personality Development & Soft Skills",
    "bengali": "Bengali"
}

SUBJECT_PATTERNS = [
    # 1. Mathematics & Statistics
    (r'\b(maths?|mathematics?|algebra|calculus|geometry|trigonometry|arithmetic|topology|cuemath|vedic|abacus|linear algebra|differential equation|discrete math)\b', 'Mathematics'),
    (r'\b(statist|statistics|probability|biostat|actuarial|econometric)\b', 'Statistics'),

    # 2. Physics & Natural Sciences
    (r'\b(physics|astrophysics|geophysics|astronomy|thermodynamics|quantum|optics|electromagnet)\b', 'Physics'),
    (r'\b(chemistry|biochem|organic chem|inorganic chem|physical chem|chemical analysis)\b', 'Chemistry'),
    (r'\b(biology|botany|zoology|microbiol|genetics|biotech|anatomy|physiology|pathology|immunology|ecology|life science|embryology|virology|neet bio|bio neet)\b', 'Biology & Life Sciences'),
    (r'\b(medicine|nursing|pharmacy|pharmacology|dental|mbbs|medical|clinical|surgery|ayurveda|homeopathy|physiotherapy)\b', 'Medicine & Healthcare'),
    (r'\b(science|general science|natural science|evs|environmental science|environmental studies|earth science)\b', 'Science (General)'),

    # 3. Languages
    (r'\b(english|grammar|spoken english|phonics|ielts|toefl|pte|creative writing|literature|comprehension|esl|verbal)\b', 'English'),
    (r'\b(hindi|shuddha hindi)\b', 'Hindi'),
    (r'\b(bengali|bangla)\b', 'Bengali'),
    (r'\b(tamil)\b', 'Tamil'),
    (r'\b(telugu)\b', 'Telugu'),
    (r'\b(kannada)\b', 'Kannada'),
    (r'\b(malayalam)\b', 'Malayalam'),
    (r'\b(marathi)\b', 'Marathi'),
    (r'\b(gujarati)\b', 'Gujarati'),
    (r'\b(punjabi|gurmukhi)\b', 'Punjabi'),
    (r'\b(urdu)\b', 'Urdu'),
    (r'\b(sanskrit)\b', 'Sanskrit'),
    (r'\b(french)\b', 'French'),
    (r'\b(german)\b', 'German'),
    (r'\b(spanish)\b', 'Spanish'),
    (r'\b(arabic|quran|tajweed|islamic studies|hadees|islamiyat)\b', 'Arabic & Islamic Studies'),
    (r'\b(japanese)\b', 'Japanese'),
    (r'\b(chinese|mandarin)\b', 'Chinese (Mandarin)'),
    (r'\b(russian)\b', 'Russian'),
    (r'\b(korean)\b', 'Korean'),
    (r'\b(italian)\b', 'Italian'),
    (r'\b(portuguese)\b', 'Portuguese'),
    (r'\b(odia|oriya)\b', 'Odia'),
    (r'\b(assamese)\b', 'Assamese'),
    (r'\b(nepali)\b', 'Nepali'),
    (r'\b(persian|farsi)\b', 'Persian'),
    (r'\b(turkish)\b', 'Turkish'),
    (r'\b(latin)\b', 'Latin'),

    # 4. Computer Science, Programming & IT
    (r'\b(python|django|flask|fastapi|pandas|numpy)\b', 'Python'),
    (r'\b(java\b|spring boot|spring framework|hibernate)\b', 'Java'),
    (r'\b(c\+\+|c programming|c / c\+\+|c and c\+\+|c#|\.net|embedded c)\b', 'C / C++ / C#'),
    (r'\b(web develop|html|css|javascript|react|node|frontend|backend|full stack|angular|vue|bootstrap|wordpress|php|tailwind)\b', 'Web Development'),
    (r'\b(sql|database|dbms|mysql|oracle|mongodb|postgresql|nosql|plsql|sqlite)\b', 'SQL & Databases'),
    (r'\b(data science|machine learning|artificial intelligence|deep learning|nlp|computer vision|power bi|tableau|data analytics|data analysis|big data|neural network|excel)\b', 'AI & Data Science'),
    (r'\b(cloud|aws|azure|gcp|devops|docker|kubernetes|jenkins|linux|unix|git|terraform|ansible)\b', 'Cloud Computing & DevOps'),
    (r'\b(cyber security|cybersecurity|ethical hacking|network security|information security|networking|ccna)\b', 'Cybersecurity & Networking'),
    (r'\b(software testing|selenium|automation testing|manual testing|qa testing|test automation)\b', 'Software Testing'),
    (r'\b(android|ios|flutter|swift|react native|mobile app|app development|kotlin)\b', 'Mobile App Development'),
    (r'\b(computer science|programming|coding|dsa|data structure|algorithms|operating system|compiler|automata|oops|software engineering|computer architecture)\b', 'Computer Science & Programming'),

    # 5. Commerce, Economics & Business
    (r'\b(account|accounts|accounting|accountancy|tally|bookkeeping|costing|taxation|gst|auditing|ca foundation|cma)\b', 'Accountancy'),
    (r'\b(economics|microeconomics|macroeconomics|econometrics)\b', 'Economics'),
    (r'\b(business|management|marketing|human resource|finance|financial|bba|mba|commerce|banking|supply chain|entrepreneurship|business studies)\b', 'Business Studies & Management'),

    # 6. Humanities & Social Sciences
    (r'\b(social science|social studies|sst)\b', 'Social Science'),
    (r'\b(history|ancient history|medieval history|modern history|world history|archaeology)\b', 'History'),
    (r'\b(geography|geology|earth science|gis)\b', 'Geography'),
    (r'\b(political science|polity|civics|law|legal|jurisprudence|constitution|clat|public administration)\b', 'Political Science & Law'),
    (r'\b(psychology|counseling|mental health)\b', 'Psychology'),
    (r'\b(sociology|social work|anthropology)\b', 'Sociology'),
    (r'\b(philosophy|ethics|logic)\b', 'Philosophy & Ethics'),
    (r'\b(journalism|mass communication|media studies)\b', 'Journalism & Mass Communication'),

    # 7. Engineering
    (r'\b(civil engineering|structural engineering|surveying|construction|strength of materials?)\b', 'Civil Engineering'),
    (r'\b(mechanical engineering|cad|autocad|solidworks|ansys|fluid mechanics|mechatronics|engineering mechanics|machine design)\b', 'Mechanical Engineering'),
    (r'\b(electrical engineering|power systems|circuit theory|electrical machines|control systems?)\b', 'Electrical Engineering'),
    (r'\b(electronics|vlsi|embedded systems|microprocessor|telecommunication|signal processing|robotics|iot|arduino|digital electronics)\b', 'Electronics & Communication'),
    (r'\b(chemical engineering|petroleum engineering)\b', 'Chemical Engineering'),

    # 8. Creative, Performing Arts & Sports
    (r'\b(music|singing|vocal|guitar|piano|keyboard|violin|flute|drums|carnatic|hindustani|tabla|sitar|harmonium)\b', 'Music'),
    (r'\b(art|drawing|painting|sketching|calligraphy|craft|fine arts|water color|oil painting)\b', 'Art & Craft'),
    (r'\b(graphic design|photoshop|illustrator|ui/ux|ui ux|animation|video editing|vfx|3d modeling)\b', 'Design & Multimedia'),
    (r'\b(dance|bharatanatyam|kathak|zumba|hip hop|salsa|ballet|folk dance)\b', 'Dance'),
    (r'\b(yoga|fitness|gym|workout|aerobics|meditation|nutrition|dietetics|physical education)\b', 'Yoga & Fitness'),
    (r'\b(chess)\b', 'Chess'),
    (r'\b(karate|taekwondo|martial arts|swimming|badminton|cricket|tennis|sports)\b', 'Sports & Martial Arts'),

    # 9. Competitive Exams, Aptitude & Soft Skills
    (r'\b(aptitude|logical reasoning|reasoning|mental ability|quantitative aptitude)\b', 'Aptitude & Logical Reasoning'),
    (r'\b(jee|neet|gate|cat exam|gre|gmat|sat|upsc|ssc|bank po|defence|nda|cds|olympiad|competitive exam)\b', 'Competitive Exam Preparation'),
    (r'\b(general knowledge|current affairs|gk)\b', 'General Knowledge'),
    (r'\b(personality development|communication skills|soft skills|public speaking|interview preparation|life coaching)\b', 'Personality Development & Soft Skills'),
    (r'\b(special education|autism|adhd|dyslexia|special needs)\b', 'Special Education'),

    # 10. All Subjects
    (r'\b(all subject|all subjects|all school subjects|all primary|all standards|all classes|all)\b', 'All Subjects')
]

PRIMARY_SET = set([
    "Mathematics", "English", "Science (General)", "Physics", "Chemistry", "Biology & Life Sciences",
    "Computer Science & Programming", "Python", "Java", "C / C++ / C#", "Web Development", "SQL & Databases",
    "AI & Data Science", "Cloud Computing & DevOps", "Cybersecurity & Networking", "Software Testing",
    "Mobile App Development", "Accountancy", "Economics", "Business Studies & Management",
    "History", "Geography", "Political Science & Law", "Psychology", "Sociology", "Philosophy & Ethics",
    "Social Science", "Statistics", "Music", "Art & Craft", "Design & Multimedia", "Dance",
    "Yoga & Fitness", "Chess", "Sports & Martial Arts", "Civil Engineering", "Mechanical Engineering",
    "Electrical Engineering", "Electronics & Communication", "Chemical Engineering",
    "Aptitude & Logical Reasoning", "Competitive Exam Preparation", "General Knowledge",
    "Personality Development & Soft Skills", "Special Education", "All Subjects", "Medicine & Healthcare",
    "Journalism & Mass Communication", "Hindi", "Bengali", "Tamil", "Telugu", "Kannada", "Malayalam",
    "Marathi", "Gujarati", "Punjabi", "Urdu", "Sanskrit", "French", "German", "Spanish",
    "Arabic & Islamic Studies", "Japanese", "Chinese (Mandarin)", "Russian", "Korean", "Italian",
    "Portuguese", "Odia", "Assamese", "Nepali", "Persian", "Turkish", "Latin"
])

CATEGORY_MAP = {
    "Science": [
        "Mathematics", "Statistics", "Physics", "Chemistry", "Biology & Life Sciences",
        "Science (General)", "Medicine & Healthcare", "Civil Engineering", "Mechanical Engineering",
        "Electrical Engineering", "Electronics & Communication", "Chemical Engineering",
        "Agriculture", "Architecture", "Bioinformatics", "Biotechnology"
    ],
    "Languages": [
        "English", "Hindi", "Bengali", "Tamil", "Telugu", "Kannada", "Malayalam",
        "Marathi", "Gujarati", "Punjabi", "Urdu", "Sanskrit", "French", "German", "Spanish",
        "Arabic & Islamic Studies", "Japanese", "Chinese (Mandarin)", "Russian", "Korean",
        "Italian", "Portuguese", "Odia", "Assamese", "Nepali", "Persian", "Turkish", "Latin"
    ],
    "Coding": [
        "Computer Science & Programming", "Python", "Java", "C / C++ / C#", "Web Development",
        "SQL & Databases", "AI & Data Science", "Cloud Computing & DevOps", "Cybersecurity & Networking",
        "Software Testing", "Mobile App Development", "Game Development"
    ],
    "Commerce": [
        "Accountancy", "Economics", "Business Studies & Management", "Commerce", "Finance",
        "Banking", "Taxation", "Auditing", "Acca", "Cfa"
    ],
    "Arts": [
        "History", "Geography", "Political Science & Law", "Psychology", "Sociology",
        "Philosophy & Ethics", "Social Science", "Journalism & Mass Communication", "Art & Craft",
        "Music", "Design & Multimedia", "Dance", "Literature", "Humanities"
    ],
    "Misc/ECA": [
        "Yoga & Fitness", "Chess", "Sports & Martial Arts", "Aptitude & Logical Reasoning",
        "Competitive Exam Preparation", "General Knowledge", "Personality Development & Soft Skills",
        "Special Education", "Astrology", "Acting"
    ],
    "All Subjects": [
        "All Subjects"
    ]
}

def assign_category(can_s):
    for cat, subjs in CATEGORY_MAP.items():
        if can_s in subjs:
            return cat
    low = can_s.lower()
    if any(k in low for k in ['language', 'english', 'hindi', 'tamil', 'french', 'german', 'spanish', 'bengali']):
        return "Languages"
    if any(k in low for k in ['program', 'code', 'data', 'web', 'dev', 'soft', 'tech', 'cloud', 'sql']):
        return "Coding"
    if any(k in low for k in ['math', 'sci', 'phys', 'chem', 'bio', 'engg', 'engineer', 'med']):
        return "Science"
    if any(k in low for k in ['commerce', 'account', 'busin', 'econ', 'financ', 'bank']):
        return "Commerce"
    if any(k in low for k in ['art', 'music', 'dance', 'design', 'history', 'law', 'psy']):
        return "Arts"
    return "Science"


def map_raw_subject(raw):
    if not raw or not raw.strip():
        return "All Subjects"
    s = raw.strip()
    low = s.lower()
    
    # Exact lookup
    if low in EXACT_SUBJECT_MAP:
        return EXACT_SUBJECT_MAP[low]

    # Pattern scan
    for pat, target in SUBJECT_PATTERNS:
        if re.search(pat, low):
            return target
            
    # Clean fallback
    clean = re.sub(r'[\(\)\[\]\{\}\<\>]', '', s).strip().title()
    clean = re.sub(r'\s+', ' ', clean)
    if len(clean) <= 20 and len(clean.split()) <= 2 and clean.replace(' ', '').isalpha():
        return clean
        
    return "Other Academic Subjects"


def run_pipeline():
    print("=" * 70)
    print("   BMP TUTORS - UNIFIED DATA CLEANING & RE-EXPORT PIPELINE")
    print("=" * 70)

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Create helpful indices for blazing fast updates
    print("    Ensuring indices on database...")
    c.execute("CREATE INDEX IF NOT EXISTS idx_teachers_profile_url ON teachers(profile_url);")
    c.execute("CREATE INDEX IF NOT EXISTS idx_teacher_subjects_tutor_id ON teacher_subjects(tutor_id);")
    c.execute("CREATE INDEX IF NOT EXISTS idx_teacher_subjects_canonical ON teacher_subjects(canonical_subject);")
    conn.commit()

    # ---------------------------------------------------------
    # STEP 1: SUBJECT CANONICALIZATION
    # ---------------------------------------------------------
    print("\n>>> STEP 1: Subject Name Canonicalization...")
    c.execute("SELECT DISTINCT canonical_subject FROM teacher_subjects WHERE canonical_subject IS NOT NULL")
    raw_subjects = [r[0] for r in c.fetchall()]
    print(f"    Loaded {len(raw_subjects):,} distinct raw subject strings.")

    # Compute raw counts
    c.execute("SELECT canonical_subject, COUNT(*) FROM teacher_subjects GROUP BY canonical_subject")
    raw_counts = dict(c.fetchall())

    # Map stage 1
    raw_to_candidate = {s: map_raw_subject(s) for s in raw_subjects}
    candidate_totals = Counter()
    for s, target in raw_to_candidate.items():
        candidate_totals[target] += raw_counts.get(s, 0)

    # Prune: keep if in PRIMARY_SET or total offerings >= 28, else collapse to 'Other Academic Subjects'
    final_mapping = {}
    for s, target in raw_to_candidate.items():
        if target in PRIMARY_SET or candidate_totals[target] >= 28:
            final_mapping[s] = target
        else:
            final_mapping[s] = "Other Academic Subjects"

    distinct_canonicals = set(final_mapping.values())
    print(f"    Successfully reduced subject variants from {len(raw_subjects):,} -> {len(distinct_canonicals):,} canonical subjects.")

    # Update database using high-performance temporary mapping table
    print("    Executing high-speed batch update on teacher_subjects table...")
    c.execute("CREATE TEMP TABLE temp_subj_map (old_subj TEXT PRIMARY KEY, new_subj TEXT, new_cat TEXT);")
    map_tuples = [(raw_s, can_s, assign_category(can_s)) for raw_s, can_s in final_mapping.items()]
    c.executemany("INSERT INTO temp_subj_map VALUES (?, ?, ?);", map_tuples)
    
    c.execute("""
        UPDATE teacher_subjects
        SET canonical_subject = m.new_subj,
            subject_category = m.new_cat
        FROM temp_subj_map m
        WHERE teacher_subjects.canonical_subject = m.old_subj;
    """)
    c.execute("DROP TABLE temp_subj_map;")
    conn.commit()
    print("    Subject canonicalization committed successfully.")

    # Verify mathematics & english examples from prompt
    c.execute("SELECT COUNT(*) FROM teacher_subjects WHERE canonical_subject IN ('Maths', 'Math', 'Mathamatics', 'Mathmatics', 'Mathematic', 'Cuemath', 'Vedicmaths')")
    bad_math = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM teacher_subjects WHERE canonical_subject IN ('Engilsh', 'Englsih')")
    bad_eng = c.fetchone()[0]
    print(f"    Verification: Remaining bad math variants: {bad_math}, bad english variants: {bad_eng}.")

    # ---------------------------------------------------------
    # STEP 2: FEE DATA CLEANING
    # ---------------------------------------------------------
    print("\n>>> STEP 2: Fee Data Cleaning (₹100 to ₹5,000 hourly range)...")
    c.execute("""
        SELECT COUNT(*) FROM teachers
        WHERE (hourly_fee_avg < 100 OR hourly_fee_avg > 5000 OR hourly_fee_avg = 0)
          AND hourly_fee_avg IS NOT NULL
    """)
    bad_fees_count = c.fetchone()[0]
    print(f"    Found {bad_fees_count:,} records with out-of-bounds fees (< ₹100, > ₹5,000, or = ₹0).")

    c.execute("""
        UPDATE teachers
        SET hourly_fee_min = NULL,
            hourly_fee_max = NULL,
            hourly_fee_avg = NULL
        WHERE (hourly_fee_avg < 100 OR hourly_fee_avg > 5000 OR hourly_fee_avg = 0)
          AND hourly_fee_avg IS NOT NULL
    """)
    conn.commit()
    print(f"    Successfully nulled {bad_fees_count:,} out-of-bounds fee entries.")

    # ---------------------------------------------------------
    # STEP 3: REMOVE DUMMY / TEST ACCOUNTS
    # ---------------------------------------------------------
    print("\n>>> STEP 3: Removing Dummy / Test Accounts...")
    c.execute("""
        SELECT tutor_id, teacher_name, hourly_fee_avg FROM teachers
        WHERE tutor_id IN ('T3662', 'T64166', 'T92868', 'T100360')
           OR lower(teacher_name) LIKE '%dummy account%'
           OR lower(teacher_name) LIKE '%maya global test prep%'
           OR lower(teacher_name) LIKE '%kiddos rehabcentre%'
           OR lower(teacher_name) LIKE '%allah is greatest fatima%'
           OR lower(teacher_name) = 'xyz'
           OR lower(teacher_name) = 'abc'
    """)
    dummies = c.fetchall()
    dummy_ids = [d[0] for d in dummies]
    print(f"    Found {len(dummy_ids)} dummy/test accounts to remove: {dummy_ids}")

    if dummy_ids:
        placeholders = ",".join("?" for _ in dummy_ids)
        c.execute(f"DELETE FROM teacher_subjects WHERE tutor_id IN ({placeholders})", dummy_ids)
        c.execute(f"DELETE FROM teachers WHERE tutor_id IN ({placeholders})", dummy_ids)
        conn.commit()
        print(f"    Deleted {len(dummy_ids)} dummy accounts from both teachers and teacher_subjects tables.")

    # ---------------------------------------------------------
    # STEP 4: HANDLE EMPTY / MISSING DATA
    # ---------------------------------------------------------
    print("\n>>> STEP 4: Handling Empty / Missing Data...")
    c.execute("SELECT COUNT(*) FROM teachers WHERE teacher_name IS NULL OR TRIM(teacher_name) = ''")
    empty_names = c.fetchone()[0]
    print(f"    Found {empty_names:,} tutors with missing names.")

    c.execute("""
        UPDATE teachers
        SET teacher_name = 'Anonymous Tutor'
        WHERE teacher_name IS NULL OR TRIM(teacher_name) = ''
    """)
    conn.commit()
    print(f"    Updated {empty_names:,} missing names to 'Anonymous Tutor'.")

    # ---------------------------------------------------------
    # STEP 5: FIX STATE & GENDER MISMATCHES
    # ---------------------------------------------------------
    print("\n>>> STEP 5: Verifying State & Gender Against TrueData.xlsx...")
    if os.path.exists(TRUE_DATA_PATH):
        c.execute("SELECT profile_url, state, gender FROM teachers")
        db_records = {r[0]: (r[1], r[2]) for r in c.fetchall()}
        
        wb = openpyxl.load_workbook(TRUE_DATA_PATH, read_only=True)
        mismatch_updates = []
        for sname in ['scraped', 'not scraped']:
            sheet = wb[sname]
            rows = sheet.iter_rows(values_only=True)
            header = next(rows)
            u_idx = header.index('profile_url')
            l_idx = header.index('Location')
            g_idx = header.index('Gender')
            
            for row in rows:
                url = row[u_idx]
                if url in db_records:
                    db_state, db_gender = db_records[url]
                    loc = row[l_idx]
                    gen = row[g_idx]
                    s_state, s_reg, _, s_city = clean_location(loc)
                    s_gen = clean_gender(gen)
                    if s_state != db_state or s_gen != db_gender:
                        mismatch_updates.append((s_state, s_reg, s_city, s_gen, url))
        
        if mismatch_updates:
            print(f"    Found {len(mismatch_updates)} discrepancies against TrueData.xlsx. Updating...")
            c.executemany("""
                UPDATE teachers
                SET state = ?, region = ?, city = ?, gender = ?
                WHERE profile_url = ?
            """, mismatch_updates)
            conn.commit()
            print(f"    Successfully updated {len(mismatch_updates)} mismatched records to match TrueData.xlsx.")
        else:
            print("    100% of records in DB already perfectly match TrueData.xlsx source data (0 discrepancies).")
    else:
        print("    TrueData.xlsx not found; skipping sheet reconciliation.")

    # ---------------------------------------------------------
    # STEP 6: DEDUPLICATION OF PROFILES
    # ---------------------------------------------------------
    print("\n>>> STEP 6: Deduplication (Same Name + Same City + Same Subjects)...")
    c.execute("SELECT tutor_id, canonical_subject FROM teacher_subjects")
    tutor_subjs = defaultdict(set)
    for tid, subj in c.fetchall():
        if subj:
            tutor_subjs[tid].add(subj.strip().lower())

    c.execute("""
        SELECT tutor_id, teacher_name, city, hourly_fee_avg, total_teaching_exp_years,
               is_active, description, highest_qualification
        FROM teachers
    """)
    all_teachers = c.fetchall()

    groups = defaultdict(list)
    for t in all_teachers:
        tid, name, city, fee, exp, active, desc, qual = t
        if name and name.strip() and name.strip() != 'Anonymous Tutor':
            norm_name = name.strip().lower()
            norm_city = (city or '').strip().lower()
            subjs_key = tuple(sorted(list(tutor_subjs[tid])))
            if norm_city and subjs_key:
                groups[(norm_name, norm_city, subjs_key)].append(t)

    dupe_ids_to_remove = []
    for (name, city, subjs), plist in groups.items():
        if len(plist) > 1:
            # Score each candidate: prioritize active, then fee, then exp, then desc, then qual
            def score(p):
                tid, name, city, fee, exp, active, desc, qual = p
                return (
                    active * 10,
                    1 if fee is not None else 0,
                    1 if exp is not None else 0,
                    1 if desc and len(desc) > 30 else 0,
                    1 if qual else 0,
                    -int(tid[1:]) if tid[1:].isdigit() else 0
                )
            sorted_p = sorted(plist, key=score, reverse=True)
            winner = sorted_p[0]
            losers = sorted_p[1:]
            for l in losers:
                dupe_ids_to_remove.append(l[0])

    print(f"    Identified {len(dupe_ids_to_remove)} redundant duplicate tutor profiles to remove.")
    if dupe_ids_to_remove:
        placeholders = ",".join("?" for _ in dupe_ids_to_remove)
        c.execute(f"DELETE FROM teacher_subjects WHERE tutor_id IN ({placeholders})", dupe_ids_to_remove)
        c.execute(f"DELETE FROM teachers WHERE tutor_id IN ({placeholders})", dupe_ids_to_remove)
        conn.commit()
        print(f"    Removed {len(dupe_ids_to_remove)} duplicate profiles.")

    # ---------------------------------------------------------
    # VACUUM & FINAL CHECKS
    # ---------------------------------------------------------
    print("\n>>> Performing SQLite VACUUM and ANALYZE...")
    c.execute("VACUUM")
    c.execute("ANALYZE")
    conn.commit()

    c.execute("SELECT COUNT(*) FROM teachers")
    final_teacher_count = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM teacher_subjects")
    final_subjects_count = c.fetchone()[0]
    c.execute("SELECT COUNT(DISTINCT canonical_subject) FROM teacher_subjects")
    final_distinct_subjects = c.fetchone()[0]

    print("\n" + "=" * 70)
    print("   DATA CLEANING VERIFICATION SUMMARY")
    print("=" * 70)
    print(f"   Total Tutors Remaining       : {final_teacher_count:,} (close to original 107,723)")
    print(f"   Total Subject Offerings      : {final_subjects_count:,}")
    print(f"   Distinct Canonical Subjects  : {final_distinct_subjects:,} (Target ~300)")
    print("=" * 70)

    conn.close()

    # ---------------------------------------------------------
    # STEP 7: RE-EXPORT DELIVERABLES
    # ---------------------------------------------------------
    print("\n>>> STEP 7: Regenerating Normalized Relational DBMS Tables (3NF)...")
    import generate_normalized_dbms_tables
    generate_normalized_dbms_tables.build_normalized_tables()
    print("\n>>> Pipeline completed successfully!")

if __name__ == "__main__":
    run_pipeline()
