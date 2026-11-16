import sqlite3
import csv
import re
from collections import Counter
import clean_database_pipeline as cdp

DB_PATH = "cleaned_tutors_data.db"
CSV_PATH = "cleaned_teacher_subjects.csv"

print(">>> Starting Full Subjects Restoration Pipeline...")

# Step 1: Connect and fetch valid tutors (107,673 tutors)
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()
cur.execute("SELECT tutor_id FROM teachers")
valid_tutors = set(r[0] for r in cur.fetchall())
print(f"Loaded {len(valid_tutors):,} valid tutors from teachers table.")

# Step 2: Define strict canonicalization patterns (Computer Science and Social Science take precedence over General Science)
patterns = [
    # 1. Mathematics & Statistics
    (r'\b(maths?|mathematics?|algebra|calculus|geometry|trigonometry|arithmetic|topology|cuemath|vedic|abacus|linear algebra|differential equation|discrete math)\b', 'Mathematics'),
    (r'\b(statist|statistics|probability|biostat|actuarial|econometric)\b', 'Statistics'),

    # 2. Computer Science, Programming, Coding & IT (HIGH PRIORITY)
    (r'\b(python|django|flask|fastapi|pandas|numpy)\b', 'Python'),
    (r'\b(java\b|spring boot|spring framework|hibernate)\b', 'Java'),
    (r'\b(web develop|html|css|javascript|react|node|frontend|backend|full stack|angular|vue|bootstrap|wordpress|php|tailwind)\b', 'Web Development'),
    (r'\b(sql|database|dbms|mysql|oracle|mongodb|postgresql|nosql|plsql|sqlite)\b', 'SQL & Databases'),
    (r'\b(data science|machine learning|artificial intelligence|deep learning|nlp|computer vision|power bi|tableau|data analytics|data analysis|big data|neural network|excel)\b', 'AI & Data Science'),
    (r'\b(cloud|aws|azure|gcp|devops|docker|kubernetes|jenkins|linux|unix|git|terraform|ansible)\b', 'Cloud Computing & DevOps'),
    (r'\b(computer network|network protocol)\b', 'Computer Science & Programming'),
    (r'\b(cyber security|cybersecurity|ethical hacking|network security|information security|networking|ccna|cryptography)\b', 'Cybersecurity & Networking'),
    (r'\b(software testing|selenium|automation testing|manual testing|qa testing|test automation)\b', 'Software Testing'),
    (r'\b(android|ios|flutter|swift|react native|mobile app|app development|kotlin)\b', 'Mobile App Development'),
    (r'\b(computer science|programming|coding|dsa|data structure|algorithms|operating system|compiler|automata|oops|software engineering|computer architecture|cse|computer|c programming|c\+\+|c / c\+\+|c and c\+\+|c#|\.net)\b', 'Computer Science & Programming'),

    # 3. Humanities & Social Sciences (HIGH PRIORITY)
    (r'\b(social science|social studies|sst)\b', 'Social Science'),
    (r'\b(political science|polity|civics|law|legal|jurisprudence|constitution|clat|public administration)\b', 'Political Science & Law'),
    (r'\b(history|ancient history|medieval history|modern history|world history|archaeology)\b', 'History'),
    (r'\b(geography|geology|earth science|gis)\b', 'Geography'),
    (r'\b(psychology|counseling|mental health)\b', 'Psychology'),
    (r'\b(sociology|social work|anthropology)\b', 'Sociology'),
    (r'\b(philosophy|ethics|logic)\b', 'Philosophy & Ethics'),
    (r'\b(journalism|mass communication|media studies)\b', 'Journalism & Mass Communication'),

    # 4. Natural Sciences
    (r'\b(physics|astrophysics|geophysics|astronomy|thermodynamics|quantum|optics|electromagnet)\b', 'Physics'),
    (r'\b(chemistry|biochem|organic chem|inorganic chem|physical chem|chemical analysis)\b', 'Chemistry'),
    (r'\b(biology|botany|zoology|microbiol|genetics|biotech|anatomy|physiology|pathology|immunology|ecology|life science|embryology|virology|neet bio|bio neet)\b', 'Biology & Life Sciences'),
    (r'\b(medicine|nursing|pharmacy|pharmacology|dental|mbbs|medical|clinical|surgery|ayurveda|homeopathy|physiotherapy)\b', 'Medicine & Healthcare'),
    (r'\b(science|general science|natural science|evs|environmental science|environmental studies)\b', 'Science (General)'),

    # 5. Commerce & Management
    (r'\b(account|accounts|accounting|accountancy|tally|bookkeeping|costing|taxation|gst|auditing|ca foundation|cma)\b', 'Accountancy'),
    (r'\b(economics|microeconomics|macroeconomics|econometrics)\b', 'Economics'),
    (r'\b(business|management|marketing|human resource|finance|financial|bba|mba|commerce|banking|supply chain|entrepreneurship|business studies)\b', 'Business Studies & Management'),

    # 6. Languages
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

    # 7. Creative & Performing Arts
    (r'\b(music|singing|vocal|guitar|piano|keyboard|violin|flute|drums|carnatic|hindustani|tabla|sitar|harmonium)\b', 'Music'),
    (r'\b(art|drawing|painting|sketching|calligraphy|craft|fine arts|water color|oil painting)\b', 'Art & Craft'),
    (r'\b(graphic design|photoshop|illustrator|ui/ux|ui ux|animation|video editing|vfx|3d modeling)\b', 'Design & Multimedia'),
    (r'\b(dance|bharatanatyam|kathak|zumba|hip hop|salsa|ballet|folk dance)\b', 'Dance'),
    (r'\b(yoga|fitness|gym|workout|aerobics|meditation|nutrition|dietetics|physical education)\b', 'Yoga & Fitness'),
    (r'\b(chess)\b', 'Chess'),
    (r'\b(karate|taekwondo|martial arts|swimming|badminton|cricket|tennis|sports)\b', 'Sports & Martial Arts'),

    # 8. All Subjects
    (r'\b(all subject|all subjects|all school subjects|all primary|all standards|all classes|all)\b', 'All Subjects')
]

for pat, can in cdp.SUBJECT_PATTERNS:
    if can not in [p[1] for p in patterns]:
        patterns.append((pat, can))

compiled = [(re.compile(p, re.I), can) for p, can in patterns]

def map_raw_str(raw_s, old_can):
    s = (raw_s or '').strip()
    s_low = s.lower()
    if s_low in cdp.EXACT_SUBJECT_MAP:
        return cdp.EXACT_SUBJECT_MAP[s_low]
    for rx, can in compiled:
        if rx.search(s_low):
            return can
    if old_can and old_can.strip():
        for rx, can in compiled:
            if rx.search(old_can.lower()):
                return can
        return old_can.strip().title()
    return 'Other Academic Subjects'

# Step 3: Read CSV and collect rows
print("Reading cleaned_teacher_subjects.csv...")
raw_records = []
cand_counts = Counter()
with open(CSV_PATH, 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    next(reader)
    for row in reader:
        t_id = row[0]
        if t_id in valid_tutors:
            s_idx = int(row[1]) if row[1] else 1
            raw_s = row[2]
            old_can = row[3]
            lvl = row[5] or None
            exam = row[6] or None
            c_cand = map_raw_str(raw_s, old_can)
            cand_counts[c_cand] += 1
            raw_records.append((t_id, s_idx, raw_s, c_cand, lvl, exam))

print(f"Collected {len(raw_records):,} offerings for valid educators.")

# Step 4: Final canonicalization and category assignment
final_rows = []
for t_id, s_idx, raw_s, c_cand, lvl, exam in raw_records:
    if c_cand in cdp.PRIMARY_SET or cand_counts[c_cand] >= 30:
        c_final = c_cand
    else:
        c_final = "Other Academic Subjects"
    cat = cdp.assign_category(c_final)
    final_rows.append((t_id, s_idx, c_final, cat, lvl, exam, raw_s))

# Step 5: Replace teacher_subjects table
print("Recreating teacher_subjects table in database...")
cur.execute("DROP TABLE IF EXISTS teacher_subjects;")
cur.execute("""
CREATE TABLE teacher_subjects (
    bridge_id INTEGER PRIMARY KEY AUTOINCREMENT,
    tutor_id TEXT,
    subject_index INTEGER,
    canonical_subject TEXT,
    subject_category TEXT,
    level_category TEXT,
    competitive_exam TEXT,
    raw_subject_string TEXT
);
""")

cur.executemany("""
INSERT INTO teacher_subjects (tutor_id, subject_index, canonical_subject, subject_category, level_category, competitive_exam, raw_subject_string)
VALUES (?, ?, ?, ?, ?, ?, ?);
""", final_rows)

print("Creating database indexes...")
cur.execute("CREATE INDEX idx_teacher_subjects_tutor_id ON teacher_subjects(tutor_id);")
cur.execute("CREATE INDEX idx_teacher_subjects_canonical ON teacher_subjects(canonical_subject);")
cur.execute("CREATE INDEX idx_teacher_subjects_category ON teacher_subjects(subject_category);")
cur.execute("CREATE INDEX idx_teacher_subjects_level ON teacher_subjects(level_category);")
conn.commit()

# Step 6: Verify
cur.execute("SELECT count(*) FROM teacher_subjects;")
total_count = cur.fetchone()[0]
cur.execute("SELECT count(DISTINCT canonical_subject) FROM teacher_subjects;")
distinct_can = cur.fetchone()[0]
cur.execute("SELECT count(DISTINCT tutor_id) FROM teacher_subjects;")
tutors_with_sub = cur.fetchone()[0]

print(f"\n==========================================")
print(f"RESTORATION COMPLETE:")
print(f"Total Offerings: {total_count:,}")
print(f"Distinct Canonical Subjects: {distinct_can:,}")
print(f"Educators with Subjects: {tutors_with_sub:,}")
print(f"Educators without Subjects: {len(valid_tutors) - tutors_with_sub:,}")
print(f"Average Offerings per Stating Educator: {round(total_count / tutors_with_sub, 2)}")
print(f"==========================================\n")

print("Checking T1 (Shubhankshi Saraf):")
cur.execute("SELECT bridge_id, canonical_subject, subject_category, level_category, raw_subject_string FROM teacher_subjects WHERE tutor_id='T1' ORDER BY subject_index ASC;")
for r in cur.fetchall():
    print(" ", r)

conn.close()
