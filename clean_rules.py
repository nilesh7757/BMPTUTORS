import re

# -------------------------------------------------------------
# 1. LOCATION & REGION MAPPINGS
# -------------------------------------------------------------

STATE_REGION_MAP = {
    # North
    "delhi": ("Delhi", "North"),
    "new delhi": ("Delhi", "North"),
    "ncr": ("Delhi", "North"),
    "chandigarh": ("Chandigarh", "North"),
    "jammu": ("Jammu and Kashmir", "North"),
    "srinagar": ("Jammu and Kashmir", "North"),
    "jammu and kashmir": ("Jammu and Kashmir", "North"),
    "jammu & kashmir": ("Jammu and Kashmir", "North"),
    "uttarakhand": ("Uttarakhand", "North"),
    "uttaranchal": ("Uttarakhand", "North"),
    "dehradun": ("Uttarakhand", "North"),
    "uttar pradesh": ("Uttar Pradesh", "North"),
    "u.p.": ("Uttar Pradesh", "North"),
    "noida": ("Uttar Pradesh", "North"),
    "greater noida": ("Uttar Pradesh", "North"),
    "ghaziabad": ("Uttar Pradesh", "North"),
    "lucknow": ("Uttar Pradesh", "North"),
    "kanpur": ("Uttar Pradesh", "North"),
    "varanasi": ("Uttar Pradesh", "North"),
    "agra": ("Uttar Pradesh", "North"),
    "prayagraj": ("Uttar Pradesh", "North"),
    "allahabad": ("Uttar Pradesh", "North"),
    "meerut": ("Uttar Pradesh", "North"),
    "bareilly": ("Uttar Pradesh", "North"),
    "aligarh": ("Uttar Pradesh", "North"),
    "moradabad": ("Uttar Pradesh", "North"),
    "haryana": ("Haryana", "North"),
    "gurgaon": ("Haryana", "North"),
    "gurugram": ("Haryana", "North"),
    "faridabad": ("Haryana", "North"),
    "panchkula": ("Haryana", "North"),
    "ambala": ("Haryana", "North"),
    "karnal": ("Haryana", "North"),
    "himachal pradesh": ("Himachal Pradesh", "North"),
    "shimla": ("Himachal Pradesh", "North"),
    "punjab": ("Punjab", "North"),
    "ludhiana": ("Punjab", "North"),
    "amritsar": ("Punjab", "North"),
    "jalandhar": ("Punjab", "North"),
    "patiala": ("Punjab", "North"),
    "ladakh": ("Ladakh", "North"),
    "leh": ("Ladakh", "North"),
    "rajasthan": ("Rajasthan", "North"),
    "jaipur": ("Rajasthan", "North"),
    "jodhpur": ("Rajasthan", "North"),
    "udaipur": ("Rajasthan", "North"),
    "kota": ("Rajasthan", "North"),
    "ajmer": ("Rajasthan", "North"),
    "bikaner": ("Rajasthan", "North"),
    
    # South
    "andhra pradesh": ("Andhra Pradesh", "South"),
    "visakhapatnam": ("Andhra Pradesh", "South"),
    "vizag": ("Andhra Pradesh", "South"),
    "vijayawada": ("Andhra Pradesh", "South"),
    "guntur": ("Andhra Pradesh", "South"),
    "tirupati": ("Andhra Pradesh", "South"),
    "kakinada": ("Andhra Pradesh", "South"),
    "telangana": ("Telangana", "South"),
    "hyderabad": ("Telangana", "South"),
    "secunderabad": ("Telangana", "South"),
    "warangal": ("Telangana", "South"),
    "karnataka": ("Karnataka", "South"),
    "bengaluru": ("Karnataka", "South"),
    "bangalore": ("Karnataka", "South"),
    "mysuru": ("Karnataka", "South"),
    "mysore": ("Karnataka", "South"),
    "mangaluru": ("Karnataka", "South"),
    "mangalore": ("Karnataka", "South"),
    "hubli": ("Karnataka", "South"),
    "belgaum": ("Karnataka", "South"),
    "kerala": ("Kerala", "South"),
    "kochi": ("Kerala", "South"),
    "cochin": ("Kerala", "South"),
    "ernakulam": ("Kerala", "South"),
    "thiruvananthapuram": ("Kerala", "South"),
    "trivandrum": ("Kerala", "South"),
    "calicut": ("Kerala", "South"),
    "kozhikode": ("Kerala", "South"),
    "thrissur": ("Kerala", "South"),
    "kollam": ("Kerala", "South"),
    "tamil nadu": ("Tamil Nadu", "South"),
    "chennai": ("Tamil Nadu", "South"),
    "madras": ("Tamil Nadu", "South"),
    "coimbatore": ("Tamil Nadu", "South"),
    "madurai": ("Tamil Nadu", "South"),
    "tiruchirappalli": ("Tamil Nadu", "South"),
    "trichy": ("Tamil Nadu", "South"),
    "salem": ("Tamil Nadu", "South"),
    "tirunelveli": ("Tamil Nadu", "South"),
    "vellore": ("Tamil Nadu", "South"),
    "puducherry": ("Puducherry", "South"),
    "pondicherry": ("Puducherry", "South"),
    "andaman and nicobar": ("Andaman and Nicobar Islands", "South"),
    "andaman": ("Andaman and Nicobar Islands", "South"),
    "port blair": ("Andaman and Nicobar Islands", "South"),
    
    # East
    "west bengal": ("West Bengal", "East"),
    "kolkata": ("West Bengal", "East"),
    "calcutta": ("West Bengal", "East"),
    "howrah": ("West Bengal", "East"),
    "durgapur": ("West Bengal", "East"),
    "siliguri": ("West Bengal", "East"),
    "asansol": ("West Bengal", "East"),
    "odisha": ("Odisha", "East"),
    "orissa": ("Odisha", "East"),
    "bhubaneswar": ("Odisha", "East"),
    "cuttack": ("Odisha", "East"),
    "rourkela": ("Odisha", "East"),
    "bihar": ("Bihar", "East"),
    "patna": ("Bihar", "East"),
    "gaya": ("Bihar", "East"),
    "muzaffarpur": ("Bihar", "East"),
    "bhagalpur": ("Bihar", "East"),
    "jharkhand": ("Jharkhand", "East"),
    "ranchi": ("Jharkhand", "East"),
    "jamshedpur": ("Jharkhand", "East"),
    "dhanbad": ("Jharkhand", "East"),
    "bokaro": ("Jharkhand", "East"),
    "assam": ("Assam", "East"),
    "guwahati": ("Assam", "East"),
    "silchar": ("Assam", "East"),
    "dibrugarh": ("Assam", "East"),
    "manipur": ("Manipur", "East"),
    "imphal": ("Manipur", "East"),
    "meghalaya": ("Meghalaya", "East"),
    "shillong": ("Meghalaya", "East"),
    "mizoram": ("Mizoram", "East"),
    "aizawl": ("Mizoram", "East"),
    "nagaland": ("Nagaland", "East"),
    "kohima": ("Nagaland", "East"),
    "dimapur": ("Nagaland", "East"),
    "tripura": ("Tripura", "East"),
    "agartala": ("Tripura", "East"),
    "sikkim": ("Sikkim", "East"),
    "gangtok": ("Sikkim", "East"),
    "arunachal pradesh": ("Arunachal Pradesh", "East"),
    "itanagar": ("Arunachal Pradesh", "East"),

    # West
    "goa": ("Goa", "West"),
    "panaji": ("Goa", "West"),
    "gujarat": ("Gujarat", "West"),
    "ahmedabad": ("Gujarat", "West"),
    "surat": ("Gujarat", "West"),
    "vadodara": ("Gujarat", "West"),
    "baroda": ("Gujarat", "West"),
    "rajkot": ("Gujarat", "West"),
    "gandhinagar": ("Gujarat", "West"),
    "bhavnagar": ("Gujarat", "West"),
    "maharashtra": ("Maharashtra", "West"),
    "mumbai": ("Maharashtra", "West"),
    "bombay": ("Maharashtra", "West"),
    "pune": ("Maharashtra", "West"),
    "nagpur": ("Maharashtra", "West"),
    "thane": ("Maharashtra", "West"),
    "navi mumbai": ("Maharashtra", "West"),
    "nashik": ("Maharashtra", "West"),
    "aurangabad": ("Maharashtra", "West"),
    "chhatrapati sambhaji nagar": ("Maharashtra", "West"),
    "solapur": ("Maharashtra", "West"),
    "kolhapur": ("Maharashtra", "West"),
    "daman and diu": ("Daman and Diu", "West"),
    "daman": ("Daman and Diu", "West"),
    "diu": ("Daman and Diu", "West"),
    "dadra and nagar haveli": ("Dadra and Nagar Haveli", "West"),
    "silvassa": ("Dadra and Nagar Haveli", "West"),
    "lakshadweep": ("Lakshadweep", "West"),

    # Central
    "madhya pradesh": ("Madhya Pradesh", "Central"),
    "m.p.": ("Madhya Pradesh", "Central"),
    "mp": ("Madhya Pradesh", "Central"),
    "bhopal": ("Madhya Pradesh", "Central"),
    "indore": ("Madhya Pradesh", "Central"),
    "gwalior": ("Madhya Pradesh", "Central"),
    "jabalpur": ("Madhya Pradesh", "Central"),
    "ujjain": ("Madhya Pradesh", "Central"),
    "chhattisgarh": ("Chhattisgarh", "Central"),
    "raipur": ("Chhattisgarh", "Central"),
    "bhilai": ("Chhattisgarh", "Central"),
    "bilaspur": ("Chhattisgarh", "Central"),
}

ALL_STATES_SET = {
    "Delhi", "Chandigarh", "Jammu and Kashmir", "Uttarakhand", "Uttar Pradesh", "Haryana", "Himachal Pradesh", "Punjab", "Ladakh", "Rajasthan",
    "Andhra Pradesh", "Telangana", "Karnataka", "Kerala", "Tamil Nadu", "Puducherry", "Andaman and Nicobar Islands",
    "West Bengal", "Odisha", "Bihar", "Jharkhand", "Assam", "Manipur", "Meghalaya", "Mizoram", "Nagaland", "Tripura", "Sikkim", "Arunachal Pradesh",
    "Goa", "Gujarat", "Maharashtra", "Daman and Diu", "Dadra and Nagar Haveli", "Lakshadweep",
    "Madhya Pradesh", "Chhattisgarh"
}

def clean_location(loc_str):
    if not loc_str or not str(loc_str).strip():
        return "Not Specified", "Not Specified", "Not Specified", "Not Specified"
    loc_clean = str(loc_str).strip()
    loc_lower = loc_clean.lower()
    
    for k in sorted(STATE_REGION_MAP.keys(), key=lambda x: len(x), reverse=True):
        pattern = r"\b" + re.escape(k) + r"\b"
        if re.search(pattern, loc_lower):
            state, region = STATE_REGION_MAP[k]
            if k.title() not in ALL_STATES_SET and k not in ["u.p.", "m.p.", "mp", "ncr"]:
                city_val = k.title()
            else:
                first_part = loc_clean.split(",")[0].strip()
                city_val = first_part.title() if first_part.lower() != k else state
            return state, region, loc_clean, city_val
            
    return "Other / International", "Other", loc_clean, loc_clean.split(",")[0].strip().title()

# -------------------------------------------------------------
# 2. EXPERIENCE ROLE & SENIORITY TIER
# -------------------------------------------------------------

EXPERIENCE_RULES = [
    (r"\b(prof|professor|lecturer|hod|head of dept|head of department|asst professor|assistant professor|associate professor)\b", "Professor/Lecturer"),
    (r"\b(school teacher|teacher|pgt|tgt|prt|faculty teacher)\b", "Teacher"),
    (r"\b(tutor|home tutor|private tutor|tuition)\b", "Tutor"),
    (r"\b(educator|faculty|coaching)\b", "Educator/Faculty"),
    (r"\b(trainer|instructor|coach)\b", "Instructor/Trainer"),
]

def clean_experience_role(exp_str):
    if not exp_str or not str(exp_str).strip():
        return "Not Specified"
    s = str(exp_str).strip().lower()
    for pattern, cat in EXPERIENCE_RULES:
        if re.search(pattern, s):
            return cat
    return "Other / Industry Professional"

def infer_tutor_seniority(exp_years, qual, role, hookline, desc):
    exp = exp_years if exp_years is not None else 0.0
    text = f"{hookline or ''} {desc or ''}".lower()
    
    if exp >= 10.0 or qual == "Advanced (PhD/Doctorate)" or "hod" in text or "head of dept" in text or "professor" in text:
        return "Master / Veteran Educator (10+ Yrs / PhD)"
    elif exp >= 5.0 or (qual == "Postgraduate (PG)" and exp >= 3.0) or "jee" in text or "neet" in text or "test prep" in text:
        return "Senior Specialist (5–10 Yrs / Test-Prep)"
    elif exp >= 2.0 or qual in ["Postgraduate (PG)", "Undergraduate (UG)"]:
        return "Experienced Practitioner (2–5 Yrs)"
    else:
        return "Early-Career / Scholar Tutor (<2 Yrs)"

# -------------------------------------------------------------
# 3. QUALIFICATIONS
# -------------------------------------------------------------

QUALIFICATION_RULES = [
    # a) Advanced
    (r"\b(ph\.?\s*d|d\.?\s*phil|doctor of philosophy|doctoral|post\s*doc|doctorate)\b", "Advanced (PhD/Doctorate)"),
    
    # b) Postgraduate (PG)
    (r"\b(m\.?\s*phil|mba|master|masters|master.*s|m\.?\s*a\b|m\.?\s*sc\b|m\.?\s*s\b|m\.?\s*tech\b|m\.?\s*f\.?\s*a\b|m\.?\s*ed\b|m\.?\s*c\.?\s*a\b|m\.?\s*com\b|ll\.?\s*m\b|pgdm|pgd\b|m\.?\s*pharm\b|m\.?\s*d\b|m\.?\s*d\.?\s*s\b|m\.?\s*e\b|pgdhhm|bsms|post\s*grad|postgraduate|chartered accountant|ca\b|m\s*tech|m\s*sc|m\s*com|m\s*ed)\b", "Postgraduate (PG)"),
    
    # c) Undergraduate (UG)
    (r"\b(bachelor|bachelors|bachelor.*s|b\.?\s*a\b|b\.?\s*sc\b|b\.?\s*com\b|b\.?\s*tech\b|b\.?\s*e\b|bsw|b\.?\s*arch\b|b\.?\s*pharm\b|ll\.?\s*b\b|bca|bfa|bpt|bhms|bams|mbbs|bds|bba|ballb|b\.?\s*ed\b|ug\b|undergrad|undergraduate|engineering|graduate|graduation|b\s*tech|b\s*sc|b\s*com|b\s*ed)\b", "Undergraduate (UG)"),
    
    # d) Less than UG
    (r"\b(diploma|certificate|higher secondary|hsc|12th|secondary|10th|school|sslc|intermediate|\+2)\b", "Less than UG"),
]

def clean_qualification(edu_str):
    if not edu_str or not str(edu_str).strip():
        return "Not Specified", 0
    s = str(edu_str).strip().lower()
    
    # Education degree flag (B.Ed / M.Ed)
    has_edu_deg = 1 if re.search(r"\b(b\.?\s*ed|m\.?\s*ed|bachelor of education|master of education|ba\(ed\)|bsc\(ed\))\b", s) else 0
    
    for pattern, cat in QUALIFICATION_RULES:
        if re.search(pattern, s):
            return cat, has_edu_deg
            
    if len(s) > 2:
        return "Undergraduate (UG)", has_edu_deg
    return "Not Specified", has_edu_deg

# -------------------------------------------------------------
# 4. COMPETITIVE EXAMS
# -------------------------------------------------------------

COMPETITIVE_EXAMS = ["JEE", "NEET", "GRE", "SAT", "IELTS", "TOEFL", "GMAT", "UPSC", "CAT", "GATE"]
COMP_EXAM_PATTERNS = [(re.compile(r"\b" + re.escape(exam) + r"\b", re.IGNORECASE), exam) for exam in COMPETITIVE_EXAMS]

# -------------------------------------------------------------
# 5. CONTEXTUAL TARGET STUDENT LEVEL (Full Profile NLP Resolution)
# -------------------------------------------------------------

def infer_target_level(raw_lvl, subj_str, hookline, desc, qual, role, exam):
    raw_s = str(raw_lvl or "").strip().lower()
    subj_s = str(subj_str or "").strip().lower()
    hook_s = str(hookline or "").strip().lower()
    desc_s = str(desc or "").strip().lower()
    full_text = f"{subj_s} {raw_s} {hook_s} {desc_s} {exam or ''}"
    
    # 1. Competitive Exam & Entrance Prep
    if exam in ["JEE", "NEET", "SAT"] or re.search(r"\b(jee|neet|iit-?jee|sat|olympiad|ntse)\b", full_text):
        return "Senior Secondary (Class 11–12 / JEE / NEET)"
    if exam in ["GATE", "CAT", "GMAT", "GRE", "UPSC"] or re.search(r"\b(gate|cat|gmat|gre|upsc|ias|civil services)\b", full_text):
        return "Higher Education (College / University)"
        
    # 2. International & Senior Secondary School (11-12, A-Levels, IBDP)
    if re.search(r"\b(as & a level|a-level|a level|ibdp|ib diploma|class 12|class 11|grade 12|grade 11|12th|11th|hsc|\+2|senior secondary|intermediate|class xi|class xii|grade xi|grade xii|cbse 12|icse 12)\b", full_text):
        return "Senior Secondary (Class 11–12 / Boards)"
        
    # 3. Secondary School (9-10, IGCSE, ICSE, 10th Boards)
    if re.search(r"\b(igcse|class 10|class 9|grade 10|grade 9|9th|10th|sslc|matriculation|secondary|class ix|class x|grade ix|grade x|icse 10|cbse 10)\b", full_text):
        return "Secondary (Class 9–10 / IGCSE / ICSE)"
        
    # 4. Middle School (6-8)
    if re.search(r"\b(class [6-8]|grade [6-8]|6th|7th|8th|middle school|class vi|class vii|class viii)\b", full_text):
        return "Middle School (Class 6–8)"
        
    # 5. Primary School (1-5)
    if re.search(r"\b(class [1-5]|grade [1-5]|1st|2nd|3rd|4th|5th|primary school|primary|class i\b|class ii\b|class iii\b|class iv\b|class v\b)\b", full_text):
        return "Primary (Class 1–5)"
        
    # 6. Pre-Primary (Early Childhood / KG)
    if re.search(r"\b(pre-?kg|nursery|kindergarten|kg|early childhood|phonics|pre-?primary)\b", full_text):
        return "Pre-Primary (Early Childhood / KG)"
        
    # 7. Higher Education (University / Engineering / Medical)
    if re.search(r"\b(engineering|b\.?tech|m\.?tech|bca|mca|b\.?sc|m\.?sc|b\.?com|m\.?com|mbbs|pharma|university|college|compiler design|data structure|algorithms|operating system|dbms|microbiology|biochemistry|anatomy|physiology|pharmacology|calculus|thermodynamics|fluid mechanics|strength of materials|digital electronics|higher education|doctorate|mphil|phd)\b", full_text):
        return "Higher Education (College / University)"
        
    # 8. Professional & Industry Upskilling (Advanced Tech & Corporate Tools)
    if re.search(r"\b(full stack|machine learning|deep learning|data science|aws|cloud|devops|react|node|docker|kubernetes|power bi|tableau|cyber security|blockchain|ai|artificial intelligence|data analytics|corporate training|software development)\b", full_text):
        return "Professional & Industry Upskilling"
        
    # 9. Foundations & Creative Skills (Arts, Music, Yoga, Hobbies, Beginner Languages)
    if re.search(r"\b(yoga|music|guitar|piano|dance|painting|drawing|chess|singing|fitness|art & craft|a1|a2|spoken english|conversation)\b", full_text):
        return "Foundations & Creative Skills"
        
    # 10. Diploma & Vocational
    if re.search(r"\b(diploma|polytechnic|vocational)\b", full_text):
        return "Diploma & Vocational"
        
    # 11. Profile-based fallback
    if role == "Professor/Lecturer" or qual == "Advanced (PhD/Doctorate)":
        return "Higher Education (College / University)"
    elif qual == "Postgraduate (PG)" or role == "Teacher":
        return "Senior Secondary (Class 11–12 / Boards)"
    elif qual == "Undergraduate (UG)":
        return "Secondary (Class 9–10 / IGCSE / ICSE)"
        
    return "Foundations & General Learning"

# -------------------------------------------------------------
# 6. SUBJECT CATEGORIES (7 Standardized Categories)
# -------------------------------------------------------------

SUBJECT_CATEGORY_RULES = [
    # Coding / CS / Tech
    (r"\b(python|java\b|sql|html|css|javascript|c\+\+|c programming|\.net|c#|react|angular|node|stack|development|data structure|algorithms|deep learning|software|programming|machine learning|ai|artificial intelligence|data science|excel|coding|computer science|cse|computer|dbms|operating system|compiler|web development|aws|cloud|power bi|autocad|cad|electronics|digital|tableau|r programming|matlab|cyber|data analytics|data analysis|vba|word|powerpoint)\b", "Coding"),
    
    # Science / STEM
    (r"\b(physics|chemistry|biology|botany|zoology|biotechnology|mathematics|math|maths|statistics|science|evs|environmental|algebra|calculus|geometry|trigonometry|anatomy|physiology|biochemistry|microbiology|pharmacology|pathology|medical|neet|jee|thermodynamics|fluid mechanics|strength of materials|mechanics|astronomy|genetics|nanotechnology|dental|mbbs|nursing)\b", "Science"),
    
    # Commerce / Business / Finance
    (r"\b(economics|accountancy|accounts|accounting|business studies|business|management|commerce|finance|marketing|taxation|ca|costing|banking|gst|auditing|corporate|macroeconomics|microeconomics|financial|mba|bba|hr|human resource)\b", "Commerce"),
    
    # Arts / Humanities / Social Sciences / Law
    (r"\b(history|philosophy|political science|polity|geography|sst|social science|social studies|psychology|sociology|international relations|design|civics|public administration|law|jurisprudence|constitution|upsc|social|humanities|anthropology)\b", "Arts"),
    
    # Languages
    (r"\b(english|hindi|odia|oriya|spanish|malayalam|kannada|telugu|tamil|sanskrit|punjabi|french|arabic|marathi|urdu|chinese|japanese|german|thai|persian|konkani|bengali|gujarati|russian|italian|korean|ielts|toefl|grammar|spoken english|phonics|quran|tajweed|quraan)\b", "Languages"),
    
    # Misc / Extra-Curricular Activities (ECA)
    (r"\b(music|dance|art|craft|sports|gym|drawing|painting|yoga|chess|singing|guitar|piano|keyboard|vocal|fitness|handwriting|calligraphy|swimming|karate|acting|theatre|flute|violin|drums)\b", "Misc/ECA"),
    
    # All Subjects
    (r"\b(all subjects|general subjects|all|general)\b", "All Subjects"),
]

CANONICAL_SUBJECT_MAP = [
    (r"\b(maths?|mathematics|algebra|calculus|geometry|trigonometry|mathematica)\b", "Mathematics"),
    (r"\b(physics)\b", "Physics"),
    (r"\b(chemistry|organic chemistry|inorganic chemistry)\b", "Chemistry"),
    (r"\b(biology|botany|zoology|microbiology|biochemistry|anatomy|physiology)\b", "Biology & Life Sciences"),
    (r"\b(science|general science)\b", "Science (General)"),
    (r"\b(evs|environmental science|environmental studies)\b", "EVS"),
    (r"\b(statistics|stats|biostatistics)\b", "Statistics"),
    (r"\b(computer science|cse|computer|coding|programming|c / c\+\+|c and c\+\+|c\+\+|c programming|c#|\.net)\b", "Computer Science & Programming"),
    (r"\b(python)\b", "Python"),
    (r"\b(java\b)\b", "Java"),
    (r"\b(sql|database|dbms|mysql|oracle)\b", "SQL & Databases"),
    (r"\b(web development|html|javascript|frontend|backend|react|node|css)\b", "Web Development"),
    (r"\b(data science|machine learning|ai|deep learning|data analysis|power bi|tableau)\b", "AI & Data Science"),
    (r"\b(economics|macroeconomics|microeconomics)\b", "Economics"),
    (r"\b(accountancy|accounts|accounting|costing)\b", "Accountancy"),
    (r"\b(business studies|business|management|marketing|finance)\b", "Business Studies & Management"),
    (r"\b(commerce)\b", "Commerce"),
    (r"\b(english|english & grammar|spoken english|phonics|ielts|toefl)\b", "English"),
    (r"\b(hindi)\b", "Hindi"),
    (r"\b(bengali)\b", "Bengali"),
    (r"\b(tamil)\b", "Tamil"),
    (r"\b(telugu)\b", "Telugu"),
    (r"\b(kannada)\b", "Kannada"),
    (r"\b(malayalam)\b", "Malayalam"),
    (r"\b(marathi)\b", "Marathi"),
    (r"\b(french)\b", "French"),
    (r"\b(german)\b", "German"),
    (r"\b(spanish)\b", "Spanish"),
    (r"\b(sanskrit)\b", "Sanskrit"),
    (r"\b(arabic|quran|tajweed|quraan)\b", "Arabic & Islamic Studies"),
    (r"\b(social studies|social science|sst|social and political life)\b", "Social Science"),
    (r"\b(history)\b", "History"),
    (r"\b(geography)\b", "Geography"),
    (r"\b(political science|polity|civics|law)\b", "Political Science & Law"),
    (r"\b(psychology)\b", "Psychology"),
    (r"\b(sociology)\b", "Sociology"),
    (r"\b(yoga|fitness|gym)\b", "Yoga & Fitness"),
    (r"\b(music|singing|guitar|piano|keyboard|vocal)\b", "Music"),
    (r"\b(art|drawing|painting|craft|calligraphy)\b", "Art & Craft"),
    (r"\b(all subjects)\b", "All Subjects"),
]

def parse_subject_and_level(raw_subj_token, raw_level_token=None, hookline=None, desc=None, qual=None, role=None):
    subj = str(raw_subj_token).strip() if raw_subj_token else ""
    if not subj:
        return None
    
    subj_lower = subj.lower()
    
    # 1. Competitive Exam
    extracted_exam = None
    for pattern, exam in COMP_EXAM_PATTERNS:
        if pattern.search(subj):
            extracted_exam = exam
            break
            
    # 2. Contextual Target Student Level
    target_lvl = infer_target_level(raw_level_token, subj, hookline, desc, qual, role, extracted_exam)
    
    # 3. Canonical Subject
    canonical_subj = None
    for pattern, csubj in CANONICAL_SUBJECT_MAP:
        if re.search(pattern, subj_lower):
            canonical_subj = csubj
            break
    if not canonical_subj:
        clean_fallback = re.sub(r"\(.*?\)", "", subj).strip()
        canonical_subj = clean_fallback.title() if clean_fallback else subj.title()
        
    # 4. Standardized Subject Category
    subj_cat = "Science"
    matched = False
    for pattern, cat in SUBJECT_CATEGORY_RULES:
        if re.search(pattern, subj_lower):
            subj_cat = cat
            matched = True
            break
    if not matched:
        subj_cat = "All Subjects" if "all" in subj_lower else "Science"

    return {
        "raw_subject": subj,
        "canonical_subject": canonical_subj,
        "subject_category": subj_cat,
        "level_category": target_lvl,
        "competitive_exam": extracted_exam
    }

# -------------------------------------------------------------
# 7. FEE, EXPERIENCES, BOOLEANS & GENDER
# -------------------------------------------------------------

def parse_fee(fee_str):
    if not fee_str or not str(fee_str).strip():
        return None, None, None, None, None
    s = str(fee_str).strip()
    
    unit = "other"
    if "/hour" in s.lower():
        unit = "hour"
    elif "/month" in s.lower():
        unit = "month"
    elif "/course" in s.lower() or "/package" in s.lower():
        unit = "course"
        
    inr_match = re.search(r"₹([\d,]+)(?:[\–\-]([\d,]+))?", s)
    if inr_match:
        f1_str = inr_match.group(1).replace(",", "")
        try:
            f1 = float(f1_str)
        except ValueError:
            f1 = None
            
        f2 = None
        if inr_match.group(2):
            f2_str = inr_match.group(2).replace(",", "")
            try:
                f2 = float(f2_str)
            except ValueError:
                f2 = None
                
        min_f = f1
        max_f = f2 if f2 is not None else f1
        avg_f = (min_f + max_f) / 2.0 if (min_f is not None and max_f is not None) else None
        
        hourly_fee = avg_f if unit == "hour" else None
        return unit, min_f, max_f, avg_f, hourly_fee
        
    return unit, None, None, None, None

def parse_numeric_exp(exp_str):
    if not exp_str or not str(exp_str).strip():
        return None
    m = re.search(r"([\d\.]+)\s*yrs", str(exp_str).lower())
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return None
    return None

def clean_bool(val):
    if not val:
        return 0
    s = str(val).lower()
    return 1 if "yes" in s or "true" in s or s == "1" else 0

def clean_gender(val):
    if not val:
        return "Not Specified"
    s = str(val).lower()
    if "female" in s:
        return "Female"
    elif "male" in s:
        return "Male"
    return "Not Specified"

def clean_works_as(val):
    if not val:
        return "Not Specified"
    return str(val).replace("Works as:", "").strip()

def clean_speaks(val):
    if not val:
        return None, 0
    s = str(val).replace("Speaks:", "").strip()
    langs = [x.strip() for x in s.split(",") if x.strip()]
    return ", ".join(langs), len(langs)
