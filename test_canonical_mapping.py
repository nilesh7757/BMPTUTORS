import sqlite3, re
from collections import Counter

conn = sqlite3.connect('cleaned_tutors_data.db')
c = conn.cursor()
c.execute("SELECT DISTINCT canonical_subject FROM teacher_subjects WHERE canonical_subject IS NOT NULL")
raw_subjects = [r[0] for r in c.fetchall()]

# Explicit canonical mapping rules
EXACT_MAPPINGS = {
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
    
    # English prompt examples
    "engilsh": "English",
    "englsih": "English",
    "spoken english": "English",
    "english grammar": "English",
    "english literature": "English",
    "phonics": "English",
    "ielts": "English",
    "toefl": "English",
    "pte": "English",
    "bengali": "Bengali",
}

def clean_subject_string(s):
    if not s or not s.strip():
        return "All Subjects"
    raw = s.strip()
    low = raw.lower()
    
    # Check exact map
    if low in EXACT_MAPPINGS:
        return EXACT_MAPPINGS[low]

    # Indic Script handling
    if any('\u0900' <= ch <= '\u097f' for ch in s):
        if any(w in s for w in ['गणित', 'अंकगणित', 'बीजगणित', 'रेखागणित']): return 'Mathematics'
        if any(w in s for w in ['हिंदी', 'हिन्दी']): return 'Hindi'
        if any(w in s for w in ['संस्कृत', 'स्तोत्र', 'हवन', 'वेद', 'वैदिक']): return 'Sanskrit'
        if any(w in s for w in ['विज्ञान', 'भौतिकी', 'रसायन', 'जीव']): return 'Science (General)'
        if any(w in s for w in ['इतिहास', 'भूगोल', 'नागरिक', 'सामाजिक']): return 'Social Science'
        return 'Hindi'
    if any('\u0980' <= ch <= '\u09ff' for ch in s):
        if 'গণিত' in s: return 'Mathematics'
        if 'হাদিস' in s or 'কোরআন' in s: return 'Arabic & Islamic Studies'
        return 'Bengali'
    if any('\u0b80' <= ch <= '\u0bff' for ch in s):
        if 'கணிதம்' in s: return 'Mathematics'
        return 'Tamil'
    if any('\u0c00' <= ch <= '\u0c7f' for ch in s):
        if 'గణితం' in s: return 'Mathematics'
        return 'Telugu'
    if any('\u0c80' <= ch <= '\u0cff' for ch in s):
        return 'Kannada'
    if any('\u0d00' <= ch <= '\u0d7f' for ch in s):
        if 'ഗണിതം' in s: return 'Mathematics'
        if 'ശാസ്ത്രം' in s: return 'Science (General)'
        return 'Malayalam'

    # Domain Pattern Rules (order matters)
    # 1. Mathematics & Statistics
    if re.search(r'\b(maths?|mathematics?|algebra|calculus|geometry|trigonometry|arithmetic|topology|cuemath|vedic|abacus|linear algebra|differential equation|discrete math)\b', low):
        return "Mathematics"
    if re.search(r'\b(statist|statistics|probability|biostat|actuarial|econometric)\b', low):
        return "Statistics"

    # 2. Physics & Natural Sciences
    if re.search(r'\b(physics|astrophysics|geophysics|astronomy|thermodynamics|quantum|optics|electromagnet)\b', low):
        return "Physics"
    if re.search(r'\b(chemistry|biochem|organic chem|inorganic chem|physical chem|chemical analysis)\b', low):
        return "Chemistry"
    if re.search(r'\b(biology|botany|zoology|microbiol|genetics|biotech|anatomy|physiology|pathology|immunology|ecology|life science|embryology|virology|neet bio|bio neet)\b', low):
        return "Biology & Life Sciences"
    if re.search(r'\b(medicine|nursing|pharmacy|pharmacology|dental|mbbs|medical|clinical|surgery|ayurveda|homeopathy|physiotherapy)\b', low):
        return "Medicine & Healthcare"
    if re.search(r'\b(science|general science|natural science|evs|environmental science|environmental studies|earth science)\b', low):
        return "Science (General)"

    # 3. Languages
    if re.search(r'\b(english|grammar|spoken english|phonics|ielts|toefl|pte|creative writing|literature|comprehension|esl|verbal)\b', low):
        return "English"
    if re.search(r'\b(hindi|shuddha hindi)\b', low):
        return "Hindi"
    if re.search(r'\b(bengali|bangla)\b', low):
        return "Bengali"
    if re.search(r'\b(tamil)\b', low):
        return "Tamil"
    if re.search(r'\b(telugu)\b', low):
        return "Telugu"
    if re.search(r'\b(kannada)\b', low):
        return "Kannada"
    if re.search(r'\b(malayalam)\b', low):
        return "Malayalam"
    if re.search(r'\b(marathi)\b', low):
        return "Marathi"
    if re.search(r'\b(gujarati)\b', low):
        return "Gujarati"
    if re.search(r'\b(punjabi|gurmukhi)\b', low):
        return "Punjabi"
    if re.search(r'\b(urdu)\b', low):
        return "Urdu"
    if re.search(r'\b(sanskrit)\b', low):
        return "Sanskrit"
    if re.search(r'\b(french)\b', low):
        return "French"
    if re.search(r'\b(german)\b', low):
        return "German"
    if re.search(r'\b(spanish)\b', low):
        return "Spanish"
    if re.search(r'\b(arabic|quran|tajweed|islamic studies|hadees|islamiyat)\b', low):
        return "Arabic & Islamic Studies"
    if re.search(r'\b(japanese)\b', low):
        return "Japanese"
    if re.search(r'\b(chinese|mandarin)\b', low):
        return "Chinese (Mandarin)"
    if re.search(r'\b(russian)\b', low):
        return "Russian"
    if re.search(r'\b(korean)\b', low):
        return "Korean"
    if re.search(r'\b(italian)\b', low):
        return "Italian"
    if re.search(r'\b(portuguese)\b', low):
        return "Portuguese"
    if re.search(r'\b(odia|oriya)\b', low):
        return "Odia"
    if re.search(r'\b(assamese)\b', low):
        return "Assamese"
    if re.search(r'\b(nepali)\b', low):
        return "Nepali"
    if re.search(r'\b(persian|farsi)\b', low):
        return "Persian"
    if re.search(r'\b(turkish)\b', low):
        return "Turkish"
    if re.search(r'\b(latin)\b', low):
        return "Latin"

    # 4. Computer Science, Programming & IT
    if re.search(r'\b(python|django|flask|fastapi|pandas|numpy)\b', low):
        return "Python"
    if re.search(r'\b(java\b|spring boot|spring framework|hibernate)\b', low) and 'javascript' not in low:
        return "Java"
    if re.search(r'\b(c\+\+|c programming|c / c\+\+|c and c\+\+|c#|\.net|embedded c)\b', low):
        return "C / C++ / C#"
    if re.search(r'\b(web develop|html|css|javascript|react|node|frontend|backend|full stack|angular|vue|bootstrap|wordpress|php|tailwind)\b', low):
        return "Web Development"
    if re.search(r'\b(sql|database|dbms|mysql|oracle|mongodb|postgresql|nosql|plsql|sqlite)\b', low):
        return "SQL & Databases"
    if re.search(r'\b(data science|machine learning|artificial intelligence|deep learning|nlp|computer vision|power bi|tableau|data analytics|data analysis|big data|neural network|excel)\b', low):
        return "AI & Data Science"
    if re.search(r'\b(cloud|aws|azure|gcp|devops|docker|kubernetes|jenkins|linux|unix|git|terraform|ansible)\b', low):
        return "Cloud Computing & DevOps"
    if re.search(r'\b(cyber security|cybersecurity|ethical hacking|network security|information security|networking|ccna)\b', low):
        return "Cybersecurity & Networking"
    if re.search(r'\b(software testing|selenium|automation testing|manual testing|qa testing|test automation)\b', low):
        return "Software Testing"
    if re.search(r'\b(android|ios|flutter|swift|react native|mobile app|app development|kotlin)\b', low):
        return "Mobile App Development"
    if re.search(r'\b(computer science|programming|coding|dsa|data structure|algorithms|operating system|compiler|automata|oops|software engineering|computer architecture)\b', low):
        return "Computer Science & Programming"

    # 5. Commerce, Economics & Business
    if re.search(r'\b(account|accounts|accounting|accountancy|tally|bookkeeping|costing|taxation|gst|auditing|ca foundation|cma)\b', low):
        return "Accountancy"
    if re.search(r'\b(economics|microeconomics|macroeconomics|econometrics)\b', low):
        return "Economics"
    if re.search(r'\b(business|management|marketing|human resource|finance|financial|bba|mba|commerce|banking|supply chain|entrepreneurship|business studies)\b', low):
        return "Business Studies & Management"

    # 6. Humanities & Social Sciences
    if re.search(r'\b(social science|social studies|sst)\b', low):
        return "Social Science"
    if re.search(r'\b(history|ancient history|medieval history|modern history|world history|archaeology)\b', low):
        return "History"
    if re.search(r'\b(geography|geology|earth science|gis)\b', low):
        return "Geography"
    if re.search(r'\b(political science|polity|civics|law|legal|jurisprudence|constitution|clat|public administration)\b', low):
        return "Political Science & Law"
    if re.search(r'\b(psychology|counseling|mental health)\b', low):
        return "Psychology"
    if re.search(r'\b(sociology|social work|anthropology)\b', low):
        return "Sociology"
    if re.search(r'\b(philosophy|ethics|logic)\b', low):
        return "Philosophy"
    if re.search(r'\b(journalism|mass communication|media studies)\b', low):
        return "Journalism & Mass Communication"

    # 7. Engineering
    if re.search(r'\b(civil engineering|structural engineering|surveying|construction|strength of materials?)\b', low):
        return "Civil Engineering"
    if re.search(r'\b(mechanical engineering|cad|autocad|solidworks|ansys|fluid mechanics|mechatronics|engineering mechanics|machine design)\b', low):
        return "Mechanical Engineering"
    if re.search(r'\b(electrical engineering|power systems|circuit theory|electrical machines|control systems?)\b', low):
        return "Electrical Engineering"
    if re.search(r'\b(electronics|vlsi|embedded systems|microprocessor|telecommunication|signal processing|robotics|iot|arduino|digital electronics)\b', low):
        return "Electronics & Communication"
    if re.search(r'\b(chemical engineering|petroleum engineering)\b', low):
        return "Chemical Engineering"

    # 8. Creative, Performing Arts & Sports
    if re.search(r'\b(music|singing|vocal|guitar|piano|keyboard|violin|flute|drums|carnatic|hindustani|tabla|sitar|harmonium)\b', low):
        return "Music"
    if re.search(r'\b(art|drawing|painting|sketching|calligraphy|craft|fine arts|water color|oil painting)\b', low):
        return "Art & Craft"
    if re.search(r'\b(graphic design|photoshop|illustrator|ui/ux|ui ux|animation|video editing|vfx|3d modeling)\b', low):
        return "Design & Multimedia"
    if re.search(r'\b(dance|bharatanatyam|kathak|zumba|hip hop|salsa|ballet|folk dance)\b', low):
        return "Dance"
    if re.search(r'\b(yoga|fitness|gym|workout|aerobics|meditation|nutrition|dietetics|physical education)\b', low):
        return "Yoga & Fitness"
    if re.search(r'\b(chess)\b', low):
        return "Chess"
    if re.search(r'\b(karate|taekwondo|martial arts|swimming|badminton|cricket|tennis|sports)\b', low):
        return "Sports & Martial Arts"

    # 9. Competitive Exams & Aptitude
    if re.search(r'\b(aptitude|logical reasoning|reasoning|mental ability|quantitative aptitude)\b', low):
        return "Aptitude & Logical Reasoning"
    if re.search(r'\b(jee|neet|gate|cat exam|gre|gmat|sat|upsc|ssc|bank po|defence|nda|cds|olympiad|competitive exam)\b', low):
        return "Competitive Exam Preparation"
    if re.search(r'\b(general knowledge|current affairs|gk)\b', low):
        return "General Knowledge"
    if re.search(r'\b(personality development|communication skills|soft skills|public speaking|interview preparation|life coaching)\b', low):
        return "Personality Development & Soft Skills"
    if re.search(r'\b(special education|autism|adhd|dyslexia|special needs)\b', low):
        return "Special Education"

    # 10. All Subjects
    if re.search(r'\b(all subject|all subjects|all school subjects|all primary|all standards|all classes|all)\b', low):
        return "All Subjects"

    # Fallback to general domain or Other Academic Subjects
    clean = re.sub(r'[\(\)\[\]\{\}\<\>]', '', raw).strip().title()
    clean = re.sub(r'\s+', ' ', clean)
    if len(clean) <= 20 and len(clean.split()) <= 2 and clean.replace(' ', '').isalpha():
        return clean
    return "Other Academic Subjects"

mapped = {s: clean_subject_string(s) for s in raw_subjects}
counts = Counter(mapped.values())

# Let's prune very rare singleton clean words (< 5 occurrences) to "Other Academic Subjects"
c.execute("SELECT canonical_subject, COUNT(*) FROM teacher_subjects GROUP BY canonical_subject")
raw_counts = dict(c.fetchall())

final_counts = Counter()
for raw_s, can_s in mapped.items():
    final_counts[can_s] += raw_counts.get(raw_s, 0)

# Filter: if can_s has fewer than 15 total tutor offerings and is not in our primary list, collapse to "Other Academic Subjects"
PRIMARY_SET = set([
    "Mathematics", "English", "Science (General)", "Physics", "Chemistry", "Biology & Life Sciences",
    "Computer Science & Programming", "Python", "Java", "C / C++ / C#", "Web Development", "SQL & Databases",
    "AI & Data Science", "Cloud Computing & DevOps", "Cybersecurity & Networking", "Software Testing",
    "Mobile App Development", "Accountancy", "Economics", "Business Studies & Management",
    "History", "Geography", "Political Science & Law", "Psychology", "Sociology", "Philosophy",
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

mapped_pruned = {}
for raw_s, can_s in mapped.items():
    if can_s in PRIMARY_SET:
        mapped_pruned[raw_s] = can_s
    elif final_counts[can_s] >= 15:
        mapped_pruned[raw_s] = can_s
    else:
        mapped_pruned[raw_s] = "Other Academic Subjects"

unique_final = set(mapped_pruned.values())
print(f"Total raw subjects: {len(raw_subjects)}")
print(f"Final distinct canonical subjects: {len(unique_final)}")
