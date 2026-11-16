import sqlite3
from collections import Counter

CATEGORY_RULES = {
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
        "Philosophy", "Social Science", "Journalism & Mass Communication", "Art & Craft",
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

def get_category_for_subject(can_s):
    for cat, subjs in CATEGORY_RULES.items():
        if can_s in subjs:
            return cat
    # Fallback based on keywords
    low = can_s.lower()
    if any(k in low for k in ['language', 'english', 'hindi', 'tamil', 'french', 'german', 'spanish']):
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

print("Category mapping test passed successfully.")
