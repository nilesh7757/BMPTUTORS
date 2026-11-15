import os
import json
import urllib.request
import re

# Load environment variables from .env if present
ENV_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"
)
if os.path.exists(ENV_PATH):
    with open(ENV_PATH, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


def call_gemini_api(prompt: str, api_key: str = None) -> dict:
    key = api_key or os.getenv("GEMINI_API_KEY")
    if not key:
        return None

    # Use Gemini 2.5 Flash (highest quality, fast, native JSON structured mode)
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={key}"

    payload = json.dumps(
        {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "responseMimeType": "application/json",
                "temperature": 0.75,
                "topP": 0.95,
            },
        }
    ).encode("utf-8")

    req = urllib.request.Request(
        url, data=payload, headers={"Content-Type": "application/json"}
    )

    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            text = data["candidates"][0]["content"]["parts"][0]["text"]
            return json.loads(text)
    except Exception as e:
        print(f"Gemini API call warning: {e}. Falling back to procedural engine.")
        return None


def generate_profile_bio(
    teacher_name: str = "Tutor",
    subjects: str = "Mathematics & Science",
    target_level: str = "Class 9–10 (CBSE/ICSE)",
    qualification: str = "Master of Science (M.Sc)",
    experience_years: int = 5,
    teaching_style: str = "Concept-First & Visual Learning",
    languages: str = "English & Hindi",
    achievements: str = "Helped 200+ students achieve 90%+ in board exams",
    usp: str = "Personalized doubt-clearing sessions & weekly mock tests",
    api_key: str = None,
):
    name = teacher_name.strip() if teacher_name.strip() else "Educator"
    subj = subjects.strip() if subjects.strip() else "Mathematics"
    level = target_level.strip() if target_level.strip() else "Secondary School"
    qual = qualification.strip() if qualification.strip() else "Postgraduate"
    exp = int(experience_years) if experience_years else 3
    style = teaching_style.strip() if teaching_style.strip() else "Concept-First"
    lang = languages.strip() if languages.strip() else "English and Hindi"
    achieve = (
        achievements.strip()
        if achievements.strip()
        else "Consistent student improvement and top board scores"
    )
    usp_val = (
        usp.strip()
        if usp.strip()
        else "Regular progress tracking and dedicated doubt support"
    )

    # 1. Try Gemini GenAI First
    gemini_prompt = f"""
You are an expert copywriter for high-converting online tutor and educator profiles in India.
Generate completely bespoke, deeply personalized, human-sounding profile descriptions for this tutor:

Tutor Profile Attributes:
- Educator Name: {name}
- Academic Degree & College: {qual}
- Total Teaching Experience: {exp} Years
- Subject(s) Specialization: {subj}
- Target Student Level / Board: {level}
- Teaching Methodology & Style: {style}
- Key Academic Track Record: {achieve}
- Unique Value Proposition (USP): {usp_val}
- Language(s) of Instruction: {lang}

Requirements:
1. Write 'bio_parent' (140-190 words):
   - Addressed to Indian parents. Warm, reassuring, direct.
   - Covers teaching approach, weekly tests, and ends with an invitation to book a trial class.
2. Write 'bio_student' (130-180 words):
   - Addressed directly to students in a warm, empathetic, relatable voice.
   - Eliminates fear of {subj}, promises zero boring lectures, emphasizes step-by-step clarity, visual shortcuts, and safe doubt clearing.
   - Concludes with an encouraging call-to-action.
3. Write 'headlines' (Array of 3 distinct, punchy one-liners under 85 characters each):
   - High click-through rate, highlighting credentials, experience, and outcome.

Return ONLY pure valid JSON in this exact structure:
{{
  "bio_parent": "...",
  "bio_student": "...",
  "headlines": ["Headline 1", "Headline 2", "Headline 3"],
  "ai_powered": true,
  "model_used": "gemini-2.5-flash"
}}
"""

    gemini_result = call_gemini_api(gemini_prompt, api_key=api_key)
    if (
        gemini_result
        and "bio_parent" in gemini_result
        and "bio_student" in gemini_result
    ):
        gemini_result["summary"] = {
            "name": name,
            "subjects": subj,
            "level": level,
            "exp_years": exp,
            "style": style,
            "engine": "Google Gemini 2.5 Flash",
        }
        return gemini_result

    # 2. Procedural Fallback Engine (natural templates if API unavailable)
    bio_parent = f"""Hi, I'm {name}. I hold a {qual} and have spent the last {exp} years teaching {subj} to {level} students.

My approach is simple: {style.lower()}. I've seen what works — {achieve}. {usp_val}.

I'm comfortable teaching in {lang}, and I keep parents updated on how things are going after each milestone.

If you'd like to try a session, I offer a first trial class so you can see if my teaching style works for your child."""

    bio_student = f"""Hey, I'm {name} ({qual}). I've been teaching {subj} for {exp} years, and I know it can feel overwhelming at first.

Here's how I do it: we skip the boring lectures and break things down into real examples you can actually use. You can ask me anything — no question is too small. We'll also go over shortcuts and exam strategies that actually save time.

{usp_val}

Whether you're trying to jump from 70% to 95% or just want to get the basics right, I'm here to help. Book a session and let's get started."""

    headlines = [
        f"{name} — {qual} with {exp}+ years teaching {subj}",
        f"Learn {subj} with {name} — {achieve[:40]}",
        f"{subj} tutor: {exp}+ years, concept-first approach",
    ]

    return {
        "bio_parent": bio_parent.strip(),
        "bio_student": bio_student.strip(),
        "headlines": headlines,
        "ai_powered": False,
        "model_used": "procedural-fallback",
        "summary": {
            "name": name,
            "subjects": subj,
            "level": level,
            "exp_years": exp,
            "style": style,
            "engine": "Procedural Rule Engine",
        },
    }


def analyze_profile_bio(text: str):
    if not text or not text.strip():
        return {
            "score": 0,
            "grade": "Incomplete",
            "word_count": 0,
            "character_count": 0,
            "checks": {
                "has_credentials": False,
                "has_experience": False,
                "has_subject": False,
                "has_methodology": False,
                "has_cta": False,
            },
            "suggestions": ["Please enter your bio text to evaluate profile strength."],
        }

    clean_text = text.strip()
    words = re.findall(r"\b[a-zA-Z0-9_\-\+\%\.\#]+\b", clean_text)
    word_count = len(words)
    char_count = len(clean_text)

    # 1. Credentials Check
    cred_keywords = [
        "phd",
        "m.sc",
        "b.sc",
        "m.tech",
        "b.tech",
        "m.a",
        "b.a",
        "b.ed",
        "m.ed",
        "degree",
        "postgraduate",
        "graduate",
        "doctorate",
        "iit",
        "nit",
        "university",
        "faculty",
        "certified",
        "scholar",
        "educator",
    ]
    has_credentials = any(
        re.search(r"\b" + re.escape(k) + r"\b", clean_text, re.I) for k in cred_keywords
    )

    # 2. Experience Check
    has_experience = bool(
        re.search(
            r"\b(\d{1,2}\+?\s*(years?|yrs?)|experience|teaching\s+since|mentored|guided)\b",
            clean_text,
            re.I,
        )
    )

    # 3. Subject / Level Specificity
    subject_keywords = [
        "math",
        "mathematics",
        "physics",
        "chemistry",
        "biology",
        "science",
        "english",
        "coding",
        "python",
        "java",
        "accounts",
        "economics",
        "cbse",
        "icse",
        "igcse",
        "ib",
        "jee",
        "neet",
        r"class\s+\d+",
        r"grade\s+\d+",
    ]
    has_subject = any(
        re.search(r"\b" + k + r"\b", clean_text, re.I) for k in subject_keywords
    )

    # 4. Teaching Methodology / Value Prop
    method_keywords = [
        "concept",
        "fundamentals",
        "problem-solving",
        "doubts",
        "pedagogy",
        "practice",
        "notes",
        "mock",
        "test",
        "structured",
        "patient",
        "interactive",
        "step-by-step",
        "visual",
    ]
    has_methodology = any(
        re.search(r"\b" + re.escape(k) + r"\b", clean_text, re.I)
        for k in method_keywords
    )

    # 5. Call To Action (CTA)
    cta_keywords = [
        "book",
        "demo",
        "trial",
        "contact",
        r"reach\s+out",
        r"get\s+in\s+touch",
        "message",
        "call",
        "enroll",
        "schedule",
        "partner",
        "session",
    ]
    has_cta = any(re.search(r"\b" + k + r"\b", clean_text, re.I) for k in cta_keywords)

    # Scoring algorithm (0 to 100)
    score = 0
    suggestions = []

    # Length evaluation (optimal: 70 to 250 words)
    if word_count < 30:
        score += 10
        suggestions.append(
            "Your bio is too short. Marketplace data shows profiles with 80–180 words receive 3.4x more student inquiries."
        )
    elif word_count <= 60:
        score += 20
        suggestions.append(
            "Add more details about your teaching structure or specific achievements."
        )
    elif word_count <= 280:
        score += 30
    else:
        score += 25
        suggestions.append(
            "Your bio is quite lengthy. Consider using concise bullet points for easier mobile skimming."
        )

    # Feature weights
    if has_credentials:
        score += 15
    else:
        suggestions.append(
            "Mention your educational qualifications (e.g. M.Sc, B.Tech, B.Ed) to build trust with parents."
        )

    if has_experience:
        score += 15
    else:
        suggestions.append(
            "Explicitly state your years of teaching experience (e.g., '5+ years experience')."
        )

    if has_subject:
        score += 15
    else:
        suggestions.append(
            "Clearly specify your core subjects and target boards (e.g., CBSE Class 9–10 Science)."
        )

    if has_methodology:
        score += 15
    else:
        suggestions.append(
            "Highlight your teaching philosophy (e.g., concept-clarity, visual problem solving, or doubt clearing)."
        )

    if has_cta:
        score += 10
    else:
        suggestions.append(
            "Add a clear Call-to-Action at the end (e.g., 'Book a free demo session today!')."
        )

    score = min(100, max(15, score))

    if score >= 85:
        grade = "Excellent (High Conversion)"
    elif score >= 65:
        grade = "Good (Market Ready)"
    elif score >= 45:
        grade = "Average (Needs Polish)"
    else:
        grade = "Weak (Needs Optimization)"

    return {
        "score": score,
        "grade": grade,
        "word_count": word_count,
        "character_count": char_count,
        "checks": {
            "has_credentials": has_credentials,
            "has_experience": has_experience,
            "has_subject": has_subject,
            "has_methodology": has_methodology,
            "has_cta": has_cta,
        },
        "suggestions": suggestions,
    }
