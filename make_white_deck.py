import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import os

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# COLOR PALETTE: Clean White / Editorial Theme - ZERO BLUE!
BG_WHITE = RGBColor(255, 255, 255)         # Pure White
CARD_BG = RGBColor(248, 250, 252)          # Very light soft gray/off-white (Slate 50)
BORDER_GRAY = RGBColor(226, 232, 240)      # Clean subtle border (Slate 200)

TEXT_DARK = RGBColor(17, 24, 39)           # Charcoal obsidian (Gray 900) - ultra legible
TEXT_MUTED = RGBColor(100, 116, 139)       # Medium slate gray (Slate 500)
TEXT_BODY = RGBColor(51, 65, 85)           # Body slate (Slate 700)

# NON-BLUE ACCENT COLORS:
ACCENT_PRIMARY = RGBColor(5, 150, 105)     # Deep Emerald Green (Emerald 600)
ACCENT_PURPLE = RGBColor(109, 40, 217)     # Royal Violet (Violet 700)
ACCENT_AMBER = RGBColor(217, 119, 6)       # Warm Amber (Amber 600)
ACCENT_ROSE = RGBColor(225, 29, 72)        # Rose Crimson (Rose 600)

BADGE_BG = RGBColor(236, 253, 245)         # Light Emerald 50
BADGE_BORDER = RGBColor(167, 243, 208)     # Emerald 200

blank_layout = prs.slide_layouts[6]

def add_white_bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_WHITE
    bg.line.fill.background()
    return bg

def add_top_bar(slide, topic, num, total=17):
    # Topic Badge (Soft Emerald pill, green text)
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(2.3), Inches(0.38))
    badge.fill.solid()
    badge.fill.fore_color.rgb = BADGE_BG
    badge.line.color.rgb = BADGE_BORDER
    tf = badge.text_frame
    tf.word_wrap = False
    tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = topic.upper()
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PRIMARY

    # Slide number at top right
    num_box = slide.shapes.add_textbox(Inches(11.0), Inches(0.4), Inches(1.5), Inches(0.38))
    tf_num = num_box.text_frame
    tf_num.word_wrap = False
    tf_num.margin_top = tf_num.margin_bottom = tf_num.margin_left = tf_num.margin_right = 0
    p_num = tf_num.paragraphs[0]
    p_num.alignment = PP_ALIGN.RIGHT
    r_num = p_num.add_run()
    r_num.text = f"{num:02d} / {total:02d}"
    r_num.font.size = Pt(11)
    r_num.font.bold = True
    r_num.font.color.rgb = TEXT_MUTED

def add_header(slide, title, subtitle):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.88), Inches(11.733), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
    
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = TEXT_DARK
    
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = subtitle
    r2.font.size = Pt(12)
    r2.font.color.rgb = TEXT_MUTED

def set_notes(slide, target_time, script, viva_q, viva_a):
    notes = slide.notes_slide.notes_text_frame
    notes.text = f"""Target Pacing: {target_time}

Word-for-Word Script (Joint Equal Authors: Nilesh Mori & Pratik Chauhan):
"{script}"

Viva Defense:
Q: {viva_q}
A: {viva_a}"""

# ----------------- SLIDE 1: Title -----------------
s1 = prs.slides.add_slide(blank_layout)
add_white_bg(s1)

tb = s1.shapes.add_textbox(Inches(1.0), Inches(1.4), Inches(11.333), Inches(4.8))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "DA-IICT  ·  B.TECH MINI PROJECT\n\n"
r.font.size = Pt(12)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE

p = tf.add_paragraph()
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "TutorMetric\n"
r.font.size = Pt(46)
r.font.bold = True
r.font.color.rgb = TEXT_DARK

p = tf.add_paragraph()
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "What 107,673 private tutors in India actually charge\n\n"
r.font.size = Pt(18)
r.font.bold = True
r.font.color.rgb = ACCENT_PRIMARY

p = tf.add_paragraph()
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "We turned raw tutor data into a clean, searchable database with fair pricing benchmarks, maps, and fee estimators.\n\n"
r.font.size = Pt(13)
r.font.color.rgb = TEXT_BODY

p = tf.add_paragraph()
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "Nilesh Mori (202301473)  ·  Pratik Chauhan (202301435)  ·  Joint Equal Authors\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = TEXT_DARK

r2 = p.add_run()
r2.text = "Mentor: Dr. Parul Gupta  ·  DA-IICT, Gandhinagar  ·  Live: bmptutors.onrender.com"
r2.font.size = Pt(11)
r2.font.color.rgb = TEXT_MUTED

set_notes(s1, "35s",
    "Respected professors and examiners, good morning. We are Nilesh Mori and Pratik Chauhan, presenting our B.Tech Mini Project under Dr. Parul Gupta: TutorMetric — What 107,673 private tutors in India actually charge. Private tutoring is a massive informal market where parents and teachers have no public benchmarks. We cleaned our mentor's verified dataset, filtered dormant accounts, standardized subjects, and built an open platform with fair pricing benchmarks, maps, and fee estimators.",
    "Why is this a Mini Project rather than commercial software?",
    "It is an academic research engineering project for our B.Tech Mini Project at DA-IICT. Our focus is empirical market transparency, clean data engineering, and econometric analysis guided by Dr. Parul Gupta's project roadmap.")

# ----------------- SLIDE 2: Motivation -----------------
s2 = prs.slides.add_slide(blank_layout)
add_white_bg(s2)
add_top_bar(s2, "Motivation", 2)
add_header(s2, "The Problem: Two parents, two wildly different quotes", "A massive market where every household is left guessing what is fair")

cards2 = [
    ("₹200 vs ₹2,000 / hour", "Wild price spread", "In the same city for Class 10 CBSE Maths, one tutor asks ₹200/hr, another asks ₹2,000. Neither number has a public benchmark to check."),
    ("30,990 Ghost Accounts", "28.8% inactive", "Commercial listing sites leave dead accounts online to inflate directory size. 28.8% of profiles in the raw data were inactive."),
    ("21,956 Messy Spellings", "Text fragmentation", "Free typing created thousands of variants for the same subject (Maths, Math, Mathematics). Parents couldn't compare reliably."),
    ("Our Mandate", "Mini Project Goal", "Verify active teachers, normalize subjects, build a clean 3NF database, and publish transparent benchmarks for the country.")
]
coords2 = [
    (Inches(0.8), Inches(2.0), Inches(5.7), Inches(2.2)),
    (Inches(6.8), Inches(2.0), Inches(5.7), Inches(2.2)),
    (Inches(0.8), Inches(4.5), Inches(5.7), Inches(2.2)),
    (Inches(6.8), Inches(4.5), Inches(5.7), Inches(2.2))
]
for (t, b, d), (l, top, w, h) in zip(cards2, coords2):
    box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, top, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = BORDER_GRAY
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.2)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = f"{t}   [{b}]\n\n"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PRIMARY
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = d
    r2.font.size = Pt(11)
    r2.font.color.rgb = TEXT_BODY

set_notes(s2, "45s",
    "We start directly with our project motivation. In India's private tuition market, severe information asymmetry harms families. In Ahmedabad for Class 10 Math, quotes range from ₹200 to ₹2,000 for identical syllabus, with zero benchmark to tell who is fair. Furthermore, commercial websites monetize phone unlocks and keep stale accounts active. Our audit showed 30,990 profiles (28.8%) were dormant. Our mandate was to clean this master data, identify active teachers, standardize subjects, and publish open pricing benchmarks.",
    "How did you verify the 28.8% inactive accounts?",
    "We ran automated Playwright headless verification and activity checks on contact signals and response timestamps to isolate dormant profiles.")

# ----------------- SLIDE 3: Value (What TutorMetric does that TeacherOn doesn't) -----------------
s3 = prs.slides.add_slide(blank_layout)
add_white_bg(s3)
add_top_bar(s3, "TutorMetric vs TeacherOn", 3)
add_header(s3, "What TutorMetric does that TeacherOn doesn't do", "TeacherOn is a paid classifieds site; TutorMetric is an open analytical utility")

# Left box (TeacherOn)
b_l = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.7), Inches(4.7))
b_l.fill.solid()
b_l.fill.fore_color.rgb = RGBColor(255, 241, 242) # Rose 50
b_l.line.color.rgb = RGBColor(254, 205, 211)     # Rose 200
tf_l = b_l.text_frame
tf_l.word_wrap = True
tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = Inches(0.25)
p = tf_l.paragraphs[0]
r = p.add_run()
r.text = "✕  What TeacherOn Does (Commercial Marketplace)\n\n"
r.font.size = Pt(13.5)
r.font.bold = True
r.font.color.rgb = ACCENT_ROSE
items_l = [
    "Opaque Pricing: Only displays an unverified fee typed by the tutor with zero context.",
    "Paywalls Contacts: Charges ₹500–₹1,500 in coins to unlock a phone number.",
    "Retains Dormant Profiles: Keeps 30,990 inactive accounts online to inflate catalog size.",
    "Text Fragmentation: 21,956 chaotic spellings with no subject taxonomy.",
    "Zero Analytics: No macro charts, no GIS maps, no city tier analysis, no fee models.",
    "Closed Data: Tutors and researchers cannot download or study the database."
]
for it in items_l:
    p = tf_l.add_paragraph()
    r = p.add_run()
    r.text = f"•  {it}\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT_DARK

# Right box (TutorMetric)
b_r = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(4.7))
b_r.fill.solid()
b_r.fill.fore_color.rgb = RGBColor(240, 253, 244) # Emerald 50
b_r.line.color.rgb = RGBColor(187, 247, 208)     # Emerald 200
tf_r = b_r.text_frame
tf_r.word_wrap = True
tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = Inches(0.25)
p = tf_r.paragraphs[0]
r = p.add_run()
r.text = "✓  What TutorMetric Adds (Research Utility)\n\n"
r.font.size = Pt(13.5)
r.font.bold = True
r.font.color.rgb = ACCENT_PRIMARY
items_r = [
    "Fair Benchmarks: State medians, level ladders, and formula predicted rates.",
    "100% Free & Open: No paywalls, no coin system, direct contact visibility.",
    "Active Verification: Explicit flag showing 76,683 active vs 30,990 inactive.",
    "Standardized 386 Subjects: Clean taxonomy across 6 macro categories with 0 loss.",
    "Interactive Analytics: 14 charts, Leaflet state map, MPC 3-way Venn diagram.",
    "ML & GenAI Tools: Live Fee Estimator and Google Gemini GenAI Bio Studio.",
    "Open SQLite / Excel: Full 3NF database and filtered cohorts downloadable."
]
for it in items_r:
    p = tf_r.add_paragraph()
    r = p.add_run()
    r.text = f"✓  {it}\n"
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(6, 78, 59)

set_notes(s3, "45s",
    "A central question is: what does TutorMetric do that TeacherOn doesn't do? TeacherOn is a commercial lead-generation marketplace. It monetizes contact unlocks, retains thousands of dormant accounts to inflate numbers, and shows self-written fees with no context. TutorMetric transforms this into an open analytical utility: we provide transparent price benchmarks, verify active teachers, normalize 21k messy subjects into 386 clean categories, offer interactive Leaflet maps and Venn diagrams, and deploy live ML fee estimators and Gemini Bio tools with 100% open data access.",
    "What is the single biggest advantage of TutorMetric over TeacherOn?",
    "Transparency. Parents get fair benchmark pricing and verified active teachers for free, while TeacherOn paywalls phone numbers and leaves families to guess what is fair.")

# ----------------- SLIDE 4: Project Tracker -----------------
s4 = prs.slides.add_slide(blank_layout)
add_white_bg(s4)
add_top_bar(s4, "Project Tracker", 4)
add_header(s4, "Our Workflow: Guided by Dr. Parul Gupta's Project Tracker", "Six structured phases from master Excel data to a deployed web platform")

steps4 = [
    ("01", "Master Data Merge", "Combined scraped (76,706) & unscraped (31,017) cohorts into 107,673 unique tutor records."),
    ("02", "Cleaning Taxonomy", "Mapped 21,956 spellings to 386 subjects, 4 qualification tiers, and 5 macro regions."),
    ("03", "3NF Database Design", "Replaced messy 78-column sheet with 2 indexed SQLite tables (107k tutors, 403k offerings)."),
    ("04", "Visual Analytics & EDA", "Built 14 interactive charts, Leaflet state map, and STEM 3-way Venn diagram."),
    ("05", "Econometric Model", "Trained Ridge regression (L2) on 72,001 tutors (R² = 9.84%, MAE = ₹370/hr)."),
    ("06", "Production Web App", "Deployed live site on Render with live ML fee estimator, Gemini Bio Studio, and Excel export.")
]
grid4 = [
    (Inches(0.8), Inches(2.0), Inches(3.7), Inches(2.2)),
    (Inches(4.8), Inches(2.0), Inches(3.7), Inches(2.2)),
    (Inches(8.8), Inches(2.0), Inches(3.7), Inches(2.2)),
    (Inches(0.8), Inches(4.5), Inches(3.7), Inches(2.2)),
    (Inches(4.8), Inches(4.5), Inches(3.7), Inches(2.2)),
    (Inches(8.8), Inches(4.5), Inches(3.7), Inches(2.2)),
]
for (num, title, desc), (l, top, w, h) in zip(steps4, grid4):
    box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, top, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = BORDER_GRAY
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.18)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = f"{num}  ·  {title}\n\n"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PURPLE
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = desc
    r2.font.size = Pt(10.5)
    r2.font.color.rgb = TEXT_BODY

set_notes(s4, "40s",
    "To ensure rigor, we followed the roadmap in ProjectTracker.xlsx across six phases. We merged the scraped and unscraped tabs in TrueData.xlsx, cleaned categories according to Categories_DataCleaning.docx, normalized into 3NF SQLite, produced visual analytics, trained an econometric pricing model, and deployed our web application at bmptutors.onrender.com.",
    "Did you complete the tracker tasks?",
    "Yes, we covered the full analytical pipeline from cleaning to modeling and deployment, prioritizing interactive web tools over static Word tables.")

# ----------------- SLIDE 5: Data Cleaning -----------------
s5 = prs.slides.add_slide(blank_layout)
add_white_bg(s5)
add_top_bar(s5, "Data Cleaning", 5)
add_header(s5, "Data Hygiene: Keeping promises while cleaning", "107,673 tutors, 403,221 offerings, and zero records lost")

metrics5 = [
    ("107,673", "Tutors in Dataset"),
    ("403,221", "Subject Offerings (0 Lost)"),
    ("₹4,000 / hr", "Winsorization Cap (99th %)"),
    ("100%", "PII Anonymized (T1..T107673)")
]
for i, (val, lbl) in enumerate(metrics5):
    m_box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.98), Inches(2.0), Inches(2.8), Inches(1.1))
    m_box.fill.solid()
    m_box.fill.fore_color.rgb = CARD_BG
    m_box.line.color.rgb = BORDER_GRAY
    tf = m_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = val + "\n"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PRIMARY
    r2 = p.add_run()
    r2.text = lbl
    r2.font.size = Pt(10)
    r2.font.color.rgb = TEXT_MUTED

clean_steps = [
    ("Privacy First", "Stripped names, phones, emails, and exact addresses. Replaced with anonymous IDs T1, T2..."),
    ("Fee Misclassifications", "Detected monthly rates (e.g. ₹15,000) typed into hourly fields. Capped at ₹4,000/hr (99th percentile)."),
    ("Standard Categories", "Grouped states into 5 regions, cities into 3 tiers, and education into 4 clean qualification levels."),
    ("Teaching Modality", "Cleaned experience into numeric years and flagged online (95.6%) vs in-person home tutoring.")
]
clean_coords = [
    (Inches(0.8), Inches(3.3), Inches(5.7), Inches(1.8)),
    (Inches(6.8), Inches(3.3), Inches(5.7), Inches(1.8)),
    (Inches(0.8), Inches(5.3), Inches(5.7), Inches(1.8)),
    (Inches(6.8), Inches(5.3), Inches(5.7), Inches(1.8)),
]
for (t, d), (l, top, w, h) in zip(clean_steps, clean_coords):
    box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, top, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = BORDER_GRAY
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.18)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = f"◈  {t}\n\n"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PURPLE
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = d
    r2.font.size = Pt(10.5)
    r2.font.color.rgb = TEXT_BODY

set_notes(s5, "45s",
    "Our data cleaning followed strict hygiene rules. We removed personal info and assigned anonymous IDs T1 to T107673. We fixed unit issues where monthly fees were typed into hourly slots, and applied Winsorization at ₹4,000/hr to handle extreme fees without deleting valid high-end tutors. All 403,221 offerings survived the cleaning pass with zero records lost.",
    "Why Winsorize instead of deleting fees above ₹4,000?",
    "Deleting them would bias our sample by removing top entrance exam specialists. Winsorization bounds their statistical influence while keeping their profile and subjects in the database.")

# ----------------- SLIDE 6: Database (3NF) -----------------
s6 = prs.slides.add_slide(blank_layout)
add_white_bg(s6)
add_top_bar(s6, "Database (3NF)", 6)
add_header(s6, "Database: Why we split one big spreadsheet into two tables", "3NF normalization eliminates empty cells and speeds up search")

# Table 1 box
t1_b = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.7), Inches(2.8))
t1_b.fill.solid()
t1_b.fill.fore_color.rgb = CARD_BG
t1_b.line.color.rgb = BORDER_GRAY
tf = t1_b.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "TABLE 1: teachers  (107,673 rows)\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE
p2 = tf.add_paragraph()
r2 = p2.add_run()
r2.text = "One row per educator. Holds demographics, fee, experience, city tier, state, active status, and delivery mode.\n\n"
r2.font.size = Pt(11)
r2.font.color.rgb = TEXT_BODY
r3 = p2.add_run()
r3.text = "Primary Key: tutor_id (T1, T2...)"
r3.font.size = Pt(10)
r3.font.color.rgb = TEXT_MUTED

# Table 2 box
t2_b = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(2.8))
t2_b.fill.solid()
t2_b.fill.fore_color.rgb = CARD_BG
t2_b.line.color.rgb = BORDER_GRAY
tf2 = t2_b.text_frame
tf2.word_wrap = True
tf2.margin_left = tf2.margin_right = tf2.margin_top = Inches(0.2)
p = tf2.paragraphs[0]
r = p.add_run()
r.text = "TABLE 2: teacher_subjects  (403,221 rows)\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PRIMARY
p2 = tf2.add_paragraph()
r2 = p2.add_run()
r2.text = "Junction table mapping tutors to subjects. Because tutors teach an average of 3.79 subjects, each offering gets its own clean row.\n\n"
r2.font.size = Pt(11)
r2.font.color.rgb = TEXT_BODY
r3 = p2.add_run()
r3.text = "Foreign Key: tutor_id -> teachers.tutor_id"
r3.font.size = Pt(10)
r3.font.color.rgb = TEXT_MUTED

adv6 = [
    ("64% Smaller", "Storage Savings", "Eliminated 20 repeated subject columns where over 60% of cells were empty NULLs."),
    ("< 50 ms", "Fast Search", "Compound B-Tree index allows instant multi-filter search across 403k offerings."),
    ("3.79 Avg", "Subjects / Tutor", "Flexible 1-to-many relationship handles generalists and specialists cleanly.")
]
for i, (st, lb, dc) in enumerate(adv6):
    box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 4.0), Inches(5.0), Inches(3.7), Inches(1.8))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = BORDER_GRAY
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = st + "\n"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PRIMARY
    r2 = p.add_run()
    r2.text = lb + "\n"
    r2.font.size = Pt(11)
    r2.font.bold = True
    r2.font.color.rgb = TEXT_DARK
    r3 = p.add_run()
    r3.text = dc
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = TEXT_MUTED

set_notes(s6, "45s",
    "In the original spreadsheet, each tutor row had 20 repeated subject columns, and 60% of cells were empty NULLs. We split this into 3NF with two tables: 'teachers' with 107k rows, and 'teacher_subjects' with 403k offerings. Because tutors teach an average of 3.79 subjects, this cut file size by 64% and enabled indexed joins that run under 50 milliseconds.",
    "Why SQLite instead of MongoDB?",
    "SQLite provides ACID guarantees and fast indexed B-Tree joins without network overhead, perfect for an analytical benchmark and web deployment.")

# ----------------- SLIDE 7: Taxonomy -----------------
s7 = prs.slides.add_slide(blank_layout)
add_white_bg(s7)
add_top_bar(s7, "Taxonomy", 7)
add_header(s7, "From 21,956 spellings to 386 clean subjects", "Zero offerings lost — all 403,221 entries grouped into 6 macro categories")

cats7 = [
    ("Science & Maths", "142,810 entries", "Mathematics, Physics, Chemistry, Biology, Statistics"),
    ("Languages", "98,450 entries", "English, Hindi, Sanskrit, French, German, Spanish"),
    ("Coding & CS", "64,220 entries", "Python, Java, C++, Web Development, Data Science"),
    ("Commerce & Econ", "45,180 entries", "Accountancy, Economics, Business Studies, Taxation"),
    ("Arts & Humanities", "32,940 entries", "History, Geography, Political Science, Psychology"),
    ("Hobbies & Other", "19,621 entries", "Music, Chess, Vedic Maths, Public Speaking, Yoga")
]
grid7 = [
    (Inches(0.8), Inches(2.0), Inches(3.7), Inches(2.3)),
    (Inches(4.8), Inches(2.0), Inches(3.7), Inches(2.3)),
    (Inches(8.8), Inches(2.0), Inches(3.7), Inches(2.3)),
    (Inches(0.8), Inches(4.5), Inches(3.7), Inches(2.3)),
    (Inches(4.8), Inches(4.5), Inches(3.7), Inches(2.3)),
    (Inches(8.8), Inches(4.5), Inches(3.7), Inches(2.3)),
]
for (name, count, ex), (l, top, w, h) in zip(cats7, grid7):
    box = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, top, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = BORDER_GRAY
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.18)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = f"{name}\n"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PRIMARY
    r2 = p.add_run()
    r2.text = f"{count}\n\n"
    r2.font.size = Pt(10.5)
    r2.font.bold = True
    r2.font.color.rgb = ACCENT_PURPLE
    p2 = tf.add_paragraph()
    r3 = p2.add_run()
    r3.text = f"Includes: {ex}"
    r3.font.size = Pt(10)
    r3.font.color.rgb = TEXT_BODY

set_notes(s7, "40s",
    "Tutors entered subjects freely, creating 21,956 unique spellings—like 'Maths', 'Math', and 'Mathematics'. Following Categories_DataCleaning.docx, we mapped these into 386 clean canonical disciplines across 6 macro-categories: Science, Languages, Coding, Commerce, Arts, and Hobbies. Not a single offering was lost; all 403,221 rows were preserved.",
    "Did you discard unrecognized subjects?",
    "No, any entry that did not match a standard school subject was classified into canonical specialized fields or placed into Hobbies/Other, keeping row loss at zero.")

# ----------------- SLIDE 8: Geography -----------------
s8 = prs.slides.add_slide(blank_layout)
add_white_bg(s8)
add_top_bar(s8, "Geography", 8)
add_header(s8, "Where tutors live: Top 5 states hold 54.4%", "Private tutoring follows coaching culture, not state population")

img8 = '/home/nilesh7757/BMPTUTORS1/assets/all_india_tutor_choropleth.png'
if os.path.exists(img8):
    s8.shapes.add_picture(img8, Inches(0.8), Inches(2.0), width=Inches(5.4))

panel8 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.0), Inches(5.9), Inches(4.8))
panel8.fill.solid()
panel8.fill.fore_color.rgb = CARD_BG
panel8.line.color.rgb = BORDER_GRAY
tf = panel8.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "TOP 5 STATES (54.4% OF ALL TUTORS):\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE
states8 = [
    ("1. Tamil Nadu", "16,806 tutors (15.6%)", "Avg ₹782/hr"),
    ("2. Kerala", "14,333 tutors (13.3%)", "Avg ₹612/hr"),
    ("3. Maharashtra", "11,544 tutors (10.7%)", "Avg ₹1,070/hr"),
    ("4. West Bengal", "8,024 tutors (7.5%)", "Avg ₹695/hr"),
    ("5. Delhi NCR", "7,869 tutors (7.3%)", "Avg ₹1,142/hr")
]
for st, cnt, fee in states8:
    p = tf.add_paragraph()
    r = p.add_run()
    r.text = f"{st}:  {cnt}  ·  {fee}\n"
    r.font.size = Pt(11)
    r.font.color.rgb = TEXT_DARK
p2 = tf.add_paragraph()
r = p2.add_run()
r.text = "\nKey Takeaway:\nKerala has 2.8% of India's population but 13.3% of its tutors. Bihar has 8.6% of population but only 1.4% of tutors (1,520). Tuition supply reflects local education culture and competitive exam drive, not raw headcount."
r.font.size = Pt(10.5)
r.font.color.rgb = TEXT_BODY

set_notes(s8, "45s",
    "Tutor supply is heavily concentrated: five states account for 54.4% of all tutors in India. Tamil Nadu and Kerala lead nationally. This reveals a clear divergence: Kerala holds 2.8% of India's population but 13.3% of tutors, while Bihar holds 8.6% of the population but only 1.4% of tutors. Tutoring density reflects household coaching investment and literacy rather than population.",
    "Why are fees in Kerala lower (₹612/hr) despite high tutor supply?",
    "High supply of qualified graduates creates healthy local competition, which exerts downward competitive pressure on hourly fees.")

# ----------------- SLIDE 9: Market Structure -----------------
s9 = prs.slides.add_slide(blank_layout)
add_white_bg(s9)
add_top_bar(s9, "Market Structure", 9)
add_header(s9, "The Central Finding: 3 out of 4 tutors work outside big cities", "74.1% of tutors are in Tier-3 semi-urban towns, enabled by online teaching")

img9 = '/home/nilesh7757/BMPTUTORS1/assets/chart_regional_distribution.png'
if os.path.exists(img9):
    s9.shapes.add_picture(img9, Inches(0.8), Inches(2.0), width=Inches(5.4))

panel9 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.0), Inches(5.9), Inches(4.8))
panel9.fill.solid()
panel9.fill.fore_color.rgb = CARD_BG
panel9.line.color.rgb = BORDER_GRAY
tf = panel9.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "CITY TIER BREAKDOWN (N = 107,673):\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE

tiers9 = [
    ("Tier-3 Semi-Urban Towns", "74.1% (79,820 tutors)", "Avg ₹823/hr", "The true backbone of Indian private tutoring."),
    ("Tier-2 Urban Cities", "13.2% (14,231 tutors)", "Avg ₹895/hr", "State capitals and industrial hubs (Pune, Ahmedabad)."),
    ("Tier-1 Metros", "12.7% (13,622 tutors)", "Avg ₹1,022/hr", "Delhi, Mumbai, Bengaluru, Chennai, Kolkata.")
]
for t, s_cnt, fee, dc in tiers9:
    p = tf.add_paragraph()
    r = p.add_run()
    r.text = f"• {t}: {s_cnt}  ·  {fee}\n"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = TEXT_DARK
    r2 = p.add_run()
    r2.text = f"   {dc}\n"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = TEXT_MUTED

p2 = tf.add_paragraph()
r = p2.add_run()
r.text = "\nOnline Revolution:\n95.6% of tutors teach online. Broadband connectivity allows small-town educators to teach students in major metros, breaking geographic boundaries."
r.font.size = Pt(10.5)
r.font.color.rgb = ACCENT_AMBER

set_notes(s9, "45s",
    "This was our most striking empirical finding: 74.1% of private tutors live outside Tier-1 metros in Tier-3 semi-urban towns. Only 12.7% are in metro cities. EdTech commercial marketing focuses almost entirely on metros, yet the vast reality of Indian tuition is a distributed, small-town phenomenon. Because 95.6% of these tutors teach online, broadband has connected small-town educators to students across the country.",
    "How were cities categorized into tiers?",
    "We followed Government of India / RBI classifications: Tier-1 for the 8 major metros, Tier-2 for state capitals and industrial centers, and Tier-3 for all other municipalities.")

# ----------------- SLIDE 10: Pricing -----------------
s10 = prs.slides.add_slide(blank_layout)
add_white_bg(s10)
add_top_bar(s10, "Pricing Dynamics", 10)
add_header(s10, "How fees climb with each school level", "National median is ₹650/hr, but mean is ₹864/hr due to top specialist rates")

img10 = '/home/nilesh7757/BMPTUTORS1/assets/chart_education_levels.png'
if os.path.exists(img10):
    s10.shapes.add_picture(img10, Inches(0.8), Inches(2.0), width=Inches(5.4))

panel10 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.0), Inches(5.9), Inches(4.8))
panel10.fill.solid()
panel10.fill.fore_color.rgb = CARD_BG
panel10.line.color.rgb = BORDER_GRAY
tf = panel10.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "THE PRICING LADDER (HOURLY RATES):\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE

ladder10 = [
    ("Pre-Primary & Primary", "₹710 / hr", "Base"),
    ("Middle School (Class 6-8)", "₹760 / hr", "+7%"),
    ("Secondary (Class 9-10)", "₹825 / hr", "+16%"),
    ("Higher Secondary (Class 11-12 STEM)", "₹950 / hr", "+34%"),
    ("College / Degree", "₹1,080 / hr", "+52%"),
    ("Professional Skills", "₹1,257 / hr", "+77% Premium")
]
for lv, fee, pct in ladder10:
    p = tf.add_paragraph()
    r = p.add_run()
    r.text = f"• {lv}:  {fee}  ({pct})\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT_DARK

p2 = tf.add_paragraph()
r = p2.add_run()
r.text = "\nFacts from the Data:\n• Median is ₹650, but mean is ₹864 because top 8.6% charge ₹2,000+.\n• Tier-1 tutors command an average ₹199/hr premium over Tier-3.\n• Fees jump sharply (+15.2%) entering Class 11-12 board exams."
r.font.size = Pt(9.5)
r.font.color.rgb = TEXT_BODY

set_notes(s10, "45s",
    "Hourly fees follow a steady ladder as student age climbs. Primary tutors average ₹710/hr, Secondary averages ₹825, Class 11-12 STEM averages ₹950, and Professional coaching reaches ₹1,257—a 77% premium over primary school. The national median is ₹650 while the mean is ₹864, because a high-earning minority charging ₹2,000+ pulls up the arithmetic average.",
    "Why does the Class 11-12 step have such a big fee jump?",
    "Because of intense board examination and competitive entrance pressures (JEE, NEET), where parents are willing to pay a premium for specialized coaching.")

# ----------------- SLIDE 11: Subject Bundling -----------------
s11 = prs.slides.add_slide(blank_layout)
add_white_bg(s11)
add_top_bar(s11, "Subject Bundling", 11)
add_header(s11, "The STEM Trio: How many tutors teach Math, Physics & Chemistry?", "Set-theoretic analysis reveals that multi-subject generalists are rare")

img11 = '/home/nilesh7757/BMPTUTORS1/assets/venn_stem_trio.png'
if os.path.exists(img11):
    s11.shapes.add_picture(img11, Inches(0.8), Inches(2.0), width=Inches(5.4))

panel11 = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.0), Inches(5.9), Inches(4.8))
panel11.fill.solid()
panel11.fill.fore_color.rgb = CARD_BG
panel11.line.color.rgb = BORDER_GRAY
tf = panel11.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "SET INTERSECTION: MATHS, PHYSICS, CHEMISTRY\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE

stats11 = [
    ("34,923 tutors", "Teach at least one of Maths, Physics, or Chem."),
    ("2,917 tutors (8.35%)", "The MPC Trio: teach all three subjects together."),
    ("16,929 tutors (48.5%)", "Pure Mathematics: teach Maths and nothing else."),
    ("4,210 tutors", "Physics & Chemistry: paired without Mathematics.")
]
for st, dc in stats11:
    p = tf.add_paragraph()
    r = p.add_run()
    r.text = f"• {st}\n"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PRIMARY
    r2 = p.add_run()
    r2.text = f"   {dc}\n"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = TEXT_BODY

p2 = tf.add_paragraph()
r = p2.add_run()
r.text = "\nAcademic Takeaway:\nCoaching ads often claim 'all subjects in one teacher', but empirical data proves only 8.35% teach the full science trio. Specialization strongly dominates secondary coaching."
r.font.size = Pt(10)
r.font.color.rgb = TEXT_MUTED

set_notes(s11, "40s",
    "Using our 3NF database, we performed set intersection across the core STEM subjects: Mathematics, Physics, and Chemistry. Across 34,923 STEM tutors, only 2,917—or 8.35%—teach all three subjects. In contrast, 16,929 tutors teach Mathematics alone. Multi-subject generalists are rare; specialization dominates secondary and higher secondary coaching.",
    "How does the web application generate this Venn diagram?",
    "Our server executes SQL INTERSECT queries on the teacher_subjects table, mapping selected subjects into disjoint subsets in real time.")

# ----------------- SLIDE 12: Demographics -----------------
s12 = prs.slides.add_slide(blank_layout)
add_white_bg(s12)
add_top_bar(s12, "Demographics", 12)
add_header(s12, "Who teaches: 58.7% women, and degrees command a premium", "A PhD earns 41.8% more; 1 in 5 tutors has a non-teaching primary job")

img12 = '/home/nilesh7757/BMPTUTORS1/assets/fig2_experience_gender.png'
if os.path.exists(img12):
    s12.shapes.add_picture(img12, Inches(0.8), Inches(2.0), width=Inches(5.4))

panel12 = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.0), Inches(5.9), Inches(4.8))
panel12.fill.solid()
panel12.fill.fore_color.rgb = CARD_BG
panel12.line.color.rgb = BORDER_GRAY
tf = panel12.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "GENDER & DEGREE PREMIUM:\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE

genders12 = [
    ("Languages", "78% Female", "English 77%, Hindi 79%"),
    ("Biology / Sciences", "73% Female", "Life sciences dominated by women"),
    ("Mathematics", "58% Female", "Nearly balanced"),
    ("Coding & CS", "61% Male", "Only male-majority subject")
]
for sub, f_pct, dc in genders12:
    p = tf.add_paragraph()
    r = p.add_run()
    r.text = f"• {sub}: {f_pct}  ({dc})\n"
    r.font.size = Pt(10)
    r.font.color.rgb = TEXT_DARK

p2 = tf.add_paragraph()
r = p2.add_run()
r.text = "\nDegree & Gig Dynamics:\n• PhD holders average ₹1,179.5/hr vs ₹832 for graduates (+41.8% premium).\n• 20,275 tutors (18.8%) hold a day job in IT, engineering, or banking."
r.font.size = Pt(10)
r.font.color.rgb = TEXT_BODY

set_notes(s12, "40s",
    "In demographics, women make up 58.7% of tutors. Female representation is highest in languages (78%) and biology (73%). Coding is the only male-majority subject at 61%, reflecting broader software industry workforce patterns. Academic degrees matter: tutors with a PhD command a 41.8% fee premium over undergraduates (₹1,179 vs ₹832). Additionally, 1 in 5 tutors teaches as a secondary gig alongside an IT or corporate job.",
    "Why is coding the only male-majority domain?",
    "It mirrors the gender distribution of software engineers who participate in evening tuition as a secondary gig occupation.")

# ----------------- SLIDE 13: Econometric Model -----------------
s13 = prs.slides.add_slide(blank_layout)
add_white_bg(s13)
add_top_bar(s13, "Econometric Model", 13)
add_header(s13, "Pricing Model: Why formulas only explain 10% of fees", "Ridge regression on 72,001 tutors: math gives a baseline, trust sets the price")

stats13 = [
    ("72,001", "Tutors Used", "Active with stated fees"),
    ("38", "Model Features", "Experience, city, subjects"),
    ("9.84%", "Model R²", "Variance explained"),
    ("₹370.67", "Avg Error (MAE)", "Prediction baseline (₹/hr)")
]
for i, (val, lbl, sub) in enumerate(stats13):
    box = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.98), Inches(2.0), Inches(2.8), Inches(1.1))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = BORDER_GRAY
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = val + "\n"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PRIMARY
    r2 = p.add_run()
    r2.text = lbl + "\n"
    r2.font.size = Pt(10.5)
    r2.font.bold = True
    r2.font.color.rgb = TEXT_DARK
    r3 = p.add_run()
    r3.text = sub
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = TEXT_MUTED

b_o = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(3.3), Inches(5.7), Inches(3.6))
b_o.fill.solid()
b_o.fill.fore_color.rgb = RGBColor(240, 253, 244) # Emerald 50
b_o.line.color.rgb = RGBColor(187, 247, 208)
tf_o = b_o.text_frame
tf_o.word_wrap = True
tf_o.margin_left = tf_o.margin_right = tf_o.margin_top = Inches(0.2)
p = tf_o.paragraphs[0]
r = p.add_run()
r.text = "✓  What the Model Captures (R² = 9.84%)\n\n"
r.font.size = Pt(12)
r.font.bold = True
r.font.color.rgb = ACCENT_PRIMARY
obs_items = [
    "Experience: Adds ~₹18.40 per year of teaching.",
    "City size: Metro location adds ~₹199 per hour.",
    "Advanced degree: PhD adds ~₹347 per hour.",
    "Student level: College coaching adds ~₹320 per hour."
]
for it in obs_items:
    p = tf_o.add_paragraph()
    r = p.add_run()
    r.text = f"•  {it}\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT_DARK

b_u = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(3.3), Inches(5.7), Inches(3.6))
b_u.fill.solid()
b_u.fill.fore_color.rgb = RGBColor(254, 243, 199) # Amber 50
b_u.line.color.rgb = RGBColor(253, 230, 138)
tf_u = b_u.text_frame
tf_u.word_wrap = True
tf_u.margin_left = tf_u.margin_right = tf_u.margin_top = Inches(0.2)
p = tf_u.paragraphs[0]
r = p.add_run()
r.text = "◈  Why R² Stays Low (The Unobservable 90%)\n\n"
r.font.size = Pt(12)
r.font.bold = True
r.font.color.rgb = ACCENT_AMBER
unobs_items = [
    "Word of mouth: Parental recommendations and trust.",
    "Teaching patience: How well a tutor handles struggling kids.",
    "Pedigree: Alumni of premier institutions (IIT/NIT).",
    "Track record: Past student board and entrance exam results."
]
for it in unobs_items:
    p = tf_u.add_paragraph()
    r = p.add_run()
    r.text = f"•  {it}\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT_DARK

set_notes(s13, "55s",
    "Fulfilling the tracker tasks, we built an econometric fee model using Ridge regression on 72,001 tutors. It achieves an R² of 9.84% with an average error of ₹370/hr. Why does a 38-feature model only explain 10% of fees? Because tutoring is an experience good governed by unobservable qualities: word of mouth, personal patience, institutional pedigree, and past student exam results. Observable data gives a reliable baseline (±₹370/hr), but individual pricing ultimately runs on trust.",
    "Why Ridge instead of Ordinary Least Squares (OLS)?",
    "One-hot encoding 386 subjects creates multicollinearity. Ridge's L2 penalty shrinks correlated coefficients, preventing overfitting and keeping coefficients economically stable.")

# ----------------- SLIDE 14: Directory -----------------
s14 = prs.slides.add_slide(blank_layout)
add_white_bg(s14)
add_top_bar(s14, "Directory", 14)
add_header(s14, "Finding the right tutor in four clicks", "Sub-50ms query latency, 9 reactive filters, and instant Excel export")

img14 = '/home/nilesh7757/BMPTUTORS1/assets/screenshot_landing_directory.png'
if os.path.exists(img14):
    s14.shapes.add_picture(img14, Inches(0.8), Inches(2.0), width=Inches(5.4))

panel14 = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.0), Inches(5.9), Inches(4.8))
panel14.fill.solid()
panel14.fill.fore_color.rgb = CARD_BG
panel14.line.color.rgb = BORDER_GRAY
tf = panel14.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "NINE REACTIVE FILTERS:\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE

filters14 = [
    ("Location & State", "Filter across all 38 states/UTs and city tiers."),
    ("386 Subjects", "Select any standardized canonical subject."),
    ("Fee Range Slider", "Dual handle slider from ₹100 to ₹4,000/hr."),
    ("Experience & Mode", "0–30+ years experience, online or home tuition."),
    ("Active Check", "Show active verified teachers (76k) with one click."),
    ("Instant Excel Export", "Download any filtered list directly to Excel.")
]
for f_t, f_d in filters14:
    p = tf.add_paragraph()
    r = p.add_run()
    r.text = f"✓ {f_t}: {f_d}\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT_DARK

set_notes(s14, "40s",
    "Here is our live directory dashboard. Users can isolate relevant educators in under four clicks using nine reactive filters: state, city tier, fee range, modality, experience, and active verification status. All 386 clean subjects are indexed, searches run in under 50ms against SQLite, and any filtered list can be exported straight to Excel.",
    "How does it handle 107k profiles without browser lag?",
    "We use server-side pagination with SQLite B-Tree indexing, transferring only 12 to 100 rows per request.")

# ----------------- SLIDE 15: Interactive Tools -----------------
s15 = prs.slides.add_slide(blank_layout)
add_white_bg(s15)
add_top_bar(s15, "Interactive Tools", 15)
add_header(s15, "Two live tools deployed on the platform", "Translating data engineering and machine learning into practical utilities")

t1_box = s15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.7), Inches(4.7))
t1_box.fill.solid()
t1_box.fill.fore_color.rgb = CARD_BG
t1_box.line.color.rgb = BORDER_GRAY
tf1 = t1_box.text_frame
tf1.word_wrap = True
tf1.margin_left = tf1.margin_right = tf1.margin_top = Inches(0.25)
p = tf1.paragraphs[0]
r = p.add_run()
r.text = "⚡  AI Fee Estimator\n\n"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE
p2 = tf1.add_paragraph()
r2 = p2.add_run()
r2.text = "Enter 7 parameters (subject, student level, city tier, experience, mode, qualification) to calculate a fair hourly rate.\n\n"
r2.font.size = Pt(11)
r2.font.color.rgb = TEXT_BODY
items15_a = [
    "Recommended hourly rate with confidence range (±₹370/hr).",
    "Monthly and annual earnings projection based on hours taught.",
    "Cost breakdown showing what each qualification and year adds."
]
for it in items15_a:
    p = tf1.add_paragraph()
    r = p.add_run()
    r.text = f"✓  {it}\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT_DARK

t2_box = s15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(4.7))
t2_box.fill.solid()
t2_box.fill.fore_color.rgb = CARD_BG
t2_box.line.color.rgb = BORDER_GRAY
tf2 = t2_box.text_frame
tf2.word_wrap = True
tf2.margin_left = tf2.margin_right = tf2.margin_top = Inches(0.25)
p = tf2.paragraphs[0]
r = p.add_run()
r.text = "✨  Gemini GenAI Bio Studio\n\n"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = ACCENT_PRIMARY
p2 = tf2.add_paragraph()
r2 = p2.add_run()
r2.text = "Turns raw credentials into engaging professional biographies that build parent trust and student conversions.\n\n"
r2.font.size = Pt(11)
r2.font.color.rgb = TEXT_BODY
items15_b = [
    "Two tailored tones: Parent-focused (trust) & Student-focused (mentorship).",
    "Generates catchy headlines tailored to the teacher's subjects.",
    "100% operational with offline rule-based fallback templates."
]
for it in items15_b:
    p = tf2.add_paragraph()
    r = p.add_run()
    r.text = f"✓  {it}\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = TEXT_DARK

set_notes(s15, "45s",
    "To make our research actionable, we built two interactive tools: First, the ML Fee Estimator. Tutors and parents enter 7 details to receive a recommended hourly rate, confidence range, income estimates, and an itemized breakdown. Second, the Bio Studio powered by Google Gemini GenAI, which writes student-converting profile bios with parent and student versions, backed by offline fallback templates.",
    "What happens if the Gemini API key is missing?",
    "The application gracefully switches to a rule-based deterministic template engine that formats tutor bios without interrupting the user experience.")

# ----------------- SLIDE 16: Key Takeaways -----------------
s16 = prs.slides.add_slide(blank_layout)
add_white_bg(s16)
add_top_bar(s16, "Findings & Impact", 16)
add_header(s16, "Five things the data showed us", "Empirical conclusions from 107,673 tutors and future research scope")

takes16 = [
    ("01", "74.1% Tier-3", "Tuition in India is predominantly a small-town reality, not a metro phenomenon."),
    ("02", "South Leads", "Southern states lead national supply (Tamil Nadu 16.8k, Kerala 14.3k tutors)."),
    ("03", "58.7% Women", "Women are the majority workforce in languages and sciences; men lead in coding."),
    ("04", "STEM Specialists", "Only 8.35% teach all 3 sciences; single-subject specialization is the norm."),
    ("05", "Trust-Based", "Formulas provide a baseline (±₹370), while reputation and trust set individual price.")
]
for i, (num, tit, dc) in enumerate(takes16):
    box = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.38), Inches(2.0), Inches(2.25), Inches(3.4))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = BORDER_GRAY
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.18)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = f"{num}\n"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = ACCENT_PURPLE
    r2 = p.add_run()
    r2.text = f"{tit}\n\n"
    r2.font.size = Pt(11)
    r2.font.bold = True
    r2.font.color.rgb = TEXT_DARK
    p2 = tf.add_paragraph()
    r3 = p2.add_run()
    r3.text = dc
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = TEXT_BODY

fw_box = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.1))
fw_box.fill.solid()
fw_box.fill.fore_color.rgb = CARD_BG
fw_box.line.color.rgb = BORDER_GRAY
tf = fw_box.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.15)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "FUTURE WORK:  "
r.font.size = Pt(11)
r.font.bold = True
r.font.color.rgb = ACCENT_PRIMARY
r2 = p.add_run()
r2.text = "Quarterly fee tracking across seasons  ·  District-level GIS maps inside states  ·  Integrating parent review ratings into fee regression"
r2.font.size = Pt(10)
r2.font.color.rgb = TEXT_BODY

set_notes(s16, "45s",
    "In conclusion, our B.Tech Mini Project delivers five verified empirical conclusions: First, 74.1% of tutoring operates in Tier-3 semi-urban markets. Second, Southern states dominate national supply. Third, women represent 58.7% of tutors. Fourth, STEM coaching is heavily specialized. Fifth, tutoring fees are anchored by trust and reputation. For future work, we plan quarterly fee tracking, district maps, and parent review scores.",
    "How can this inform educational policy in India?",
    "By showing that 74% of tutoring is outside metros and that fees vary tenfold in the same zip code, education authorities can establish fair coaching guidelines.")

# ----------------- SLIDE 17: Live Demo -----------------
s17 = prs.slides.add_slide(blank_layout)
add_white_bg(s17)
add_top_bar(s17, "Live Demo", 17)
add_header(s17, "Conclusion & Live Platform Demonstration", "bmptutors.onrender.com")

left_box = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.7), Inches(4.7))
left_box.fill.solid()
left_box.fill.fore_color.rgb = CARD_BG
left_box.line.color.rgb = BORDER_GRAY
tf_l = left_box.text_frame
tf_l.word_wrap = True
tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = Inches(0.25)
p = tf_l.paragraphs[0]
r = p.add_run()
r.text = "PROJECT CREDENTIALS\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PURPLE

p = tf_l.add_paragraph()
r = p.add_run()
r.text = "TutorMetric  ·  B.Tech Mini Project\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = TEXT_DARK

p = tf_l.add_paragraph()
r = p.add_run()
r.text = "Joint Equal Authors:\nNilesh Mori (202301473)  ·  Pratik Chauhan (202301435)\n\nMentor:\nDr. Parul Gupta  ·  DA-IICT, Gandhinagar\n\nLive Platform:\nbmptutors.onrender.com"
r.font.size = Pt(11)
r.font.color.rgb = TEXT_BODY

right_box = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(4.7))
right_box.fill.solid()
right_box.fill.fore_color.rgb = CARD_BG
right_box.line.color.rgb = BORDER_GRAY
tf_r = right_box.text_frame
tf_r.word_wrap = True
tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = Inches(0.25)
p = tf_r.paragraphs[0]
r = p.add_run()
r.text = "LIVE DEMO PROTOCOL (5 MINS)\n\n"
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = ACCENT_PRIMARY

demo17 = [
    ("1. Directory Search", "Filter Ahmedabad -> Class 10 -> Mathematics -> Active only -> Export Excel."),
    ("2. Visual Analytics", "Inspect 74.1% Tier-3 pie chart, Leaflet state map, and MPC 3-way Venn diagram."),
    ("3. Live Fee AI", "Calculate recommended rate and cost breakdown for a Class 10 Math teacher."),
    ("4. Gemini Bio Studio", "Generate conversion-tailored profile bio live on screen.")
]
for st, dc in demo17:
    p = tf_r.add_paragraph()
    r = p.add_run()
    r.text = f"✓  {st}\n"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = TEXT_DARK
    r2 = p.add_run()
    r2.text = f"    {dc}\n\n"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = TEXT_MUTED

set_notes(s17, "5 Mins",
    "Thank you, professors. We invite your questions on our database schema, cleaning methodology, or econometric model. We will now switch to the live deployed application at bmptutors.onrender.com for our interactive demonstration.",
    "Is the deployed system running locally or on cloud infrastructure?",
    "It is running live on Render's container platform with Docker and automated continuous integration from GitHub.")

out_paths = [
    '/home/nilesh7757/Downloads/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/Downloads/BMP_Tutors_Deck_Full.pptx',
    '/home/nilesh7757/BMPTUTORS1/project_showcase_bundle/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/BMPTUTORS1/project_showcase_bundle/BMP_Tutors_Deck_Full.pptx',
    '/home/nilesh7757/BMPTUTORS1/TutorMetric_Master_Presentation.pptx'
]
for p in out_paths:
    prs.save(p)
    print(f"Saved clean white deck: {p} ({os.path.getsize(p):,} bytes)")

print("Clean white presentation deck (no blue) built successfully!")
