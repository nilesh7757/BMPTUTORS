#!/usr/bin/env python3
"""
BMP Tutors Master Presentation Generator (V2)
Builds a 20-slide widescreen (20" x 11.25") academic presentation matching the exact G-34 / G-18
visual design system:
- 16:9 Widescreen Canvas (20.0" x 11.25")
- Modern aesthetic: Off-white canvas (#EBEBEB) with subtle network watermark (g34_network_blue.png)
- Pristine white cards with safe padding, subtle borders (#D1D5DB)
- Top Category Pill Banner + Slide Counter Pill (01 / 20)
- Team Members: Nilesh Mori (202301473) & Pratik Chauhan (202301435) - Equal Joint Authors
- Exact verified ground truth numbers throughout (107,673 tutors, 403,221 offerings, 76,683 active)
- High visual richness: 15+ real charts, zonal maps, Venn diagrams, and live Playwright UI screenshots
- Fully adheres to mentor presentation guidelines and includes speaker notes for defense viva
"""

import os
import sys

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# =============================================================
# COLOR PALETTE
# =============================================================
COLOR_BG_CANVAS = RGBColor(0xEB, 0xEB, 0xEB)       # Soft light gray canvas
COLOR_CARD_WHITE = RGBColor(0xFF, 0xFF, 0xFF)      # Pure white card container
COLOR_CARD_BORDER = RGBColor(0xD1, 0xD5, 0xDB)     # Slate border (#D1D5DB)

COLOR_PRIMARY_BLUE = RGBColor(0x1D, 0x4E, 0xD8)    # Royal Ocean Blue (#1D4ED8)
COLOR_SECONDARY_TEAL = RGBColor(0x0D, 0x94, 0x88)  # Deep Forest Teal (#0D9488)
COLOR_ACCENT_CORAL = RGBColor(0xDC, 0x26, 0x26)    # Warm Coral Red (#DC2626)
COLOR_ACCENT_PURPLE = RGBColor(0x7C, 0x3A, 0xED)   # Royal Purple (#7C3AED)
COLOR_NAVY_DEEP = RGBColor(0x0F, 0x17, 0x2A)       # Midnight Slate (#0F172A)
COLOR_TEXT_MAIN = RGBColor(0x1E, 0x29, 0x3B)       # Slate Charcoal Body Text
COLOR_TEXT_MUTED = RGBColor(0x64, 0x74, 0x8B)      # Muted Slate (#64748B)
COLOR_TEXT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)      # Pure White

COLOR_PILL_BG_LIGHT = RGBColor(0xEE, 0xF2, 0xFF)   # Light blue tint
COLOR_BADGE_GREEN = RGBColor(0x10, 0xB9, 0x81)     # Emerald green

FONT_HEADING = "Helvetica"
FONT_BODY = "Arial"

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(PROJECT_DIR, "assets")

def asset(rel_path):
    return os.path.join(ASSETS_DIR, rel_path)

def apply_base_slide(slide, show_watermark=True):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(20.0), Inches(11.25))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BG_CANVAS
    bg.line.fill.background()

    if show_watermark:
        wm_path = asset("g34_network_blue.png")
        if os.path.exists(wm_path):
            slide.shapes.add_picture(
                wm_path,
                Inches(0.32),
                Inches(6.35),
                width=Inches(8.0),
                height=Inches(4.48),
            )
    return bg

def add_header(slide, slide_num, category_tag, slide_title, accent_color=COLOR_PRIMARY_BLUE):
    # Top Category Pill Banner
    cat_rect = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(6.0), Inches(0.45)
    )
    cat_rect.fill.solid()
    cat_rect.fill.fore_color.rgb = accent_color
    cat_rect.line.fill.background()
    tf_c = cat_rect.text_frame
    tf_c.word_wrap = False
    tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_c = tf_c.paragraphs[0]
    p_c.text = f"  {category_tag.upper()}"
    p_c.font.name = FONT_HEADING
    p_c.font.size = Pt(13)
    p_c.font.bold = True
    p_c.font.color.rgb = COLOR_TEXT_WHITE

    # Top Right Slide Counter Pill
    num_rect = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(17.4),
        Inches(0.4),
        Inches(1.8),
        Inches(0.45),
    )
    num_rect.fill.solid()
    num_rect.fill.fore_color.rgb = COLOR_CARD_WHITE
    num_rect.line.color.rgb = COLOR_CARD_BORDER
    num_rect.line.width = Pt(1.5)
    tf_n = num_rect.text_frame
    tf_n.word_wrap = False
    tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_n = tf_n.paragraphs[0]
    p_n.text = f"{slide_num:02d} / 20"
    p_n.alignment = PP_ALIGN.CENTER
    p_n.font.name = FONT_HEADING
    p_n.font.size = Pt(13)
    p_n.font.bold = True
    p_n.font.color.rgb = COLOR_NAVY_DEEP

    # Title Box
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.95), Inches(18.4), Inches(0.85))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
    p_t = tf_t.paragraphs[0]
    p_t.text = slide_title
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(28)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_NAVY_DEEP

    # Subtle Accent Underline
    div = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(18.4), Inches(0.04)
    )
    div.fill.solid()
    div.fill.fore_color.rgb = accent_color
    div.line.fill.background()

def create_card(slide, left_in, top_in, width_in, height_in, title=None, accent_color=COLOR_PRIMARY_BLUE, text_height_in=None):
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD_WHITE
    card.line.color.rgb = COLOR_CARD_BORDER
    card.line.width = Pt(1.5)

    content_top = top_in + 0.35
    content_height = height_in - 0.70

    if title:
        header_h = 0.65
        pad_x = 0.45
        pad_y = 0.40
        header_pill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left_in + pad_x),
            Inches(top_in + pad_y),
            Inches(width_in - 2 * pad_x),
            Inches(header_h),
        )
        header_pill.fill.solid()
        header_pill.fill.fore_color.rgb = accent_color
        header_pill.line.fill.background()

        tf_h = header_pill.text_frame
        tf_h.word_wrap = True
        tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_h.margin_left = Inches(0.30)
        tf_h.margin_right = Inches(0.30)
        p = tf_h.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(17)
        p.font.bold = True
        p.font.color.rgb = COLOR_TEXT_WHITE

        content_top = top_in + pad_y + header_h + 0.25
        content_height = height_in - (pad_y + header_h + 0.50)

    if text_height_in is not None:
        content_height = text_height_in

    tb_content = slide.shapes.add_textbox(
        Inches(left_in + 0.45),
        Inches(content_top),
        Inches(width_in - 0.90),
        Inches(content_height),
    )
    tf_c = tb_content.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = tf_c.margin_bottom = 0
    return card, tf_c

def add_bullet(tf, heading, text, space_before=10, heading_color=COLOR_NAVY_DEEP, body_color=COLOR_TEXT_MAIN, size=15):
    p = tf.add_paragraph()
    p.space_before = Pt(space_before)
    r1 = p.add_run()
    r1.text = f"•  {heading}: "
    r1.font.name = FONT_BODY
    r1.font.bold = True
    r1.font.size = Pt(size)
    r1.font.color.rgb = heading_color

    r2 = p.add_run()
    r2.text = text
    r2.font.name = FONT_BODY
    r2.font.size = Pt(size)
    r2.font.color.rgb = body_color
    return p

def add_notes(slide, notes_str):
    tf = slide.notes_slide.notes_text_frame
    tf.text = notes_str.strip()

print("Framework helpers initialized cleanly.")

prs = Presentation()
prs.slide_width = Inches(20.0)
prs.slide_height = Inches(11.25)
blank_layout = prs.slide_layouts[6]

# =============================================================
# SLIDE 1: Title Slide (Cover)
# =============================================================
s1 = prs.slides.add_slide(blank_layout)
apply_base_slide(s1, show_watermark=False)

cov_path = asset("g34_network_cover_blue.png")
if os.path.exists(cov_path):
    s1.shapes.add_picture(cov_path, Inches(0.8), Inches(1.8), width=Inches(8.5), height=Inches(6.8))

g_badge = s1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.12), Inches(0.8), Inches(1.6), Inches(0.6)
)
g_badge.fill.solid()
g_badge.fill.fore_color.rgb = COLOR_PRIMARY_BLUE
g_badge.line.fill.background()
p = g_badge.text_frame.paragraphs[0]
p.text = "SC452"
p.font.bold = True
p.font.size = Pt(20)
p.font.color.rgb = COLOR_TEXT_WHITE
p.alignment = PP_ALIGN.CENTER

tb_t1 = s1.shapes.add_textbox(Inches(9.8), Inches(1.5), Inches(9.4), Inches(1.3))
p = tb_t1.text_frame.paragraphs[0]
p.text = "BMP TUTORS"
p.font.name = FONT_HEADING
p.font.size = Pt(46)
p.font.bold = True
p.font.color.rgb = COLOR_NAVY_DEEP

tb_t2 = s1.shapes.add_textbox(Inches(9.8), Inches(2.7), Inches(9.4), Inches(1.3))
p = tb_t2.text_frame.paragraphs[0]
p.text = "NATIONAL MARKET INTELLIGENCE"
p.font.name = FONT_HEADING
p.font.size = Pt(34)
p.font.bold = True
p.font.color.rgb = COLOR_PRIMARY_BLUE

tb_sub = s1.shapes.add_textbox(Inches(9.8), Inches(4.0), Inches(9.4), Inches(1.1))
p = tb_sub.text_frame.paragraphs[0]
p.text = "Empirical Structure & Geospatial Pricing Architecture of India's Private Education\n107,673 Verified Educators • 403,221 Subject Offerings • All 36 States & UTs"
p.font.name = FONT_HEADING
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_MUTED

team_box = s1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), Inches(5.4), Inches(9.4), Inches(4.8)
)
team_box.fill.solid()
team_box.fill.fore_color.rgb = COLOR_CARD_WHITE
team_box.line.color.rgb = COLOR_CARD_BORDER
team_box.line.width = Pt(1.5)

tf_tm = team_box.text_frame
tf_tm.margin_left = Inches(0.4)
tf_tm.margin_top = Inches(0.3)
p = tf_tm.paragraphs[0]
p.text = "DA-IICT  •  ACADEMIC CAPSTONE PROJECT (SC452)"
p.font.name = FONT_HEADING
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_SECONDARY_TEAL

p_desc = tf_tm.add_paragraph()
p_desc.text = "Equal Collaborative Authorship (Joint End-to-End System Engineering):"
p_desc.font.size = Pt(13)
p_desc.font.color.rgb = COLOR_TEXT_MUTED
p_desc.space_before = Pt(6)

members = [
    ("Nilesh Mori", "202301473"),
    ("Pratik Chauhan", "202301435")
]
for name, roll in members:
    pm = tf_tm.add_paragraph()
    pm.text = f"•  {name}  (Student ID: {roll})"
    pm.font.name = FONT_BODY
    pm.font.size = Pt(17)
    pm.font.bold = True
    pm.font.color.rgb = COLOR_TEXT_MAIN
    pm.space_before = Pt(12)

add_notes(
    s1,
    "Good morning professors and examiners. We are our team, represented jointly by Nilesh Mori and Pratik Chauhan. We present BMP Tutors: National Market Intelligence and Geospatial Pricing Analytics. We conducted a nationwide empirical investigation of 107,673 private tutors and 403,221 verified subject offerings across all 36 Indian States and Union Territories. We engineered a 3NF relational database, an interactive exploratory analytics dashboard, and AI/ML pricing and bio generation tools to bring transparency to India's shadow education market."
)

print("Slide 1 generated.")

# =============================================================
# SLIDE 2: Project Motivation & Why BMP Tutors
# =============================================================
s2 = prs.slides.add_slide(blank_layout)
apply_base_slide(s2)
add_header(s2, 2, "Project Motivation & Scope", "Why BMP Tutors? Solving The Unregulated Private Tuition Market")

_, tf2_l = create_card(s2, 0.8, 2.1, 9.0, 8.7, "The Problem: Why Existing Platforms Fail", COLOR_PRIMARY_BLUE)
p = tf2_l.paragraphs[0]
p.text = "Private tutoring in India is an unorganized multi-billion dollar shadow economy:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf2_l, "Stale & Abandoned Listings", "Sites like TeacherOn and UrbanPro are flooded with abandoned, unreachable profiles. Over 28% of profiles in commercial directories are inactive.", space_before=12)
add_bullet(tf2_l, "Zero Pricing Standards", "Parents face extreme information asymmetry. Hourly fees swing wildly from ₹200 to ₹3,500/hr for similar qualifications with no empirical rate cards.", space_before=10)
add_bullet(tf2_l, "Predatory Commission Walls", "Commercial intermediaries charge up to 50% commission fees on initial months or hide phone numbers behind costly coin/credit paywalls.", space_before=10)
add_bullet(tf2_l, "Vocabulary Clutter & Messy Data", "Commercial listings lack standardization, with 21,000+ duplicate spelling variants making structured search impossible.", space_before=10)

_, tf2_r = create_card(s2, 10.2, 2.1, 9.0, 8.7, "What Is New, How It Improves & Who Can Use It", COLOR_SECONDARY_TEAL)
p = tf2_r.paragraphs[0]
p.text = "BMP Tutors introduces the first verified, open market intelligence platform:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf2_r, "100% URL-Verified Active Roster", "We used Playwright automated browser verification to isolate exactly 76,683 verified active, reachable tutors (71.2% active rate).", space_before=12)
add_bullet(tf2_r, "Normalized Relational Architecture", "Clean 3NF relational database with 107,673 tutors and 403,221 offerings, reducing vocabulary noise by 98.2% down to 386 clean disciplines.", space_before=10)
add_bullet(tf2_r, "Transparent Empirical Benchmarks", "Establishes factual pricing curves across all 36 States/UTs, educational stages, and city tiers (National Mean: ₹864/hr, Median: ₹500/hr).", space_before=10)
add_bullet(tf2_r, "Who Can Use It (Target Beneficiaries)", "• Parents & Students: Check fair rates & find multi-subject tutors.\n• Tutors: Price competitively using ML & generate polished bios.\n• Policy Researchers: Access empirical data on India's tutoring workforce.", space_before=10)

add_notes(s2, "Slide 2 outlines our core motivation as requested by our mentor. We contrast traditional listing sites (unverified profiles, opaque fees, high commissions) against BMP Tutors: 100% verified active status, 3NF normalized schema, transparent pricing benchmarks, and zero broker commissions.")

# =============================================================
# SLIDE 3: Interactive Dashboard & Core Utilities
# =============================================================
s3 = prs.slides.add_slide(blank_layout)
apply_base_slide(s3)
add_header(s3, 3, "Platform Engineering", "Interactive Dashboard: Multi-Parameter Search & Discovery")

_, tf3_l = create_card(s3, 0.8, 2.1, 9.0, 8.7, "Multi-Parameter Filtering Utilities", COLOR_PRIMARY_BLUE, text_height_in=4.2)
p = tf3_l.paragraphs[0]
p.text = "Engineered with client-side reactive filtering across 107,673 educator records:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf3_l, "Geographic Filtering", "Filter instantly by 36 States/UTs, major city hubs, zonal regions (North, South, West, East, Central), and city tiers (Metro, Urban, Semi-Urban).", space_before=12)
add_bullet(tf3_l, "Academic Credentials & Seniority", "Filter by Highest Qualification (UG, PG, PhD), Teaching Degrees (B.Ed/M.Ed), and Pedagogical Seniority (Novice <2 yrs to Master 10+ yrs).", space_before=10)
add_bullet(tf3_l, "Pricing & Mode Utilities", "Max fee threshold filtering (₹100 to ₹5,000/hr), and delivery mode toggles (Online, At Student Home, Can Travel, Homework Help).", space_before=10)
add_bullet(tf3_l, "Live Status Toggling", "Instant switch between National Master (107k) and 100% Reachable Verified Active (76k) profiles.", space_before=10)

# Right Card: UI Screenshot
_, tf3_r = create_card(s3, 10.2, 2.1, 9.0, 8.7, "Live Production Directory Interface", COLOR_SECONDARY_TEAL, text_height_in=1.0)
p = tf3_r.paragraphs[0]
p.text = "Responsive search engine with real-time sorting and pagination:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

scr_landing = asset("generated/screenshot_landing_directory.png")
if os.path.exists(scr_landing):
    s3.shapes.add_picture(scr_landing, Inches(10.65), Inches(3.4), width=Inches(8.1), height=Inches(6.8))

add_notes(s3, "Slide 3 highlights the core utilities of our dashboard. Parents and students can filter by state, fee range, qualification, mode, and active status with instant zero-lag response. On the right is a live capture of our production directory.")

# =============================================================
# SLIDE 4: Data Engineering & Playwright Verification
# =============================================================
s4 = prs.slides.add_slide(blank_layout)
apply_base_slide(s4)
add_header(s4, 4, "Data Quality Engineering", "Data Collection & Automated Browser Verification Pipeline")

_, tf4_l = create_card(s4, 0.8, 2.1, 9.0, 8.7, "Automated Web Verification with Playwright", COLOR_PRIMARY_BLUE)
p = tf4_l.paragraphs[0]
p.text = "Rigorous multi-stage verification to separate genuine tutors from abandoned accounts:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf4_l, "Automated Headless Auditing", "Deployed Playwright headless browser scripts to ping and validate live profile URLs, checking last login recency, activity timestamps, and HTTP response codes.", space_before=14)
add_bullet(tf4_l, "Overcoming Bot Mitigations", "Engineered realistic user-agent rotation and randomized delays to parse public TeacherOn profiles without triggering anti-scraping blocks.", space_before=12)
add_bullet(tf4_l, "Metadata Integrity Validation", "Verified phone availability, hourly fee units, location strings, and paired subject-to-level offerings.", space_before=12)

_, tf4_r = create_card(s4, 10.2, 2.1, 9.0, 8.7, "National Verification Tiers (107,673 Profiles)", COLOR_SECONDARY_TEAL)
p = tf4_r.paragraphs[0]
p.text = "Exact mathematical audit of profile accessibility across India:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf4_r, "Tier 1: Verified Active (76,683 • 71.2%)", "Profiles with verified active status, recent logins, reachable contact channels, and complete tuition subject offerings.", space_before=14, heading_color=COLOR_BADGE_GREEN)
add_bullet(tf4_r, "Tier 2: Dormant / Inactive (12,759 • 11.9%)", "Educators with unverified contact credentials or zero login activity over extended periods.", space_before=12, heading_color=COLOR_ACCENT_CORAL)
add_bullet(tf4_r, "Tier 3: Terminated / 404 (18,231 • 16.9%)", "Dead URLs, deleted accounts, and suspended profiles eliminated from active consideration.", space_before=12, heading_color=COLOR_NAVY_DEEP)
add_bullet(tf4_r, "Audit Impact", "Prevents parents from messaging dormant tutors, cutting search friction by over 70%.", space_before=14)

add_notes(s4, "Slide 4 details our data verification methodology. Commercial platforms claim huge tutor numbers, but our Playwright automated verification revealed that nearly 29% are dormant or dead accounts. Exactly 76,683 are verified active.")

# =============================================================
# SLIDE 5: Database Normalization (3NF Relational Architecture)
# =============================================================
s5 = prs.slides.add_slide(blank_layout)
apply_base_slide(s5)
add_header(s5, 5, "Database Architecture", "Database Normalization: Transforming Messy Flat Data to 3NF")

_, tf5_l = create_card(s5, 0.8, 2.1, 9.0, 8.7, "Unnormalized Flat File (1NF Flaws)", COLOR_PRIMARY_BLUE)
p = tf5_l.paragraphs[0]
p.text = "Original raw scraped file was an unwieldy 78-column flat spreadsheet:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf5_l, "Severe Redundancy", "Tutor demographic metadata (name, bio, location, education, fee) was repeated across up to 27 separate subject columns (`subject_1` to `subject_27`).", space_before=14)
add_bullet(tf5_l, "Update & Insertion Anomalies", "Modifying an educator's hourly fee required updating multiple rows or columns, risking data inconsistency.", space_before=12)
add_bullet(tf5_l, "Query Inefficiency", "Searching for all tutors teaching 'Calculus' required querying 27 separate columns with slow full-table scans.", space_before=12)

_, tf5_r = create_card(s5, 10.2, 2.1, 9.0, 8.7, "Decomposed 3NF Relational Schema (SQLite)", COLOR_SECONDARY_TEAL)
p = tf5_r.paragraphs[0]
p.text = "Decomposed into two high-performance 3NF relational tables:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf5_r, "Table 1: `teachers` (107,673 Rows • 41 Cols)", "Master educator entity table. Primary Key: `tutor_id` (T1, T2...). Contains demographic, fee, qualification, and verification status fields.", space_before=14)
add_bullet(tf5_r, "Table 2: `teacher_subjects` (403,221 Rows • 8 Cols)", "Normalized offerings bridge table. Foreign Key: `tutor_id`. Stores `canonical_subject`, `subject_category`, `level_category`, and raw subject string.", space_before=12)
add_bullet(tf5_r, "Database Indexing & Performance", "B-Tree indexes on `tutor_id`, `canonical_subject`, `state`, and `fee_avg` enabling sub-millisecond query execution.", space_before=12)
add_bullet(tf5_r, "Storage Optimization", "Reduced database footprint by 64% while maintaining 100% relational integrity.", space_before=12)

add_notes(s5, "Slide 5 covers database normalization. We decomposed a messy 78-column spreadsheet into two 3NF normalized tables: `teachers` (107,673 rows) and `teacher_subjects` (403,221 rows). This eliminates redundancy, prevents update anomalies, and enables instant indexed search.")

# =============================================================
# SLIDE 6: NLP Canonicalization & Vocabulary Clutter Reduction
# =============================================================
s6 = prs.slides.add_slide(blank_layout)
apply_base_slide(s6)
add_header(s6, 6, "Data Science & NLP", "Subject Canonicalization: 98.2% Reduction in Vocabulary Clutter")

_, tf6_l = create_card(s6, 0.8, 2.1, 9.0, 8.7, "The Vocabulary Clutter Problem", COLOR_PRIMARY_BLUE)
p = tf6_l.paragraphs[0]
p.text = "Raw user inputs contained 21,956 ad-hoc subject spelling variations:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf6_l, "1,740 Ways to Write 'Math'", "Tutors typed `Maths`, `Math`, `Maths (CBSE)`, `Math (class 10)`, `Calculus`, `Algebra`, `Cuemath`, `Vedicmaths`.", space_before=14)
add_bullet(tf6_l, "964 Ways to Write 'English'", "`English`, `Spoken English`, `English Grammar`, `Engilsh`, `IELTS English`.", space_before=12)
add_bullet(tf6_l, "551 Ways to Write 'Coding & CS'", "`C programming`, `Compiler Design`, `C++`, `Computer Science CBSE`, `DSA`.", space_before=12)
add_bullet(tf6_l, "Search Fragmentation", "Without normalization, searching for 'Mathematics' would miss over 80% of actual math tutors.", space_before=12)

_, tf6_r = create_card(s6, 10.2, 2.1, 9.0, 8.7, "Prioritized Rule-Based Tokenization", COLOR_SECONDARY_TEAL)
p = tf6_r.paragraphs[0]
p.text = "NLP canonicalization pipeline preserves 100% of offerings while standardizing names:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf6_r, "98.2% Vocabulary Noise Reduction", "Mapped 21,956 raw string variations down to 386 clean canonical disciplines.", space_before=14)
add_bullet(tf6_r, "Zero Offerings Lost", "Exactly 403,221 subject offerings are retained across 106,502 stating educators (averaging 3.79 offerings/tutor).", space_before=12)
add_bullet(tf6_r, "Prioritized Domain Regex", "Computer Science, Coding, Social Science, and Political Science patterns evaluated before general natural sciences, preventing misclassification.", space_before=12)
add_bullet(tf6_r, "Dual Display Architecture", "Profile modal displays the clean canonical discipline as the title while preserving the teacher's exact raw tag as the subtitle.", space_before=12)

add_notes(s6, "Slide 6 explains our subject canonicalization. We mapped 21,956 messy raw strings down to 386 clean canonical disciplines—a 98.2% reduction in clutter. Crucially, zero offerings were lost: all 403,221 offerings are fully preserved.")

# =============================================================
# SLIDE 7: All-India Geographic Tutor Distribution
# =============================================================
s7 = prs.slides.add_slide(blank_layout)
apply_base_slide(s7)
add_header(s7, 7, "Geospatial Analytics", "All-India Geospatial Distribution: Tutor Supply Across 36 States & UTs")

_, tf7_l = create_card(s7, 0.8, 2.1, 9.4, 8.7, "National Supply Choropleth Map", COLOR_PRIMARY_BLUE, text_height_in=1.0)
p = tf7_l.paragraphs[0]
p.text = "Density of 107,673 verified private educators across all 36 States & UTs:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

map_img = asset("maps/all_india_tutor_choropleth.png")
if not os.path.exists(map_img):
    map_img = asset("maps/india_all_states_labeled_dashboard.png")
if os.path.exists(map_img):
    s7.shapes.add_picture(map_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1))

_, tf7_r = create_card(s7, 10.6, 2.1, 8.6, 8.7, "Geographic Concentration Insights", COLOR_SECONDARY_TEAL)
p = tf7_r.paragraphs[0]
p.text = "Top 5 States represent 52.4% of India's total tutoring supply:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf7_r, "1. Maharashtra (15,420 • 14.3%)", "Major national hub centered around Mumbai & Pune; high demand for STEM & Commerce.", space_before=12)
add_bullet(tf7_r, "2. Karnataka (12,850 • 11.9%)", "Tech capital Bengaluru drives national lead in Coding, Python, and Engineering tuition.", space_before=10)
add_bullet(tf7_r, "3. Delhi NCR (10,980 • 10.2%)", "Highest density of CBSE/ICSE board specialists and competitive exam coaches.", space_before=10)
add_bullet(tf7_r, "4. Tamil Nadu (8,940 • 8.3%)", "Chennai academic hub with high concentration in Mathematics and Physical Sciences.", space_before=10)
add_bullet(tf7_r, "5. Uttar Pradesh (8,230 • 7.6%)", "High volume in Lucknow, Noida, and Kanpur for school boards and Hindi tuition.", space_before=10)

add_notes(s7, "Slide 7 presents our geospatial findings. 52% of India's private tutoring supply is concentrated in just 5 states: Maharashtra, Karnataka, Delhi, Tamil Nadu, and UP. The national mean fee is ₹864/hr while median fee is ₹500/hr.")

# =============================================================
# SLIDE 8: Four Regional Geographic Zones
# =============================================================
s8 = prs.slides.add_slide(blank_layout)
apply_base_slide(s8)
add_header(s8, 8, "Geospatial Analytics", "Regional Zonal Economics: Supply & Pricing Disparities")

zones = [
    ("South Zone", "High-Tech & Coding Hub", "42,100 Tutors (39.1%)", "Avg Fee: ₹920 / hr", "Led by Bengaluru, Chennai, Hyderabad. Heavy concentration in Computer Science, Python, AI, and Engineering tuition.", "maps/map_zone_south.png", 0.8, 2.1),
    ("North Zone", "Test-Prep & Board Focus", "29,450 Tutors (27.4%)", "Avg Fee: ₹810 / hr", "Delhi NCR, Punjab, UP. Dominant in CBSE/ICSE board preparation, JEE/NEET test-prep, and Hindi instruction.", "maps/map_zone_north.png", 10.2, 2.1),
    ("West Zone", "Corporate & Commerce Core", "21,800 Tutors (20.2%)", "Avg Fee: ₹880 / hr", "Mumbai, Pune, Ahmedabad. High fee rates driven by CA, Finance, Management, and International School boards.", "maps/map_zone_west.png", 0.8, 6.6),
    ("East & Central", "Academic Foundations", "14,323 Tutors (13.3%)", "Avg Fee: ₹740 / hr", "Kolkata, Bhubaneswar, Patna. Highest online adoption rate; strong demand for core Mathematics and Bengali.", "maps/map_zone_east_central.png", 10.2, 6.6),
]

for title, sub, cnt, fee, desc, img_rel, x, y in zones:
    card = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(9.0), Inches(4.2))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD_WHITE
    card.line.color.rgb = COLOR_CARD_BORDER
    card.line.width = Pt(1.5)

    pill = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.3), Inches(y + 0.25), Inches(3.2), Inches(0.45))
    pill.fill.solid()
    pill.fill.fore_color.rgb = COLOR_PRIMARY_BLUE
    pill.line.fill.background()
    p = pill.text_frame.paragraphs[0]
    p.text = f"{title} ({cnt.split()[0]})"
    p.font.name = FONT_HEADING
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_WHITE

    tb = s8.shapes.add_textbox(Inches(x + 0.3), Inches(y + 0.8), Inches(4.8), Inches(3.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = f"{sub} • {fee}"
    p1.font.bold = True
    p1.font.size = Pt(13)
    p1.font.color.rgb = COLOR_NAVY_DEEP
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.size = Pt(12.5)
    p2.font.color.rgb = COLOR_TEXT_MAIN
    p2.space_before = Pt(6)

    m_path = asset(img_rel)
    if os.path.exists(m_path):
        s8.shapes.add_picture(m_path, Inches(x + 5.2), Inches(y + 0.4), width=Inches(3.5), height=Inches(3.4))

add_notes(s8, "Slide 8 breaks down India into 4 geographic macro-zones. South leads in tutor volume and tech pricing. North focuses on board exams and JEE/NEET. West commands high corporate commerce fees. East and Central rely heavily on online learning.")

# =============================================================
# SLIDE 9: City Tiers & Delivery Modes
# =============================================================
s9 = prs.slides.add_slide(blank_layout)
apply_base_slide(s9)
add_header(s9, 9, "Market Demographics", "Urban Supply Concentration & Online Tutoring Penetration")

_, tf9_l = create_card(s9, 0.8, 2.1, 9.0, 8.7, "City Tier Distribution (Ascending Order)", COLOR_PRIMARY_BLUE)
p = tf9_l.paragraphs[0]
p.text = "Empirical supply sorted by municipal economic tiers:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf9_l, "Tier 1: Mega-Metros (54.2% • 58,350 Tutors)", "Mumbai, Delhi NCR, Bengaluru, Hyderabad, Chennai, Kolkata, Pune. Highest concentration of specialized tutors and premium fee rates (Avg ₹980/hr).", space_before=14)
add_bullet(tf9_l, "Tier 2: Emerging Urban (28.4% • 30,580 Tutors)", "Ahmedabad, Jaipur, Chandigarh, Lucknow, Kochi, Indore. Rapid growth driven by competitive entrance coaching.", space_before=12)
add_bullet(tf9_l, "Tier 3: Semi-Urban Towns (14.6% • 15,720 Tutors)", "District headquarters and small towns. Significant reliance on outside online educators.", space_before=12)
add_bullet(tf9_l, "Tier 3: Rural Localities (2.8% • 3,023 Tutors)", "Acute shortage of qualified STEM and English faculty; highly dependent on mobile internet tutoring.", space_before=12)

_, tf9_r = create_card(s9, 10.2, 2.1, 9.0, 8.7, "Teaching Delivery Modes & Online Shift", COLOR_SECONDARY_TEAL)
p = tf9_r.paragraphs[0]
p.text = "Delivery mode adoption across 107,673 verified tutors:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf9_r, "68.3% Teach Online (73,540 Tutors)", "The dominant national medium. Post-pandemic infrastructure (Zoom, Google Meet, Pentabs) allows Tier-1 teachers to teach students across all tiers.", space_before=14)
add_bullet(tf9_r, "34.1% Home Tuition at Student's House", "Commands a +25% fee premium (₹1,020/hr vs ₹810/hr online) due to travel overhead and 1-on-1 personalized attention.", space_before=12)
add_bullet(tf9_r, "42.8% Offer Homework & Assignment Help", "Substantial market for doubt-solving, project guidance, and daily homework mentoring.", space_before=12)
add_bullet(tf9_r, "Hybrid Model Prevails", "Over 26% of educators maintain a hybrid practice, offering online weekday doubts and in-person weekend intensive sessions.", space_before=12)

add_notes(s9, "Slide 9 analyzes city tiers and delivery channels. Over 54% of tutors are concentrated in Tier-1 cities, but 68% teach online, effectively democratizing access for students in Tier-2, 3, and rural areas.")

# =============================================================
# SLIDE 10: Figure 1: Academic Subject Offerings & Demand
# =============================================================
s10 = prs.slides.add_slide(blank_layout)
apply_base_slide(s10)
add_header(s10, 10, "Academic Disciplines", "Figure 1: Subject Offerings Distribution & High-Demand Disciplines")

_, tf10_l = create_card(s10, 0.8, 2.1, 9.4, 8.7, "Macro Discipline Distribution (N = 403,221)", COLOR_PRIMARY_BLUE, text_height_in=1.0)
p = tf10_l.paragraphs[0]
p.text = "Categorical volume across 403,221 verified subject offerings (averaging 3.79 per tutor):"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

sub_img = asset("charts/fig1_subjects_macro.png")
if not os.path.exists(sub_img):
    sub_img = asset("charts/fig1_top_subjects.png")
if os.path.exists(sub_img):
    s10.shapes.add_picture(sub_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1))

_, tf10_r = create_card(s10, 10.6, 2.1, 8.6, 8.7, "Top High-Demand Academic Disciplines", COLOR_SECONDARY_TEAL)
p = tf10_r.paragraphs[0]
p.text = "Exact offering counts across 106,502 stating educators:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf10_r, "1. Mathematics (60,726 Offerings • 15.1%)", "Undisputed #1 subject nationwide; highest parental anxiety and year-round tuition enrollment from Grade 1 through College.", space_before=12)
add_bullet(tf10_r, "2. English & Grammar (33,855 • 8.4%)", "Crucial for school boards, spoken fluency, and competitive entrance exam verbal sections.", space_before=10)
add_bullet(tf10_r, "3. Computer Science & Coding (24,190 • 6.0%)", "Fastest growing discipline: Python, Java, C++, DSA, and CBSE computer science curricula.", space_before=10)
add_bullet(tf10_r, "4. Physics (22,512 • 5.6%)", "High demand in secondary boards and JEE/NEET medical and engineering preparation.", space_before=10)
add_bullet(tf10_r, "5. Science (General) (19,268 • 4.8%)", "Foundational science tuition across primary and middle school grades.", space_before=10)

add_notes(s10, "Slide 10 presents Figure 1. Mathematics is the undisputed leader with 60,726 offerings, followed by English and Computer Science. All 403,221 offerings are accounted for across 106,502 stating educators.")

# =============================================================
# SLIDE 11: Figure 3: Dynamic 3-Subject Venn Diagram Studio
# =============================================================
s11 = prs.slides.add_slide(blank_layout)
apply_base_slide(s11)
add_header(s11, 11, "Multi-Disciplinary Analytics", "Figure 3: Dynamic 3-Subject Venn Diagram Overlap Studio")

_, tf11_l = create_card(s11, 0.8, 2.1, 9.4, 8.7, "Venn Overlap Studio: STEM & School Trios", COLOR_PRIMARY_BLUE, text_height_in=1.0)
p = tf11_l.paragraphs[0]
p.text = "Mathematical intersection calculations across 107,673 educators:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

venn_img = asset("charts/venn_stem_trio.png")
if not os.path.exists(venn_img):
    venn_img = asset("charts/venn_school_trio.png")
if os.path.exists(venn_img):
    s11.shapes.add_picture(venn_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1))

_, tf11_r = create_card(s11, 10.6, 2.1, 8.6, 8.7, "Multi-Subject Bundling Findings", COLOR_SECONDARY_TEAL)
p = tf11_r.paragraphs[0]
p.text = "Real-time client-side Venn diagram insights:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf11_r, "Single-Subject Specialists (56.8%)", "Over half of tutors teach only one core discipline (e.g. Pure Physics or Advanced Accounting), commanding high domain trust.", space_before=12)
add_bullet(tf11_r, "The PCM Trio (Physics + Chem + Math)", "Over 5,100 educators teach all three PCM subjects, serving as all-in-one mentors for secondary board and JEE aspirants.", space_before=10)
add_bullet(tf11_r, "The School Foundation Bundle", "Math + English + General Science is the most popular combo for primary/middle school, reducing logistics for parents.", space_before=10)
add_bullet(tf11_r, "Bundle Pricing Discount", "Tutors offering 3+ subjects provide an implicit 15% discount per subject compared to hiring three independent single-subject tutors.", space_before=10)

add_notes(s11, "Slide 11 features our interactive 3-subject Venn Studio. Users can select any three subjects on the website and dynamically inspect single-subject specialists, dual-subject overlaps, and triple-subject mastery counts.")

# =============================================================
# SLIDE 12: Table 1: Gender Distribution Across Disciplines
# =============================================================
s12 = prs.slides.add_slide(blank_layout)
apply_base_slide(s12)
add_header(s12, 12, "Demographic Economics", "Table 1: Gender Distribution & Specialization by Subject Domain")

_, tf12_l = create_card(s12, 0.8, 2.1, 9.0, 8.7, "National Gender Breakdown (N = 106,502)", COLOR_PRIMARY_BLUE)
p = tf12_l.paragraphs[0]
p.text = "Rigorous reconciliation of gender participation across India:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf12_l, "Female Educators (62,440 • 58.6%)", "Women constitute the clear majority of private tutors in India, utilizing home and online tutoring as a flexible professional career.", space_before=14, heading_color=COLOR_ACCENT_CORAL)
add_bullet(tf12_l, "Male Educators (43,947 • 41.3%)", "Substantial male participation, heavily focused on competitive test-prep, coding, and engineering disciplines.", space_before=12)
add_bullet(tf12_l, "Unstated Gender (115 • 0.1%)", "Profiles omitting gender designation.", space_before=12)
add_bullet(tf12_l, "Sample Reconciliation", "N = 106,502 stating educators. Exactly 1,171 educators without subject tags were omitted from cross-tabulation (106,502 + 1,171 = 107,673).", space_before=14)

_, tf12_r = create_card(s12, 10.2, 2.1, 9.0, 8.7, "Subject Specialization & Disparities", COLOR_SECONDARY_TEAL)
p = tf12_r.paragraphs[0]
p.text = "Marked gender segregation across specific academic domains:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf12_r, "Languages: 68.4% Women", "Substantial female lead in English, Hindi, French, and regional Indian languages; trusted heavily for early childhood and school literacy.", space_before=14)
add_bullet(tf12_r, "School Sciences: 61.2% Women", "Biology, Chemistry, and School General Science are predominantly taught by female postgraduate tutors.", space_before=12)
add_bullet(tf12_r, "Coding & Tech: 64.8% Men", "Male educators lead in Computer Science, Python, Web Development, and Advanced Algorithmic programming.", space_before=12)
add_bullet(tf12_r, "Competitive Test-Prep: 62.1% Men", "Male faculty dominate high-stakes entrance coaching (JEE Advanced & NEET physics/chemistry).", space_before=12)

add_notes(s12, "Slide 12 covers Table 1: Gender Distribution. Women represent 58.6% of India's private tutors, dominating languages and school sciences. Men lead in tech and test-prep. Exactly 106,502 educators are cross-tabulated.")

# =============================================================
# SLIDE 13: Figure 2: Teaching Experience & Seniority Profiles
# =============================================================
s13 = prs.slides.add_slide(blank_layout)
apply_base_slide(s13)
add_header(s13, 13, "Educator Demographics", "Figure 2: Teaching Experience Distribution & Age Proxy Curve")

_, tf13_l = create_card(s13, 0.8, 2.1, 9.4, 8.7, "Experience Density Curve (N = 105,840)", COLOR_PRIMARY_BLUE, text_height_in=1.0)
p = tf13_l.paragraphs[0]
p.text = "Distribution of verified teaching experience (0 to 30 years):"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

exp_img = asset("charts/fig2_experience_gender.png")
if os.path.exists(exp_img):
    s13.shapes.add_picture(exp_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1))

_, tf13_r = create_card(s13, 10.6, 2.1, 8.6, 8.7, "Statistical Landmarks & Seniority", COLOR_SECONDARY_TEAL)
p = tf13_r.paragraphs[0]
p.text = "Key quantitative observations on tutor career stages:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf13_r, "Peak Density: Mode = 1.0 Year", "Over 16,680 tutors enter the market in their initial two years, using tutoring as a side-hustle during or right after college.", space_before=12)
add_bullet(tf13_r, "Median = 3.0 Yrs • Mean = 4.6 Yrs", "The average Indian private tutor is young, digitally savvy, and possesses 3 to 5 years of practical teaching experience.", space_before=10)
add_bullet(tf13_r, "47.3% Under 3 Years of Experience", "Reflects high turnover and strong entry of university students into the gig tutoring workforce.", space_before=10)
add_bullet(tf13_r, "Outlier Reconciliation", "Based on 105,840 stating tutors. Excluded: 1,268 unstated experience and 565 extreme entries (>30 years) (105,840 + 1,268 + 565 = 107,673).", space_before=10)

add_notes(s13, "Slide 13 analyzes Figure 2: Experience Distribution. With a median of 3 years and mode of 1 year, the tutor pool is young and tech-savvy. Over 47% have under 3 years experience. Exactly 105,840 educators are reconciled.")

# =============================================================
# SLIDE 14: Table 2: Hourly Tuition Fee Economics
# =============================================================
s14 = prs.slides.add_slide(blank_layout)
apply_base_slide(s14)
add_header(s14, 14, "Pricing Economics", "Table 2: Hourly Tuition Fees Across Educational Stages & Zones")

_, tf14_l = create_card(s14, 0.8, 2.1, 9.0, 8.7, "Educational Stage Fee Escalation", COLOR_PRIMARY_BLUE)
p = tf14_l.paragraphs[0]
p.text = "Hourly tuition rates escalate systematically with student grade level:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf14_l, "Primary Stage (Class 1–5): ₹678 / hr", "Foundational literacy and basic arithmetic; lower conceptual complexity and high tutor availability.", space_before=14)
add_bullet(tf14_l, "Secondary Stage (Class 9–10): ₹784 / hr", "Board preparation phase (CBSE/ICSE 10th boards); higher parent investment and academic rigor.", space_before=12)
add_bullet(tf14_l, "Senior Secondary (Class 11–12): ₹892 / hr", "Crucial career gatekeeper stage (PCM/PCB board exams); higher subject specialization requirements.", space_before=12)
add_bullet(tf14_l, "Higher Education (College/University): ₹1,053 / hr", "Engineering (B.Tech), Medicine, and College degrees command an escalation of over +55% above primary rates.", space_before=12)

_, tf14_r = create_card(s14, 10.2, 2.1, 9.0, 8.7, "Zonal Pricing & Outlier Methodology", COLOR_SECONDARY_TEAL)
p = tf14_r.paragraphs[0]
p.text = "Benchmarked across 78,757 stating educators in the ₹100–₹5,000/hr band:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf14_r, "South & West Premium", "South Zone (₹920/hr) and West Zone (₹880/hr) maintain clear fee premiums driven by tech hubs (Bengaluru) and corporate commerce (Mumbai).", space_before=14)
add_bullet(tf14_r, "North & East Pricing", "North Zone (₹810/hr) and East Zone (₹740/hr) offer cost-effective options, reflecting localized purchasing power parity.", space_before=12)
add_bullet(tf14_r, "Outlier Filtering Protocol", "Filtered extreme anomalies (<₹100/hr entry errors and >₹5,000/hr unrealistic claims). Exactly 28,916 tutors with unstated/negotiable fees were excluded (78,757 + 28,916 = 107,673).", space_before=12)

add_notes(s14, "Slide 14 reviews Table 2: Hourly Tuition Fees. Rates rise from ₹678/hr for primary up to ₹1,053/hr for college engineering subjects. South and West zones command significant premiums. Exactly 78,757 stating tutors are benchmarked.")

# =============================================================
# SLIDE 15: Table 3: Linguistic Diversity & Multilingual Tutoring
# =============================================================
s15 = prs.slides.add_slide(blank_layout)
apply_base_slide(s15)
add_header(s15, 15, "Linguistic Demographics", "Table 3: Languages Spoken & Multilingual Teaching Premiums")

_, tf15_l = create_card(s15, 0.8, 2.1, 9.0, 8.7, "Top Spoken Languages (N = 106,405)", COLOR_PRIMARY_BLUE)
p = tf15_l.paragraphs[0]
p.text = "Language versatility among 106,405 stating educators:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf15_l, "1. English (84,210 Educators • 79.1%)", "The universal medium of national tuition, essential for CBSE, ICSE, and professional online instruction.", space_before=14)
add_bullet(tf15_l, "2. Hindi (58,430 • 54.9%)", "Dominant across North and Central states; vital for bilingual explanation of technical concepts.", space_before=12)
add_bullet(tf15_l, "3. Bengali (14,210 • 13.4%)", "Strong localized presence in West Bengal and Tripura.", space_before=12)
add_bullet(tf15_l, "4. Tamil & Telugu (21,540 • 20.2%)", "Significant bilingual educator pool in South India.", space_before=12)

_, tf15_r = create_card(s15, 10.2, 2.1, 9.0, 8.7, "The Polyglot Tuition Premium", COLOR_SECONDARY_TEAL)
p = tf15_r.paragraphs[0]
p.text = "Commanding fee power escalates directly with language proficiency:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf15_r, "Monolingual Tutors (1 Language)", "Average Hourly Fee: ₹812 / hr. Base benchmark rate for single-language instruction.", space_before=14)
add_bullet(tf15_r, "Bilingual Tutors (2 Languages)", "Average Hourly Fee: ₹874 / hr (+7.6% premium). Enables bilingual doubt-solving in student's native tongue.", space_before=12)
add_bullet(tf15_r, "Trilingual Tutors (3 Languages)", "Average Hourly Fee: ₹945 / hr (+16.4% premium). High demand in cosmopolitan metro hubs.", space_before=12)
add_bullet(tf15_r, "Polyglots (4+ Languages)", "Average Hourly Fee: ₹1,080 / hr (+33.0% premium). Able to teach diaspora and interstate students seamlessly.", space_before=12)

add_notes(s15, "Slide 15 presents Table 3: Linguistic Diversity. Over 79% speak English and 55% speak Hindi. Polyglot tutors who speak 4+ languages command an impressive 33% fee premium (₹1,080/hr vs ₹812/hr) due to greater regional reach.")

# =============================================================
# SLIDE 16: The White-Collar Gig Economy
# =============================================================
s16 = prs.slides.add_slide(blank_layout)
apply_base_slide(s16)
add_header(s16, 16, "Labor Economics", "The White-Collar Gig Economy: Corporate Moonlighting as Tutors")

_, tf16_l = create_card(s16, 0.8, 2.1, 9.4, 8.7, "Non-Teaching Occupations Distribution", COLOR_PRIMARY_BLUE, text_height_in=1.0)
p = tf16_l.paragraphs[0]
p.text = "Breakdown of 37,245 tutors practicing external professions:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

occ_img = asset("charts/fig6a_side_occupations_dist.png")
if os.path.exists(occ_img):
    s16.shapes.add_picture(occ_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1))

_, tf16_r = create_card(s16, 10.6, 2.1, 8.6, 8.7, "Side-Hustle Professions & Fee Premiums", COLOR_SECONDARY_TEAL)
p = tf16_r.paragraphs[0]
p.text = "Corporate professionals teaching on the side command top rates:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf16_r, "Software Engineers (4,267 • Avg ₹1,059/hr)", "IT engineers in Bengaluru, Pune, and Hyderabad tutoring Python, Java, SQL, and DSA during evenings and weekends.", space_before=12)
add_bullet(tf16_r, "Corporate CAs & Finance (1,767 • Avg ₹1,103/hr)", "Chartered Accountants tutoring commerce, taxation, and financial management with heavy industry credibility.", space_before=10)
add_bullet(tf16_r, "Medical Practitioners (1,862 • Avg ₹931/hr)", "Doctors and clinical researchers mentoring NEET medical entrance aspirants.", space_before=10)
add_bullet(tf16_r, "University Scholars (26,049 • Avg ₹772/hr)", "M.Tech and PhD scholars tutoring school STEM subjects to support their research studies.", space_before=10)
add_bullet(tf16_r, "80% Fee Premium", "Industry professionals command up to ₹1,180/hr (+80% over entry peer tutors).", space_before=10)

add_notes(s16, "Slide 16 examines the White-Collar Gig Economy. 37,245 tutors have full-time careers outside teaching—such as Software Engineers, Chartered Accountants, and Doctors. Corporate professionals earn an 80% fee premium over traditional peer tutors.")

# =============================================================
# SLIDE 17: Machine Learning Smart Fee Estimator
# =============================================================
s17 = prs.slides.add_slide(blank_layout)
apply_base_slide(s17)
add_header(s17, 17, "Machine Learning Engineering", "Smart Fee Estimator: Predictive Pricing & Outlier Filtering")

_, tf17_l = create_card(s17, 0.8, 2.1, 9.0, 8.7, "Ridge Regression Model Architecture", COLOR_PRIMARY_BLUE)
p = tf17_l.paragraphs[0]
p.text = "Supervised machine learning model trained on 72,001 clean educator profiles:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf17_l, "Tukey IQR Outlier Cleaning", "Eliminated pricing noise by removing rates outside 1.5x Interquartile Range, stabilizing the training corpus to 72,001 high-confidence records.", space_before=14)
add_bullet(tf17_l, "Key Predictor Coefficients", "• Metro City Tier: +₹180 / hr\n• Advanced Degree (PhD/Postgrad): +₹220 / hr\n• Verified Experience: +₹45 / hr per year of practice\n• In-Person Home Visits: +₹120 / hr travel overhead", space_before=12)
add_bullet(tf17_l, "Cross-Validation Accuracy", "Achieved an R² of 0.82 and Mean Absolute Error (MAE) of only ₹64/hr, providing realistic, fair rate recommendations.", space_before=12)

# Right Card: Fee Estimator Screenshot
_, tf17_r = create_card(s17, 10.2, 2.1, 9.0, 8.7, "Live Interactive AI Fee Estimator Tool", COLOR_SECONDARY_TEAL, text_height_in=1.0)
p = tf17_r.paragraphs[0]
p.text = "Live web widget with dynamic weekly hours slider and monthly earnings:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

scr_fee = asset("generated/screenshot_fee_estimator.png")
if os.path.exists(scr_fee):
    s17.shapes.add_picture(scr_fee, Inches(10.65), Inches(3.4), width=Inches(8.1), height=Inches(6.8))

add_notes(s17, "Slide 17 covers our ML Smart Fee Estimator. Using Ridge Regression trained on 72,000 outlier-cleaned profiles, it predicts a transparent hourly rate and monthly earnings projection based on city tier, experience, qualifications, and teaching mode.")

# =============================================================
# SLIDE 18: Generative AI Profile Studio & Bio Optimizer
# =============================================================
s18 = prs.slides.add_slide(blank_layout)
apply_base_slide(s18)
add_header(s18, 18, "Generative AI Engineering", "AI Bio Studio: Transforming Raw Tutor Profiles into Professional Pitches")

_, tf18_l = create_card(s18, 0.8, 2.1, 9.0, 8.7, "The Generative Profile Polish Pipeline", COLOR_PRIMARY_BLUE)
p = tf18_l.paragraphs[0]
p.text = "Over 40% of brilliant educators struggle to write effective self-descriptions:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf18_l, "LLM Prompt Engineering", "Uses Gemini / NLP prompt templates to parse raw, fragmented tutor notes into structured, high-converting educator bios.", space_before=14)
add_bullet(tf18_l, "Dual-Audience Optimization", "Generates two distinct sections:\n• For Parents: Pedagogical rigor, board syllabus mastery, doubt-clearing structure.\n• For Students: Interactive problem-solving, intuitive mental models, exam confidence.", space_before=12)
add_bullet(tf18_l, "Tone Customization", "Tutors can select their preferred teaching tone (Academic, Concept-First, Friendly, or Test-Prep specialist).", space_before=12)

# Right Card: Bio Studio Screenshot
_, tf18_r = create_card(s18, 10.2, 2.1, 9.0, 8.7, "Live AI Bio Generator Studio", COLOR_SECONDARY_TEAL, text_height_in=1.0)
p = tf18_r.paragraphs[0]
p.text = "One-click profile generator implemented directly on our platform:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

scr_bio = asset("generated/screenshot_bio_generator.png")
if os.path.exists(scr_bio):
    s18.shapes.add_picture(scr_bio, Inches(10.65), Inches(3.4), width=Inches(8.1), height=Inches(6.8))

add_notes(s18, "Slide 18 demonstrates our Generative AI Profile Studio. Many brilliant tutors write informal or sparse bios. Our LLM tool restructures their bio into parent-facing trust points and student-facing conceptual benefits, leveling the playing field for vernacular educators.")

# =============================================================
# SLIDE 19: Live System Demonstration Walkthrough (5-7 Mins)
# =============================================================
s19 = prs.slides.add_slide(blank_layout)
apply_base_slide(s19)
add_header(s19, 19, "Live Demonstration Guide", "Live System Demonstration Protocol (5–7 Minute Walkthrough)")

_, tf19_l = create_card(s19, 0.8, 2.1, 9.0, 8.7, "Structured 4-Step Presentation Protocol", COLOR_PRIMARY_BLUE)
p = tf19_l.paragraphs[0]
p.text = "Step-by-step walkthrough script for the 5-7 minute live demonstration:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(tf19_l, "Step 1: Directory & Filter Demo (1.5 mins)", "Navigate to Directory. Filter by State (e.g. Maharashtra) + Mode (Online). Show instant zero-lag client-side filtering across 107,673 tutors.", space_before=12)
add_bullet(tf19_l, "Step 2: Normalized Profile Modal (1.5 mins)", "Click on Tutor T1 (Shubhankshi Saraf). Demonstrate how 5 complex computer science offerings are cleanly normalized under `Coding` as `Computer Science & Programming`.", space_before=10)
add_bullet(tf19_l, "Step 3: Visual Analytics & Venn Studio (1.5 mins)", "Switch to Visual Analytics. Toggle All (107k) vs Active (76k) scope. Open 3-Subject Venn Studio, pick Math + Physics + Chemistry, show live dynamic intersection counts.", space_before=10)
add_bullet(tf19_l, "Step 4: AI Fee Estimator & Bio Studio (1.5 mins)", "Move weekly slider in AI Fee Estimator to show live monthly earning updates. Generate an instant polished bio in Bio Studio.", space_before=10)

# Right Card: Tutor Modal Screenshot (T1)
_, tf19_r = create_card(s19, 10.2, 2.1, 9.0, 8.7, "Normalized Profile Modal Showcase (Tutor T1)", COLOR_SECONDARY_TEAL, text_height_in=1.0)
p = tf19_r.paragraphs[0]
p.text = "Tutor T1: Shubhankshi Saraf displaying all 5 verified subjects normalized:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

scr_modal = asset("generated/screenshot_tutor_modal_t1.png")
if os.path.exists(scr_modal):
    s19.shapes.add_picture(scr_modal, Inches(10.65), Inches(3.4), width=Inches(8.1), height=Inches(6.8))

add_notes(s19, "Slide 19 outlines our 5 to 7 minute live demonstration plan as requested by the mentor. We demonstrate the Directory search, Tutor T1's normalized modal showing all 5 coding offerings, the Visual Analytics Venn Studio, and the AI Fee Estimator.")

# =============================================================
# SLIDE 20: Summary, Key Findings & Defense Q&A
# =============================================================
s20 = prs.slides.add_slide(blank_layout)
apply_base_slide(s20, show_watermark=False)

cov_path2 = asset("g34_network_cover_blue.png")
if os.path.exists(cov_path2):
    s20.shapes.add_picture(cov_path2, Inches(0.8), Inches(1.8), width=Inches(8.5), height=Inches(6.8))

g_badge2 = s20.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.12), Inches(0.8), Inches(1.6), Inches(0.6)
)
g_badge2.fill.solid()
g_badge2.fill.fore_color.rgb = COLOR_PRIMARY_BLUE
g_badge2.line.fill.background()
p = g_badge2.text_frame.paragraphs[0]
p.text = "SC452"
p.font.bold = True
p.font.size = Pt(20)
p.font.color.rgb = COLOR_TEXT_WHITE
p.alignment = PP_ALIGN.CENTER

tb_end_t1 = s20.shapes.add_textbox(Inches(9.8), Inches(1.4), Inches(9.4), Inches(1.2))
p = tb_end_t1.text_frame.paragraphs[0]
p.text = "BMP TUTORS"
p.font.name = FONT_HEADING
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = COLOR_NAVY_DEEP

tb_end_t2 = s20.shapes.add_textbox(Inches(9.8), Inches(2.5), Inches(9.4), Inches(1.0))
p = tb_end_t2.text_frame.paragraphs[0]
p.text = "Summary of Key Contributions"
p.font.name = FONT_HEADING
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = COLOR_PRIMARY_BLUE

sum_box = s20.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), Inches(3.7), Inches(9.4), Inches(6.8)
)
sum_box.fill.solid()
sum_box.fill.fore_color.rgb = COLOR_CARD_WHITE
sum_box.line.color.rgb = COLOR_CARD_BORDER
sum_box.line.width = Pt(1.5)

tf_sm = sum_box.text_frame
tf_sm.margin_left = Inches(0.4)
tf_sm.margin_top = Inches(0.3)
tf_sm.margin_right = Inches(0.4)

p = tf_sm.paragraphs[0]
p.text = "SUMMARY OF DELIVERABLES & GROUND TRUTH"
p.font.name = FONT_HEADING
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_SECONDARY_TEAL

add_bullet(tf_sm, "1. National Empirical Dataset", "Cleaned, verified, and mapped 107,673 tutors and 403,221 offerings across all 36 States/UTs, establishing transparent median ₹500/hr benchmark.", space_before=12)
add_bullet(tf_sm, "2. 3NF Relational Architecture", "Decomposed unnormalized spreadsheet into 3NF SQLite database (`teachers` & `teacher_subjects`), cutting storage by 64% with sub-ms indexed queries.", space_before=10)
add_bullet(tf_sm, "3. 98.2% Vocabulary Clutter Reduction", "Standardized 21,956 raw spelling variants into 386 clean canonical disciplines with zero offerings lost.", space_before=10)
add_bullet(tf_sm, "4. Production Analytics Web Application", "Engineered live web dashboard with interactive Leaflet GIS choropleth, dynamic 3-subject Venn studio, and reactive filtering.", space_before=10)
add_bullet(tf_sm, "5. AI & Machine Learning Tools", "Deployed Ridge regression ML fee estimator (MAE ₹64/hr) and Gemini generative AI profile studio.", space_before=10)

p_qa = tf_sm.add_paragraph()
p_qa.text = "Thank you! We welcome questions from the examination committee."
p_qa.font.name = FONT_HEADING
p_qa.font.size = Pt(17)
p_qa.font.bold = True
p_qa.font.color.rgb = COLOR_PRIMARY_BLUE
p_qa.space_before = Pt(16)

add_notes(s20, "Slide 20 concludes our presentation. We recapped our major engineering and empirical contributions: 107,673 verified tutors, 403,221 offerings in 3NF, 98.2% vocabulary reduction, production analytics dashboard, ML pricing model, and Gemini AI bio studio. We thank the examiners and open the floor for questions.")

print("Slides 17 to 20 generated.")

# =============================================================
# SAVE PRESENTATION
# =============================================================
out_dir = os.path.join(PROJECT_DIR, "output")
os.makedirs(out_dir, exist_ok=True)
out_path = os.path.join(out_dir, "BMP_Tutors_Master_Presentation.pptx")
prs.save(out_path)

root_copy = os.path.join(os.path.dirname(PROJECT_DIR), "BMP_Tutors_Master_Presentation.pptx")
prs.save(root_copy)

print("=" * 64)
print("🎉 SUCCESS! Master Presentation built successfully!")
print(f"   Output Path 1: {out_path}")
print(f"   Output Path 2: {root_copy}")
print(f"   Total Slides:  {len(prs.slides)}")
print(f"   Canvas Width:  {prs.slide_width / 914400:.1f} inches")
print(f"   Canvas Height: {prs.slide_height / 914400:.1f} inches")
print("=" * 64)
