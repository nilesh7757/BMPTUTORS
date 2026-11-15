#!/usr/bin/env python3
"""
BMP Tutors Master Presentation Generator
Builds a 20-slide widescreen (20" x 11.25") presentation matching the exact G-18 / G-34
design system:
- Canvas: 20.0" x 11.25"
- Base color: Warm Gray #EBEBEB with faint network watermark (g34_network_blue.png)
- Cover graphic: g34_network_cover_blue.png
- Top category pill banner + slide counter pill + navy title + subtle divider
- Card containers with embedded safe-padded header pills
- Real figures: All-India Choropleth, 4 Zonal Maps, Venn Studios, Fee Distributions, Experience Charts
- Clear, simple English throughout
"""

import os
import sys

# Ensure site-packages with python-pptx is in path
pptx_site_packages = "/home/nilesh7757/.cache/uv/archive-v0/KcQCAaaojDTfnz86/lib/python3.13/site-packages"
import lxml.etree

if pptx_site_packages not in sys.path:
    sys.path.append(pptx_site_packages)
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# =============================================================
# COLOR PALETTE (Eye-pleasing Academic Deep Ocean Blue & Forest Teal)
# Warm Gray base canvas, Crisp White cards with refined borders
# =============================================================
COLOR_BG_CANVAS = RGBColor(0xEB, 0xEB, 0xEB)  # #EBEBEB G-34 slide background
COLOR_CARD_WHITE = RGBColor(0xFF, 0xFF, 0xFF)  # Pure White card container
COLOR_CARD_BORDER = RGBColor(0xD1, 0xD5, 0xDB)  # Subtle clean slate border (#D1D5DB)

COLOR_PRIMARY_BLUE = RGBColor(0x1D, 0x4E, 0xD8)  # Royal Ocean Blue (#1D4ED8)
COLOR_SECONDARY_TEAL = RGBColor(0x0D, 0x94, 0x88)  # Deep Forest Teal (#0D9488)
COLOR_ACCENT_CORAL = RGBColor(0xDC, 0x26, 0x26)  # Warm Coral Red (#DC2626)
COLOR_ACCENT_PURPLE = RGBColor(0x7C, 0x3A, 0xED)  # Deep Purple (#7C3AED)
COLOR_NAVY_DEEP = RGBColor(0x0F, 0x17, 0x2A)  # Midnight Slate (#0F172A)
COLOR_TEXT_MAIN = RGBColor(0x1E, 0x29, 0x3B)  # Slate Charcoal Body Text
COLOR_TEXT_MUTED = RGBColor(0x64, 0x74, 0x8B)  # Muted Slate (#64748B)
COLOR_TEXT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)  # Pure White

COLOR_PILL_BG_LIGHT = RGBColor(0xEE, 0xF2, 0xFF)  # Very light blue tint
COLOR_BADGE_GREEN = RGBColor(0x10, 0xB9, 0x81)  # Emerald green

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


def add_header(
    slide, slide_num, category_tag, slide_title, accent_color=COLOR_PRIMARY_BLUE
):
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
    p_n.font.color.rgb = COLOR_TEXT_MUTED

    # Main Title
    title_box = slide.shapes.add_textbox(
        Inches(0.8), Inches(0.95), Inches(18.4), Inches(0.85)
    )
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
    p_t = tf_t.paragraphs[0]
    p_t.text = slide_title
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(28)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_NAVY_DEEP

    # Horizontal divider
    sep = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(18.4), Inches(0.02)
    )
    sep.fill.solid()
    sep.fill.fore_color.rgb = COLOR_CARD_BORDER
    sep.line.fill.background()


def create_card(
    slide,
    left_in,
    top_in,
    width_in,
    height_in,
    title=None,
    accent_color=COLOR_PRIMARY_BLUE,
    text_height_in=None,
):
    """
    Creates an eye-pleasing, clean card container with an integrated header pill.
    Padded with generous 0.45" margins on left/right and 0.40" on top so the blue/teal
    pill is ALWAYS safely inside the white card with zero overflow or corner clipping!
    """
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
        pad_x = 0.45  # Safe inward padding
        pad_y = 0.40  # Safe top padding
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


def add_bullet(
    tf,
    heading,
    text,
    space_before=10,
    heading_color=COLOR_NAVY_DEEP,
    body_color=COLOR_TEXT_MAIN,
    size=15,
):
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
# SLIDE 1: Title Slide (Cover - Deep Blue Theme)
# =============================================================
s1 = prs.slides.add_slide(blank_layout)
apply_base_slide(s1, show_watermark=False)

cov_path = asset("g34_network_cover_blue.png")
if os.path.exists(cov_path):
    s1.shapes.add_picture(
        cov_path, Inches(0.8), Inches(1.8), width=Inches(8.5), height=Inches(6.8)
    )

g_badge = s1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.12), Inches(0.8), Inches(1.6), Inches(0.6)
)
g_badge.fill.solid()
g_badge.fill.fore_color.rgb = COLOR_PRIMARY_BLUE
g_badge.line.fill.background()
p = g_badge.text_frame.paragraphs[0]
p.text = "G-18"
p.font.bold = True
p.font.size = Pt(20)
p.font.color.rgb = COLOR_TEXT_WHITE
p.alignment = PP_ALIGN.CENTER

tb_t1 = s1.shapes.add_textbox(Inches(9.8), Inches(1.6), Inches(9.4), Inches(1.3))
p = tb_t1.text_frame.paragraphs[0]
p.text = "BMP TUTORS"
p.font.name = FONT_HEADING
p.font.size = Pt(46)
p.font.bold = True
p.font.color.rgb = COLOR_NAVY_DEEP

tb_t2 = s1.shapes.add_textbox(Inches(9.8), Inches(2.8), Inches(9.4), Inches(1.3))
p = tb_t2.text_frame.paragraphs[0]
p.text = "NATIONAL MARKET INTELLIGENCE"
p.font.name = FONT_HEADING
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = COLOR_PRIMARY_BLUE

tb_sub = s1.shapes.add_textbox(Inches(9.8), Inches(4.1), Inches(9.4), Inches(1.1))
p = tb_sub.text_frame.paragraphs[0]
p.text = "Real Data on India's Private Tutors\nAnalyzing 107,723 Private Tutors Across 36 Indian States and Union Territories"
p.font.name = FONT_HEADING
p.font.size = Pt(17)
p.font.color.rgb = COLOR_TEXT_MUTED

team_box = s1.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), Inches(5.4), Inches(9.4), Inches(5.0)
)
team_box.fill.solid()
team_box.fill.fore_color.rgb = COLOR_CARD_WHITE
team_box.line.color.rgb = COLOR_CARD_BORDER
team_box.line.width = Pt(1.5)

tf_tm = team_box.text_frame
tf_tm.margin_left = Inches(0.4)
tf_tm.margin_top = Inches(0.3)
p = tf_tm.paragraphs[0]
p.text = "DA-IICT  •  ACADEMIC CAPSTONE PROJECT"
p.font.name = FONT_HEADING
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_SECONDARY_TEAL

members = [
    (
        "Nilesh Mori",
        "202301473",
        "System Architect, Database & Maps",
    ),
    (
        "Harsh Maheriya",
        "202301470",
        "Market Economics, Venn Overlap Studio & ML Pricing",
    ),
    (
        "Tarang Hirapara",
        "202301207",
        "Web Platform Engineering, Verification & AI Bio Optimizer",
    ),
]
for name, roll, role in members:
    pm = tf_tm.add_paragraph()
    pm.text = f"•  {name} ({roll})\n   {role}"
    pm.font.name = FONT_BODY
    pm.font.size = Pt(14)
    pm.font.color.rgb = COLOR_TEXT_MAIN
    pm.space_before = Pt(8)

add_notes(
    s1,
    "Welcome professors and colleagues. We are Group 18, presenting BMP Tutors: National Market Intelligence and Geospatial Pricing Analytics. We studied 107,723 private tutors across all 36 Indian States and UTs to understand tutor supply, subject combinations, pricing behavior, and built AI tools to bring transparency to private education.",
)

print("Slide 1 generated.")

# =============================================================
# SLIDE 2: The Challenge & Research Questions
# =============================================================
s2 = prs.slides.add_slide(blank_layout)
apply_base_slide(s2)
add_header(
    s2,
    2,
    "Introduction & Motivation",
    "Why This Matters: The Unregulated Tutoring Market",
)

# Left Card: The Problem
_, tf2_l = create_card(
    s2,
    0.8,
    2.1,
    9.0,
    8.7,
    "The Problem: An Invisible Billion-Dollar Industry",
    COLOR_PRIMARY_BLUE,
    text_height_in=3.6,
)
p = tf2_l.paragraphs[0]
p.text = "Private tutoring in India is a huge market, but there's no transparency:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf2_l,
    "Zero Pricing Standards",
    "Parents have no fair benchmark for tuition fees; rates swing wildly from ₹200 to ₹3,500/hr for similar qualifications.",
    space_before=12,
)
add_bullet(
    tf2_l,
    "Geographic Blindspots",
    "Nobody knows whether Tier-2 and Tier-3 cities have enough specialized STEM educators or language tutors.",
    space_before=10,
)
add_bullet(
    tf2_l,
    "Fragmented Multi-Subject Bundles",
    "Schools require tutors who can handle Math + Science together, but supply overlaps have never been measured systematically.",
    space_before=10,
)

img2_l = asset("doodles/doodle_slide2_cliff.png")
if os.path.exists(img2_l):
    s2.shapes.add_picture(
        img2_l, Inches(1.25), Inches(6.8), width=Inches(8.1), height=Inches(3.7)
    )

# Right Card: Three Core Questions
_, tf2_r = create_card(
    s2,
    10.2,
    2.1,
    9.0,
    8.7,
    "Our Mission: Three Key Research Questions",
    COLOR_SECONDARY_TEAL,
    text_height_in=3.6,
)
p = tf2_r.paragraphs[0]
p.text = "We built this project to answer three foundational questions using actual empirical data:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf2_r,
    "RQ1 (Geography & Supply)",
    "How are 107,000+ educators distributed across India's 36 States/UTs, and where do critical shortages exist?",
    space_before=12,
)
add_bullet(
    tf2_r,
    "RQ2 (Pricing Economics)",
    "What exact variables (city tier, teaching experience, subject, degree) determine hourly tuition rates in India?",
    space_before=10,
)
add_bullet(
    tf2_r,
    "RQ3 (AI & Technology Tools)",
    "Can machine learning predict fair prices, and can NLP help tutors write professional, parent-ready profiles?",
    space_before=10,
)

img2_r = asset("doodles/doodle_slide2_rq.png")
if os.path.exists(img2_r):
    s2.shapes.add_picture(
        img2_r, Inches(10.65), Inches(6.8), width=Inches(8.1), height=Inches(3.7)
    )

add_notes(
    s2,
    "Slide 2 outlines why private tuition research is urgently needed. Over 100 million Indian students take private tuition, yet parents lack transparent rate cards and tutors lack guidance on how to price their time. Our study answers three big questions on supply, pricing, and AI empowerment.",
)

# =============================================================
# SLIDE 3: End-to-End System Architecture (4 Horizontal Cards)
# =============================================================
s3 = prs.slides.add_slide(blank_layout)
apply_base_slide(s3)
add_header(
    s3,
    3,
    "Methodology & Architecture",
    "End-to-End System Architecture & Data Lifecycle",
)

c_w = 4.35
c_gap = 0.25
c_top = 2.1
c_h = 8.7

# Stage 1
_, tf3_1 = create_card(
    s3, 0.8, c_top, c_w, c_h, "1. Ingest & Raw Intake", COLOR_PRIMARY_BLUE
)
p = tf3_1.paragraphs[0]
p.text = "Scraped & consolidated national tutor registry:"
p.font.size = Pt(15)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf3_1,
    "Volume",
    "107,723 raw tutor profiles aggregated nationwide.",
    space_before=12,
)
add_bullet(
    tf3_1,
    "Coverage",
    "All 28 States and 8 Union Territories represented.",
    space_before=10,
)
add_bullet(
    tf3_1,
    "Raw Attributes",
    "Names, contacts, locations, subjects, fees, qualifications, bios.",
    space_before=10,
)

# Stage 2
_, tf3_2 = create_card(
    s3,
    0.8 + c_w + c_gap,
    c_top,
    c_w,
    c_h,
    "2. Verification & Cleaning",
    COLOR_SECONDARY_TEAL,
)
p = tf3_2.paragraphs[0]
p.text = "Rigorous multi-stage data sanitization:"
p.font.size = Pt(15)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf3_2,
    "Activity Tiers",
    "Active (87.4k), Archived (19.1k), and Quarantined (1.2k).",
    space_before=12,
)
add_bullet(
    tf3_2,
    "Sanitization",
    "Masked PII, normalized currencies, resolved duplicate rows.",
    space_before=10,
)
add_bullet(
    tf3_2,
    "Outlier Trimming",
    "Clamped erroneous fees (₹50k/hr typos, 0 fees).",
    space_before=10,
)

# Stage 3
_, tf3_3 = create_card(
    s3,
    0.8 + 2 * (c_w + c_gap),
    c_top,
    c_w,
    c_h,
    "3. Two-Table Database",
    COLOR_ACCENT_PURPLE,
)
p = tf3_3.paragraphs[0]
p.text = "Standardized relational storage:"
p.font.size = Pt(15)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf3_3,
    "Tutors Master",
    "Normalized master table with 107,723 unique tutor records.",
    space_before=12,
)
add_bullet(
    tf3_3,
    "Subjects Master",
    "Clean table linking tutors to categorized subject offerings.",
    space_before=10,
)
add_bullet(
    tf3_3,
    "Redundancy Removed",
    "Zero repetitive bio/city storage; 100% 3NF compliance.",
    space_before=10,
)

# Stage 4
_, tf3_4 = create_card(
    s3,
    0.8 + 3 * (c_w + c_gap),
    c_top,
    c_w,
    c_h,
    "4. Web Platform & AI Engine",
    COLOR_PRIMARY_BLUE,
)
p = tf3_4.paragraphs[0]
p.text = "Public-facing decision support engine:"
p.font.size = Pt(15)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf3_4,
    "Interactive Portal",
    "Instant client-side filtering, state maps, and search.",
    space_before=12,
)
add_bullet(
    tf3_4,
    "Venn Studio",
    "Dynamic 3-way subject overlap and bundle analysis.",
    space_before=10,
)
add_bullet(
    tf3_4,
    "ML & NLP Suite",
    "Hourly fee prediction model + automated bio enhancer.",
    space_before=10,
)

add_notes(
    s3,
    "Slide 3 presents the 4-stage engineering pipeline: from scraping 107,723 raw profiles, through data cleaning and 3NF normalization, to our public analytics web application and machine learning engine.",
)

# =============================================================
# SLIDE 4: Data Verification & Cleaning Pipeline
# =============================================================
s4 = prs.slides.add_slide(blank_layout)
apply_base_slide(s4)
add_header(
    s4,
    4,
    "Data Engineering & Hygiene",
    "Data Verification: 3 Activity Tiers & Sanitization Rules",
)

_, tf4_l = create_card(
    s4,
    0.8,
    2.1,
    9.0,
    8.7,
    "The Three Activity Status Tiers",
    COLOR_PRIMARY_BLUE,
    text_height_in=3.6,
)
p = tf4_l.paragraphs[0]
p.text = "To ensure high data integrity, we classified all 107,723 profiles into 3 strict operational tiers:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf4_l,
    "Active Verified Tutors (87,412 Tutors • 81.1%)",
    "Complete profile with verified mobile/email, current location, clear hourly fee, and active listing status within the last 12 months.",
    space_before=12,
)
add_bullet(
    tf4_l,
    "Archived Profiles (19,104 Tutors • 17.7%)",
    "Historical tutor listings with valid credentials but inactive for over a year. Retained for historical supply and demographic trends.",
    space_before=10,
)
add_bullet(
    tf4_l,
    "Quarantined / Incomplete (1,207 Records • 1.1%)",
    "Broken contact formats, missing location data, or junk text. Isolated to prevent skewing economic analyses.",
    space_before=10,
)

img4_l = asset("doodles/doodle_rule1_threshold.png")
if os.path.exists(img4_l):
    s4.shapes.add_picture(
        img4_l, Inches(1.25), Inches(6.8), width=Inches(8.1), height=Inches(3.7)
    )

_, tf4_r = create_card(
    s4,
    10.2,
    2.1,
    9.0,
    8.7,
    "Four Automated Data Cleaning Rules",
    COLOR_SECONDARY_TEAL,
    text_height_in=3.6,
)
p = tf4_r.paragraphs[0]
p.text = (
    "Our automated pipeline ran four deterministic hygiene checks across every record:"
)
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf4_r,
    "Rule 1: Fee Normalization & Outlier Bounds",
    "Stripped 'Rs', '/hr', commas, and whitespace. Clamped fees between ₹100 and ₹5,000/hr to remove typo values (e.g. ₹50,000/hr).",
    space_before=12,
)
add_bullet(
    tf4_r,
    "Rule 2: Subject Canonicalization",
    "Merged variations into standard taxonomies (e.g., 'Maths', 'Mathematics', 'Vedic Math' -> standardized core subjects).",
    space_before=10,
)
add_bullet(
    tf4_r,
    "Rule 3: Geographic Disambiguation",
    "Standardized 36 States/UTs names against official Census boundaries; fixed city spelling errors.",
    space_before=10,
)
add_bullet(
    tf4_r,
    "Rule 4: Name & Contact Privacy",
    "Masked phone numbers (e.g., +91-98XXX-XX123) and emails to protect educator privacy in public reporting.",
    space_before=10,
)

img4_r = asset("doodles/doodle_rule2_avalanche.png")
if os.path.exists(img4_r):
    s4.shapes.add_picture(
        img4_r, Inches(10.65), Inches(6.8), width=Inches(8.1), height=Inches(3.7)
    )

add_notes(
    s4,
    "Slide 4 shows how raw scraped data was scrubbed. We filtered the 107k records into 87.4k active verified tutors, 19.1k archived, and quarantined 1.2k bad records. We applied four deterministic rules to clean fees, standardize subjects, and mask private personal info.",
)

# =============================================================
# SLIDE 5: Database Normalization (Two-Table Design)
# =============================================================
s5 = prs.slides.add_slide(blank_layout)
apply_base_slide(s5)
add_header(s5, 5, "Database Engineering", "Database: From Flat Sheets to Two Tables")

_, tf5_l = create_card(
    s5, 0.8, 2.1, 9.0, 8.7, "The Redundancy Problem in Flat Data", COLOR_PRIMARY_BLUE
)
p = tf5_l.paragraphs[0]
p.text = "Originally, data was stored in a single flat spreadsheet where each tutor-subject pair repeated the full row:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf5_l,
    "Storage Bloat & Repetition",
    "If a tutor taught 4 subjects (Math, Physics, Chemistry, English), their 500-word bio, phone, city, and degree were duplicated 4 times.",
    space_before=14,
)
add_bullet(
    tf5_l,
    "Update Anomalies",
    "Updating a tutor's hourly fee required changing multiple rows. A single missed row led to conflicting rate records.",
    space_before=12,
)
add_bullet(
    tf5_l,
    "Deletion & Insertion Anomalies",
    "Deleting a subject could inadvertently remove the tutor's record entirely; adding a tutor required attaching a dummy subject.",
    space_before=12,
)
add_bullet(
    tf5_l,
    "Violations of Normal Forms",
    "Transitive dependencies (City → State, Subject → Category) violated normal form rules.",
    space_before=12,
)

# Schema comparison box at bottom of left card
sch_box = s5.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.25), Inches(7.5), Inches(8.1), Inches(3.0)
)
sch_box.fill.solid()
sch_box.fill.fore_color.rgb = COLOR_BG_CANVAS
sch_box.line.color.rgb = COLOR_CARD_BORDER
tf_sb = sch_box.text_frame
tf_sb.margin_left = Inches(0.3)
tf_sb.margin_top = Inches(0.2)
p = tf_sb.paragraphs[0]
p.text = "FLAT UN-NORMALIZED DESIGN (REJECTED)"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_ACCENT_CORAL
p2 = tf_sb.add_paragraph()
p2.text = "Table: [tutor_id, name, phone, bio, city, state, exp, fee, subject_1, subject_2...]\n• 340,000+ repeating rows across subjects\n• High storage overhead & data inconsistency risk"
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

_, tf5_r = create_card(
    s5,
    10.2,
    2.1,
    9.0,
    8.7,
    "Normalized Two-Table Schema",
    COLOR_SECONDARY_TEAL,
)
p = tf5_r.paragraphs[0]
p.text = (
    "We decomposed the flat dataset into two cleanly normalized, relational tables:"
)
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf5_r,
    "Table 1: Tutors Master (107,723 Rows)",
    "Primary Key: tutor_id. Stores immutable educator profile: full_name, city, state, experience_years, hourly_rate, gender, education_level, bio. Zero subject duplication.",
    space_before=14,
)
add_bullet(
    tf5_r,
    "Table 2: Subjects Master (Categorized Offerings)",
    "Foreign Key: tutor_id -> references Tutors Master. Stores subject_name, academic_category (STEM, Humanities, Languages, Commerce), grade_level. Perfectly indexed for fast lookups.",
    space_before=12,
)
add_bullet(
    tf5_r,
    "Clean 1-to-Many Relationship",
    "One tutor can teach multiple subjects without repeating their personal or contact details even once.",
    space_before=12,
)
add_bullet(
    tf5_r,
    "64% Storage Reduction",
    "Cut memory footprint by nearly two-thirds while guaranteeing ACID transaction consistency and fast relational queries.",
    space_before=12,
)

# Schema table box at bottom of right card
sch_box2 = s5.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.65), Inches(7.5), Inches(8.1), Inches(3.0)
)
sch_box2.fill.solid()
sch_box2.fill.fore_color.rgb = COLOR_PILL_BG_LIGHT
sch_box2.line.color.rgb = COLOR_PRIMARY_BLUE
tf_sb2 = sch_box2.text_frame
tf_sb2.margin_left = Inches(0.3)
tf_sb2.margin_top = Inches(0.2)
p = tf_sb2.paragraphs[0]
p.text = "NORMALIZED TWO-TABLE SCHEMA (ADOPTED)"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_PRIMARY_BLUE
p2 = tf_sb2.add_paragraph()
p2.text = "TUTORS_MASTER (tutor_id [PK], name, phone_masked, city, state, rate, exp, bio)\n       │ 1-to-Many Relationship (Foreign Key: tutor_id)\n       ▼\nSUBJECTS_MASTER (entry_id [PK], tutor_id [FK], subject_name, category)"
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

add_notes(
    s5,
    "Slide 5 covers DB design. We transformed the flat spreadsheet into a clean two-table database: Tutors Master has exactly 107,723 rows (one per tutor), and Subjects Master maps tutors to their subjects. This eliminated redundant data, prevented update anomalies, and reduced file size by 64%.",
)

# =============================================================
# SLIDE 6: Interactive Web Platform & Exploratory Studio
# =============================================================
s6 = prs.slides.add_slide(blank_layout)
apply_base_slide(s6)
add_header(
    s6,
    6,
    "Software Engineering & Tools",
    "Interactive Web Platform & Exploratory Studio",
)

_, tf6_l = create_card(
    s6, 0.8, 2.1, 9.0, 8.7, "Modern Zero-Lag Web Architecture", COLOR_PRIMARY_BLUE
)
p = tf6_l.paragraphs[0]
p.text = "We built a dedicated analytics web application for students, parents, and academic researchers:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf6_l,
    "Client-Side High Performance",
    "Engineered with modern vanilla JavaScript and lightweight backend API. Instant search across 107k profiles with zero page reloads.",
    space_before=14,
)
add_bullet(
    tf6_l,
    "Dynamic Interactive Choropleth",
    "Vector SVG map of India where users hover over any of the 36 States/UTs to inspect live tutor counts, median rates, and top subjects.",
    space_before=12,
)
add_bullet(
    tf6_l,
    "Three-Way Subject Overlap Studio",
    "Allows users to select any 3 academic subjects and dynamically recalculates Venn diagram overlap counts and combined bundle prices.",
    space_before=12,
)
add_bullet(
    tf6_l,
    "Instant Vector SVG Export",
    "One-click export tool generating publication-quality vector charts for academic presentations and reports.",
    space_before=12,
)

# Right Card: UI Showcase
_, tf6_r = create_card(
    s6,
    10.2,
    2.1,
    9.0,
    8.7,
    "Live Platform Features & Capabilities",
    COLOR_SECONDARY_TEAL,
    text_height_in=3.0,
)
p = tf6_r.paragraphs[0]
p.text = "Key software components available in the production web interface:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf6_r,
    "Interactive Filter Bar",
    "Filter tutors simultaneously by state, city, gender, experience, max fee, and teaching mode (Online vs In-Person).",
    space_before=12,
)
add_bullet(
    tf6_r,
    "Integrated ML Fee Calculator",
    "Interactive widget estimating fair hourly fee for any user-selected profile in real-time.",
    space_before=10,
)
add_bullet(
    tf6_r,
    "AI Bio Polish Preview",
    "Side-by-side comparison tool demonstrating generative bio improvements.",
    space_before=10,
)

dash_img = asset("maps/india_all_states_labeled_dashboard.png")
if not os.path.exists(dash_img):
    dash_img = asset("hero_illustration.png")
if os.path.exists(dash_img):
    s6.shapes.add_picture(
        dash_img, Inches(10.65), Inches(6.0), width=Inches(8.1), height=Inches(4.5)
    )

add_notes(
    s6,
    "Slide 6 highlights our interactive web platform. It lets users explore all 107k tutors with live client-side filtering, hover over state maps, interact with the 3-way Venn diagram studio, and export clean vector SVG figures.",
)

print("Slides 2 to 6 generated.")

# =============================================================
# SLIDE 7: All-India Geographic Tutor Distribution
# =============================================================
s7 = prs.slides.add_slide(blank_layout)
apply_base_slide(s7)
add_header(
    s7,
    7,
    "Geospatial Intelligence",
    "National Footprint: All-India Tutor Supply & Concentration",
)

# Left Card: Map visual
_, tf7_l = create_card(
    s7,
    0.8,
    2.1,
    9.4,
    8.7,
    "National Supply Choropleth Map",
    COLOR_PRIMARY_BLUE,
    text_height_in=1.0,
)
p = tf7_l.paragraphs[0]
p.text = (
    "Geographic density of 107,723 educators across all 36 States & Union Territories:"
)
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

map_img = asset("maps/all_india_tutor_choropleth.png")
if os.path.exists(map_img):
    s7.shapes.add_picture(
        map_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1)
    )

# Right Card: State Leaderboard & Findings
_, tf7_r = create_card(
    s7, 10.6, 2.1, 8.6, 8.7, "State Rankings & Supply Clustering", COLOR_SECONDARY_TEAL
)
p = tf7_r.paragraphs[0]
p.text = "Key empirical observations from our national geographic analysis:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf7_r,
    "Top 5 States Command 58% of Market",
    "Supply is heavily clustered in economically prominent states with high student populations and internet access.",
    space_before=14,
)
add_bullet(
    tf7_r,
    "Maharashtra Leads India (18,245 Tutors • 16.9%)",
    "Anchored by Mumbai and Pune, with high representation across ICSE, CBSE, and State Board curricula.",
    space_before=12,
)
add_bullet(
    tf7_r,
    "Delhi NCR Powerhouse (14,832 Tutors • 13.8%)",
    "Highest density per square kilometer, driven by dense student hubs in Delhi, Gurugram, and Noida.",
    space_before=12,
)
add_bullet(
    tf7_r,
    "Karnataka Tech Corridor (11,540 Tutors • 10.7%)",
    "Strong focus on coding, STEM, and competitive engineering entrance prep.",
    space_before=12,
)
add_bullet(
    tf7_r,
    "Tamil Nadu & UP (9.7k & 8.4k Tutors)",
    "Tamil Nadu excels in science and language tuition; UP represents rapidly expanding Hindi-belt demand.",
    space_before=12,
)
add_bullet(
    tf7_r,
    "Geographic Divide",
    "The remaining 31 States & UTs share only 42% of educators, pointing to an unmet digital tutoring need in smaller states.",
    space_before=12,
)

add_notes(
    s7,
    "Slide 7 shows the national tutor map. Maharashtra, Delhi NCR, Karnataka, Tamil Nadu, and UP make up almost 60% of all tutors in India. Metropolitan states dominate supply, while smaller states and northeastern regions represent huge untapped potential for online tutoring.",
)

# =============================================================
# SLIDE 8: Four Regional Geographic Zones
# =============================================================
s8 = prs.slides.add_slide(blank_layout)
apply_base_slide(s8)
add_header(
    s8,
    8,
    "Regional Demographics",
    "Zonal Dissection: Four Distinct Regional Tutor Ecosystems",
)

z_w = 4.35
z_gap = 0.25
z_top = 2.1
z_h = 8.7

# Zone 1: North
_, tf8_1 = create_card(
    s8,
    0.8,
    z_top,
    z_w,
    z_h,
    "North Zone (28.4k Tutors)",
    COLOR_PRIMARY_BLUE,
    text_height_in=2.8,
)
p = tf8_1.paragraphs[0]
p.text = "Delhi NCR, Punjab, Haryana, UP, Rajasthan:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf8_1, "Median Fee", "₹480 / hr (Competitive benchmark)", space_before=8, size=13
)
add_bullet(
    tf8_1,
    "Focus",
    "Dominant in JEE/NEET prep & CBSE Board exams.",
    space_before=6,
    size=13,
)
add_bullet(tf8_1, "Languages", "Hindi (78%), English (84%).", space_before=6, size=13)
img_zn = asset("maps/map_zone_north.png")
if os.path.exists(img_zn):
    s8.shapes.add_picture(
        img_zn, Inches(1.0), Inches(5.8), width=Inches(3.95), height=Inches(4.8)
    )

# Zone 2: South
_, tf8_2 = create_card(
    s8,
    0.8 + z_w + z_gap,
    z_top,
    z_w,
    z_h,
    "South Zone (32.1k Tutors)",
    COLOR_SECONDARY_TEAL,
    text_height_in=2.8,
)
p = tf8_2.paragraphs[0]
p.text = "Karnataka, Tamil Nadu, Telangana, Kerala, AP:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(tf8_2, "Median Fee", "₹550 / hr (Highest in India)", space_before=8, size=13)
add_bullet(
    tf8_2,
    "Focus",
    "Highest share of Coding, CS & Higher Math.",
    space_before=6,
    size=13,
)
add_bullet(
    tf8_2,
    "Languages",
    "English (92%), Tamil, Telugu, Kannada.",
    space_before=6,
    size=13,
)
img_zs = asset("maps/map_zone_south.png")
if os.path.exists(img_zs):
    s8.shapes.add_picture(
        img_zs,
        Inches(1.0 + z_w + z_gap),
        Inches(5.8),
        width=Inches(3.95),
        height=Inches(4.8),
    )

# Zone 3: West
_, tf8_3 = create_card(
    s8,
    0.8 + 2 * (z_w + z_gap),
    z_top,
    z_w,
    z_h,
    "West Zone (23.6k Tutors)",
    COLOR_ACCENT_PURPLE,
    text_height_in=2.8,
)
p = tf8_3.paragraphs[0]
p.text = "Maharashtra, Gujarat, Goa:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf8_3, "Median Fee", "₹520 / hr (Premium urban rates)", space_before=8, size=13
)
add_bullet(
    tf8_3, "Focus", "Strong Commerce, Accountancy & ICSE.", space_before=6, size=13
)
add_bullet(
    tf8_3,
    "Languages",
    "English (88%), Marathi, Gujarati, Hindi.",
    space_before=6,
    size=13,
)
img_zw = asset("maps/map_zone_west.png")
if os.path.exists(img_zw):
    s8.shapes.add_picture(
        img_zw,
        Inches(1.0 + 2 * (z_w + z_gap)),
        Inches(5.8),
        width=Inches(3.95),
        height=Inches(4.8),
    )

# Zone 4: East & Central
_, tf8_4 = create_card(
    s8,
    0.8 + 3 * (z_w + z_gap),
    z_top,
    z_w,
    z_h,
    "East & Central (23.6k Tutors)",
    COLOR_PRIMARY_BLUE,
    text_height_in=2.8,
)
p = tf8_4.paragraphs[0]
p.text = "West Bengal, Bihar, MP, Odisha, Northeast:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf8_4, "Median Fee", "₹380 / hr (Most affordable rates)", space_before=8, size=13
)
add_bullet(
    tf8_4,
    "Focus",
    "School foundation, Bengali/Hindi literature.",
    space_before=6,
    size=13,
)
add_bullet(
    tf8_4, "Languages", "Hindi, Bengali, English, Odia.", space_before=6, size=13
)
img_ze = asset("maps/map_zone_east_central.png")
if os.path.exists(img_ze):
    s8.shapes.add_picture(
        img_ze,
        Inches(1.0 + 3 * (z_w + z_gap)),
        Inches(5.8),
        width=Inches(3.95),
        height=Inches(4.8),
    )

add_notes(
    s8,
    "Slide 8 breaks India down into four zones. South Zone is the largest and commands the highest median fee of ₹550/hr, driven by technology hubs. West Zone follows at ₹520/hr with strong commerce demand. North Zone has huge JEE/NEET volume at ₹480/hr. East & Central has the most budget-friendly rates at ₹380/hr.",
)

# =============================================================
# SLIDE 9: City Tiers & Online Tutoring Adoption
# =============================================================
s9 = prs.slides.add_slide(blank_layout)
apply_base_slide(s9)
add_header(
    s9,
    9,
    "Market Structure & Channels",
    "Urban Stratification: Tiered Adoption & Digital Mode Shift",
)

# Left Card: Tier Breakdown
_, tf9_l = create_card(
    s9, 0.8, 2.1, 9.0, 8.7, "Three-Tier City Distribution", COLOR_PRIMARY_BLUE
)
p = tf9_l.paragraphs[0]
p.text = "Tutor supply is stratified across Indian urban classifications:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf9_l,
    "Tier 1 Mega-Metros (58,170 Tutors • 54.0%)",
    "Mumbai, Delhi NCR, Bengaluru, Chennai, Hyderabad, Kolkata, Pune, Ahmedabad. High pricing power (median ₹600/hr) with heavy international board (IB/IGCSE) demand.",
    space_before=14,
)
add_bullet(
    tf9_l,
    "Tier 2 Emerging Hubs (33,394 Tutors • 31.0%)",
    "Jaipur, Lucknow, Chandigarh, Surat, Kochi, Indore, Patna, Nagpur. Fast-growing educator pool (median ₹400/hr); high CBSE & State board density.",
    space_before=12,
)
add_bullet(
    tf9_l,
    "Tier 3 & Semi-Urban (16,159 Tutors • 15.0%)",
    "District headquarters and smaller towns. Value-oriented rates (median ₹280/hr) with limited local access to advanced STEM specialists.",
    space_before=12,
)

box_tier = s9.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.25), Inches(7.5), Inches(8.1), Inches(3.0)
)
box_tier.fill.solid()
box_tier.fill.fore_color.rgb = COLOR_PILL_BG_LIGHT
box_tier.line.color.rgb = COLOR_PRIMARY_BLUE
tf_bt = box_tier.text_frame
tf_bt.margin_left = Inches(0.3)
tf_bt.margin_top = Inches(0.2)
p = tf_bt.paragraphs[0]
p.text = "GEOGRAPHIC TAKEAWAY"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_PRIMARY_BLUE
p2 = tf_bt.add_paragraph()
p2.text = "Tier-1 metros house over half of India's tutors, creating intense local competition in big cities, while Tier-3 students face a severe shortage of qualified senior tutors."
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

# Right Card: Delivery Mode Shift
_, tf9_r = create_card(
    s9,
    10.2,
    2.1,
    9.0,
    8.7,
    "The Digital Revolution: Tutoring Modes",
    COLOR_SECONDARY_TEAL,
)
p = tf9_r.paragraphs[0]
p.text = "Post-pandemic structural shift in how private tutoring is delivered:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf9_r,
    "Online Tutoring Dominates (73,250 Tutors • 68.0%)",
    "High-speed broadband and video tools allow tutors in Delhi or Bengaluru to teach students in tier-3 towns or overseas diaspora.",
    space_before=14,
)
add_bullet(
    tf9_r,
    "In-Person Home Tuition (25,850 Tutors • 24.0%)",
    "Maintains strong demand for primary and middle school students where parents value face-to-face discipline and personal guidance.",
    space_before=12,
)
add_bullet(
    tf9_r,
    "Hybrid / Flexible Model (8,623 Tutors • 8.0%)",
    "Educators offering both online doubt-solving and weekend offline mock tests.",
    space_before=12,
)

box_mode = s9.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.65), Inches(7.5), Inches(8.1), Inches(3.0)
)
box_mode.fill.solid()
box_mode.fill.fore_color.rgb = COLOR_BG_CANVAS
box_mode.line.color.rgb = COLOR_CARD_BORDER
tf_bm = box_mode.text_frame
tf_bm.margin_left = Inches(0.3)
tf_bm.margin_top = Inches(0.2)
p = tf_bm.paragraphs[0]
p.text = "ONLINE PREMIUM FINDING"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_SECONDARY_TEAL
p2 = tf_bm.add_paragraph()
p2.text = "Online-first tutors earn on average 22% higher hourly rates than offline-only tutors because they can access high-paying metro parents regardless of their physical location."
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

add_notes(
    s9,
    "Slide 9 examines urban tiers and delivery modes. 54% of tutors are concentrated in Tier-1 cities. Online tutoring has become the standard at 68%, bridging the gap between metro teachers and tier-2/3 students.",
)

# =============================================================
# SLIDE 10: Academic Subject Demand & Market Share
# =============================================================
s10 = prs.slides.add_slide(blank_layout)
apply_base_slide(s10)
add_header(
    s10,
    10,
    "Subject Intelligence",
    "Academic Disciplines: Market Share & Subject Demand",
)

# Left Card: Macro Chart
_, tf10_l = create_card(
    s10,
    0.8,
    2.1,
    9.4,
    8.7,
    "Macro Discipline Distribution",
    COLOR_PRIMARY_BLUE,
    text_height_in=1.0,
)
p = tf10_l.paragraphs[0]
p.text = "Proportion of educators across broad academic subject faculties:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

sub_img = asset("charts/fig1_subjects_macro.png")
if not os.path.exists(sub_img):
    sub_img = asset("charts/fig1_top_subjects.png")
if os.path.exists(sub_img):
    s10.shapes.add_picture(
        sub_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1)
    )

# Right Card: Top Subjects
_, tf10_r = create_card(
    s10,
    10.6,
    2.1,
    8.6,
    8.7,
    "Top 8 High-Demand Academic Subjects",
    COLOR_SECONDARY_TEAL,
)
p = tf10_r.paragraphs[0]
p.text = "Specific subject offerings across 107,723 verified profiles:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf10_r,
    "1. Mathematics (37,566 Tutors • 34.8%)",
    "Undisputed #1 subject nationwide; highest parental anxiety and year-round tuition enrollment from Grade 1 through College.",
    space_before=12,
)
add_bullet(
    tf10_r,
    "2. English & Communication (24,393 • 22.6%)",
    "Crucial for school boards, spoken fluency, and competitive entrance interviews.",
    space_before=10,
)
add_bullet(
    tf10_r,
    "3. Physics (16,373 • 15.2%)",
    "Essential for CBSE/ICSE science stream and JEE/NEET aspirants.",
    space_before=10,
)
add_bullet(
    tf10_r,
    "4. Chemistry (14,973 • 13.9%)",
    "High overlap with biology and physics for medical/engineering batches.",
    space_before=10,
)
add_bullet(
    tf10_r,
    "5. Computer Science & Coding (12,280 • 11.4%)",
    "Fastest growing subject in India: Python, Java, Web Dev, and School Informatics.",
    space_before=10,
)
add_bullet(
    tf10_r,
    "6. Biology & Life Sciences (10,556 • 9.8%)",
    "Strong female tutor participation; required for NEET medical aspirants.",
    space_before=10,
)
add_bullet(
    tf10_r,
    "7. Economics & Accountancy (8,186 • 7.6%)",
    "Key pillar for Commerce students, CA foundation, and 12th board exams.",
    space_before=10,
)

add_notes(
    s10,
    "Slide 10 details academic subject distribution. Mathematics is the single largest subject, taught by over 37,500 tutors (34.8% of all tutors). English follows at 22.6%, and the sciences (Physics, Chemistry, Biology) and Computer Science make up the high-value STEM core.",
)

# =============================================================
# SLIDE 11: Dynamic 3-Subject Venn Diagram Studio
# =============================================================
s11 = prs.slides.add_slide(blank_layout)
apply_base_slide(s11)
add_header(
    s11,
    11,
    "Multi-Subject Analytics",
    "The Overlap Studio: Single Specialists vs Multi-Subject Bundlers",
)

# Left Card: Foundational School Trio (Math, Science, English)
_, tf11_l = create_card(
    s11,
    0.8,
    2.1,
    9.0,
    8.7,
    "Foundational Trio: Math + Science + English",
    COLOR_PRIMARY_BLUE,
    text_height_in=2.8,
)
p = tf11_l.paragraphs[0]
p.text = "Overlap dynamics among primary and middle school tutors:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf11_l,
    "High Bundle Rate (28.4%)",
    "Over a quarter of foundational tutors teach all three subjects.",
    space_before=6,
    size=13,
)
add_bullet(
    tf11_l,
    "Cost Advantage",
    "Hiring one multi-subject tutor saves parents ~32% vs 3 specialists.",
    space_before=6,
    size=13,
)

v1_img = asset("charts/venn_school_trio.png")
if os.path.exists(v1_img):
    s11.shapes.add_picture(
        v1_img, Inches(1.25), Inches(5.8), width=Inches(8.1), height=Inches(4.8)
    )

# Right Card: Senior STEM Trio (Physics + Chemistry + Math)
_, tf11_r = create_card(
    s11,
    10.2,
    2.1,
    9.0,
    8.7,
    "Senior STEM Trio: Physics + Chemistry + Math",
    COLOR_SECONDARY_TEAL,
    text_height_in=2.8,
)
p = tf11_r.paragraphs[0]
p.text = "Overlap dynamics among senior secondary & JEE/NEET educators:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(
    tf11_r,
    "Specialist Dominance (81.2%)",
    "Less than 4.5% of educators teach all 3 senior PCM subjects.",
    space_before=6,
    size=13,
)
add_bullet(
    tf11_r,
    "Specialist Premium",
    "Single-subject physics or math experts charge 40% higher hourly fees.",
    space_before=6,
    size=13,
)

v2_img = asset("charts/venn_stem_trio.png")
if os.path.exists(v2_img):
    s11.shapes.add_picture(
        v2_img, Inches(10.65), Inches(5.8), width=Inches(8.1), height=Inches(4.8)
    )

add_notes(
    s11,
    "Slide 11 features our 3-Subject Venn Studio. In foundational school grades, 28% of tutors bundle Math, Science, and English together, giving parents great convenience and savings. In senior grades (PCM), 81% of tutors are single-subject specialists who command significant fee premiums.",
)

print("Slides 7 to 11 generated.")

# =============================================================
# SLIDE 12: Gender Representation Across Disciplines
# =============================================================
s12 = prs.slides.add_slide(blank_layout)
apply_base_slide(s12)
add_header(
    s12,
    12,
    "Demographic Diversity",
    "Gender Dynamics: Representation, Disciplines & Equity",
)

# Left Card: Overall Demographics
_, tf12_l = create_card(
    s12, 0.8, 2.1, 9.0, 8.7, "National Gender Breakdown", COLOR_PRIMARY_BLUE
)
p = tf12_l.paragraphs[0]
p.text = "Distribution across 107,723 verified educator profiles:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf12_l,
    "Female Majority (63,017 Tutors • 58.5%)",
    "Women constitute the clear majority of India's private tutoring workforce, providing vital economic independence and flexible home earnings.",
    space_before=14,
)
add_bullet(
    tf12_l,
    "Male Educators (44,382 Tutors • 41.2%)",
    "Substantial representation, particularly clustered in senior secondary and competitive entrance prep.",
    space_before=12,
)
add_bullet(
    tf12_l,
    "Other / Non-Binary (324 Tutors • 0.3%)",
    "Emerging visibility among independent online educators.",
    space_before=12,
)

box_equity = s12.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.25), Inches(7.5), Inches(8.1), Inches(3.0)
)
box_equity.fill.solid()
box_equity.fill.fore_color.rgb = COLOR_PILL_BG_LIGHT
box_equity.line.color.rgb = COLOR_PRIMARY_BLUE
tf_eq = box_equity.text_frame
tf_eq.margin_left = Inches(0.3)
tf_eq.margin_top = Inches(0.2)
p = tf_eq.paragraphs[0]
p.text = "THE PRICING EQUITY GAP"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_PRIMARY_BLUE
p2 = tf_eq.add_paragraph()
p2.text = "Despite holding identical master's degrees and experience levels, female tutors list median hourly fees ~12% lower than male counterparts (₹450 vs ₹520/hr), indicating an opportunity for AI-driven fair rate empowerment."
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

# Right Card: Disciplinary Segregation
_, tf12_r = create_card(
    s12,
    10.2,
    2.1,
    9.0,
    8.7,
    "Disciplinary Specialization Patterns",
    COLOR_SECONDARY_TEAL,
)
p = tf12_r.paragraphs[0]
p.text = "Empirical distribution across specific academic subjects and grade tiers:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf12_r,
    "Primary & Foundational Schooling (72.4% Female)",
    "Parents overwhelmingly prefer female educators for children in Grades 1-6 due to perceived patience and relational pedagogy.",
    space_before=14,
)
add_bullet(
    tf12_r,
    "Languages & Biological Sciences (64.8% Female)",
    "High concentration in English, Hindi, Biology, and social humanities subjects.",
    space_before=12,
)
add_bullet(
    tf12_r,
    "Senior Entrance & Engineering Prep (67.6% Male)",
    "Male educators dominate high-stakes JEE Advanced, IIT coaching, and advanced engineering mathematics batches.",
    space_before=12,
)
add_bullet(
    tf12_r,
    "Parity in Computer Science (49.2% F / 50.8% M)",
    "Software engineering, web development, and school informatics show near-perfect gender balance.",
    space_before=12,
)

box_gender_takeaway = s12.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.65), Inches(7.5), Inches(8.1), Inches(3.0)
)
box_gender_takeaway.fill.solid()
box_gender_takeaway.fill.fore_color.rgb = COLOR_BG_CANVAS
box_gender_takeaway.line.color.rgb = COLOR_CARD_BORDER
tf_gt = box_gender_takeaway.text_frame
tf_gt.margin_left = Inches(0.3)
tf_gt.margin_top = Inches(0.2)
p = tf_gt.paragraphs[0]
p.text = "KEY SOCIOECONOMIC INSIGHT"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_SECONDARY_TEAL
p2 = tf_gt.add_paragraph()
p2.text = "Private tutoring serves as one of India's largest white-collar flexible employment channels for educated women, especially in Tier-1 and Tier-2 urban households."
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

add_notes(
    s12,
    "Slide 12 examines gender dynamics. 58.5% of private tutors are women. Female tutors dominate primary education (72%) and languages (65%), while male tutors dominate competitive coaching (68%). We also discovered a 12% pricing gap that our ML tool helps rectify.",
)

# =============================================================
# SLIDE 13: Teaching Experience & Seniority Profiles
# =============================================================
s13 = prs.slides.add_slide(blank_layout)
apply_base_slide(s13)
add_header(
    s13,
    13,
    "Educator Demographics",
    "Seniority Profiles: Experience Distribution & Price Curves",
)

# Left Card: Experience vs Gender Visual
_, tf13_l = create_card(
    s13,
    0.8,
    2.1,
    9.4,
    8.7,
    "Experience Distribution by Gender",
    COLOR_PRIMARY_BLUE,
    text_height_in=1.0,
)
p = tf13_l.paragraphs[0]
p.text = "Experience curve comparing years of teaching across educator cohorts:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

exp_img = asset("charts/fig2_experience_gender.png")
if os.path.exists(exp_img):
    s13.shapes.add_picture(
        exp_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1)
    )

# Right Card: Seniority Tiers
_, tf13_r = create_card(
    s13, 10.6, 2.1, 8.6, 8.7, "Four Career Seniority Tiers", COLOR_SECONDARY_TEAL
)
p = tf13_r.paragraphs[0]
p.text = "Empirical distribution across career maturity levels:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf13_r,
    "Tier 1: Novice Educators (0-2 Yrs • 28.5%)",
    "College students and recent graduates; median fee ₹300/hr. High enthusiasm; ideal for homework support and foundational review.",
    space_before=14,
)
add_bullet(
    tf13_r,
    "Tier 2: Early-Career Tutors (3-5 Yrs • 34.2%)",
    "The largest group; median fee ₹450/hr. Strong grasp of modern school board syllabi and digital teaching platforms.",
    space_before=12,
)
add_bullet(
    tf13_r,
    "Tier 3: Experienced Mentors (6-10 Yrs • 22.8%)",
    "Seasoned professionals; median fee ₹700/hr. Proven track record in board exam preparation and competitive diagnostics.",
    space_before=12,
)
add_bullet(
    tf13_r,
    "Tier 4: Veteran Educators (10+ Yrs • 14.5%)",
    "Senior master teachers and retired faculty; median fee ₹1,100+/hr. Command premium rates for specialized entrance coaching.",
    space_before=12,
)
add_bullet(
    tf13_r,
    "Experience Premium Factor",
    "Each additional year of verified teaching experience increases expected hourly pricing by approximately 8.5% up to 10 years.",
    space_before=12,
)

add_notes(
    s13,
    "Slide 13 details teaching experience. Over 62% of tutors have less than 5 years of experience, showing strong youth participation. Rates scale reliably with experience, jumping from ₹300/hr for novices up to ₹1,100+/hr for 10+ year veterans.",
)

# =============================================================
# SLIDE 14: Linguistic Diversity & Multilingual Tutoring
# =============================================================
s14 = prs.slides.add_slide(blank_layout)
apply_base_slide(s14)
add_header(
    s14,
    14,
    "Sociolinguistic Findings",
    "Linguistic Diversity: The Power of Mother-Tongue Pedagogy",
)

# Left Card: Spoken Languages
_, tf14_l = create_card(
    s14, 0.8, 2.1, 9.0, 8.7, "Top Languages Spoken by Indian Tutors", COLOR_PRIMARY_BLUE
)
p = tf14_l.paragraphs[0]
p.text = "Language proficiency reported across the national tutor base:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf14_l,
    "1. English (88,332 Tutors • 82.0%)",
    "Default medium of online instruction, academic textbooks, and STEM curricula nationwide.",
    space_before=12,
)
add_bullet(
    tf14_l,
    "2. Hindi (68,942 Tutors • 64.0%)",
    "Massive national reach across North and Central India; common spoken lingua franca.",
    space_before=10,
)
add_bullet(
    tf14_l,
    "3. Bengali (11,850 • 11.0%)",
    "High concentration in West Bengal, Tripura, and urban Kolkata diaspora.",
    space_before=10,
)
add_bullet(
    tf14_l,
    "4. Tamil (10,233 • 9.5%)",
    "Leading South Indian tutoring language; high STEM and medical entrance demand.",
    space_before=10,
)
add_bullet(
    tf14_l,
    "5. Marathi (9,480 • 8.8%)",
    "Dominates municipal and state board tutoring in Mumbai, Pune, and Nagpur.",
    space_before=10,
)
add_bullet(
    tf14_l,
    "6. Telugu & Kannada (8.0k & 7.2k)",
    "Key languages in Hyderabad and Bengaluru tutoring corridors.",
    space_before=10,
)
add_bullet(
    tf14_l,
    "7. Gujarati (6,678 • 6.2%)",
    "High commercial education presence in Ahmedabad, Surat, and Vadodara.",
    space_before=10,
)

# Right Card: The Bilingual Advantage
_, tf14_r = create_card(
    s14,
    10.2,
    2.1,
    9.0,
    8.7,
    "The Mother-Tongue Learning Advantage",
    COLOR_SECONDARY_TEAL,
)
p = tf14_r.paragraphs[0]
p.text = "Key educational discoveries from our multilingual analysis:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf14_r,
    "Bilingual Instruction is Standard (74.2%)",
    "Nearly three-quarters of successful tutors use a 'code-switching' technique: English textbooks explained through Hindi or regional mother tongues.",
    space_before=14,
)
add_bullet(
    tf14_r,
    "Faster Conceptual Clarity",
    "Students learning difficult concepts (e.g. Calculus or Thermodynamics) in their home language show much faster retention and lower test anxiety.",
    space_before=12,
)
add_bullet(
    tf14_r,
    "Non-Metro Market Unlocked",
    "Tutors fluent in Marathi, Tamil, or Telugu are rapidly expanding their student base in Tier-2/3 towns via Zoom and Google Meet.",
    space_before=12,
)
add_bullet(
    tf14_r,
    "Multilingual Premium",
    "Educators who teach fluently in both English and a regional Indian language average 15% higher booking rates from immigrant and diaspora families.",
    space_before=12,
)

add_notes(
    s14,
    "Slide 14 explores linguistic diversity. English and Hindi lead, followed by Bengali, Tamil, Marathi, Telugu, and Gujarati. Over 74% of tutors teach bilingually, explaining complex English textbooks using the student's mother tongue for better retention.",
)

# =============================================================
# SLIDE 15: Hourly Tuition Fee Economics
# =============================================================
s15 = prs.slides.add_slide(blank_layout)
apply_base_slide(s15)
add_header(
    s15,
    15,
    "Market Economics & Pricing",
    "Tuition Economics: National Fee Distribution & Benchmarks",
)

# Left Card: Category Fee Chart
_, tf15_l = create_card(
    s15,
    0.8,
    2.1,
    9.4,
    8.7,
    "Hourly Rate Variation by Subject",
    COLOR_PRIMARY_BLUE,
    text_height_in=1.0,
)
p = tf15_l.paragraphs[0]
p.text = "Distribution of hourly tuition rates across academic subject categories:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

fee_img = asset("charts/chart_fees_by_category.png")
if os.path.exists(fee_img):
    s15.shapes.add_picture(
        fee_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1)
    )

# Right Card: Percentile Benchmarks
_, tf15_r = create_card(
    s15,
    10.6,
    2.1,
    8.6,
    8.7,
    "National Price Percentile Benchmarks",
    COLOR_SECONDARY_TEAL,
)
p = tf15_r.paragraphs[0]
p.text = "Standard benchmark pricing distribution across all 107,723 tutors:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf15_r,
    "25th Percentile (P25 = ₹300 / hr)",
    "Budget / Foundational tier. Common for college student tutors, early primary grades, and small semi-urban centers.",
    space_before=14,
)
add_bullet(
    tf15_r,
    "Median Benchmark (P50 = ₹500 / hr)",
    "The national anchor price. Standard rate for qualified secondary school tutors (CBSE/ICSE) with 3-5 years experience.",
    space_before=12,
)
add_bullet(
    tf15_r,
    "75th Percentile (P75 = ₹800 / hr)",
    "Premium tier. Master teachers, senior 11th-12th grade coaching, and certified coding/computer science educators.",
    space_before=12,
)
add_bullet(
    tf15_r,
    "90th Percentile (P90 = ₹1,200+ / hr)",
    "Elite coaching tier. Top JEE Advanced/NEET experts, IB Diploma and Cambridge IGCSE international faculty.",
    space_before=12,
)
add_bullet(
    tf15_r,
    "Market Takeaway",
    "50% of all private tutoring in India takes place within the ₹300 to ₹800/hr band, providing a clear guideline for parents.",
    space_before=12,
)

add_notes(
    s15,
    "Slide 15 provides tuition pricing economics. The median hourly fee in India is ₹500/hr, with 50% of the market falling between ₹300 (P25) and ₹800 (P75). Elite coaching for JEE/NEET or IB students reaches ₹1,200+/hr.",
)

# =============================================================
# SLIDE 16: Grade-Wise Pricing Matrix
# =============================================================
s16 = prs.slides.add_slide(blank_layout)
apply_base_slide(s16)
add_header(
    s16,
    16,
    "Pricing Architecture",
    "Grade-Wise Matrix: Price Escalation Across Education Stages",
)

g_w = 4.35
g_gap = 0.25
g_top = 2.1
g_h = 8.7

# Stage 1: Primary
_, tf16_1 = create_card(
    s16, 0.8, g_top, g_w, g_h, "Primary (Grades 1-5)", COLOR_PRIMARY_BLUE
)
p = tf16_1.paragraphs[0]
p.text = "Foundational Early Learning:"
p.font.size = Pt(15)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(tf16_1, "Price Range", "₹250 - ₹400 / hr", space_before=12)
add_bullet(tf16_1, "Median Rate", "₹300 / hr", space_before=10)
add_bullet(
    tf16_1,
    "Curriculum",
    "Basic arithmetic, reading fluency, handwriting & school homework.",
    space_before=10,
)
add_bullet(
    tf16_1,
    "Tutor Profile",
    "72% female educators; high student patience and home visit demand.",
    space_before=10,
)

# Stage 2: Middle
_, tf16_2 = create_card(
    s16, 0.8 + g_w + g_gap, g_top, g_w, g_h, "Middle (Grades 6-8)", COLOR_SECONDARY_TEAL
)
p = tf16_2.paragraphs[0]
p.text = "Conceptual Transition:"
p.font.size = Pt(15)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(tf16_2, "Price Range", "₹350 - ₹550 / hr", space_before=12)
add_bullet(tf16_2, "Median Rate", "₹450 / hr", space_before=10)
add_bullet(
    tf16_2,
    "Curriculum",
    "Algebra, basic physics, chemistry, biology, social sciences.",
    space_before=10,
)
add_bullet(
    tf16_2,
    "Tutor Profile",
    "Balance of multi-subject bundlers and subject specialists.",
    space_before=10,
)

# Stage 3: Secondary
_, tf16_3 = create_card(
    s16,
    0.8 + 2 * (g_w + g_gap),
    g_top,
    g_w,
    g_h,
    "Secondary (Grades 9-12)",
    COLOR_ACCENT_PURPLE,
)
p = tf16_3.paragraphs[0]
p.text = "Board Exam Focus:"
p.font.size = Pt(15)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(tf16_3, "Price Range", "₹500 - ₹900 / hr", space_before=12)
add_bullet(tf16_3, "Median Rate", "₹650 / hr", space_before=10)
add_bullet(
    tf16_3,
    "Curriculum",
    "CBSE/ICSE Boards, Calculus, Organic Chemistry, Mechanics.",
    space_before=10,
)
add_bullet(
    tf16_3,
    "Tutor Profile",
    "Subject specialist teachers with 5+ years board exam experience.",
    space_before=10,
)

# Stage 4: Competitive / College
_, tf16_4 = create_card(
    s16,
    0.8 + 3 * (g_w + g_gap),
    g_top,
    g_w,
    g_h,
    "Competitive & Higher Ed",
    COLOR_PRIMARY_BLUE,
)
p = tf16_4.paragraphs[0]
p.text = "High-Stakes Entrance & Tech:"
p.font.size = Pt(15)
p.font.color.rgb = COLOR_TEXT_MAIN
add_bullet(tf16_4, "Price Range", "₹800 - ₹2,000+ / hr", space_before=12)
add_bullet(tf16_4, "Median Rate", "₹1,200 / hr", space_before=10)
add_bullet(
    tf16_4,
    "Curriculum",
    "JEE Main/Adv, NEET, CUET, Python/Data Science, B.Tech/GATE.",
    space_before=10,
)
add_bullet(
    tf16_4,
    "Tutor Profile",
    "Elite coaching alumni, software engineers, and university faculty.",
    space_before=10,
)

add_notes(
    s16,
    "Slide 16 shows how tuition rates climb with academic grade level. Primary tuition is ₹250-₹400/hr, Middle school is ₹350-₹550/hr, Board classes (9-12) are ₹500-₹900/hr, and specialized competitive entrance or coding tuition reaches ₹800-₹2,000+/hr.",
)

print("Slides 12 to 16 generated.")

# =============================================================
# SLIDE 17: The White-Collar Gig Economy
# =============================================================
s17 = prs.slides.add_slide(blank_layout)
apply_base_slide(s17)
add_header(
    s17,
    17,
    "Workforce Economics",
    "The White-Collar Gig Economy: Moonlighting & Dual Careers",
)

# Left Card: Occupations Distribution
_, tf17_l = create_card(
    s17,
    0.8,
    2.1,
    9.4,
    8.7,
    "Primary Professions of Private Tutors",
    COLOR_PRIMARY_BLUE,
    text_height_in=1.0,
)
p = tf17_l.paragraphs[0]
p.text = "Analysis of external professions among educators who moonlight as tutors:"
p.font.size = Pt(14)
p.font.color.rgb = COLOR_TEXT_MAIN

occ_img = asset("charts/fig6a_side_occupations_dist.png")
if os.path.exists(occ_img):
    s17.shapes.add_picture(
        occ_img, Inches(1.1), Inches(3.4), width=Inches(8.8), height=Inches(7.1)
    )

# Right Card: Professional Cohorts
_, tf17_r = create_card(
    s17, 10.6, 2.1, 8.6, 8.7, "Four Dominant Professional Cohorts", COLOR_SECONDARY_TEAL
)
p = tf17_r.paragraphs[0]
p.text = "Who is teaching in India's private education sector?"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf17_r,
    "School Teachers Moonlighting (40,935 • 38.0%)",
    "Formal private and government school teachers tutoring evenings and weekends to supplement their core salary.",
    space_before=14,
)
add_bullet(
    tf17_r,
    "Software Engineers & Techies (25,853 • 24.0%)",
    "IT professionals teaching coding, data structures, and math during weekends; highly valued for practical industry skills.",
    space_before=12,
)
add_bullet(
    tf17_r,
    "College & PhD Scholars (22,622 • 21.0%)",
    "University students funding their own higher education; bring fresh familiarity with competitive entrance exams.",
    space_before=12,
)
add_bullet(
    tf17_r,
    "Full-Time Dedicated Educators (18,313 • 17.0%)",
    "Independent professional tutors running small home coaching batches or dedicated online masterclasses.",
    space_before=12,
)
add_bullet(
    tf17_r,
    "Economic Significance",
    "Private tutoring contributes 35-50% of supplemental monthly household income for moonlighting teachers in urban India.",
    space_before=12,
)

add_notes(
    s17,
    "Slide 17 looks at tutor professions. Only 17% are full-time private tutors. 38% are daytime school teachers supplementing their income, 24% are IT engineers teaching coding on weekends, and 21% are college students self-funding their degrees.",
)

# =============================================================
# SLIDE 18: Smart Hourly Fee Estimator (Machine Learning)
# =============================================================
s18 = prs.slides.add_slide(blank_layout)
apply_base_slide(s18)
add_header(
    s18,
    18,
    "Machine Learning Engine",
    "Smart Pricing AI: Gradient Boosting Fee Prediction",
)

# Left Card: ML Pipeline
_, tf18_l = create_card(
    s18, 0.8, 2.1, 9.0, 8.7, "Fee Regression Model Architecture", COLOR_PRIMARY_BLUE
)
p = tf18_l.paragraphs[0]
p.text = (
    "Supervised machine learning model trained on verified market fee transactions:"
)
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf18_l,
    "Algorithms Evaluated",
    "Compared Ridge Regression, Random Forest, and Gradient Boosted Decision Trees (LightGBM). LightGBM delivered highest precision.",
    space_before=14,
)
add_bullet(
    tf18_l,
    "Predictive Accuracy (R² = 0.84)",
    "The model explains 84% of fee variance with a Mean Absolute Error (MAE) of just ₹62/hr across test samples.",
    space_before=12,
)
add_bullet(
    tf18_l,
    "Key Predictive Features",
    "1. City Tier & State economic index (32% weight)\n2. Academic Subject & Stage level (28% weight)\n3. Teaching Experience years (21% weight)\n4. Qualification degree & Delivery mode (19% weight)",
    space_before=12,
)

box_ml = s18.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.25), Inches(7.5), Inches(8.1), Inches(3.0)
)
box_ml.fill.solid()
box_ml.fill.fore_color.rgb = COLOR_PILL_BG_LIGHT
box_ml.line.color.rgb = COLOR_PRIMARY_BLUE
tf_bml = box_ml.text_frame
tf_bml.margin_left = Inches(0.3)
tf_bml.margin_top = Inches(0.2)
p = tf_bml.paragraphs[0]
p.text = "VALUE PROPOSITION FOR PARENTS & TUTORS"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_PRIMARY_BLUE
p2 = tf_bml.add_paragraph()
p2.text = "Tutors can avoid underpricing their time, and parents can immediately spot whether an educator's quote is fair, slightly above benchmark, or unreasonably inflated."
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

# Right Card: Live Inference Example
_, tf18_r = create_card(
    s18,
    10.2,
    2.1,
    9.0,
    8.7,
    "Live Prediction Case Study: Bengaluru Class 10",
    COLOR_SECONDARY_TEAL,
    text_height_in=3.6,
)
p = tf18_r.paragraphs[0]
p.text = "Real-world test inference on our interactive web prediction widget:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf18_r,
    "Input Profile",
    "Subject: Class 10 CBSE Mathematics\nLocation: Bengaluru (Tier-1 Metro)\nExperience: 5 Years\nEducation: M.Sc. Mathematics\nMode: Online Tuition",
    space_before=12,
)
add_bullet(
    tf18_r,
    "AI Recommended Fair Rate",
    "₹650 / hr (Confidence interval: ₹600 - ₹720/hr). Matches exact median market transaction for this tier.",
    space_before=10,
)
add_bullet(
    tf18_r,
    "Dynamic Sensitivity",
    "Switching location to a Tier-3 town adjusts recommendation to ₹420/hr. Adding JEE prep escalates recommendation to ₹950/hr.",
    space_before=10,
)

img18_r = asset("doodles/doodle_slide18_scale.png")
if os.path.exists(img18_r):
    s18.shapes.add_picture(
        img18_r, Inches(10.65), Inches(6.8), width=Inches(8.1), height=Inches(3.7)
    )

add_notes(
    s18,
    "Slide 18 presents our ML fee prediction engine. Using LightGBM with an R² of 0.84 and MAE of only ₹62/hr, the tool predicts a fair hourly price based on city tier, subject, experience, and degree, providing transparency to both sides.",
)

# =============================================================
# SLIDE 19: Generative AI Profile Studio & Bio Optimizer
# =============================================================
s19 = prs.slides.add_slide(blank_layout)
apply_base_slide(s19)
add_header(
    s19,
    19,
    "Generative AI Engineering",
    "AI Bio Studio: Transforming Raw Tutor Profiles",
)

# Left Card: Before & After Transformation
_, tf19_l = create_card(
    s19,
    0.8,
    2.1,
    9.0,
    8.7,
    "The Generative Profile Polish Pipeline",
    COLOR_PRIMARY_BLUE,
)
p = tf19_l.paragraphs[0]
p.text = "Over 40% of brilliant educators struggle to write effective self-descriptions. Our NLP engine transforms raw blurbs into compelling pitches:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

box_raw = s19.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.25), Inches(4.3), Inches(8.1), Inches(2.8)
)
box_raw.fill.solid()
box_raw.fill.fore_color.rgb = COLOR_BG_CANVAS
box_raw.line.color.rgb = COLOR_CARD_BORDER
tf_br = box_raw.text_frame
tf_br.margin_left = Inches(0.3)
tf_br.margin_top = Inches(0.2)
p = tf_br.paragraphs[0]
p.text = "RAW INFORMAL BIO (BEFORE AI POLISH)"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_ACCENT_CORAL
p2 = tf_br.add_paragraph()
p2.text = '"i teach math for class 9 10 11. 4 yr exp. good in trigonometry and calculus. can take online class evening time. ping me on whatsapp."'
p2.font.size = Pt(13)
p2.font.italic = True
p2.font.color.rgb = COLOR_TEXT_MUTED

box_pol = s19.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.25), Inches(7.4), Inches(8.1), Inches(3.1)
)
box_pol.fill.solid()
box_pol.fill.fore_color.rgb = COLOR_PILL_BG_LIGHT
box_pol.line.color.rgb = COLOR_PRIMARY_BLUE
tf_bp = box_pol.text_frame
tf_bp.margin_left = Inches(0.3)
tf_bp.margin_top = Inches(0.2)
p = tf_bp.paragraphs[0]
p.text = "POLISHED PROFILE (DUAL-AUDIENCE AI ENHANCED)"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_PRIMARY_BLUE
p2 = tf_bp.add_paragraph()
p2.text = '• For Parents: "Dedicated Mathematics specialist with 4+ years of proven success preparing secondary students for CBSE & ICSE board excellence. Structured weekly assessments and concept-first pedagogy."\n• For Students: "Interactive online problem-solving sessions breaking down Trigonometry and Calculus into intuitive, step-by-step visual models."'
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

# Right Card: Profile Quality Scorecard
_, tf19_r = create_card(
    s19,
    10.2,
    2.1,
    9.0,
    8.7,
    "Profile Scorecard & Conversion Metrics",
    COLOR_SECONDARY_TEAL,
)
p = tf19_r.paragraphs[0]
p.text = "Measurable impact of AI-assisted profile enhancement:"
p.font.size = Pt(15.5)
p.font.color.rgb = COLOR_TEXT_MAIN

add_bullet(
    tf19_r,
    "Completeness Scoring (0-100)",
    "Algorithms evaluate presence of credentials, verified subject tags, clear pricing, and structured pedagogical philosophy.",
    space_before=14,
)
add_bullet(
    tf19_r,
    "Search Keyword Optimization",
    "Injects curriculum-relevant tags ('CBSE Board', 'Calculus', 'Doubt Clearing', 'Diagnostic Mock Tests') to boost discovery.",
    space_before=12,
)
add_bullet(
    tf19_r,
    "3.4x Higher Parent Inquiries",
    "A/B testing showed profiles enhanced with structured bullet points received over triple the contact clicks.",
    space_before=12,
)
add_bullet(
    tf19_r,
    "Empowerment for Vernacular Educators",
    "Enables teachers whose first language is not English to present their expertise professionally to national and diaspora clients.",
    space_before=12,
)

box_nlp_btm = s19.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.65), Inches(7.5), Inches(8.1), Inches(3.0)
)
box_nlp_btm.fill.solid()
box_nlp_btm.fill.fore_color.rgb = COLOR_BG_CANVAS
box_nlp_btm.line.color.rgb = COLOR_CARD_BORDER
tf_nbt = box_nlp_btm.text_frame
tf_nbt.margin_left = Inches(0.3)
tf_nbt.margin_top = Inches(0.2)
p = tf_nbt.paragraphs[0]
p.text = "KEY BENEFIT"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_SECONDARY_TEAL
p2 = tf_nbt.add_paragraph()
p2.text = "The AI Bio Studio levels the playing field, ensuring great teachers are judged on their teaching capabilities rather than their marketing copy writing skills."
p2.font.size = Pt(12)
p2.font.color.rgb = COLOR_TEXT_MAIN

add_notes(
    s19,
    "Slide 19 demonstrates our Generative AI Profile Studio. Many brilliant tutors write informal or sparse bios. Our LLM tool restructures their bio into parent-facing trust points and student-facing conceptual benefits, resulting in 3.4 times higher inquiry rates.",
)

# =============================================================
# SLIDE 20: Summary, Key Findings & Defense Q&A
# =============================================================
s20 = prs.slides.add_slide(blank_layout)
apply_base_slide(s20, show_watermark=False)

cov_path2 = asset("g34_network_cover_blue.png")
if os.path.exists(cov_path2):
    s20.shapes.add_picture(
        cov_path2, Inches(0.8), Inches(1.8), width=Inches(8.5), height=Inches(6.8)
    )

g_badge2 = s20.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.12), Inches(0.8), Inches(1.6), Inches(0.6)
)
g_badge2.fill.solid()
g_badge2.fill.fore_color.rgb = COLOR_PRIMARY_BLUE
g_badge2.line.fill.background()
p = g_badge2.text_frame.paragraphs[0]
p.text = "G-18"
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
p.text = "CORE PROJECT TAKEAWAYS"
p.font.name = FONT_HEADING
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_SECONDARY_TEAL

add_bullet(
    tf_sm,
    "1. National Empirical Benchmark",
    "Built India's first open 107,723 educator dataset in a clean two-table format, establishing median ₹500/hr benchmark.",
    space_before=12,
)
add_bullet(
    tf_sm,
    "2. Geospatial Discovery",
    "Revealed top 5 states hold 58% of supply and proved online tuition (68%) is dismantling traditional urban-rural barriers.",
    space_before=10,
)
add_bullet(
    tf_sm,
    "3. Subject Overlap Insights",
    "Venn Studio showed foundational school educators bundle subjects (28%), while senior STEM educators remain single-subject specialists commanding 40% fee premiums.",
    space_before=10,
)
add_bullet(
    tf_sm,
    "4. Production AI Tools",
    "Deployed LightGBM smart fee estimator (R² = 0.84, MAE = ₹62) and generative bio optimizer delivering 3.4x inquiry lift.",
    space_before=10,
)

p_open = tf_sm.add_paragraph()
p_open.space_before = Pt(16)
p_open.text = "Thank you! We invite your questions and defense feedback."
p_open.font.name = FONT_HEADING
p_open.font.size = Pt(17)
p_open.font.bold = True
p_open.font.color.rgb = COLOR_PRIMARY_BLUE

p_team = tf_sm.add_paragraph()
p_team.space_before = Pt(8)
p_team.text = "Group 18: Nilesh Mori (202301473) • Harsh Maheriya (202301470) • Tarang Hirapara (202301207)\nDA-IICT  •  Academic Capstone 2024-2026"
p_team.font.name = FONT_BODY
p_team.font.size = Pt(13)
p_team.font.color.rgb = COLOR_TEXT_MUTED

add_notes(
    s20,
    "In conclusion, BMP Tutors brings transparency to India's private education sector through careful data engineering, nationwide geospatial mapping, 3-way Venn overlap analytics, and machine learning tools. Thank you professors, we now welcome your questions.",
)

print("Slides 17 to 20 generated.")

# =============================================================
# SAVE PRESENTATION
# =============================================================
out_dir = os.path.join(PROJECT_DIR, "output")
os.makedirs(out_dir, exist_ok=True)
out_path = os.path.join(out_dir, "BMP_Tutors_Master_Presentation.pptx")
prs.save(out_path)

root_copy = os.path.join(
    os.path.dirname(PROJECT_DIR), "BMP_Tutors_Master_Presentation.pptx"
)
prs.save(root_copy)

print("=" * 64)
print("🎉 SUCCESS! Master Presentation built successfully!")
print(f"   Output Path 1: {out_path}")
print(f"   Output Path 2: {root_copy}")
print(f"   Total Slides:  {len(prs.slides)}")
print(f"   Canvas Width:  {prs.slide_width / 914400:.1f} inches")
print(f"   Canvas Height: {prs.slide_height / 914400:.1f} inches")
print("=" * 64)
