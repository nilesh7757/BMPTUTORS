"""
Slide 4: Data Cleaning & Live Verification Methodology.
Matches Group 18 reference: 2 large container cards with clear bullet comparisons & KPI chips.
"""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    add_slide_base,
    add_slide_header,
    add_rectangle,
    add_textbox,
    add_pill,
    add_card_box,
    ACCENT_CORAL,
    ACCENT_INDIGO,
    ACCENT_EMERALD,
    ACCENT_AMBER,
    TEXT_DARK,
    TEXT_MEDIUM,
    TEXT_LIGHT,
    WHITE,
    PILL_BLUE,
    PILL_MINT,
    PILL_CORAL,
    BG_SUBTLE,
    FONT_TITLE,
    FONT_BODY,
)


def add_slide(prs, assets_dir):
    slide, _ = add_slide_base(
        prs,
        category_pill_text="DATA PIPELINE",
        slide_num_str="04 / 20",
        top_tag="HOW WE CLEANED THE DATA",
        assets_dir=assets_dir,
        pill_color=ACCENT_EMERALD,
    )

    add_slide_header(
        slide,
        title="Cleaning the Data — Removing Ghost Accounts and Bad Entries",
        subtitle="Raw tutor directories have test accounts and abandoned profiles. Here is how we audited every one.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT CARD: VERIFICATION AUDIT & COHORTS (Width: 5.9in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(5.85)

    add_card_box(slide, left_x, top_y, left_w, card_h, fill_color=WHITE)
    add_rectangle(
        slide,
        left_x,
        top_y,
        left_w,
        Inches(0.10),
        ACCENT_EMERALD,
        shape=MSO_SHAPE.RECTANGLE,
    )

    add_textbox(
        slide,
        left_x + Inches(0.35),
        top_y + Inches(0.25),
        left_w - Inches(0.7),
        Inches(0.32),
        text="TUTOR STATUS BREAKDOWN",
        font_size=11,
        font_color=ACCENT_EMERALD,
        bold=True,
    )
    add_textbox(
        slide,
        left_x + Inches(0.35),
        top_y + Inches(0.58),
        left_w - Inches(0.7),
        Inches(0.38),
        text="Which Tutors Are Still Active",
        font_size=16,
        font_color=TEXT_DARK,
        bold=True,
    )

    # 3 Activity Breakdown Mini Cards
    tiers = [
        (
            "76,706 Tutors (71.2%)",
            "VERIFIED ACTIVE",
            "Tutors whose last login differs from registration — they still log in and respond to students.",
            ACCENT_EMERALD,
            PILL_MINT,
        ),
        (
            "31,017 Tutors (28.8%)",
            "DORMANT OR REMOVED",
            "Accounts created once and never used, or profiles that no longer exist. Kept in the archive but not in pricing.",
            ACCENT_AMBER,
            PILL_CORAL,
        ),
        (
            "1,268 Tutors (1.2%)",
            "NO ACTIVITY DATA",
            "Profiles with missing login dates, verified via headless browser automation.",
            ACCENT_CORAL,
            PILL_BLUE,
        ),
    ]

    for j, (count_label, badge_label, desc, bar_col, pill_bg) in enumerate(tiers):
        item_y = top_y + Inches(1.10 + j * 1.15)
        add_card_box(
            slide,
            left_x + Inches(0.30),
            item_y,
            left_w - Inches(0.60),
            Inches(1.05),
            fill_color=BG_SUBTLE,
        )
        add_rectangle(
            slide,
            left_x + Inches(0.30),
            item_y,
            Inches(0.08),
            Inches(1.05),
            bar_col,
            shape=MSO_SHAPE.RECTANGLE,
        )

        add_textbox(
            slide,
            left_x + Inches(0.48),
            item_y + Inches(0.12),
            Inches(3.2),
            Inches(0.26),
            text=count_label,
            font_size=13,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_pill(
            slide,
            left_x + Inches(3.60),
            item_y + Inches(0.12),
            Inches(1.8),
            Inches(0.24),
            badge_label,
            pill_bg,
            text_color=bar_col,
            font_size=8,
            bold=True,
        )
        add_textbox(
            slide,
            left_x + Inches(0.48),
            item_y + Inches(0.42),
            left_w - Inches(1.0),
            Inches(0.55),
            text=desc,
            font_size=10.5,
            font_color=TEXT_MEDIUM,
            line_spacing=1.1,
        )

    # -------------------------------------------------------------
    # RIGHT CARD: DATA HYGIENE & OUTLIER RULES (Width: 6.0in)
    # -------------------------------------------------------------
    right_x = Inches(6.70)
    right_w = Inches(6.05)

    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)
    add_rectangle(
        slide,
        right_x,
        top_y,
        right_w,
        Inches(0.10),
        ACCENT_INDIGO,
        shape=MSO_SHAPE.RECTANGLE,
    )

    add_textbox(
        slide,
        right_x + Inches(0.35),
        top_y + Inches(0.25),
        right_w - Inches(0.7),
        Inches(0.32),
        text="CLEANING RULES WE APPLIED",
        font_size=11,
        font_color=ACCENT_INDIGO,
        bold=True,
    )
    add_textbox(
        slide,
        right_x + Inches(0.35),
        top_y + Inches(0.58),
        right_w - Inches(0.7),
        Inches(0.38),
        text="Fixing 107k Records",
        font_size=16,
        font_color=TEXT_DARK,
        bold=True,
    )

    rules = [
        (
            "Fake Fees (₹50 to ₹5,000/hr)",
            "Tossed out dummy entries like ₹99,999/hr or ₹0/hr using statistical rules so fee averages stay realistic.",
        ),
        (
            "Subject Name Mapping",
            "Raw listings had thousands of spelling variants ('Math', 'Maths', 'Mathematics 10th'). We grouped them into 12,327 standard subjects.",
        ),
        (
            "City Tiers & Regions",
            "Sorted every educator into Tier 1 (Metros), Tier 2, or Tier 3 / Semi-Urban, using Census city classifications.",
        ),
        (
            "Class Levels",
            "Split offerings into 4 stages: Primary (Class 1–5), Middle (Class 6–8), Senior Secondary (Class 9–12), and Higher Ed / Competitive Exams.",
        ),
    ]

    for j, (rtitle, rdesc) in enumerate(rules):
        rule_y = top_y + Inches(1.10 + j * 0.88)
        add_pill(
            slide,
            right_x + Inches(0.35),
            rule_y,
            Inches(0.40),
            Inches(0.28),
            str(j + 1),
            ACCENT_INDIGO,
            text_color=WHITE,
            font_size=11,
            bold=True,
        )
        add_textbox(
            slide,
            right_x + Inches(0.90),
            rule_y,
            right_w - Inches(1.1),
            Inches(0.25),
            text=rtitle,
            font_size=12,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_textbox(
            slide,
            right_x + Inches(0.90),
            rule_y + Inches(0.26),
            right_w - Inches(1.1),
            Inches(0.52),
            text=rdesc,
            font_size=10.5,
            font_color=TEXT_MEDIUM,
            line_spacing=1.1,
        )
