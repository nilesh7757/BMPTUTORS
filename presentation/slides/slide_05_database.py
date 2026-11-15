"""
Slide 5: Relational Normalization (3NF) — Two Clean Tables.
Matches Group 18 reference: Relational entity cards + 3 takeaway benefit cards.
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
    BORDER_SOFT,
    FONT_TITLE,
    FONT_BODY,
)


def add_slide(prs, assets_dir):
    slide, _ = add_slide_base(
        prs,
        category_pill_text="DATABASE DESIGN",
        slide_num_str="05 / 20",
        top_tag="3NF NORMALIZATION",
        assets_dir=assets_dir,
        pill_color=ACCENT_INDIGO,
    )

    add_slide_header(
        slide,
        title="Relational Normalization — Eliminating 327,000 Repeated Rows",
        subtitle="Instead of repeating tutor names and fees thousands of times in a giant sheet, we normalized into two clean tables.",
    )

    # -------------------------------------------------------------
    # TOP SECTION: THE TWO NORMALIZED TABLES (Side-by-side)
    # -------------------------------------------------------------
    top_y = Inches(2.05)
    tbl_h = Inches(2.60)
    tbl_w = Inches(5.95)

    # Table 1: Tutors Master
    t1_x = Inches(0.55)
    add_card_box(slide, t1_x, top_y, tbl_w, tbl_h, fill_color=WHITE)
    add_rectangle(slide, t1_x, top_y, tbl_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_pill(slide, t1_x + Inches(0.30), top_y + Inches(0.20), Inches(1.8), Inches(0.30), "PRIMARY ENTITY", PILL_BLUE, text_color=ACCENT_INDIGO, font_size=9, bold=True)
    add_textbox(slide, t1_x + Inches(2.25), top_y + Inches(0.18), Inches(3.4), Inches(0.32), text="107,723 Unique Rows", font_size=12, font_color=TEXT_LIGHT, bold=True, align=PP_ALIGN.RIGHT)
    add_textbox(slide, t1_x + Inches(0.30), top_y + Inches(0.58), tbl_w - Inches(0.6), Inches(0.35), text="Table 1: Tutors Master (Tutors_Master)", font_size=16, font_color=TEXT_DARK, bold=True)

    t1_desc = (
        "• Primary Key: Tutor_ID (T1, T2, T3...)\n"
        "• Stores educator personal facts exactly once: Name, City, State, Tier, Gender, Experience.\n"
        "• Hourly Fee Avg, Verification Status, Profile URL.\n"
        "• Relational Foreign Key: Subject_IDs_Taught (e.g., SUB-00001; SUB-00003)."
    )
    add_textbox(slide, t1_x + Inches(0.30), top_y + Inches(0.98), tbl_w - Inches(0.6), Inches(1.50), text=t1_desc, font_size=11, font_color=TEXT_MEDIUM, line_spacing=1.2)

    # Table 2: Subjects Master
    t2_x = Inches(6.80)
    add_card_box(slide, t2_x, top_y, tbl_w, tbl_h, fill_color=WHITE)
    add_rectangle(slide, t2_x, top_y, tbl_w, Inches(0.08), ACCENT_EMERALD, shape=MSO_SHAPE.RECTANGLE)

    add_pill(slide, t2_x + Inches(0.30), top_y + Inches(0.20), Inches(1.8), Inches(0.30), "LOOKUP ENTITY", PILL_MINT, text_color=ACCENT_EMERALD, font_size=9, bold=True)
    add_textbox(slide, t2_x + Inches(2.25), top_y + Inches(0.18), Inches(3.4), Inches(0.32), text="12,327 Unique Rows", font_size=12, font_color=TEXT_LIGHT, bold=True, align=PP_ALIGN.RIGHT)
    add_textbox(slide, t2_x + Inches(0.30), top_y + Inches(0.58), tbl_w - Inches(0.6), Inches(0.35), text="Table 2: Subjects Master (Subjects_Master)", font_size=16, font_color=TEXT_DARK, bold=True)

    t2_desc = (
        "• Primary Key: Subject_ID (SUB-00001, SUB-00002...)\n"
        "• Stores academic discipline facts: Canonical Name, Domain (Science, Coding, Languages).\n"
        "• Target Educational Level, Target Competitive Exam.\n"
        "• Market-wide benchmarks: Total Tutors Offering, Avg Hourly Fee, Min/Max Fee."
    )
    add_textbox(slide, t2_x + Inches(0.30), top_y + Inches(0.98), tbl_w - Inches(0.6), Inches(1.50), text=t2_desc, font_size=11, font_color=TEXT_MEDIUM, line_spacing=1.2)

    # -------------------------------------------------------------
    # BOTTOM SECTION: 3 KEY NORMALIZATION BENEFITS
    # -------------------------------------------------------------
    bot_y = Inches(4.85)
    card_h = Inches(2.00)
    card_w = Inches(3.85)

    benefits = [
        (
            "01",
            "Zero Data Redundancy",
            "In the old sheet, a tutor teaching 5 subjects had their full bio and phone repeated 5 times across 327k rows. Now, every tutor appears exactly once.",
            ACCENT_CORAL,
        ),
        (
            "02",
            "100% Relational Integrity",
            "Every subject offering is mapped directly to a verified Subject_ID. There are zero broken references, mismatched spelling variants, or orphaned records.",
            ACCENT_INDIGO,
        ),
        (
            "03",
            "70% File Size Reduction",
            "Eliminating redundant text compressed the spreadsheet footprint from an unwieldy 80 MB flat file to a fast, clean 24 MB normalized workbook.",
            ACCENT_EMERALD,
        ),
    ]

    for i, (bnum, btitle, bbody, baccent) in enumerate(benefits):
        bx = Inches(0.55 + i * (3.85 + 0.32))
        add_card_box(slide, bx, bot_y, card_w, card_h, fill_color=BG_SUBTLE)
        add_rectangle(slide, bx, bot_y, card_w, Inches(0.06), baccent, shape=MSO_SHAPE.RECTANGLE)

        add_pill(slide, bx + Inches(0.20), bot_y + Inches(0.20), Inches(0.50), Inches(0.32), bnum, baccent, text_color=WHITE, font_size=11, bold=True)
        add_textbox(slide, bx + Inches(0.80), bot_y + Inches(0.20), card_w - Inches(1.0), Inches(0.30), text=btitle, font_size=13, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, bx + Inches(0.20), bot_y + Inches(0.62), card_w - Inches(0.4), Inches(1.25), text=bbody, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)
