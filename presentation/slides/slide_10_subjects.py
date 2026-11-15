"""
Slide 10: Academic Subject Demand & Top Disciplines.
Matches Group 18 reference: Macro domain chart on left + Top subjects breakdown on right.
"""

import os
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
    add_image_fit,
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
        category_pill_text="SUBJECT DEMAND",
        slide_num_str="10 / 20",
        top_tag="ACADEMIC OFFERINGS",
        assets_dir=assets_dir,
        pill_color=ACCENT_INDIGO,
    )

    add_slide_header(
        slide,
        title="Academic Subject Demand — What Do Indian Students Study?",
        subtitle="School foundational subjects and STEM dominate private tutoring, led by Mathematics, English, and General Science.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT COLUMN: MACRO DOMAIN BAR CHART (Width: 6.5in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(6.40)

    add_card_box(slide, left_x, top_y, left_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, left_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.20), left_w - Inches(0.6), Inches(0.25), text="MACRO DOMAIN DISTRIBUTION", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.48), left_w - Inches(0.6), Inches(0.35), text="STEM & Languages Command 72% of Market", font_size=15, font_color=TEXT_DARK, bold=True)

    chart_path = os.path.join(assets_dir, "charts", "fig1_subjects_macro.png")
    if os.path.exists(chart_path):
        add_image_fit(
            slide,
            chart_path,
            left=left_x + Inches(0.20),
            top=top_y + Inches(0.90),
            max_width=left_w - Inches(0.40),
            max_height=Inches(3.40),
            align_center=True,
            valign_center=True,
        )

    # -------------------------------------------------------------
    # RIGHT COLUMN: TOP 8 INDIVIDUAL SUBJECTS (Width: 5.6in)
    # -------------------------------------------------------------
    right_x = Inches(7.15)
    right_w = Inches(5.60)

    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, right_w, Inches(0.08), ACCENT_CORAL, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), right_w - Inches(0.6), Inches(0.25), text="TOP TUTORING SUBJECTS", font_size=11, font_color=ACCENT_CORAL, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), right_w - Inches(0.6), Inches(0.35), text="Single Subject Rankings", font_size=15, font_color=TEXT_DARK, bold=True)

    subjects = [
        ("1. Mathematics", "37,566 Tutors", "34.9% of all tutors", "₹790 / hr"),
        ("2. English", "24,393 Tutors", "22.6% of all tutors", "₹733 / hr"),
        ("3. Science (General)", "23,229 Tutors", "21.6% of all tutors", "₹758 / hr"),
        ("4. Physics", "14,732 Tutors", "13.7% of all tutors", "₹831 / hr"),
        ("5. Biology & Life Sciences", "10,952 Tutors", "10.2% of all tutors", "₹786 / hr"),
        ("6. Chemistry", "9,840 Tutors", "9.1% of all tutors", "₹820 / hr"),
        ("7. Hindi (Language)", "9,601 Tutors", "8.9% of all tutors", "₹640 / hr"),
        ("8. Computer Science & Coding", "6,475 Tutors", "6.0% of all tutors", "₹920 / hr"),
    ]

    for j, (sub_name, sub_count, sub_share, sub_fee) in enumerate(subjects):
        sy = top_y + Inches(0.92 + j * 0.44)
        add_textbox(slide, right_x + Inches(0.30), sy, Inches(2.2), Inches(0.25), text=sub_name, font_size=11, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, right_x + Inches(2.55), sy, Inches(1.5), Inches(0.25), text=f"{sub_count} ({sub_share.split()[0]})", font_size=10, font_color=TEXT_LIGHT)
        add_pill(slide, right_x + Inches(4.30), sy, Inches(0.95), Inches(0.22), sub_fee, PILL_BLUE, text_color=ACCENT_INDIGO, font_size=8.5, bold=True)

    # Key insight card at bottom
    add_card_box(slide, right_x + Inches(0.20), top_y + Inches(4.15), right_w - Inches(0.40), Inches(0.50), fill_color=BG_SUBTLE)
    add_textbox(slide, right_x + Inches(0.28), top_y + Inches(4.20), right_w - Inches(0.56), Inches(0.40), text="💡 Takeaway: 1 out of every 3 tutors in India teaches Mathematics, making it the single largest anchor discipline.", font_size=9.5, font_color=TEXT_MEDIUM)
