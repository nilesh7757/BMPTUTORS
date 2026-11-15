"""
Slide 13: Teaching Experience Profile & Seniority Escalation.
Matches Group 18 reference: Experience chart on left + 4 seniority tiers on right.
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
        category_pill_text="EXPERIENCE",
        slide_num_str="13 / 20",
        top_tag="CAREER SENIORITY",
        assets_dir=assets_dir,
        pill_color=ACCENT_INDIGO,
    )

    add_slide_header(
        slide,
        title="Teaching Experience — Young Graduates vs Veteran Educators",
        subtitle="Private tutoring is driven by young educators, but hourly rates increase steadily as teaching experience deepens.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT COLUMN: EXPERIENCE PYRAMID CHART (Width: 6.4in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(6.35)

    add_card_box(slide, left_x, top_y, left_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, left_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.20), left_w - Inches(0.6), Inches(0.25), text="EXPERIENCE DISTRIBUTION", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.48), left_w - Inches(0.6), Inches(0.35), text="Tenure Breakdown by Gender", font_size=15, font_color=TEXT_DARK, bold=True)

    chart_path = os.path.join(assets_dir, "charts", "fig2_experience_gender.png")
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
    # RIGHT COLUMN: 4 SENIORITY TIERS (Width: 5.7in)
    # -------------------------------------------------------------
    right_x = Inches(7.10)
    right_w = Inches(5.65)

    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, right_w, Inches(0.08), ACCENT_EMERALD, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), right_w - Inches(0.6), Inches(0.25), text="SENIORITY & PRICING ESCALATION", font_size=11, font_color=ACCENT_EMERALD, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), right_w - Inches(0.6), Inches(0.35), text="The Experience Premium", font_size=15, font_color=TEXT_DARK, bold=True)

    tiers = [
        ("Early Career (0–3 Yrs)", "38,420 Tutors (35.7%)", "₹540 / hr", "Recent university graduates and college seniors offering budget-friendly foundational coaching.", ACCENT_INDIGO, PILL_BLUE),
        ("Mid-Career (4–8 Yrs)", "42,150 Tutors (39.1%)", "₹780 / hr", "The core backbone: full-time professional tutors with strong board syllabus mastery.", ACCENT_EMERALD, PILL_MINT),
        ("Experienced (9–15 Yrs)", "18,340 Tutors (17.0%)", "₹1,050 / hr", "Senior school teachers, department heads, and entrance exam specialists.", ACCENT_AMBER, PILL_CORAL),
        ("Veteran Masters (16+ Yrs)", "8,813 Tutors (8.2%)", "₹1,420 / hr", "Elite mentors and board examiners charging top-of-market hourly rates.", ACCENT_CORAL, PILL_BLUE),
    ]

    for j, (tname, tvolume, tfee, tdesc, taccent, tpill) in enumerate(tiers):
        ty = top_y + Inches(0.95 + j * 0.90)
        add_textbox(slide, right_x + Inches(0.30), ty, Inches(2.2), Inches(0.24), text=tname, font_size=11.5, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, right_x + Inches(2.55), ty, Inches(1.5), Inches(0.24), text=tvolume, font_size=9.5, font_color=TEXT_LIGHT)
        add_pill(slide, right_x + Inches(4.30), ty, Inches(1.00), Inches(0.22), tfee, tpill, text_color=taccent, font_size=8.5, bold=True)
        add_textbox(slide, right_x + Inches(0.30), ty + Inches(0.25), right_w - Inches(0.60), Inches(0.50), text=tdesc, font_size=9.5, font_color=TEXT_MEDIUM, line_spacing=1.1)
