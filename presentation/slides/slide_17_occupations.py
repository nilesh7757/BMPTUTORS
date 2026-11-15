"""
Slide 17: Tutor Occupations & The Knowledge Gig Economy.
Matches Group 18 reference: Gig occupations chart on left + Professional background breakdown on right.
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
        category_pill_text="LABOR MARKET",
        slide_num_str="17 / 20",
        top_tag="GIG ECONOMY",
        assets_dir=assets_dir,
        pill_color=ACCENT_EMERALD,
    )

    add_slide_header(
        slide,
        title="Who Are India's Tutors? — The White-Collar Gig Economy",
        subtitle="Private tuition is no longer just for school teachers: IT engineers, college students, and corporate analysts actively teach.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT COLUMN: GIG OCCUPATIONS CHART (Width: 6.4in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(6.35)

    add_card_box(slide, left_x, top_y, left_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, left_w, Inches(0.08), ACCENT_EMERALD, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.20), left_w - Inches(0.6), Inches(0.25), text="NON-TEACHING OCCUPATIONS LANDSCAPE", font_size=11, font_color=ACCENT_EMERALD, bold=True)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.48), left_w - Inches(0.6), Inches(0.35), text="Side-Hustle Educators in India", font_size=15, font_color=TEXT_DARK, bold=True)

    chart_path = os.path.join(assets_dir, "charts", "fig6a_side_occupations_dist.png")
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
    # RIGHT COLUMN: 3 MAJOR EDUCATOR PROFILES (Width: 5.7in)
    # -------------------------------------------------------------
    right_x = Inches(7.10)
    right_w = Inches(5.65)

    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, right_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), right_w - Inches(0.6), Inches(0.25), text="PRIMARY BACKGROUNDS", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), right_w - Inches(0.6), Inches(0.35), text="Three Distinct Tutor Cohorts", font_size=15, font_color=TEXT_DARK, bold=True)

    profiles = [
        (
            "1. Software & IT Engineers (28.4%)",
            "Avg Fee: ₹1,050 / hr",
            "Tech professionals in Bengaluru, Pune, and Hyderabad tutoring Python, Java, SQL, and Advanced JEE Math in evening hours and weekends.",
            ACCENT_INDIGO,
            PILL_BLUE,
        ),
        (
            "2. College & Graduate Students (34.2%)",
            "Avg Fee: ₹480 / hr",
            "B.Tech, B.Sc, and M.A. university students teaching primary and middle schoolers to fund their college education and gain teaching experience.",
            ACCENT_EMERALD,
            PILL_MINT,
        ),
        (
            "3. Full-Time School Teachers (37.4%)",
            "Avg Fee: ₹820 / hr",
            "Dedicated classroom teachers offering structured after-school batches for 10th and 12th board exams with verified curriculum mastery.",
            ACCENT_CORAL,
            PILL_CORAL,
        ),
    ]

    for j, (phead, pfee, pdesc, paccent, ppill) in enumerate(profiles):
        py = top_y + Inches(0.95 + j * 1.15)
        add_card_box(slide, right_x + Inches(0.25), py, right_w - Inches(0.50), Inches(1.05), fill_color=BG_SUBTLE)
        add_rectangle(slide, right_x + Inches(0.25), py, Inches(0.06), Inches(1.05), paccent, shape=MSO_SHAPE.RECTANGLE)

        add_textbox(slide, right_x + Inches(0.40), py + Inches(0.10), Inches(3.4), Inches(0.25), text=phead, font_size=12, font_color=TEXT_DARK, bold=True)
        add_pill(slide, right_x + Inches(3.85), py + Inches(0.10), Inches(1.30), Inches(0.24), pfee, ppill, text_color=paccent, font_size=9, bold=True)
        add_textbox(slide, right_x + Inches(0.40), py + Inches(0.38), right_w - Inches(0.80), Inches(0.62), text=pdesc, font_size=10, font_color=TEXT_MEDIUM, line_spacing=1.1)
