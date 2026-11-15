"""
Slide 7: Geospatial Footprint — All-India Choropleth & State Distribution.
Matches Group 18 reference: Large visual map on left (7.0in) + State rankings card on right (5.0in).
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
        category_pill_text="GEOSPATIAL",
        slide_num_str="07 / 20",
        top_tag="ALL-INDIA COVERAGE",
        assets_dir=assets_dir,
        pill_color=ACCENT_CORAL,
    )

    add_slide_header(
        slide,
        title="Where Tutors Live — A Truly Pan-India Study",
        subtitle="Private tutoring is concentrated around metropolitan tech hubs and state capitals, but online tutoring spans all 36 States & UTs.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT COLUMN: HIGH-RES ALL-INDIA MAP (Width: 7.2in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(7.2)

    add_card_box(slide, left_x, top_y, left_w, card_h, fill_color=WHITE)

    # Map image
    map_path = os.path.join(assets_dir, "maps", "all_india_tutor_choropleth.png")
    if os.path.exists(map_path):
        add_image_fit(
            slide,
            map_path,
            left=left_x + Inches(0.15),
            top=top_y + Inches(0.15),
            max_width=left_w - Inches(0.30),
            max_height=card_h - Inches(0.30),
            align_center=True,
            valign_center=True,
        )

    # -------------------------------------------------------------
    # RIGHT COLUMN: STATE LEADERBOARD & INSIGHTS (Width: 4.8in)
    # -------------------------------------------------------------
    right_x = Inches(8.0)
    right_w = Inches(4.75)

    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, right_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), right_w - Inches(0.6), Inches(0.25), text="STATE CONCENTRATION", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), right_w - Inches(0.6), Inches(0.35), text="Top 5 Tutoring States", font_size=16, font_color=TEXT_DARK, bold=True)

    top_states = [
        ("1. Maharashtra", "14,820 Tutors", "Mumbai, Pune (13.8% of India)", ACCENT_INDIGO),
        ("2. Karnataka", "12,410 Tutors", "Bengaluru Tech Hub (11.5%)", ACCENT_INDIGO),
        ("3. Delhi NCR", "11,940 Tutors", "Competitive Exam Hub (11.1%)", ACCENT_INDIGO),
        ("4. Tamil Nadu", "11,230 Tutors", "Chennai & Engineering (10.4%)", ACCENT_INDIGO),
        ("5. Uttar Pradesh", "9,850 Tutors", "Rapid Online Growth (9.1%)", ACCENT_INDIGO),
    ]

    for j, (sname, scount, sdesc, scol) in enumerate(top_states):
        sy = top_y + Inches(0.95 + j * 0.58)
        add_textbox(slide, right_x + Inches(0.30), sy, Inches(2.2), Inches(0.25), text=sname, font_size=12, font_color=TEXT_DARK, bold=True)
        add_pill(slide, right_x + Inches(2.40), sy, Inches(1.10), Inches(0.24), scount, PILL_BLUE, text_color=scol, font_size=9, bold=True)
        add_textbox(slide, right_x + Inches(0.30), sy + Inches(0.25), right_w - Inches(0.6), Inches(0.25), text=sdesc, font_size=10, font_color=TEXT_LIGHT)

    # Key Takeaway Card at bottom
    add_card_box(slide, right_x + Inches(0.25), top_y + Inches(3.90), right_w - Inches(0.50), Inches(0.70), fill_color=BG_SUBTLE)
    add_textbox(slide, right_x + Inches(0.35), top_y + Inches(3.96), right_w - Inches(0.70), Inches(0.58), text="💡 Takeaway: Just 5 States account for over 55% of all private tutors nationwide. High urban density drives heavy tutoring demand.", font_size=10, font_color=TEXT_MEDIUM, line_spacing=1.1)
