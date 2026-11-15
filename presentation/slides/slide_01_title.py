"""
Slide 1: Title — High-impact opener matching Group18_SC452 reference style.
Textured background + white card overlay + coral pill + author cards + hero graphic.
"""

import os
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    add_slide_base,
    add_textbox,
    add_pill,
    add_author_card,
    add_image_fit,
    CARD_COLOR,
    ACCENT_CORAL,
    ACCENT_INDIGO,
    ACCENT_EMERALD,
    TEXT_DARK,
    TEXT_MEDIUM,
    TEXT_LIGHT,
    WHITE,
    PILL_BLUE,
    PILL_MINT,
    PILL_CORAL,
    FONT_TITLE,
    FONT_BODY,
)


def add_slide(prs, assets_dir):
    slide, _ = add_slide_base(
        prs,
        category_pill_text="BMP TUTORS",
        slide_num_str="",
        top_tag="DA-IICT  •  ACADEMIC PROJECT PRESENTATION",
        assets_dir=assets_dir,
        pill_color=ACCENT_CORAL,
    )

    # Main Project Title
    add_textbox(
        slide,
        Inches(0.65),
        Inches(1.25),
        Inches(7.2),
        Inches(0.65),
        text="BMP TUTORS",
        font_size=38,
        font_color=ACCENT_INDIGO,
        bold=True,
        font_name=FONT_TITLE,
    )

    add_textbox(
        slide,
        Inches(0.65),
        Inches(1.95),
        Inches(7.2),
        Inches(1.4),
        text="Real Numbers on India's Private Tutors",
        font_size=32,
        font_color=TEXT_DARK,
        bold=True,
        font_name=FONT_TITLE,
        line_spacing=1.1,
    )

    # Subtitle / Mission statement in simple English
    add_textbox(
        slide,
        Inches(0.65),
        Inches(3.50),
        Inches(6.8),
        Inches(0.85),
        text="We scraped 107,723 tutor profiles from TeacherOn, checked which ones are still active, and put fair pricing benchmarks in one place — across all 36 Indian States and Union Territories.",
        font_size=13,
        font_color=TEXT_MEDIUM,
        font_name=FONT_BODY,
        line_spacing=1.2,
    )

    # Feature badges
    add_pill(
        slide,
        Inches(0.65),
        Inches(4.65),
        Inches(2.1),
        Inches(0.38),
        "107,723 TUTORS",
        PILL_BLUE,
        text_color=ACCENT_INDIGO,
        font_size=10,
        bold=True,
    )
    add_pill(
        slide,
        Inches(2.90),
        Inches(4.65),
        Inches(2.2),
        Inches(0.38),
        "36 STATES & UTS",
        PILL_MINT,
        text_color=ACCENT_EMERALD,
        font_size=10,
        bold=True,
    )
    add_pill(
        slide,
        Inches(5.25),
        Inches(4.65),
        Inches(2.2),
        Inches(0.38),
        "FEE CALCULATOR",
        PILL_CORAL,
        text_color=ACCENT_CORAL,
        font_size=10,
        bold=True,
    )

    # Hero visual on the right
    hero_path = os.path.join(assets_dir, "hero_illustration.png")
    if not os.path.exists(hero_path):
        hero_path = os.path.join(
            assets_dir, "maps", "india_all_states_labeled_dashboard.png"
        )
    if os.path.exists(hero_path):
        add_image_fit(
            slide,
            hero_path,
            left=Inches(7.8),
            top=Inches(1.2),
            max_width=Inches(4.8),
            max_height=Inches(4.2),
            align_center=True,
            valign_center=True,
        )

    # 3 Author Cards side by side at the bottom (matches reference slide 1)
    card_w = Inches(3.85)
    card_h = Inches(0.92)
    card_y = Inches(5.85)

    add_author_card(
        slide,
        Inches(0.65),
        card_y,
        "01",
        "Nilesh Mori",
        "Student ID: 202301473",
        width=card_w,
        height=card_h,
        accent_color=ACCENT_INDIGO,
    )

    add_author_card(
        slide,
        Inches(4.75),
        card_y,
        "02",
        "Harsh Maheriya",
        "Student ID: 202301470",
        width=card_w,
        height=card_h,
        accent_color=ACCENT_EMERALD,
    )

    add_author_card(
        slide,
        Inches(8.85),
        card_y,
        "03",
        "Tarang Hirapara",
        "Student ID: 202301207",
        width=card_w,
        height=card_h,
        accent_color=ACCENT_CORAL,
    )
