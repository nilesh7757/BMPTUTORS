"""
Slide 7: Who Are the Tutors? — gender, qualifications, experience, occupations.
"""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    BG_COLOR,
    CARD_COLOR,
    ACCENT_INDIGO,
    TEXT_DARK,
    TEXT_MEDIUM,
    TEXT_LIGHT,
    FONT_TITLE,
    FONT_BODY,
    SLIDE_WIDTH,
    SLIDE_HEIGHT,
    CARD_LEFT,
    CARD_TOP,
    CARD_WIDTH,
    CARD_HEIGHT,
    set_bg_solid,
    add_rectangle,
    add_textbox,
    add_pill,
    add_image_fit,
    ACCENT_VIOLET,
    ACCENT_EMERALD,
    ACCENT_AMBER,
    ACCENT_CORAL,
    WHITE,
    PILL_BLUE,
    PILL_MINT,
    PILL_CORAL,
    PILL_SLATE,
)
import os


def add_slide(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg_solid(slide, BG_COLOR)

    # White card
    add_rectangle(
        slide,
        CARD_LEFT,
        CARD_TOP,
        CARD_WIDTH,
        CARD_HEIGHT,
        CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )

    # Title
    add_textbox(
        slide,
        Inches(0.55),
        Inches(0.35),
        Inches(12.0),
        Inches(0.6),
        text="Who Are the Tutors?",
        font_size=26,
        font_color=TEXT_DARK,
        bold=True,
        font_name=FONT_TITLE,
    )

    # Subtitle
    add_textbox(
        slide,
        Inches(0.55),
        Inches(0.95),
        Inches(12.0),
        Inches(0.4),
        text="Demographics: gender mix, qualifications, experience levels, and side occupations",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # Left column: demographics stats
    stats = [
        ("👨‍🏫", "Male Tutors", "~55%"),
        ("👩‍🏫", "Female Tutors", "~45%"),
        ("🎓", "Postgraduate+", "Majority hold PG or higher"),
        ("📚", "3.7 Subjects", "Avg subjects taught per educator"),
    ]
    y = 1.70
    for emoji, label, val in stats:
        add_textbox(
            slide,
            Inches(0.66),
            Inches(y),
            Inches(0.5),
            Inches(0.4),
            text=emoji,
            font_size=15,
            font_color=ACCENT_INDIGO,
        )
        add_textbox(
            slide,
            Inches(1.15),
            Inches(y - 0.02),
            Inches(2.6),
            Inches(0.35),
            text=label,
            font_size=13,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_textbox(
            slide,
            Inches(3.7),
            Inches(y - 0.02),
            Inches(1.2),
            Inches(0.35),
            text=val,
            font_size=13,
            font_color=ACCENT_EMERALD,
            bold=True,
            align=PP_ALIGN.RIGHT,
        )
        y += 0.72

    # Right side: charts (2x2 grid of images)
    charts = [
        ("fig2_experience_gender.png", "Experience & Gender"),
        ("chart_qualifications.png", "Qualifications"),
        ("chart_education_levels.png", "Education Levels"),
        ("fig6a_side_occupations_dist.png", "Side Occupations"),
    ]
    cw = Inches(2.85)
    ch = Inches(1.85)
    cx = Inches(5.55)
    cy = Inches(1.75)
    for i, (fname, label) in enumerate(charts):
        col = i % 2
        row = i // 2
        l = cx + (col * (cw + Inches(0.15)))
        t = cy + (row * (ch + Inches(0.18)))
        fpath = os.path.join(assets_dir, "charts", fname)
        if os.path.exists(fpath):
            add_image_fit(slide, fpath, l, t, cw, ch)
        add_textbox(
            slide,
            l,
            t + ch - Inches(0.05),
            cw,
            Inches(0.25),
            text=label,
            font_size=9,
            font_color=TEXT_MEDIUM,
            align=PP_ALIGN.CENTER,
        )

    # Footer
    add_textbox(
        slide,
        Inches(0.78),
        Inches(7.00),
        Inches(5.0),
        Inches(0.25),
        text="BMP Tutors Presentation  •  Slide 7 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
