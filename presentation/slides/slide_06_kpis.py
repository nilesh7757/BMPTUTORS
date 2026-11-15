"""
Slide 6: Data Scale at a Glance — 4 KPI cards with emojis.
"""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    BG_COLOR,
    CARD_COLOR,
    ACCENT_INDIGO,
    ACCENT_VIOLET,
    ACCENT_EMERALD,
    ACCENT_AMBER,
    TEXT_DARK,
    TEXT_MEDIUM,
    TEXT_LIGHT,
    WHITE,
    FONT_TITLE,
    FONT_BODY,
    SLIDE_WIDTH,
    SLIDE_HEIGHT,
    CARD_LEFT,
    CARD_TOP,
    CARD_WIDTH,
    CARD_HEIGHT,
    PILL_BLUE,
    set_bg_solid,
    add_rectangle,
    add_textbox,
    add_pill,
    add_kpi_card,
    ACCENT_CORAL,
    PILL_MINT,
    PILL_CORAL,
    PILL_SLATE,
)


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
        text="Data Scale at a Glance",
        font_size=28,
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
        text="Key metrics from the cleaned and normalized BMP Tutors dataset",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # 4 KPI cards (2x2 grid)
    card_w = Inches(6.0)
    card_h = Inches(1.55)
    gap_x = Inches(0.25)
    gap_y = Inches(0.25)
    start_x = Inches(0.66)
    start_y = Inches(1.75)

    add_kpi_card(
        slide,
        start_x,
        start_y,
        card_w,
        card_h,
        "Verified Unique Tutors",
        "107,723",
        "👥",
        accent_color=ACCENT_INDIGO,
        value_size=34,
    )

    add_kpi_card(
        slide,
        start_x + card_w + gap_x,
        start_y,
        card_w,
        card_h,
        "Subject-Grade Linkages",
        "403,353",
        "📚",
        accent_color=ACCENT_VIOLET,
        value_size=34,
    )

    add_kpi_card(
        slide,
        start_x,
        start_y + card_h + gap_y,
        card_w,
        card_h,
        "Benchmark Hourly Fee",
        "₹416.4 /hr",
        "💰",
        accent_color=ACCENT_EMERALD,
        value_size=34,
    )

    add_kpi_card(
        slide,
        start_x + card_w + gap_x,
        start_y + card_h + gap_y,
        card_w,
        card_h,
        "Live Active Profiles",
        "76,706 (71.2%)",
        "⚡",
        accent_color=ACCENT_AMBER,
        value_size=30,
    )

    # Bottom insight strip
    add_rectangle(
        slide,
        Inches(0.66),
        Inches(5.30),
        Inches(12.0),
        Inches(0.85),
        PILL_BLUE if hasattr(__import__("theme"), "PILL_BLUE") else CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_textbox(
        slide,
        Inches(0.95),
        Inches(5.42),
        Inches(11.5),
        Inches(0.65),
        text="📍 Coverage: 37 States & UTs  |  📈 Avg 3.7 subjects per educator  |  🇮🇳 Pan-India private tutoring market intelligence",
        font_size=12.5,
        font_color=TEXT_MEDIUM,
        font_name=FONT_BODY,
    )

    # Footer
    add_textbox(
        slide,
        Inches(0.78),
        Inches(7.00),
        Inches(5.0),
        Inches(0.25),
        text="BMP Tutors Presentation  •  Slide 6 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
