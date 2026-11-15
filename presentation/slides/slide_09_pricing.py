"""
Slide 9: Pricing Benchmarks — fees by category, avg fee insights.
"""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    BG_COLOR,
    CARD_COLOR,
    ACCENT_INDIGO,
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
    set_bg_solid,
    add_rectangle,
    add_textbox,
    add_pill,
    add_image_fit,
    add_kpi_card,
    ACCENT_VIOLET,
    ACCENT_CORAL,
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
        text="Pricing Benchmarks",
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
        text="Average hourly fees by subject category — outlier-trimmed mean (₹150–₹2,800/hr)",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # Top KPI cards (2)
    card_w = Inches(6.0)
    card_h = Inches(1.35)
    add_kpi_card(
        slide,
        Inches(0.66),
        Inches(1.70),
        card_w,
        card_h,
        "Market Average Fee",
        "₹416.4 /hr",
        "💰",
        accent_color=ACCENT_EMERALD,
        value_size=30,
    )
    add_kpi_card(
        slide,
        Inches(6.67),
        Inches(1.70),
        card_w,
        card_h,
        "Outlier-Trimmed Range",
        "₹150 – ₹2,800",
        "📊",
        accent_color=ACCENT_INDIGO,
        value_size=24,
    )

    # Fees by category chart (right side)
    add_textbox(
        slide,
        Inches(0.66),
        Inches(3.35),
        Inches(4.0),
        Inches(0.35),
        text="Avg Fee by Subject Category",
        font_size=13,
        font_color=TEXT_DARK,
        bold=True,
    )
    fpath = os.path.join(assets_dir, "charts", "chart_fees_by_category.png")
    if os.path.exists(fpath):
        add_image_fit(
            slide, fpath, Inches(0.66), Inches(3.75), Inches(5.4), Inches(2.55)
        )

    # Right: fee insights bullets
    insights = [
        ("💵", "Coding & IT", "Premium rates — tech professionals charge more"),
        ("🔬", "Science", "High demand for JEE/NEET prep drives fees up"),
        ("🗣️", "Languages", "Competitive pricing for school-level tutoring"),
        ("📈", "Commerce", "Steady demand for accounts & economics"),
    ]
    y = 3.35
    for emoji, label, desc in insights:
        add_textbox(
            slide,
            Inches(6.6),
            Inches(y),
            Inches(0.5),
            Inches(0.4),
            text=emoji,
            font_size=14,
            font_color=ACCENT_INDIGO,
        )
        add_textbox(
            slide,
            Inches(7.1),
            Inches(y - 0.02),
            Inches(3.0),
            Inches(0.35),
            text=label,
            font_size=12.5,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_textbox(
            slide,
            Inches(7.1),
            Inches(y + 0.30),
            Inches(5.2),
            Inches(0.45),
            text=desc,
            font_size=10.5,
            font_color=TEXT_MEDIUM,
        )
        y += 0.70

    # Footer
    add_textbox(
        slide,
        Inches(0.78),
        Inches(7.00),
        Inches(5.0),
        Inches(0.25),
        text="BMP Tutors Presentation  •  Slide 9 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
