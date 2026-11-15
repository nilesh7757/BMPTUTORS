"""
Slide 5: Website Platform — Starlette app with filters, pagination, tutor explorer.
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
    PILL_SLATE,
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
        text="Website Platform — Tutor Explorer",
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
        text="Starlette + uvicorn web app with faceted filters, pagination, and downloadable exports",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # Left: feature list
    features = [
        (
            "🔍",
            "Faceted Filters",
            "Region, state, subject, qualification, occupation, gender, online mode",
        ),
        ("📄", "Pagination", "Browse 107,723 tutors in pages of 15–100 rows"),
        ("📥", "Custom Exports", "Download filtered Excel or editable DOCX reports"),
        ("🗺️", "Live Maps", "India choropleth + district-level drilldown"),
        ("💰", "Price Predictor", "ML-powered hourly fee estimation"),
        ("✨", "Bio Generator", "NLP + Gemini AI tutor profile bios"),
    ]
    y = 1.70
    for emoji, title, desc in features:
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
            Inches(3.4),
            Inches(0.35),
            text=title,
            font_size=13,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_textbox(
            slide,
            Inches(1.15),
            Inches(y + 0.30),
            Inches(4.0),
            Inches(0.45),
            text=desc,
            font_size=10.5,
            font_color=TEXT_MEDIUM,
        )
        y += 0.68

    # Right: website screenshot placeholder (use a dashboard chart image as visual)
    # Use regional distribution chart as the "website" visual
    chart_path = os.path.join(assets_dir, "charts", "chart_regional_distribution.png")
    if os.path.exists(chart_path):
        add_image_fit(
            slide, chart_path, Inches(7.1), Inches(1.75), Inches(5.5), Inches(3.6)
        )

    # Bottom: tech stack chips
    add_textbox(
        slide,
        Inches(0.55),
        Inches(5.75),
        Inches(3.0),
        Inches(0.35),
        text="Tech Stack:",
        font_size=12,
        font_color=TEXT_MEDIUM,
        bold=True,
    )
    chips = ["Python", "Starlette", "SQLite", "Chart.js", "Leaflet", "OpenPyXL"]
    chip_x = 0.55
    for chip in chips:
        add_pill(
            slide,
            Inches(chip_x),
            Inches(6.15),
            Inches(1.05),
            Inches(0.38),
            chip,
            PILL_SLATE if hasattr(__import__("theme"), "PILL_SLATE") else CARD_COLOR,
            text_color=TEXT_DARK,
            font_size=9.5,
            bold=True,
        )
        chip_x += 1.15

    # Footer
    add_textbox(
        slide,
        Inches(0.78),
        Inches(7.00),
        Inches(5.0),
        Inches(0.25),
        text="BMP Tutors Presentation  •  Slide 5 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
