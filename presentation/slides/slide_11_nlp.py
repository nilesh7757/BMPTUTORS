"""
Slide 11: NLP + Gemini Bio Generation — AI tutor bios.
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
    set_bg_solid,
    add_rectangle,
    add_textbox,
    add_pill,
    ACCENT_CORAL,
    PILL_BLUE,
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
        text="NLP + Gemini: AI Tutor Bios",
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
        text="Gemini 2.5 Flash generates parent/student bios + headlines, with a procedural fallback engine",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # Left: bio cards (parent / student)
    # Parent bio card
    add_rectangle(
        slide,
        Inches(0.66),
        Inches(1.70),
        Inches(5.9),
        Inches(2.9),
        PILL_BLUE if hasattr(__import__("theme"), "PILL_BLUE") else CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_rectangle(
        slide,
        Inches(0.66),
        Inches(1.70),
        Inches(5.9),
        Inches(0.4),
        ACCENT_INDIGO,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_textbox(
        slide,
        Inches(0.85),
        Inches(1.76),
        Inches(5.5),
        Inches(0.3),
        text="👨‍👩‍👧‍👦 Bio for Parents",
        font_size=13,
        font_color=WHITE,
        bold=True,
    )
    add_textbox(
        slide,
        Inches(0.85),
        Inches(2.25),
        Inches(5.55),
        Inches(2.25),
        text="Bio written for Indian parents — reassuring, results-focused. Covers teaching approach, weekly tests, and ends with an invitation to book a trial class.",
        font_size=10.5,
        font_color=TEXT_DARK,
        font_name=FONT_BODY,
    )

    # Student bio card
    add_rectangle(
        slide,
        Inches(6.77),
        Inches(1.70),
        Inches(5.9),
        Inches(2.9),
        PILL_MINT if hasattr(__import__("theme"), "PILL_MINT") else CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_rectangle(
        slide,
        Inches(6.77),
        Inches(1.70),
        Inches(5.9),
        Inches(0.4),
        ACCENT_VIOLET,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_textbox(
        slide,
        Inches(6.96),
        Inches(1.76),
        Inches(5.5),
        Inches(0.3),
        text="🎓 Bio for Students",
        font_size=13,
        font_color=WHITE,
        bold=True,
    )
    add_textbox(
        slide,
        Inches(6.96),
        Inches(2.25),
        Inches(5.55),
        Inches(2.25),
        text="Warm, empathetic, relatable voice addressed to students. Eliminates fear of the subject, promises zero boring lectures, emphasizes step-by-step clarity and visual shortcuts, ends with an encouraging call-to-action.",
        font_size=10.5,
        font_color=TEXT_DARK,
        font_name=FONT_BODY,
    )

    # Right: Gemini + fallback info
    add_textbox(
        slide,
        Inches(0.66),
        Inches(4.85),
        Inches(4.0),
        Inches(0.35),
        text="How It Works",
        font_size=13,
        font_color=TEXT_DARK,
        bold=True,
    )
    flow_items = [
        ("🤖", "Gemini 2.5 Flash API", "Generates bespoke bios from tutor attributes"),
        ("🔄", "Procedural Fallback", "Template-based engine if API unavailable"),
        ("📝", "Headlines (3)", "Punchy one-liners under 85 chars each"),
        ("📊", "Bio Analyzer", "Scores profile strength (0–100) with suggestions"),
    ]
    y = 5.05
    for emoji, label, desc in flow_items:
        add_textbox(
            slide,
            Inches(0.66),
            Inches(y),
            Inches(0.45),
            Inches(0.4),
            text=emoji,
            font_size=13,
            font_color=ACCENT_INDIGO,
        )
        add_textbox(
            slide,
            Inches(1.1),
            Inches(y - 0.02),
            Inches(2.8),
            Inches(0.35),
            text=label,
            font_size=11.5,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_textbox(
            slide,
            Inches(1.1),
            Inches(y + 0.28),
            Inches(5.3),
            Inches(0.4),
            text=desc,
            font_size=9.5,
            font_color=TEXT_MEDIUM,
        )
        y += 0.58

    # Footer
    add_textbox(
        slide,
        Inches(0.78),
        Inches(7.00),
        Inches(5.0),
        Inches(0.25),
        text="BMP Tutors Presentation  •  Slide 11 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
