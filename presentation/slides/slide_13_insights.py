"""
Slide 13: Key Insights & Takeaways — checkmark callout boxes.
"""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    BG_COLOR,
    CARD_COLOR,
    ACCENT_INDIGO,
    ACCENT_EMERALD,
    ACCENT_VIOLET,
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
        text="Key Insights & Takeaways",
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
        text="Top findings from the BMP Tutors data analysis",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # Insight cards (2x2 grid)
    insights = [
        (
            "📈",
            "STEM Dominates",
            "Science, Coding & Math lead subject offerings — high demand for JEE/NEET & tech tutoring.",
            ACCENT_INDIGO,
        ),
        (
            "💰",
            "Fee Benchmark",
            "Avg hourly rate ₹416.4/hr (outlier-trimmed) — coding & IT command premium pricing.",
            ACCENT_EMERALD,
        ),
        (
            "🗺️",
            "Pan-India Reach",
            "107,723 tutors across 37 States & UTs — strong presence in metros and Tier-3 towns.",
            ACCENT_VIOLET,
        ),
        (
            "⚡",
            "Active Rate",
            "71.2% of tutors are live & verified — a healthy, active tutor marketplace.",
            ACCENT_AMBER,
        ),
    ]
    cw = Inches(6.0)
    ch = Inches(1.65)
    gap_x = Inches(0.25)
    gap_y = Inches(0.25)
    start_x = Inches(0.66)
    start_y = Inches(1.75)

    for i, (emoji, title, desc, color) in enumerate(insights):
        col = i % 2
        row = i // 2
        l = start_x + (col * (cw + gap_x))
        t = start_y + (row * (ch + gap_y))
        add_rectangle(
            slide, l, t, cw, ch, CARD_COLOR, shape=MSO_SHAPE.ROUNDED_RECTANGLE
        )
        # left accent bar
        add_rectangle(slide, l, t, Inches(0.09), ch, color, shape=MSO_SHAPE.RECTANGLE)
        add_textbox(
            slide,
            l + Inches(0.25),
            t + Inches(0.15),
            Inches(0.6),
            Inches(0.6),
            text=emoji,
            font_size=20,
            font_color=color,
        )
        add_textbox(
            slide,
            l + Inches(0.95),
            t + Inches(0.15),
            cw - Inches(1.2),
            Inches(0.4),
            text=title,
            font_size=14,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_textbox(
            slide,
            l + Inches(0.95),
            t + Inches(0.62),
            cw - Inches(1.2),
            Inches(0.9),
            text=desc,
            font_size=11,
            font_color=TEXT_MEDIUM,
            font_name=FONT_BODY,
        )

    # Bottom closing bar
    add_rectangle(
        slide,
        Inches(0.66),
        Inches(5.55),
        Inches(12.0),
        Inches(0.85),
        PILL_BLUE if hasattr(__import__("theme"), "PILL_BLUE") else CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_textbox(
        slide,
        Inches(0.95),
        Inches(5.68),
        Inches(11.5),
        Inches(0.6),
        text="💡  BMP Tutors turns raw tutor data into actionable market intelligence — from cleaning and normalization to ML pricing and AI-generated bios.",
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
        text="BMP Tutors Presentation  •  Slide 13 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
