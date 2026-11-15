"""
Slide 14: Thank You / Q&A — closing slide matching reference style.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    BG_COLOR,
    CARD_COLOR,
    ACCENT_CORAL,
    ACCENT_INDIGO,
    ACCENT_VIOLET,
    ACCENT_EMERALD,
    TEXT_DARK,
    TEXT_MEDIUM,
    TEXT_LIGHT,
    WHITE,
    PILL_BLUE,
    PILL_MINT,
    PILL_CORAL,
    PILL_SLATE,
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
    add_author_card,
    ACCENT_AMBER,
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

    # Coral accent pill top-left
    add_pill(
        slide,
        Inches(0.55),
        Inches(0.50),
        Inches(1.45),
        Inches(0.42),
        "THANK YOU",
        ACCENT_CORAL,
        text_color=WHITE,
        font_size=13,
        bold=True,
    )

    # Top-right
    add_textbox(
        slide,
        Inches(8.5),
        Inches(0.52),
        Inches(4.2),
        Inches(0.32),
        text="DA-IICT  •  BMP TUTORS VISUAL ANALYTICS",
        font_size=11,
        font_color=TEXT_MEDIUM,
        align=PP_ALIGN.RIGHT,
    )

    # Big thank you
    add_textbox(
        slide,
        Inches(0.62),
        Inches(1.50),
        Inches(12.0),
        Inches(0.85),
        text="Thank You!",
        font_size=48,
        font_color=TEXT_DARK,
        bold=True,
        font_name=FONT_TITLE,
        align=PP_ALIGN.CENTER,
    )

    # Subtitle
    add_textbox(
        slide,
        Inches(0.62),
        Inches(2.45),
        Inches(12.0),
        Inches(0.5),
        text="Questions & Discussion",
        font_size=20,
        font_color=TEXT_MEDIUM,
        align=PP_ALIGN.CENTER,
    )

    # Three feature recap pills
    add_pill(
        slide,
        Inches(3.0),
        Inches(3.25),
        Inches(2.0),
        Inches(0.45),
        "🧹 Data Cleaning",
        PILL_BLUE,
        font_size=12,
    )
    add_pill(
        slide,
        Inches(5.15),
        Inches(3.25),
        Inches(2.0),
        Inches(0.45),
        "🤖 ML Price Model",
        PILL_MINT,
        font_size=12,
    )
    add_pill(
        slide,
        Inches(7.3),
        Inches(3.25),
        Inches(2.0),
        Inches(0.45),
        "✨ Gemini NLP Bios",
        PILL_CORAL,
        font_size=12,
    )

    # Divider line
    add_rectangle(
        slide,
        Inches(0.66),
        Inches(4.15),
        Inches(12.0),
        Inches(0.02),
        RGBColor(0xC9, 0xD6, 0xE2),
        shape=MSO_SHAPE.RECTANGLE,
    )

    # Contact / presenter info
    add_textbox(
        slide,
        Inches(0.62),
        Inches(4.45),
        Inches(6.0),
        Inches(0.4),
        text="Presented by: Nilesh Mori",
        font_size=15,
        font_color=TEXT_DARK,
        bold=True,
    )
    add_textbox(
        slide,
        Inches(0.62),
        Inches(4.85),
        Inches(6.0),
        Inches(0.35),
        text="DA-IICT Gandhinagar  •  Roll: 202301473  •  ICT Batch 2023–27",
        font_size=11,
        font_color=TEXT_MEDIUM,
    )

    add_textbox(
        slide,
        Inches(7.0),
        Inches(4.45),
        Inches(5.5),
        Inches(0.4),
        text="Project: BMP Tutors — National Market Intelligence Platform",
        font_size=13,
        font_color=TEXT_DARK,
        bold=True,
        align=PP_ALIGN.RIGHT,
    )
    add_textbox(
        slide,
        Inches(7.0),
        Inches(4.85),
        Inches(5.5),
        Inches(0.35),
        text="Cleaned Data  •  DBMS Normalization  •  Website  •  DOCX Report  •  ML  •  NLP",
        font_size=10,
        font_color=TEXT_MEDIUM,
        align=PP_ALIGN.RIGHT,
    )

    # Footer
    add_textbox(
        slide,
        Inches(0.78),
        Inches(7.00),
        Inches(5.0),
        Inches(0.25),
        text="BMP Tutors Presentation  •  Slide 14 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
