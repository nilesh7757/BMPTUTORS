"""
Slide 8: What Subjects Are Taught — subjects distribution + Venn trios.
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
        text="What Subjects Are Taught?",
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
        text="Subject distribution across the tutor network, with subject overlap visualized via Venn diagrams",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # Left: top subjects chart
    add_textbox(
        slide,
        Inches(0.66),
        Inches(1.65),
        Inches(4.0),
        Inches(0.35),
        text="Top Subjects (Macro View)",
        font_size=13,
        font_color=TEXT_DARK,
        bold=True,
    )
    fpath = os.path.join(assets_dir, "charts", "fig1_top_subjects.png")
    if os.path.exists(fpath):
        add_image_fit(
            slide, fpath, Inches(0.66), Inches(2.05), Inches(4.2), Inches(3.1)
        )

    # Middle: subjects macro chart
    add_textbox(
        slide,
        Inches(5.1),
        Inches(1.65),
        Inches(4.0),
        Inches(0.35),
        text="Subject Categories",
        font_size=13,
        font_color=TEXT_DARK,
        bold=True,
    )
    fpath = os.path.join(assets_dir, "charts", "fig1_subjects_macro.png")
    if os.path.exists(fpath):
        add_image_fit(slide, fpath, Inches(5.1), Inches(2.05), Inches(4.0), Inches(2.6))

    # Right: Venn trios side by side
    add_textbox(
        slide,
        Inches(9.35),
        Inches(1.65),
        Inches(3.2),
        Inches(0.35),
        text="Subject Overlap",
        font_size=13,
        font_color=TEXT_DARK,
        bold=True,
    )
    fpath1 = os.path.join(assets_dir, "charts", "venn_school_trio.png")
    fpath2 = os.path.join(assets_dir, "charts", "venn_stem_trio.png")
    if os.path.exists(fpath1):
        add_image_fit(
            slide, fpath1, Inches(9.35), Inches(2.05), Inches(3.2), Inches(1.75)
        )
    if os.path.exists(fpath2):
        add_image_fit(
            slide, fpath2, Inches(9.35), Inches(3.95), Inches(3.2), Inches(1.75)
        )

    # Bottom insight
    add_rectangle(
        slide,
        Inches(0.66),
        Inches(5.45),
        Inches(12.0),
        Inches(0.85),
        PILL_BLUE if hasattr(__import__("theme"), "PILL_BLUE") else CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_textbox(
        slide,
        Inches(0.95),
        Inches(5.58),
        Inches(11.5),
        Inches(0.65),
        text="💡  STEM subjects (Science, Coding, Math) dominate the tutor network — 3.7 subjects per educator shows strong multi-subject teaching capability.",
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
        text="BMP Tutors Presentation  •  Slide 8 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
