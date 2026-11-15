"""
Slide 12: Where Tutors Are — India choropleth + zone maps.
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
        text="Where Tutors Are — India Map Distribution",
        font_size=24,
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
        text="Choropleth maps showing tutor concentration and average fees across 37 States & UTs",
        font_size=13,
        font_color=TEXT_MEDIUM,
    )

    # Main India choropleth (large, left)
    map_path = os.path.join(assets_dir, "maps", "all_india_tutor_choropleth.png")
    if not os.path.exists(map_path):
        map_path = os.path.join(
            assets_dir, "maps", "india_all_states_labeled_dashboard.png"
        )
    if os.path.exists(map_path):
        add_image_fit(
            slide, map_path, Inches(0.66), Inches(1.65), Inches(5.6), Inches(4.4)
        )
    add_textbox(
        slide,
        Inches(0.66),
        Inches(6.10),
        Inches(5.6),
        Inches(0.3),
        text="All-India Tutor Choropleth",
        font_size=9.5,
        font_color=TEXT_MEDIUM,
        align=PP_ALIGN.CENTER,
    )

    # Zone maps (right side, 2x2 grid)
    zone_maps = [
        ("map_zone_north.png", "North Zone"),
        ("map_zone_south.png", "South Zone"),
        ("map_zone_west.png", "West Zone"),
        ("map_zone_east_central.png", "East & Central Zone"),
    ]
    zw = Inches(3.25)
    zh = Inches(1.95)
    zx = Inches(6.55)
    zy = Inches(1.65)
    for i, (fname, label) in enumerate(zone_maps):
        col = i % 2
        row = i // 2
        l = zx + (col * (zw + Inches(0.12)))
        t = zy + (row * (zh + Inches(0.15)))
        fpath = os.path.join(assets_dir, "maps", fname)
        if os.path.exists(fpath):
            add_image_fit(slide, fpath, l, t, zw, zh)
        add_textbox(
            slide,
            l,
            t + zh - Inches(0.02),
            zw,
            Inches(0.25),
            text=label,
            font_size=9.5,
            font_color=TEXT_MEDIUM,
            align=PP_ALIGN.CENTER,
        )

    # Bottom insight
    add_rectangle(
        slide,
        Inches(0.66),
        Inches(6.45),
        Inches(12.0),
        Inches(0.45),
        PILL_BLUE if hasattr(__import__("theme"), "PILL_BLUE") else CARD_COLOR,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )
    add_textbox(
        slide,
        Inches(0.95),
        Inches(6.55),
        Inches(11.5),
        Inches(0.3),
        text="🗺️  Tutors span all of India — from metros to Tier-3 towns, enabling pan-India online and offline tutoring demand analysis.",
        font_size=11,
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
        text="BMP Tutors Presentation  •  Slide 12 of 14",
        font_size=10,
        font_color=TEXT_LIGHT,
    )

    return slide
