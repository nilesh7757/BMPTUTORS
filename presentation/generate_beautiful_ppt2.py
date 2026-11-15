#!/usr/bin/env python3
"""
BMP Tutors - Beautiful Presentation Generator
Design inspired by modern slide decks (arena.ai, GitHub, presentation design systems):
- Full-bleed background images with opacity overlays
- Properly sized white card containers on every slide
- Doodle/illustration accents as decorative elements
- Relevant images on content slides (charts, maps, website mockups)
- No text overlap (precise grid-based positioning)
- Character and personality through color, imagery, and layout
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os
from PIL import Image as PILImage

# ============================================================================
# THEME & COLORS
# ============================================================================
BG_COLOR = RGBColor(0xF3, 0xF6, 0xFA)  # Light blue-gray slide background
CARD_COLOR = RGBColor(0xFF, 0xFF, 0xFF)  # Pure white card
CARD_TINT = RGBColor(0xF8, 0xFA, 0xFC)  # Slightly tinted white
ACCENT_CORAL = RGBColor(0xF2, 0x6B, 0x5B)  # Coral accent
ACCENT_INDIGO = RGBColor(0x4F, 0x46, 0xE5)  # Indigo primary
ACCENT_VIOLET = RGBColor(0x7E, 0x22, 0xCE)  # Violet secondary
ACCENT_EMERALD = RGBColor(0x10, 0xB9, 0x81)  # Emerald success
ACCENT_AMBER = RGBColor(0xF5, 0x9E, 0x0B)  # Amber warning
TEXT_DARK = RGBColor(0x0F, 0x17, 0x2A)  # Slate-900
TEXT_MEDIUM = RGBColor(0x33, 0x41, 0x55)  # Slate-700
TEXT_LIGHT = RGBColor(0x64, 0x74, 0x8B)  # Slate-500
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

PILL_BLUE = RGBColor(0xEA, 0xF1, 0xFF)
PILL_MINT = RGBColor(0xE8, 0xF7, 0xF3)
PILL_CORAL = RGBColor(0xFF, 0xF0, 0xED)
PILL_SLATE = RGBColor(0xD9, 0xE2, 0xEC)
PILL_AMBER = RGBColor(0xFF, 0xFB, 0xEB)

FONT_TITLE = "Calibri"
FONT_BODY = "Calibri"

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

ASSETS_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "presentation", "assets"
)


# ============================================================================
# HELPER: Opacity
# ============================================================================
def set_shape_opacity(shape, opacity_pct):
    if opacity_pct >= 100:
        return
    try:
        sp = shape._element.spPr
        alpha = sp.find(qn("a:alpha"))
        if alpha is None:
            alpha = sp.makeelement(qn("a:alpha"), {})
            sp.append(alpha)
        alpha.set("val", str(int(opacity_pct * 1000)))
    except Exception:
        pass


# ============================================================================
# HELPER: Rectangle
# ============================================================================
def add_rect(
    slide,
    left,
    top,
    width,
    height,
    fill_color,
    opacity=100,
    line_color=None,
    line_width=None,
    shape=MSO_SHAPE.RECTANGLE,
):
    shp = slide.shapes.add_shape(shape, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if opacity < 100:
        set_shape_opacity(shp, opacity)
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = line_width or Pt(1)
    shp.shadow.inherit = False
    return shp


# ============================================================================
# HELPER: Text box
# ============================================================================
def add_text(
    slide,
    left,
    top,
    width,
    height,
    text,
    font_size=14,
    font_color=TEXT_DARK,
    bold=False,
    align=PP_ALIGN.LEFT,
    font_name=FONT_BODY,
    anchor=MSO_ANCHOR.TOP,
    italic=False,
    line_spacing=1.0,
    space_after=0,
):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    lines = text.split("\n") if isinstance(text, str) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if space_after:
            p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = font_name
        run.font.color.rgb = font_color
    return tb


# ============================================================================
# HELPER: Image with fit + opacity
# ============================================================================
def add_image(
    slide,
    image_path,
    left,
    top,
    max_width,
    max_height,
    align_center=True,
    valign_center=True,
    opacity=100,
):
    if not image_path or not os.path.exists(image_path):
        return None
    SLIDE_W_IN = 13.333
    SLIDE_H_IN = 7.5
    left = max(0.0, min(left, SLIDE_W_IN - 0.1))
    top = max(0.0, min(top, SLIDE_H_IN - 0.1))
    max_width = min(max_width, SLIDE_W_IN - left - 0.05)
    max_height = min(max_height, SLIDE_H_IN - top - 0.05)
    if max_width <= 0 or max_height <= 0:
        return None
    with PILImage.open(image_path) as img:
        img_w, img_h = img.size
    aspect = img_w / img_h
    fit_w = max_width
    fit_h = fit_w / aspect
    if fit_h > max_height:
        fit_h = max_height
        fit_w = fit_h * aspect
    final_left = left
    final_top = top
    if align_center:
        final_left = left + (max_width - fit_w) / 2
    if valign_center:
        final_top = top + (max_height - fit_h) / 2
    final_left = max(0.0, min(final_left, SLIDE_W_IN - fit_w - 0.02))
    final_top = max(0.0, min(final_top, SLIDE_H_IN - fit_h - 0.02))
    pic = slide.shapes.add_picture(
        image_path, Inches(final_left), Inches(final_top), Inches(fit_w), Inches(fit_h)
    )
    if opacity < 100:
        set_shape_opacity(pic, opacity)
    return pic


# ============================================================================
# HELPER: Pill badge
# ============================================================================
def add_pill(
    slide,
    left,
    top,
    width,
    height,
    text,
    fill_color,
    text_color=TEXT_DARK,
    font_size=11,
    bold=True,
):
    shp = add_rect(
        slide, left, top, width, height, fill_color, shape=MSO_SHAPE.ROUNDED_RECTANGLE
    )
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = FONT_BODY
    run.font.color.rgb = text_color
    return shp


# ============================================================================
# HELPER: Base slide (background + white card + header elements)
# ============================================================================
def add_base(
    slide,
    category_pill="",
    slide_num="",
    top_tag="DA-IICT  •  BMP TUTORS",
    pill_color=ACCENT_CORAL,
    bg_image=None,
    bg_opacity=40,
    card_opacity=100,
    card_tint=CARD_TINT,
):
    # Background image or solid color
    if bg_image and os.path.exists(bg_image):
        add_image(
            slide,
            bg_image,
            Inches(0),
            Inches(0),
            SLIDE_WIDTH,
            SLIDE_HEIGHT,
            opacity=bg_opacity,
        )
    else:
        add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, BG_COLOR)

    # Main white card container (properly sized - almost full slide)
    card = add_rect(
        slide,
        Inches(0.25),
        Inches(0.25),
        Inches(12.83),
        Inches(7.0),
        card_tint,
        opacity=card_opacity,
        line_color=RGBColor(0xE2, 0xE8, 0xF0),
        line_width=Pt(1),
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )

    # Category pill (top-left inside card)
    if category_pill:
        pill_w = max(1.2, min(3.0, 0.12 * len(category_pill) + 0.4))
        add_pill(
            slide,
            Inches(0.55),
            Inches(0.45),
            Inches(pill_w),
            Inches(0.38),
            category_pill.upper(),
            pill_color,
            text_color=WHITE,
            font_size=11,
        )

    # Metadata (top-right inside card)
    meta_text = top_tag
    if slide_num:
        meta_text = f"{top_tag}   •   {slide_num}" if top_tag else slide_num
    add_text(
        slide,
        Inches(8.0),
        Inches(0.48),
        Inches(4.8),
        Inches(0.34),
        meta_text,
        font_size=11,
        font_color=TEXT_LIGHT,
        align=PP_ALIGN.RIGHT,
    )

    return slide, card


# ============================================================================
# HELPER: Header inside card
# ============================================================================
def header(slide, title, subtitle="", y_offset=0.55):
    add_text(
        slide,
        Inches(0.55),
        Inches(y_offset),
        Inches(11.8),
        Inches(0.55),
        title,
        font_size=26,
        font_color=TEXT_DARK,
        bold=True,
    )
    if subtitle:
        add_text(
            slide,
            Inches(0.55),
            Inches(y_offset + 0.55),
            Inches(11.8),
            Inches(0.4),
            subtitle,
            font_size=13,
            font_color=TEXT_MEDIUM,
        )
