"""
Theme module for BMP Tutors Presentation.
Matches the Group18_SC452 reference visual style:
- Slide background: #F3F6FA (light blue-gray)
- White card overlay: #FFFFFF
- Accent/coral: #F26B5B
- Feature pills: pastel blue #EAF1FF, mint #E8F7F3, light coral #FFF0ED
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

# ============================================================================
# COLOR PALETTE (from Group18_SC452 reference)
# ============================================================================
BG_COLOR = RGBColor(0xF3, 0xF6, 0xFA)  # slide background (light blue-gray)
CARD_COLOR = RGBColor(0xFF, 0xFF, 0xFF)  # white card overlay
ACCENT_CORAL = RGBColor(0xF2, 0x6B, 0x5B)  # coral accent (G-18 pill)
ACCENT_INDIGO = RGBColor(0x4F, 0x46, 0xE5)  # indigo primary
ACCENT_VIOLET = RGBColor(0x7E, 0x22, 0xCE)  # violet secondary
ACCENT_EMERALD = RGBColor(0x10, 0xB9, 0x81)  # emerald success/active
ACCENT_AMBER = RGBColor(0xF5, 0x9E, 0x0B)  # amber benchmark/warning
TEXT_DARK = RGBColor(0x0F, 0x17, 0x2A)  # slate-900 dark text
TEXT_MEDIUM = RGBColor(0x33, 0x41, 0x55)  # slate-700 medium text
TEXT_LIGHT = RGBColor(0x64, 0x74, 0x8B)  # slate-500 light text
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# Feature pill pastel colors (from reference)
PILL_BLUE = RGBColor(0xEA, 0xF1, 0xFF)  # light blue pill
PILL_MINT = RGBColor(0xE8, 0xF7, 0xF3)  # mint pill
PILL_CORAL = RGBColor(0xFF, 0xF0, 0xED)  # light coral pill
PILL_SLATE = RGBColor(0xD9, 0xE2, 0xEC)  # slate author-card bg

# Fonts
FONT_TITLE = "Calibri"
FONT_BODY = "Calibri"

# Slide size (16:9 widescreen, matches reference)
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Card layout constants (from reference slide 1)
CARD_LEFT = Inches(0.10)
CARD_TOP = Inches(0.10)
CARD_WIDTH = Inches(13.13)
CARD_HEIGHT = Inches(7.30)
BORDER_SOFT = RGBColor(0xE2, 0xE8, 0xF0)  # slate-200 border
BG_SUBTLE = RGBColor(0xF8, 0xFA, 0xFC)    # slate-50 light card fill


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================


def add_slide_base(
    prs,
    category_pill_text="ANALYTICS",
    slide_num_str="",
    top_tag="DA-IICT  •  BMP TUTORS",
    assets_dir=None,
    pill_color=ACCENT_CORAL,
):
    """
    Standard base layout for all presentation slides:
    1. Base textured background pattern (from Group 18 reference).
    2. Large white rounded rectangle card overlay (0.10in margin).
    3. Category pill top-left (e.g., 'DATA PIPELINE', 'GEOSPATIAL', etc.).
    4. Top-right project metadata and slide number pill.
    """
    import os

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # Background pattern or solid color
    bg_pattern_path = os.path.join(assets_dir, "bg_pattern.png") if assets_dir else None
    if bg_pattern_path and os.path.exists(bg_pattern_path):
        slide.shapes.add_picture(
            bg_pattern_path, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT
        )
    else:
        set_bg_solid(slide, BG_COLOR)

    # Large white rounded rectangle card overlay
    card = add_rectangle(
        slide,
        CARD_LEFT,
        CARD_TOP,
        CARD_WIDTH,
        CARD_HEIGHT,
        CARD_COLOR,
        line_color=BORDER_SOFT,
        line_width=Pt(1),
        shape=MSO_SHAPE.ROUNDED_RECTANGLE,
    )

    # Top-left Category Pill
    if category_pill_text:
        # Approximate width based on text length
        pill_w = max(1.2, min(3.0, 0.12 * len(category_pill_text) + 0.4))
        add_pill(
            slide,
            Inches(0.55),
            Inches(0.46),
            Inches(pill_w),
            Inches(0.38),
            category_pill_text.upper(),
            pill_color,
            text_color=WHITE,
            font_size=11,
            bold=True,
        )

    # Top-right metadata & slide counter
    if top_tag or slide_num_str:
        meta_text = top_tag
        if slide_num_str:
            meta_text = f"{top_tag}   •   {slide_num_str}" if top_tag else slide_num_str
        add_textbox(
            slide,
            Inches(7.5),
            Inches(0.48),
            Inches(5.2),
            Inches(0.34),
            text=meta_text,
            font_size=11,
            font_color=TEXT_LIGHT,
            align=PP_ALIGN.RIGHT,
        )

    return slide, card


def add_slide_header(slide, title, subtitle=""):
    """
    Standard clean slide header with bold title and simple English subtitle.
    """
    add_textbox(
        slide,
        Inches(0.55),
        Inches(0.98),
        Inches(12.2),
        Inches(0.52),
        text=title,
        font_size=24,
        font_color=TEXT_DARK,
        bold=True,
        font_name=FONT_TITLE,
    )
    if subtitle:
        add_textbox(
            slide,
            Inches(0.55),
            Inches(1.50),
            Inches(12.2),
            Inches(0.36),
            text=subtitle,
            font_size=12.5,
            font_color=TEXT_MEDIUM,
            font_name=FONT_BODY,
        )


def add_card_box(
    slide,
    left,
    top,
    width,
    height,
    fill_color=CARD_COLOR,
    line_color=BORDER_SOFT,
    shape=MSO_SHAPE.ROUNDED_RECTANGLE,
):
    """Add a card container with rounded corners and subtle border."""
    return add_rectangle(
        slide,
        left,
        top,
        width,
        height,
        fill_color,
        line_color=line_color,
        line_width=Pt(1),
        shape=shape,
    )


def set_bg_solid(slide, color=BG_COLOR):
    """Set slide background to a solid color."""
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def add_rectangle(
    slide,
    left,
    top,
    width,
    height,
    fill_color,
    line_color=None,
    line_width=None,
    shape=MSO_SHAPE.RECTANGLE,
    shadow=False,
):
    """Add a rounded/regular rectangle shape with fill and optional border."""
    shp = slide.shapes.add_shape(shape, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = line_width or Pt(1)
    shp.shadow.inherit = shadow
    return shp


def add_textbox(
    slide,
    left,
    top,
    width,
    height,
    text="",
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
    """Add a textbox with formatted text."""
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


def add_rich_text(
    slide,
    left,
    top,
    width,
    height,
    runs,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    line_spacing=1.0,
    space_after=0,
):
    """
    Add a textbox with multiple styled runs.
    runs: list of (text, font_size, font_color, bold, italic) tuples
    """
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    if space_after:
        p.space_after = Pt(space_after)

    for item in runs:
        text, size, color, bold, italic = item
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = FONT_BODY
        run.font.color.rgb = color
    return tb


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
    shape=MSO_SHAPE.ROUNDED_RECTANGLE,
):
    """Add a small rounded-pill label (like reference feature pills)."""
    shp = add_rectangle(
        slide, left, top, width, height, fill_color, shape=shape, shadow=False
    )
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = FONT_BODY
    run.font.color.rgb = text_color
    return shp


def add_kpi_card(
    slide,
    left,
    top,
    width,
    height,
    label,
    value,
    icon_emoji,
    card_color=CARD_COLOR,
    accent_color=ACCENT_INDIGO,
    value_size=32,
    label_size=11,
    icon_size=30,
):
    """Add a KPI metric card with icon emoji, large value, and label."""
    # Card background
    card = add_rectangle(slide, left, top, width, height, card_color, shadow=False)
    # Accent bar on top
    add_rectangle(slide, left, top, width, Inches(0.07), accent_color, shadow=False)
    # Icon emoji
    add_textbox(
        slide,
        left + Inches(0.25),
        top + Inches(0.35),
        Inches(0.8),
        Inches(0.8),
        text=icon_emoji,
        font_size=icon_size,
        font_color=accent_color,
        align=PP_ALIGN.LEFT,
    )
    # Value
    add_textbox(
        slide,
        left + Inches(1.15),
        top + Inches(0.30),
        width - Inches(1.4),
        Inches(1.0),
        text=value,
        font_size=value_size,
        font_color=TEXT_DARK,
        bold=True,
        align=PP_ALIGN.LEFT,
    )
    # Label
    add_textbox(
        slide,
        left + Inches(1.15),
        top + height - Inches(0.55),
        width - Inches(1.4),
        Inches(0.4),
        text=label,
        font_size=label_size,
        font_color=TEXT_MEDIUM,
        align=PP_ALIGN.LEFT,
    )
    return card


def add_image_fit(
    slide,
    image_path,
    left,
    top,
    max_width,
    max_height,
    align_center=True,
    valign_center=True,
):
    """
    Add an image preserving aspect ratio, fitting within max_width x max_height.
    Clamps position/size so nothing goes off-slide.
    Returns the picture shape.
    """
    from PIL import Image

    # Slide bounds (13.33 x 7.5 in)
    SLIDE_W_IN = 13.333
    SLIDE_H_IN = 7.5

    # Clamp input box to slide bounds
    left = max(0.0, min(left, SLIDE_W_IN - 0.1))
    top = max(0.0, min(top, SLIDE_H_IN - 0.1))
    max_width = min(max_width, SLIDE_W_IN - left - 0.05)
    max_height = min(max_height, SLIDE_H_IN - top - 0.05)
    if max_width <= 0 or max_height <= 0:
        return None, 0, 0

    with Image.open(image_path) as img:
        img_w, img_h = img.size
    aspect = img_w / img_h

    # Calculate fit
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

    # Final clamp so picture never goes off-slide
    final_left = max(0.0, min(final_left, SLIDE_W_IN - fit_w - 0.02))
    final_top = max(0.0, min(final_top, SLIDE_H_IN - fit_h - 0.02))

    pic = slide.shapes.add_picture(
        image_path, Inches(final_left), Inches(final_top), Inches(fit_w), Inches(fit_h)
    )
    return pic, fit_w, fit_h


def add_author_card(
    slide,
    left,
    top,
    number,
    name,
    roll,
    width=Inches(3.82),
    height=Inches(0.90),
    accent_color=ACCENT_INDIGO,
):
    """Add an author card like the reference (numbered pill + name + roll)."""
    # Outer slate card
    add_rectangle(slide, left, top, width, height, PILL_SLATE, shadow=False)
    # Inner white card
    add_rectangle(
        slide,
        left + Inches(0.03),
        top + Inches(0.03),
        width - Inches(0.06),
        height - Inches(0.06),
        CARD_COLOR,
        shadow=False,
    )
    # Number pill
    add_textbox(
        slide,
        left + Inches(0.18),
        top + Inches(0.30),
        Inches(0.5),
        Inches(0.4),
        text=number,
        font_size=16,
        font_color=accent_color,
        bold=True,
    )
    # Name
    add_textbox(
        slide,
        left + Inches(0.75),
        top + Inches(0.20),
        width - Inches(1.0),
        Inches(0.3),
        text=name,
        font_size=12,
        font_color=TEXT_DARK,
        bold=True,
    )
    # Roll
    add_textbox(
        slide,
        left + Inches(0.75),
        top + Inches(0.52),
        width - Inches(1.0),
        Inches(0.28),
        text=roll,
        font_size=10,
        font_color=TEXT_LIGHT,
    )
    return


def add_section_divider(slide, title, subtitle="", accent_color=ACCENT_INDIGO):
    """Add a section divider strip at top of content slides."""
    add_rectangle(
        slide,
        Inches(0.0),
        Inches(0.0),
        SLIDE_WIDTH,
        Inches(0.12),
        accent_color,
        shadow=False,
    )
    add_textbox(
        slide,
        Inches(0.55),
        Inches(0.35),
        Inches(11.0),
        Inches(0.6),
        text=title,
        font_size=28,
        font_color=TEXT_DARK,
        bold=True,
    )
    if subtitle:
        add_textbox(
            slide,
            Inches(0.55),
            Inches(0.95),
            Inches(12.0),
            Inches(0.4),
            text=subtitle,
            font_size=14,
            font_color=TEXT_MEDIUM,
        )
