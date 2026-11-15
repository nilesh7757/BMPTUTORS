#!/usr/bin/env python3
"""
BMP Tutors - Beautiful Presentation Generator
Design: Full-bleed background images, properly sized white cards,
relevant images/icons on every slide, doodle accents, no text overlap.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os
from PIL import Image as PILImage

# Paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(SCRIPT_DIR, "assets")
GEN_DIR = os.path.join(ASSETS_DIR, "generated")

# Theme
BG_COLOR = RGBColor(0xF3, 0xF6, 0xFA)
CARD_COLOR = RGBColor(0xFF, 0xFF, 0xFF)
CARD_TINT = RGBColor(0xF8, 0xFA, 0xFC)
ACCENT_CORAL = RGBColor(0xF2, 0x6B, 0x5B)
ACCENT_INDIGO = RGBColor(0x4F, 0x46, 0xE5)
ACCENT_VIOLET = RGBColor(0x7E, 0x22, 0xCE)
ACCENT_EMERALD = RGBColor(0x10, 0xB9, 0x81)
ACCENT_AMBER = RGBColor(0xF5, 0x9E, 0x0B)
TEXT_DARK = RGBColor(0x0F, 0x17, 0x2A)
TEXT_MEDIUM = RGBColor(0x33, 0x41, 0x55)
TEXT_LIGHT = RGBColor(0x64, 0x74, 0x8B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PILL_BLUE = RGBColor(0xEA, 0xF1, 0xFF)
PILL_MINT = RGBColor(0xE8, 0xF7, 0xF3)
PILL_CORAL = RGBColor(0xFF, 0xF0, 0xED)
PILL_SLATE = RGBColor(0xD9, 0xE2, 0xEC)
FONT_TITLE = "Calibri"
FONT_BODY = "Calibri"
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


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


def add_base(
    slide,
    category_pill="",
    slide_num="",
    top_tag="DA-IICT  •  BMP TUTORS",
    pill_color=ACCENT_CORAL,
    bg_image=None,
    bg_opacity=40,
    card_tint=CARD_TINT,
    card_opacity=100,
):
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
    add_rect(
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
    return slide


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

# ============================================================================
# SLIDE 1: TITLE
# ============================================================================
def slide_title(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    hero_bg = os.path.join(assets_dir, "generated", "hero_bg.png")
    pattern = os.path.join(assets_dir, "bg_pattern.png")

    # Full background - hero gradient
    add_image(slide, hero_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    if pattern and os.path.exists(pattern):
        add_image(slide, pattern, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=25)

    # Main white card - left side (content area)
    add_rect(slide, Inches(0.50), Inches(0.80), Inches(6.30), Inches(5.90), CARD_COLOR,
             line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, Inches(0.50), Inches(0.80), Inches(0.10), Inches(5.90), ACCENT_INDIGO)

    # Category pill
    add_pill(slide, Inches(0.75), Inches(1.10), Inches(1.9), Inches(0.38),
             "BMP TUTORS", ACCENT_CORAL, text_color=WHITE, font_size=10)

    # Main title
    add_text(slide, Inches(0.75), Inches(1.65), Inches(5.9), Inches(1.6),
             "Real Numbers on India's Private Tutors",
             font_size=38, font_color=TEXT_DARK, bold=True, line_spacing=1.05)

    # Subtitle
    add_text(slide, Inches(0.75), Inches(3.35), Inches(5.7), Inches(1.4),
             "We scraped 107,723 tutor profiles from TeacherOn, checked which ones "
             "are still active, and put fair pricing benchmarks in one place — "
             "across all 36 Indian States and Union Territories.",
             font_size=15, font_color=TEXT_MEDIUM, line_spacing=1.2)

    # Stat row
    stats_y = Inches(4.95)
    add_text(slide, Inches(0.75), stats_y, Inches(1.7), Inches(0.55),
             "107,723", font_size=32, font_color=ACCENT_INDIGO, bold=True)
    add_text(slide, Inches(0.75), stats_y + Inches(0.60), Inches(1.7), Inches(0.3),
             "tutors scraped", font_size=10, font_color=TEXT_LIGHT)

    add_text(slide, Inches(2.60), stats_y, Inches(1.7), Inches(0.55),
             "76,706", font_size=32, font_color=ACCENT_EMERALD, bold=True)
    add_text(slide, Inches(2.60), stats_y + Inches(0.60), Inches(1.7), Inches(0.3),
             "active tutors", font_size=10, font_color=TEXT_LIGHT)

    add_text(slide, Inches(4.45), stats_y, Inches(1.7), Inches(0.55),
             "31,017", font_size=32, font_color=ACCENT_CORAL, bold=True)
    add_text(slide, Inches(4.45), stats_y + Inches(0.60), Inches(1.7), Inches(0.3),
             "removed / dormant", font_size=10, font_color=TEXT_LIGHT)

    # Hero image on right side - properly sized card
    add_rect(slide, Inches(7.15), Inches(0.80), Inches(5.58), Inches(5.90),
             WHITE, opacity=95, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_image(slide, hero_bg, Inches(7.35), Inches(1.00), Inches(5.20), Inches(5.50))

    # Doodle accent bottom-left
    doodle = os.path.join(assets_dir, "doodles", "doodle_card_a.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(0.50), Inches(5.95), Inches(1.7), Inches(1.15),
                  opacity=50)

    # Author cards at bottom - three cards with proper spacing
    authors = [
        ("01", "Nilesh Mori", "202301473", ACCENT_INDIGO),
        ("02", "Harsh Maheriya", "202301470", ACCENT_EMERALD),
        ("03", "Tarang Hirapara", "202301207", ACCENT_CORAL),
    ]
    ax = Inches(0.75)
    for num, name, roll, accent in authors:
        add_rect(slide, ax, Inches(6.55), Inches(3.70), Inches(0.62), PILL_SLATE,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(slide, ax + Inches(0.15), Inches(6.62), Inches(0.5), Inches(0.4),
                 num, font_size=13, font_color=accent, bold=True)
        add_text(slide, ax + Inches(0.70), Inches(6.60), Inches(2.80), Inches(0.30),
                 name, font_size=11.5, font_color=TEXT_DARK, bold=True)
        add_text(slide, ax + Inches(0.70), Inches(6.88), Inches(2.80), Inches(0.25),
                 roll, font_size=8.5, font_color=TEXT_LIGHT)
        ax += Inches(3.92)

    return slide

# ============================================================================
# SLIDE 2: WHY WE BUILT IT
# ============================================================================
def slide_why(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    problem_bg = os.path.join(assets_dir, "generated", "problem_bg.png")
    add_image(slide, problem_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(1.9), Inches(0.38),
             "PROBLEM", ACCENT_CORAL, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   02 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "Why We Built BMP Tutors",
           "Private tutoring in India is a messy market — no fair pricing benchmarks, "
           "ghost profiles everywhere, and no way to find multi-subject tutors.", y_offset=0.55)

    cards = [
        (Inches(0.55), Inches(2.10), Inches(3.95), Inches(2.40), ACCENT_CORAL,
         "No Pricing Benchmarks",
         "Tutors quote ₹200 to ₹3,500/hr for the same subject. Parents overpay. "
         "Tutors underprice. No fair market rate exists.",
         os.path.join(assets_dir, "generated", "icon_filters.png")),
        (Inches(4.70), Inches(2.10), Inches(3.95), Inches(2.40), ACCENT_INDIGO,
         "Ghost / Dormant Profiles",
         "TeacherOn shows profiles but won't tell you which are still active. "
         "31,017 of 107,723 accounts are dormant or removed.",
         os.path.join(assets_dir, "generated", "icon_pages.png")),
        (Inches(8.85), Inches(2.10), Inches(3.93), Inches(2.40), ACCENT_EMERALD,
         "Multi-Subject Tutors",
         "Parents want one tutor for Math + Science together, but multi-subject "
         "bundling rates were never measured.",
         os.path.join(assets_dir, "generated", "icon_price.png")),
    ]
    for l, t, w, h, accent, title, desc, img_path in cards:
        add_rect(slide, l, t, w, h, CARD_COLOR,
                 line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, l, t, Inches(0.10), h, accent)
        if img_path and os.path.exists(img_path):
            add_image(slide, img_path, l + Inches(w - 1.15), t + Inches(0.15),
                      Inches(1.0), Inches(0.85), opacity=80)
        add_text(slide, l + Inches(0.25), t + Inches(0.20), w - Inches(1.40), Inches(0.32),
                 title, font_size=13, font_color=TEXT_DARK, bold=True)
        add_text(slide, l + Inches(0.25), t + Inches(0.58), w - Inches(0.50), h - Inches(0.75),
                 desc, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.12)

    add_rect(slide, Inches(0.55), Inches(4.85), Inches(12.23), Inches(1.95), PILL_BLUE,
             opacity=45, line_color=ACCENT_INDIGO, line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(slide, Inches(0.85), Inches(5.00), Inches(3.8), Inches(0.55),
             "107,723", font_size=34, font_color=ACCENT_INDIGO, bold=True)
    add_text(slide, Inches(0.85), Inches(5.55), Inches(3.8), Inches(0.30),
             "tutor profiles scraped & verified", font_size=10, font_color=TEXT_LIGHT)
    add_text(slide, Inches(4.75), Inches(5.00), Inches(3.8), Inches(0.55),
             "76,706", font_size=34, font_color=ACCENT_EMERALD, bold=True)
    add_text(slide, Inches(4.75), Inches(5.55), Inches(3.8), Inches(0.30),
             "active tutors (71.2%)", font_size=10, font_color=TEXT_LIGHT)
    add_text(slide, Inches(8.65), Inches(5.00), Inches(3.8), Inches(0.55),
             "31,017", font_size=34, font_color=ACCENT_CORAL, bold=True)
    add_text(slide, Inches(8.65), Inches(5.55), Inches(3.8), Inches(0.30),
             "removed / dormant (28.8%)", font_size=10, font_color=TEXT_LIGHT)

    doodle = os.path.join(assets_dir, "doodles", "doodle_card_c.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(10.90), Inches(5.85), Inches(2.0), Inches(1.25),
                  opacity=45)

    return slide

# ============================================================================
# SLIDE 3: HOW WE BUILT IT (PIPELINE)
# ============================================================================
def slide_pipeline(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    pipeline_bg = os.path.join(assets_dir, "generated", "pipeline_bg.png")
    add_image(slide, pipeline_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.4), Inches(0.38),
             "OUR ROADMAP", ACCENT_INDIGO, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   03 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "How We Built BMP Tutors — Four Steps",
           "From raw TeacherOn listings to a live website with verified tutors "
           "and fair pricing.", y_offset=0.55)

    steps = [
        ("01", "SCRAPE & VERIFY", "Collecting and checking every profile",
         ["Scraped 107,723 tutor profiles from TeacherOn.",
          "Checked login dates — 76,706 tutors are still active.",
          "31,017 accounts were dormant or removed.",
          "Used a headless browser to verify profiles are live."],
         ACCENT_CORAL, PILL_CORAL,
         os.path.join(assets_dir, "generated", "icon_export.png")),
        ("02", "CLEANED INTO TWO TABLES", "Simple database structure",
         ["One table for tutors (107,723 rows, one per educator).",
          "One table for subjects (12,327 distinct subjects).",
          "No duplicate facts — each tutor appears exactly once."],
         ACCENT_INDIGO, PILL_BLUE,
         os.path.join(assets_dir, "generated", "icon_map.png")),
        ("03", "MAPS & PRICING TOOLS", "Interactive visuals anyone can use",
         ["India map showing tutor counts by state.",
          "3-subject Venn diagram showing subject overlaps.",
          "Searchable web directory with filters.",
          "Export to Word (.docx) or PDF reports."],
         ACCENT_EMERALD, PILL_MINT,
         os.path.join(assets_dir, "generated", "icon_bio.png")),
        ("04", "FAIR PRICING CALCULATOR", "ML-based fee estimator",
         ["Predicts hourly fees from experience, city, and subject.",
          "Removes extreme outlier fees (e.g., ₹99,999/hr).",
          "AI writes parent and student profile bios.",
          "Bio scorecard checks how strong your profile is."],
         ACCENT_AMBER, PILL_CORAL,
         os.path.join(assets_dir, "generated", "icon_price.png")),
    ]
    col_w = Inches(3.03)
    gap = Inches(0.08)
    top_y = Inches(2.20)
    col_h = Inches(4.45)

    for i, (num, stage_title, subtitle, bullets, accent, pill_bg, doodle_path) in enumerate(steps):
        left_x = Inches(0.55 + i * (col_w.inches + gap.inches))
        add_rect(slide, left_x, top_y, col_w, col_h, CARD_COLOR,
                 line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, left_x, top_y, col_w, Inches(0.10), accent)
        add_pill(slide, left_x + Inches(0.20), top_y + Inches(0.22),
                 Inches(0.55), Inches(0.38), num, accent,
                 text_color=WHITE, font_size=13)
        add_text(slide, left_x + Inches(0.20), top_y + Inches(0.72),
                 col_w - Inches(0.40), Inches(0.30), stage_title,
                 font_size=11.5, font_color=accent, bold=True)
        add_text(slide, left_x + Inches(0.20), top_y + Inches(1.02),
                 col_w - Inches(0.40), Inches(0.35), subtitle,
                 font_size=12.5, font_color=TEXT_DARK, bold=True)
        add_rect(slide, left_x + Inches(0.20), top_y + Inches(1.47),
                 col_w - Inches(0.40), Pt(1), RGBColor(0xEB, 0xF0, 0xF5))
        bullet_text = "\n".join([f"•  {b}" for b in bullets])
        add_text(slide, left_x + Inches(0.20), top_y + Inches(1.60),
                 col_w - Inches(0.55), Inches(2.70), bullet_text,
                 font_size=10, font_color=TEXT_MEDIUM, line_spacing=1.15)
        if doodle_path and os.path.exists(doodle_path):
            add_image(slide, doodle_path, left_x + Inches(col_w.inches - 1.15),
                      top_y + Inches(col_h.inches - 1.05),
                      Inches(1.05), Inches(0.95), opacity=35)

    return slide

# ============================================================================
# SLIDE 4: WEBSITE PLATFORM
# ============================================================================
def slide_website(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    website_bg = os.path.join(assets_dir, "generated", "website_bg.png")
    add_image(slide, website_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.2), Inches(0.38),
             "WEBSITE", ACCENT_EMERALD, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   04 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "Website Platform — Tutor Explorer",
           "Starlette + uvicorn web app with faceted filters, pagination, "
           "and downloadable exports.", y_offset=0.55)

    features = [
        ("🔍", "Faceted Filters", "Region, state, subject, qualification, occupation, gender, online mode", ACCENT_INDIGO),
        ("📄", "Pagination", "Browse 107,723 tutors in pages of 15–100 rows", ACCENT_EMERALD),
        ("📥", "Custom Exports", "Download filtered Excel or editable DOCX reports", ACCENT_CORAL),
        ("🗺️", "Live Maps", "India choropleth + district-level drilldown", ACCENT_VIOLET),
        ("💰", "Price Predictor", "ML-powered hourly fee estimation", ACCENT_AMBER),
        ("✨", "Bio Generator", "NLP + AI tutor profile bios for parents & students", ACCENT_INDIGO),
    ]
    y = Inches(2.15)
    col1_x = Inches(0.55)
    col2_x = Inches(4.65)
    card_w = Inches(3.95)
    card_h = Inches(1.55)
    for i, (emoji, title, desc, accent) in enumerate(features):
        x = col1_x if i % 2 == 0 else col2_x
        y_pos = y if i < 2 else (y + Inches(1.65)) if i < 4 else (y + Inches(3.30))
        add_rect(slide, x, y_pos, card_w, card_h, CARD_COLOR,
                 line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, x, y_pos, Inches(0.08), card_h, accent)
        add_text(slide, x + Inches(0.25), y_pos + Inches(0.12), Inches(0.45), Inches(0.45),
                 emoji, font_size=18, font_color=accent)
        add_text(slide, x + Inches(0.75), y_pos + Inches(0.12), card_w - Inches(1.0), Inches(0.32),
                 title, font_size=12, font_color=TEXT_DARK, bold=True)
        add_text(slide, x + Inches(0.75), y_pos + Inches(0.48), card_w - Inches(0.95), Inches(0.95),
                 desc, font_size=9.5, font_color=TEXT_MEDIUM, line_spacing=1.05)

    chart_path = os.path.join(assets_dir, "charts", "chart_regional_distribution.png")
    if chart_path and os.path.exists(chart_path):
        add_rect(slide, Inches(8.75), Inches(2.15), Inches(4.03), Inches(4.40),
                 PILL_BLUE, opacity=25, line_color=RGBColor(0xE2, 0xE8, 0xF0),
                 line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_image(slide, chart_path, Inches(9.00), Inches(2.40), Inches(3.55), Inches(3.80))
        add_rect(slide, Inches(8.75), Inches(2.15), Inches(4.03), Inches(0.50),
                 CARD_COLOR, opacity=95)
        add_text(slide, Inches(8.95), Inches(2.22), Inches(3.65), Inches(0.35),
                 "Live Dashboard Preview", font_size=11.5, font_color=ACCENT_INDIGO,
                 bold=True)

    doodle = os.path.join(assets_dir, "doodles", "doodle_slide15_web.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(10.70), Inches(5.80), Inches(2.05), Inches(1.25),
                  opacity=40)

    return slide

# ============================================================================
# SLIDE 5: DATA CLEANING
# ============================================================================
def slide_cleaning(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cleaning_bg = os.path.join(assets_dir, "generated", "cleaning_bg.png")
    add_image(slide, cleaning_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.8), Inches(0.38),
             "DATA PIPELINE", ACCENT_EMERALD, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   05 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "Cleaning the Data — Removing Ghost Accounts and Bad Entries",
           "Raw tutor directories have test accounts and abandoned profiles. "
           "Here is how we audited every one.", y_offset=0.55)

    left_x = Inches(0.55)
    left_w = Inches(5.85)
    top_y = Inches(2.20)
    card_h = Inches(4.55)

    add_rect(slide, left_x, top_y, left_w, card_h, CARD_COLOR,
             line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, left_x, top_y, left_w, Inches(0.10), ACCENT_EMERALD)
    add_text(slide, left_x + Inches(0.30), top_y + Inches(0.25),
             left_w - Inches(0.6), Inches(0.32),
             "TUTOR STATUS BREAKDOWN", font_size=11, font_color=ACCENT_EMERALD, bold=True)
    add_text(slide, left_x + Inches(0.30), top_y + Inches(0.58),
             left_w - Inches(0.6), Inches(0.38),
             "Which Tutors Are Still Active", font_size=16, font_color=TEXT_DARK, bold=True)
    tiers = [
        ("76,706 Tutors (71.2%)", "VERIFIED ACTIVE",
         "Tutors whose last login differs from registration — they still log in and respond to students.",
         ACCENT_EMERALD, PILL_MINT),
        ("31,017 Tutors (28.8%)", "DORMANT OR REMOVED",
         "Accounts created once and never used, or profiles that no longer exist. Kept in archive but not in pricing.",
         ACCENT_AMBER, PILL_CORAL),
        ("1,268 Tutors (1.2%)", "NO ACTIVITY DATA",
         "Profiles with missing login dates, verified via headless browser automation.",
         ACCENT_CORAL, PILL_BLUE),
    ]
    for j, (count_label, badge_label, desc, bar_col, pill_bg) in enumerate(tiers):
        item_y = top_y + Inches(1.10 + j * 1.25)
        add_rect(slide, left_x + Inches(0.30), item_y, left_w - Inches(0.60),
                 Inches(1.10), PILL_BLUE, opacity=15)
        add_rect(slide, left_x + Inches(0.30), item_y, Inches(0.08), Inches(1.10), bar_col)
        add_text(slide, left_x + Inches(0.48), item_y + Inches(0.12),
                 Inches(3.0), Inches(0.26), count_label,
                 font_size=13, font_color=TEXT_DARK, bold=True)
        add_pill(slide, left_x + Inches(3.55), item_y + Inches(0.12),
                 Inches(1.9), Inches(0.24), badge_label, pill_bg,
                 text_color=bar_col, font_size=8, bold=True)
        add_text(slide, left_x + Inches(0.48), item_y + Inches(0.42),
                 left_w - Inches(1.10), Inches(0.62), desc,
                 font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.05)

    right_x = Inches(6.70)
    right_w = Inches(6.08)
    add_rect(slide, right_x, top_y, right_w, card_h, CARD_COLOR,
             line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, right_x, top_y, right_w, Inches(0.10), ACCENT_INDIGO)
    add_text(slide, right_x + Inches(0.30), top_y + Inches(0.25),
             right_w - Inches(0.6), Inches(0.32),
             "CLEANING RULES WE APPLIED", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_text(slide, right_x + Inches(0.30), top_y + Inches(0.58),
             right_w - Inches(0.6), Inches(0.38),
             "Fixing 107k Records", font_size=16, font_color=TEXT_DARK, bold=True)
    rules = [
        ("Fake Fees (₹50 to ₹5,000/hr)",
         "Tossed out dummy entries like ₹99,999/hr or ₹0/hr using statistical rules so fee averages stay realistic."),
        ("Subject Name Mapping",
         "Raw listings had thousands of spelling variants ('Math', 'Maths', 'Mathematics 10th'). We grouped them into 12,327 standard subjects."),
        ("City Tiers & Regions",
         "Sorted every educator into Tier 1 (Metros), Tier 2, or Tier 3 / Semi-Urban, using Census city classifications."),
        ("Class Levels",
         "Split offerings into 4 stages: Primary (Class 1–5), Middle (Class 6–8), Senior Secondary (Class 9–12), and Higher Ed / Competitive Exams."),
    ]
    for j, (rtitle, rdesc) in enumerate(rules):
        rule_y = top_y + Inches(1.10 + j * 0.82)
        add_pill(slide, right_x + Inches(0.30), rule_y,
                 Inches(0.40), Inches(0.28), str(j + 1), ACCENT_INDIGO,
                 text_color=WHITE, font_size=11, bold=True)
        add_text(slide, right_x + Inches(0.80), rule_y,
                 right_w - Inches(1.2), Inches(0.25), rtitle,
                 font_size=11.5, font_color=TEXT_DARK, bold=True)
        add_text(slide, right_x + Inches(0.80), rule_y + Inches(0.26),
                 right_w - Inches(1.2), Inches(0.48), rdesc,
                 font_size=10, font_color=TEXT_MEDIUM, line_spacing=1.05)

    doodle = os.path.join(assets_dir, "doodles", "doodle_slide6_chsc.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(0.55), Inches(5.85), Inches(2.2), Inches(1.10),
                  opacity=40)

    return slide

# ============================================================================
# SLIDE 6: PRICING CALCULATOR
# ============================================================================
def slide_pricing(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    pricing_bg = os.path.join(assets_dir, "generated", "pricing_bg.png")
    add_image(slide, pricing_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.2), Inches(0.38),
             "PRICING", ACCENT_AMBER, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   06 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "Fair Pricing Calculator",
           "ML-based fee estimator — predicts hourly rates from experience, city, "
           "and subject.", y_offset=0.55)

    stats = [
        ("₹500/hr", "Median hourly rate", ACCENT_INDIGO),
        ("₹300/hr", "25th percentile (budget)", ACCENT_EMERALD),
        ("₹800/hr", "75th percentile (premium)", ACCENT_AMBER),
        ("R² = 0.84", "ML model accuracy", ACCENT_CORAL),
    ]
    sx = Inches(0.55)
    sy = Inches(2.20)
    sw = Inches(2.90)
    sh = Inches(1.30)
    gap = Inches(0.13)
    for i, (val, label, accent) in enumerate(stats):
        x = sx if i % 2 == 0 else sx + Inches(3.03)
        y = sy if i < 2 else sy + Inches(1.43)
        add_rect(slide, x, y, sw, sh, PILL_BLUE, opacity=35)
        add_rect(slide, x, y, Inches(0.08), sh, accent)
        add_text(slide, x + Inches(0.25), y + Inches(0.10),
                 sw - Inches(0.45), Inches(0.55), val,
                 font_size=22, font_color=TEXT_DARK, bold=True)
        add_text(slide, x + Inches(0.25), y + Inches(0.72),
                 sw - Inches(0.45), Inches(0.40), label,
                 font_size=10, font_color=TEXT_MEDIUM)

    rx = Inches(6.70)
    add_rect(slide, rx, Inches(2.20), Inches(6.08), Inches(3.10), CARD_COLOR,
             line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, rx, Inches(2.20), Inches(6.08), Inches(0.10), ACCENT_AMBER)
    add_text(slide, rx + Inches(0.30), Inches(2.35),
             Inches(5.6), Inches(0.35),
             "How the Fee Calculator Works", font_size=14, font_color=TEXT_DARK, bold=True)
    bands = [
        ("Experience", "+₹50–₹150 per year of teaching"),
        ("City Tier", "Tier-1 metros pay 20–30% more than Tier-3"),
        ("Subject", "STEM & coding command higher rates than arts"),
        ("Online Mode", "Online tutoring adds reach, slightly lower base rate"),
    ]
    by = Inches(2.80)
    for btitle, bdesc in bands:
        add_text(slide, rx + Inches(0.35), by, Inches(2.5), Inches(0.30),
                 btitle, font_size=11, font_color=TEXT_DARK, bold=True)
        add_text(slide, rx + Inches(2.90), by, Inches(3.4), Inches(0.55),
                 bdesc, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.0)
        by += Inches(0.58)

    chart_path = os.path.join(assets_dir, "charts", "chart_fees_by_category.png")
    if chart_path and os.path.exists(chart_path):
        add_rect(slide, Inches(0.55), Inches(5.55), Inches(12.23), Inches(1.45),
                 PILL_MINT, opacity=20, line_color=ACCENT_EMERALD, line_width=Pt(1),
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_image(slide, chart_path, Inches(0.80), Inches(5.65),
                  Inches(11.7), Inches(1.30))
        add_rect(slide, Inches(0.55), Inches(5.55), Inches(0.08), Inches(1.45), ACCENT_EMERALD)

    doodle = os.path.join(assets_dir, "doodles", "doodle_metric_vcliff.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(10.70), Inches(5.90), Inches(1.95), Inches(1.20),
                  opacity=45)

    return slide

# ============================================================================
# SLIDE 7: BIO GENERATOR
# ============================================================================
def slide_bio(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bio_bg = os.path.join(assets_dir, "generated", "bio_bg.png")
    add_image(slide, bio_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.2), Inches(0.38),
             "TOOLS", ACCENT_VIOLET, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   07 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "AI Bio Generator for Tutors",
           "Turns rough tutor descriptions into two versions — one for parents, "
           "one for students.", y_offset=0.55)

    add_rect(slide, Inches(0.55), Inches(2.15), Inches(5.95), Inches(2.60),
             PILL_BLUE, opacity=18, line_color=ACCENT_INDIGO, line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, Inches(0.55), Inches(2.15), Inches(0.08), Inches(2.60), ACCENT_INDIGO)
    add_text(slide, Inches(0.80), Inches(2.30), Inches(5.5), Inches(0.30),
             "For Parents (Reassuring & Results-Focused)", font_size=13,
             font_color=TEXT_DARK, bold=True)
    add_text(slide, Inches(0.80), Inches(2.65), Inches(5.5), Inches(2.0),
             "Hi, I'm Anita S. I hold an M.Sc in Mathematics and have spent the last 6 years "
             "teaching Math to CBSE Class 9–10 students.\n\nMy approach is simple: concept-first "
             "learning with real examples. I've helped 200+ students achieve 90%+ in board exams.\n\n"
             "I'm comfortable teaching in English and Hindi, and I keep parents updated after "
             "every milestone.\n\nI offer a first trial class so you can see if my teaching "
             "style works for your child.",
             font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)

    add_rect(slide, Inches(6.83), Inches(2.15), Inches(5.95), Inches(2.60),
             PILL_MINT, opacity=18, line_color=ACCENT_EMERALD, line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, Inches(6.83), Inches(2.15), Inches(0.08), Inches(2.60), ACCENT_EMERALD)
    add_text(slide, Inches(7.08), Inches(2.30), Inches(5.5), Inches(0.30),
             "For Students (Friendly & Encouraging)", font_size=13,
             font_color=TEXT_DARK, bold=True)
    add_text(slide, Inches(7.08), Inches(2.65), Inches(5.5), Inches(2.0),
             "Hey, I'm Anita (M.Sc Math). I've been teaching Math for 6 years, and I know "
             "it can feel overwhelming at first.\n\nHere's how I do it: we skip boring "
             "lectures and break things down into real examples. You can ask me anything — "
             "no question is too small.\n\nWe'll go over shortcuts and exam strategies that "
             "actually save time.\n\nBook a session and let's get started!",
             font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)

    add_rect(slide, Inches(0.55), Inches(4.95), Inches(12.23), Inches(1.95), CARD_COLOR,
             line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, Inches(0.55), Inches(4.95), Inches(0.08), Inches(1.95), ACCENT_VIOLET)
    add_text(slide, Inches(0.80), Inches(5.10), Inches(11.8), Inches(0.35),
             "How the Bio Generator Works", font_size=14, font_color=TEXT_DARK, bold=True)
    steps = [
        ("1", "Paste your rough bio",
         "Enter your qualifications, experience, subjects, and teaching style in plain text."),
        ("2", "AI generates two versions",
         "One bio written for parents, one written for students — different tones, same facts."),
        ("3", "Get a quality score",
         "Bio scorecard checks credentials, experience, subject clarity, and call-to-action."),
    ]
    sx = Inches(1.05)
    for n, t, d in steps:
        add_text(slide, sx, Inches(5.45), Inches(0.45), Inches(0.35),
                 n, font_size=14, font_color=ACCENT_VIOLET, bold=True)
        add_text(slide, sx + Inches(0.55), Inches(5.43), Inches(3.6), Inches(0.30),
                 t, font_size=11.5, font_color=TEXT_DARK, bold=True)
        add_text(slide, sx + Inches(0.55), Inches(5.73), Inches(3.6), Inches(0.55),
                 d, font_size=10, font_color=TEXT_MEDIUM, line_spacing=1.0)
        sx += Inches(4.00)

    doodle = os.path.join(assets_dir, "doodles", "doodle_card_s.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(10.75), Inches(5.85), Inches(2.0), Inches(1.15),
                  opacity=40)

    return slide

# ============================================================================
# SLIDE 8: GEOSPATIAL MAP
# ============================================================================
def slide_map(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    map_bg = os.path.join(assets_dir, "generated", "map_bg.png")
    add_image(slide, map_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.8), Inches(0.38),
             "GEOSPATIAL", ACCENT_CORAL, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   08 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "All-India Tutor Map",
           "State boundaries colored by tutor supply volume and average hourly rates.", y_offset=0.55)

    map_path = os.path.join(assets_dir, "maps", "all_india_tutor_choropleth.png")
    if map_path and os.path.exists(map_path):
        add_rect(slide, Inches(0.55), Inches(2.20), Inches(7.0), Inches(4.60),
                 WHITE, opacity=95, line_color=RGBColor(0xE2, 0xE8, 0xF0),
                 line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_image(slide, map_path, Inches(0.80), Inches(2.40), Inches(6.5), Inches(4.25))

    rx = Inches(7.80)
    add_rect(slide, rx, Inches(2.20), Inches(4.98), Inches(4.60), CARD_COLOR,
             line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, rx, Inches(2.20), Inches(4.98), Inches(0.10), ACCENT_CORAL)
    add_text(slide, rx + Inches(0.25), Inches(2.35),
             Inches(4.6), Inches(0.35),
             "Key Insights from the Map", font_size=13, font_color=TEXT_DARK, bold=True)
    insights = [
        ("Maharashtra leads", "Highest tutor concentration — Mumbai & Pune are major hubs."),
        ("Tier-3 gaps", "Smaller cities have fewer tutors per student — opportunity for new educators."),
        ("Fee variation", "Tier-1 metros average ₹650–₹800/hr; Tier-3 averages ₹350–₹500/hr."),
        ("Online reach", "Online tutoring lets Tier-3 educators serve Tier-1 students at metro rates."),
    ]
    iy = Inches(2.85)
    for t, d in insights:
        add_text(slide, rx + Inches(0.35), iy, Inches(2.3), Inches(0.35),
                 t, font_size=11.5, font_color=TEXT_DARK, bold=True)
        add_text(slide, rx + Inches(0.35), iy + Inches(0.35),
                 Inches(4.4), Inches(0.55), d, font_size=10, font_color=TEXT_MEDIUM,
                 line_spacing=1.05)
        iy += Inches(0.95)

    add_text(slide, Inches(0.55), Inches(7.00), Inches(3.0), Inches(0.30),
             "Regional Breakdown", font_size=11, font_color=TEXT_DARK, bold=True)
    zone_paths = [
        os.path.join(assets_dir, "maps", "map_zone_north.png"),
        os.path.join(assets_dir, "maps", "map_zone_south.png"),
        os.path.join(assets_dir, "maps", "map_zone_west.png"),
        os.path.join(assets_dir, "maps", "map_zone_east_central.png"),
    ]
    zone_names = ["North", "South", "West", "East & Central"]
    zx = Inches(0.55)
    for zp, zn in zip(zone_paths, zone_names):
        if zp and os.path.exists(zp):
            add_image(slide, zp, zx, Inches(7.02), Inches(2.85), Inches(0.42), opacity=75)
            add_text(slide, zx, Inches(7.50), Inches(2.95), Inches(0.25),
                     zn + " Zone", font_size=8.5, font_color=TEXT_LIGHT, align=PP_ALIGN.CENTER)
            zx += Inches(3.08)

    doodle = os.path.join(assets_dir, "doodles", "doodle_triangles_and_hubs.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(10.75), Inches(2.15), Inches(2.10), Inches(1.30),
                  opacity=40)

    return slide

# ============================================================================
# SLIDE 9: VENN DIAGRAM
# ============================================================================
def slide_venn(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    venn_bg = os.path.join(assets_dir, "generated", "venn_bg.png")
    add_image(slide, venn_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.8), Inches(0.38),
             "ANALYTICS", ACCENT_EMERALD, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   09 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "3-Subject Overlap Tool",
           "See which tutors teach multiple subjects together — and what that means "
           "for pricing.", y_offset=0.55)

    venn_path = os.path.join(assets_dir, "charts", "venn_stem_trio.png")
    if venn_path and os.path.exists(venn_path):
        add_rect(slide, Inches(0.55), Inches(2.20), Inches(5.70), Inches(4.60),
                 PILL_MINT, opacity=20, line_color=ACCENT_EMERALD, line_width=Pt(1),
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, Inches(0.55), Inches(2.20), Inches(5.70), Inches(0.10), ACCENT_EMERALD)
        add_text(slide, Inches(0.80), Inches(2.32), Inches(5.2), Inches(0.30),
                 "Subject Combination Analysis", font_size=13, font_color=TEXT_DARK, bold=True)
        add_image(slide, venn_path, Inches(0.85), Inches(2.75), Inches(5.1), Inches(3.95))

    rx = Inches(6.50)
    add_rect(slide, rx, Inches(2.20), Inches(6.28), Inches(4.60), CARD_COLOR,
             line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, rx, Inches(2.20), Inches(6.28), Inches(0.10), ACCENT_INDIGO)
    add_text(slide, rx + Inches(0.25), Inches(2.35),
             Inches(5.9), Inches(0.35),
             "What the Overlap Reveals", font_size=13, font_color=TEXT_DARK, bold=True)
    overlaps = [
        ("Math + Science", "28% bundle these — school parents want one tutor for both."),
        ("Physics + Chemistry", "Senior secondary specialists command a 40% fee premium."),
        ("English + Coding", "Growing demand for spoken English + basic programming."),
        ("Single Subject", "Most tutors specialize — only 1 in 5 teaches 3+ subjects."),
    ]
    iy = Inches(2.85)
    for t, d in overlaps:
        add_text(slide, rx + Inches(0.35), iy, Inches(2.5), Inches(0.35),
                 t, font_size=11.5, font_color=TEXT_DARK, bold=True)
        add_text(slide, rx + Inches(0.35), iy + Inches(0.35),
                 Inches(5.7), Inches(0.55), d, font_size=10, font_color=TEXT_MEDIUM,
                 line_spacing=1.05)
        iy += Inches(0.95)

    add_text(slide, Inches(0.55), Inches(6.95), Inches(12.0), Inches(0.35),
             "Multi-subject tutors can offer bundle discounts — parents save, tutors earn more per hour.",
             font_size=11, font_color=TEXT_MEDIUM, italic=True)

    doodle = os.path.join(assets_dir, "doodles", "doodle_triangles_and_hubs.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(10.75), Inches(5.85), Inches(2.05), Inches(1.25),
                  opacity=40)

    return slide

# ============================================================================
# SLIDE 10: TEAM
# ============================================================================
def slide_team(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    team_bg = os.path.join(assets_dir, "generated", "team_bg.png")
    add_image(slide, team_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.2), Inches(0.38),
             "TEAM", ACCENT_CORAL, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   10 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "Project Authors",
           "DA-IICT Capstone Group 18 — three students, one shared mission: transparency in tutoring.",
           y_offset=0.55)

    authors = [
        ("01", "Nilesh Mori", "202301473", ACCENT_INDIGO,
         "Built the 3NF database normalization, scraping pipeline, PII protection, "
         "and India state map cartography."),
        ("02", "Harsh Maheriya", "202301470", ACCENT_EMERALD,
         "Built the 3-subject Venn diagram studio, fee regression model, "
         "and grade-wise pricing."),
        ("03", "Tarang Hirapara", "202301207", ACCENT_CORAL,
         "Built the client dashboard, tutor verification system, "
         "and AI bio generator."),
    ]
    ax = Inches(0.55)
    for num, name, roll, accent, desc in authors:
        add_rect(slide, ax, Inches(2.15), Inches(3.95), Inches(2.70), CARD_COLOR,
                 line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, ax, Inches(2.15), Inches(0.08), Inches(2.70), accent)
        add_pill(slide, ax + Inches(0.20), Inches(2.35), Inches(0.55), Inches(0.40),
                 num, accent, text_color=WHITE, font_size=14)
        add_text(slide, ax + Inches(0.90), Inches(2.33), Inches(2.9), Inches(0.35),
                 name, font_size=15, font_color=TEXT_DARK, bold=True)
        add_text(slide, ax + Inches(0.90), Inches(2.67), Inches(2.9), Inches(0.25),
                 roll, font_size=10, font_color=TEXT_LIGHT)
        add_text(slide, ax + Inches(0.20), Inches(2.95), Inches(3.55), Inches(1.75),
                 desc, font_size=10, font_color=TEXT_MEDIUM, line_spacing=1.15)
        ax += Inches(4.12)

    add_rect(slide, Inches(0.55), Inches(5.10), Inches(12.23), Inches(1.85),
             RGBColor(0xEE, 0xF2, 0xF7), opacity=85, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, Inches(0.55), Inches(5.10), Inches(0.08), Inches(1.85), ACCENT_CORAL)
    add_text(slide, Inches(0.85), Inches(5.25), Inches(3.0), Inches(0.30),
             "FACULTY MENTOR", font_size=10, font_color=ACCENT_CORAL, bold=True)
    add_text(slide, Inches(0.85), Inches(5.55), Inches(3.0), Inches(0.35),
             "Dr. Parul Gupta", font_size=14, font_color=TEXT_DARK, bold=True)
    add_text(slide, Inches(0.85), Inches(5.90), Inches(3.0), Inches(0.90),
             "DA-IICT, Gandhinagar — guided our research approach, data verification, "
             "and pricing validation.",
             font_size=10, font_color=TEXT_MEDIUM, line_spacing=1.1)

    add_text(slide, Inches(4.20), Inches(5.25), Inches(8.5), Inches(0.30),
             "Dhirubhai Ambani Institute of Information and Communication Technology, Gandhinagar, Gujarat",
             font_size=10, font_color=TEXT_LIGHT)

    doodle = os.path.join(assets_dir, "doodles", "doodle_squad_healer.png")
    if doodle and os.path.exists(doodle):
        add_image(slide, doodle, Inches(10.75), Inches(5.85), Inches(2.05), Inches(1.25),
                  opacity=40)

    return slide

# ============================================================================
# SLIDE 11: CONCLUSION / TAKEAWAYS
# ============================================================================
def slide_conclusion(prs, assets_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    conclusion_bg = os.path.join(assets_dir, "generated", "conclusion_bg.png")
    add_image(slide, conclusion_bg, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, opacity=100)
    add_rect(slide, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, CARD_TINT, opacity=70)
    add_rect(slide, Inches(0.25), Inches(0.25), Inches(12.83), Inches(7.0),
             CARD_COLOR, opacity=92, line_color=RGBColor(0xE2, 0xE8, 0xF0),
             line_width=Pt(1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_pill(slide, Inches(0.55), Inches(0.45), Inches(2.8), Inches(0.38),
             "CONCLUSION", ACCENT_AMBER, text_color=WHITE, font_size=11)
    add_text(slide, Inches(8.0), Inches(0.48), Inches(4.8), Inches(0.34),
             "DA-IICT  •  BMP TUTORS   •   11 / 12", font_size=11,
             font_color=TEXT_LIGHT, align=PP_ALIGN.RIGHT)
    header(slide, "Key Takeaways",
           "What we built, what we learned, and why it matters.", y_offset=0.55)

    takeaways = [
        ("01", "Real Transparency",
         "Turned messy tutor listings into a clean public list of 107k educators "
         "and 12.3k subjects.", ACCENT_INDIGO),
        ("02", "Fair Pricing",
         "Trained on 72,001 records with bad data removed — parents and new "
         "tutors get fair hourly rate estimates.", ACCENT_EMERALD),
        ("03", "Interactive Tools",
         "India state maps, 440+ district directories, 3-subject overlap tool, "
         "and bio generator.", ACCENT_CORAL),
        ("04", "Verified Data",
         "76,706 active tutors confirmed live — 31,017 dormant/removed accounts "
         "excluded from pricing.", ACCENT_AMBER),
    ]
    col_w = Inches(2.95)
    gap = Inches(0.13)
    top_y = Inches(2.20)
    col_h = Inches(3.90)

    for i, (tnum, thead, tbody, taccent) in enumerate(takeaways):
        tx = Inches(0.55 + i * (col_w.inches + gap.inches))
        add_rect(slide, tx, top_y, col_w, col_h, CARD_COLOR,
                 line_color=RGBColor(0xE2, 0xE8, 0xF0), line_width=Pt(1),
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(slide, tx, top_y, col_w, Inches(0.08), taccent)
        add_pill(slide, tx + Inches(0.20), top_y + Inches(0.25),
                 Inches(0.55), Inches(0.40), tnum, taccent,
                 text_color=WHITE, font_size=14)
        add_text(slide, tx + Inches(0.90), top_y + Inches(0.25),
                 col_w - Inches(1.1), Inches(0.35), thead,
                 font_size=13, font_color=TEXT_DARK, bold=True)
        add_text(slide, tx + Inches(0.20), top_y + Inches(0.75),
                 col_w - Inches(0.40), Inches(3.0), tbody,
                 font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.2)

    add_rect(slide, Inches(0.55), Inches(6.35), Inches(12.23), Inches(0.65),
             PILL_CORAL, opacity=55, line_color=ACCENT_CORAL, line_width=Pt(1),
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(slide, Inches(0.85), Inches(6.42), Inches(11.8), Inches(0.50),
             "\u201cWe now welcome your questions.\u201d — Thank you, professors.",
             font_size=14, font_color=ACCENT_CORAL, bold=True, align=PP_ALIGN.CENTER)

    return slide

# ============================================================================
# MAIN
# ============================================================================
def main():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    assets_dir = ASSETS_DIR

    slide_title(prs, assets_dir)
    slide_why(prs, assets_dir)
    slide_pipeline(prs, assets_dir)
    slide_website(prs, assets_dir)
    slide_cleaning(prs, assets_dir)
    slide_pricing(prs, assets_dir)
    slide_bio(prs, assets_dir)
    slide_map(prs, assets_dir)
    slide_venn(prs, assets_dir)
    slide_team(prs, assets_dir)
    slide_conclusion(prs, assets_dir)

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "BMP_Tutors_Beautiful_Presentation.pptx")
    prs.save(output_path)
    print("Presentation saved: " + output_path)
    print("Total slides: " + str(len(prs.slides._sldIdLst)))

if __name__ == "__main__":
    main()
