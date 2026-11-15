"""
Slide 6: Live Web Analytics Application Architecture & Features.
Matches Group 18 reference: Feature checklist + Browser mockup container.
"""

import os
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    add_slide_base,
    add_slide_header,
    add_rectangle,
    add_textbox,
    add_pill,
    add_card_box,
    add_image_fit,
    ACCENT_CORAL,
    ACCENT_INDIGO,
    ACCENT_EMERALD,
    ACCENT_AMBER,
    TEXT_DARK,
    TEXT_MEDIUM,
    TEXT_LIGHT,
    WHITE,
    PILL_BLUE,
    PILL_MINT,
    PILL_CORAL,
    BG_SUBTLE,
    BORDER_SOFT,
    FONT_TITLE,
    FONT_BODY,
)


def add_slide(prs, assets_dir):
    slide, _ = add_slide_base(
        prs,
        category_pill_text="WEB APPLICATION",
        slide_num_str="06 / 20",
        top_tag="PRODUCTION PLATFORM",
        assets_dir=assets_dir,
        pill_color=ACCENT_INDIGO,
    )

    add_slide_header(
        slide,
        title="Interactive Web Platform — Built for Speed and Simplicity",
        subtitle="A full-featured web application that makes all 107,723 records instantly searchable, filterable, and visual.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT COLUMN: 4 PLATFORM FEATURES (Width: 5.6in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(5.6)

    features = [
        (
            "⚡ High-Speed Search (<50ms)",
            "Instant autocomplete search across 107k tutor names, IDs, subjects, and cities using indexed SQLite queries.",
            ACCENT_INDIGO,
        ),
        (
            "🛡️ Verified Profile Audit",
            "Clear status indicators show whether each educator is Verified Active (87.4k) or Inactive Dormant (19.0k).",
            ACCENT_EMERALD,
        ),
        (
            "🗺️ Real-Time Visual Analytics",
            "Interactive All-India state choropleths, 440+ district boundary directory, and dynamic 3-Subject Venn studio.",
            ACCENT_CORAL,
        ),
        (
            "📥 Multi-Format Reports & Exports",
            "Download custom normalized Excel workbooks, native DrawingML Word reports (.docx), and printable PDF dossiers.",
            ACCENT_AMBER,
        ),
    ]

    for i, (ftitle, fdesc, faccent) in enumerate(features):
        fy = top_y + Inches(i * 1.18)
        add_card_box(slide, left_x, fy, left_w, Inches(1.08), fill_color=WHITE)
        add_rectangle(slide, left_x, fy, Inches(0.08), Inches(1.08), faccent, shape=MSO_SHAPE.RECTANGLE)

        add_textbox(slide, left_x + Inches(0.25), fy + Inches(0.12), left_w - Inches(0.4), Inches(0.30), text=ftitle, font_size=13, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, left_x + Inches(0.25), fy + Inches(0.44), left_w - Inches(0.4), Inches(0.55), text=fdesc, font_size=11, font_color=TEXT_MEDIUM, line_spacing=1.15)

    # -------------------------------------------------------------
    # RIGHT COLUMN: BROWSER MOCKUP CONTAINER (Width: 6.4in)
    # -------------------------------------------------------------
    right_x = Inches(6.45)
    right_w = Inches(6.30)

    # Browser Window Outer
    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)

    # Browser Top Bar
    bar_h = Inches(0.40)
    add_rectangle(slide, right_x, top_y, right_w, bar_h, RGBColor(0x0F, 0x17, 0x2A), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Window dots
    add_rectangle(slide, right_x + Inches(0.20), top_y + Inches(0.14), Inches(0.12), Inches(0.12), RGBColor(0xEF, 0x44, 0x44), shape=MSO_SHAPE.OVAL)
    add_rectangle(slide, right_x + Inches(0.38), top_y + Inches(0.14), Inches(0.12), Inches(0.12), RGBColor(0xF5, 0x9E, 0x0B), shape=MSO_SHAPE.OVAL)
    add_rectangle(slide, right_x + Inches(0.56), top_y + Inches(0.14), Inches(0.12), Inches(0.12), RGBColor(0x10, 0xB9, 0x81), shape=MSO_SHAPE.OVAL)

    # Address bar text
    add_textbox(slide, right_x + Inches(0.85), top_y + Inches(0.08), Inches(3.5), Inches(0.26), text="https://bmptutors.onrender.com", font_size=10, font_color=RGBColor(0x94, 0xA3, 0xB8))

    # Dashboard screenshot / graphic
    dash_img = os.path.join(assets_dir, "maps", "india_all_states_labeled_dashboard.png")
    if os.path.exists(dash_img):
        add_image_fit(
            slide,
            dash_img,
            left=right_x + Inches(0.15),
            top=top_y + Inches(0.50),
            max_width=right_w - Inches(0.30),
            max_height=card_h - Inches(0.60),
            align_center=True,
            valign_center=True,
        )
