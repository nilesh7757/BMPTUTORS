"""
Slide 11: Dynamic 3-Set Venn Diagram Studio & Multi-Subject Bundles.
Matches Group 18 reference: Dual visual Venn diagram cards (School vs STEM Trio).
"""

import os
from pptx.util import Inches, Pt
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
    FONT_TITLE,
    FONT_BODY,
)


def add_slide(prs, assets_dir):
    slide, _ = add_slide_base(
        prs,
        category_pill_text="VENN ANALYTICS",
        slide_num_str="11 / 20",
        top_tag="SUBJECT OVERLAPS",
        assets_dir=assets_dir,
        pill_color=ACCENT_EMERALD,
    )

    add_slide_header(
        slide,
        title="Subject Overlaps — The 3-Way Venn Diagram Studio",
        subtitle="Analyzing tutors who teach multiple subjects simultaneously: Foundational School bundles vs Senior STEM trios.",
    )

    card_w = Inches(5.95)
    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT CARD: FOUNDATIONAL SCHOOL TRIO (Math + Science + English)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    add_card_box(slide, left_x, top_y, card_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, card_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.25), top_y + Inches(0.18), card_w - Inches(0.5), Inches(0.25), text="FOUNDATIONAL SCHOOL BUNDLE", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, left_x + Inches(0.25), top_y + Inches(0.44), card_w - Inches(0.5), Inches(0.35), text="Mathematics + Science + English", font_size=15, font_color=TEXT_DARK, bold=True)

    # Venn Diagram Image
    v1_path = os.path.join(assets_dir, "charts", "venn_school_trio.png")
    if os.path.exists(v1_path):
        add_image_fit(
            slide,
            v1_path,
            left=left_x + Inches(0.20),
            top=top_y + Inches(0.85),
            max_width=card_w - Inches(0.40),
            max_height=Inches(2.50),
            align_center=True,
            valign_center=True,
        )

    # Metrics Box
    v1_metrics = (
        "• Total Combined Pool: 58,051 unique educators across the 3 subjects.\n"
        "• All 3 Subjects Masters: 5,512 tutors (9.5% of the pool) teach all three.\n"
        "• Average Hourly Rate: ₹694 / hr (12% lower than hiring 3 separate specialists)."
    )
    add_card_box(slide, left_x + Inches(0.20), top_y + Inches(3.45), card_w - Inches(0.40), Inches(1.15), fill_color=BG_SUBTLE)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(3.52), card_w - Inches(0.60), Inches(1.00), text=v1_metrics, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)

    # -------------------------------------------------------------
    # RIGHT CARD: SENIOR STEM TRIO (Physics + Chemistry + Math)
    # -------------------------------------------------------------
    right_x = Inches(6.80)
    add_card_box(slide, right_x, top_y, card_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, card_w, Inches(0.08), ACCENT_CORAL, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.25), top_y + Inches(0.18), card_w - Inches(0.5), Inches(0.25), text="SENIOR SECONDARY STEM TRIO", font_size=11, font_color=ACCENT_CORAL, bold=True)
    add_textbox(slide, right_x + Inches(0.25), top_y + Inches(0.44), card_w - Inches(0.5), Inches(0.35), text="Physics + Chemistry + Mathematics (PCM)", font_size=15, font_color=TEXT_DARK, bold=True)

    v2_path = os.path.join(assets_dir, "charts", "venn_stem_trio.png")
    if os.path.exists(v2_path):
        add_image_fit(
            slide,
            v2_path,
            left=right_x + Inches(0.20),
            top=top_y + Inches(0.85),
            max_width=card_w - Inches(0.40),
            max_height=Inches(2.50),
            align_center=True,
            valign_center=True,
        )

    v2_metrics = (
        "• Total Combined Pool: 42,390 unique senior STEM educators.\n"
        "• All 3 PCM Masters: 3,840 tutors (9.1% of pool) teach full high-school science.\n"
        "• Average Hourly Rate: ₹865 / hr (+24.6% premium due to competitive JEE prep)."
    )
    add_card_box(slide, right_x + Inches(0.20), top_y + Inches(3.45), card_w - Inches(0.40), Inches(1.15), fill_color=BG_SUBTLE)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(3.52), card_w - Inches(0.60), Inches(1.00), text=v2_metrics, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)
