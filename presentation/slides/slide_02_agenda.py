"""
Slide 2: Why This Project Matters — Problem Statement, Market Context & Scope.
Matches Group 18 reference Slide 2: Category pill + KPI numbers + 3 problem/solution cards.
"""

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
        category_pill_text="WHY THIS MATTERS",
        slide_num_str="02 / 20",
        top_tag="SC452  •  BMP TUTORS",
        assets_dir=assets_dir,
        pill_color=ACCENT_CORAL,
    )

    add_slide_header(
        slide,
        title="India's Tutoring Market — The Need for Open Intelligence",
        subtitle="Millions of students take private tuition every day, yet parents, students, and tutors lack basic pricing benchmarks.",
    )

    # ---------------------------------------------------------
    # LEFT COLUMN: 3 BIG STAT CARDS (Width: 4.8in)
    # ---------------------------------------------------------
    left_x = Inches(0.55)
    card_w = Inches(4.7)

    stats = [
        ("107,723", "EDUCATORS MAPPED", "Complete national footprint across 28 States and 8 Union Territories", ACCENT_INDIGO, PILL_BLUE),
        ("₹788 / hr", "NATIONAL BENCHMARK FEE", "Average hourly rate, spanning from ₹250 in towns to ₹2,500+ in Metros", ACCENT_CORAL, PILL_CORAL),
        ("74.8%", "TEACHES ONLINE", "Over 80,000 tutors now reach students digitally beyond their hometown", ACCENT_EMERALD, PILL_MINT),
    ]

    for i, (val, title, desc, accent, pill_bg) in enumerate(stats):
        top_y = Inches(2.05 + i * 1.62)
        add_card_box(slide, left_x, top_y, card_w, Inches(1.48), fill_color=BG_SUBTLE)

        # Left color bar
        add_rectangle(slide, left_x, top_y, Inches(0.12), Inches(1.48), accent, shape=MSO_SHAPE.RECTANGLE)

        # Big Number
        add_textbox(slide, left_x + Inches(0.35), top_y + Inches(0.18), Inches(4.0), Inches(0.48), text=val, font_size=28, font_color=accent, bold=True)

        # Title badge
        add_textbox(slide, left_x + Inches(0.35), top_y + Inches(0.68), Inches(4.0), Inches(0.24), text=title, font_size=10, font_color=TEXT_LIGHT, bold=True)

        # Description
        add_textbox(slide, left_x + Inches(0.35), top_y + Inches(0.94), Inches(4.0), Inches(0.45), text=desc, font_size=11, font_color=TEXT_MEDIUM)

    # ---------------------------------------------------------
    # RIGHT COLUMN: 3 PROBLEM & SOLUTION CARDS (Width: 7.2in)
    # ---------------------------------------------------------
    right_x = Inches(5.55)
    right_w = Inches(7.2)

    cards = [
        (
            "01",
            "The Problem: Complete Price Opacity",
            "Parents frequently overpay or struggle to know what tuition is worth. Tutors entering the gig market have no clue whether to charge ₹300/hr or ₹1,200/hr for Class 10 CBSE Math.",
            ACCENT_CORAL,
        ),
        (
            "02",
            "The Challenge: Inactive & Fake Directory Listings",
            "Online tutor platforms are flooded with ghost signups. Almost 20% of tutor profiles are dormant one-time registrations with unstated fees and deleted contact links.",
            ACCENT_AMBER,
        ),
        (
            "03",
            "Our Solution: Open Analytics & Fair Price ML",
            "We cleaned 107k records into a verified relational database, built an interactive web dashboard, and trained a smart pricing engine to predict fair hourly fees instantly.",
            ACCENT_EMERALD,
        ),
    ]

    for i, (num, header_text, body_text, accent) in enumerate(cards):
        top_y = Inches(2.05 + i * 1.62)
        add_card_box(slide, right_x, top_y, right_w, Inches(1.48), fill_color=WHITE)

        # Number pill
        add_pill(slide, right_x + Inches(0.25), top_y + Inches(0.22), Inches(0.55), Inches(0.35), num, accent, text_color=WHITE, font_size=12, bold=True)

        # Header
        add_textbox(slide, right_x + Inches(0.95), top_y + Inches(0.22), right_w - Inches(1.2), Inches(0.32), text=header_text, font_size=14, font_color=TEXT_DARK, bold=True)

        # Body
        add_textbox(slide, right_x + Inches(0.95), top_y + Inches(0.60), right_w - Inches(1.2), Inches(0.75), text=body_text, font_size=11.5, font_color=TEXT_MEDIUM, line_spacing=1.15)
