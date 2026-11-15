"""
Slide 20: Summary, Impact & Thank You — Matching Group 18 reference Slide 18.
Punchy bold typography + author cards + live repository & web application links.
"""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from theme import (
    add_slide_base,
    add_textbox,
    add_pill,
    add_author_card,
    add_card_box,
    add_rectangle,
    CARD_COLOR,
    ACCENT_CORAL,
    ACCENT_INDIGO,
    ACCENT_EMERALD,
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
        category_pill_text="THANK YOU",
        slide_num_str="20 / 20",
        top_tag="PROJECT SUMMARY & OPEN SOURCE",
        assets_dir=assets_dir,
        pill_color=ACCENT_CORAL,
    )

    # Big Punchy Statements (matches reference Slide 18)
    add_textbox(
        slide,
        Inches(0.65),
        Inches(1.20),
        Inches(12.0),
        Inches(0.65),
        text="107,723 EDUCATORS MAPPED.",
        font_size=36,
        font_color=ACCENT_INDIGO,
        bold=True,
        font_name=FONT_TITLE,
    )

    add_textbox(
        slide,
        Inches(0.65),
        Inches(1.85),
        Inches(12.0),
        Inches(0.65),
        text="FAIR PRICING STANDARDIZED.",
        font_size=36,
        font_color=TEXT_DARK,
        bold=True,
        font_name=FONT_TITLE,
    )

    add_textbox(
        slide,
        Inches(0.65),
        Inches(2.50),
        Inches(12.0),
        Inches(0.70),
        text="THANK YOU",
        font_size=42,
        font_color=ACCENT_CORAL,
        bold=True,
        font_name=FONT_TITLE,
    )

    # 3 Summary Pillar Cards
    card_y = Inches(3.40)
    card_w = Inches(3.85)
    card_h = Inches(2.15)

    takeaways = [
        (
            "01",
            "Real Transparency",
            "Turned messy tutor listings into a clean public list of 107k educators and 12.3k subjects.",
            ACCENT_INDIGO,
        ),
        (
            "02",
            "Fee Calculator",
            "Trained on 72,001 records with bad data removed, so parents and new tutors get fair hourly rate estimates.",
            ACCENT_EMERALD,
        ),
        (
            "03",
            "Interactive Tools",
            "India state maps, 440+ district directories, 3-subject overlap tool, and bio generator for parents and students.",
            ACCENT_CORAL,
        ),
    ]

    for i, (tnum, thead, tbody, taccent) in enumerate(takeaways):
        tx = Inches(0.65 + i * (3.85 + 0.25))
        add_card_box(slide, tx, card_y, card_w, card_h, fill_color=BG_SUBTLE)
        add_rectangle(
            slide, tx, card_y, card_w, Inches(0.06), taccent, shape=MSO_SHAPE.RECTANGLE
        )

        add_pill(
            slide,
            tx + Inches(0.20),
            card_y + Inches(0.20),
            Inches(0.50),
            Inches(0.32),
            tnum,
            taccent,
            text_color=WHITE,
            font_size=11,
            bold=True,
        )
        add_textbox(
            slide,
            tx + Inches(0.80),
            card_y + Inches(0.20),
            card_w - Inches(1.0),
            Inches(0.32),
            text=thead,
            font_size=13,
            font_color=TEXT_DARK,
            bold=True,
        )
        add_textbox(
            slide,
            tx + Inches(0.20),
            card_y + Inches(0.65),
            card_w - Inches(0.4),
            Inches(1.35),
            text=tbody,
            font_size=10.5,
            font_color=TEXT_MEDIUM,
            line_spacing=1.15,
        )

    # Author Cards at bottom
    ay = Inches(5.85)
    add_author_card(
        slide,
        Inches(0.65),
        ay,
        "01",
        "Nilesh Mori",
        "202301473",
        width=card_w,
        height=Inches(0.92),
        accent_color=ACCENT_INDIGO,
    )
    add_author_card(
        slide,
        Inches(4.75),
        ay,
        "02",
        "Harsh Maheriya",
        "202301470",
        width=card_w,
        height=Inches(0.92),
        accent_color=ACCENT_EMERALD,
    )
    add_author_card(
        slide,
        Inches(8.85),
        ay,
        "03",
        "Tarang Hirapara",
        "202301207",
        width=card_w,
        height=Inches(0.92),
        accent_color=ACCENT_CORAL,
    )
