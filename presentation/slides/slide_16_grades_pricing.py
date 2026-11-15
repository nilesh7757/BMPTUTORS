"""
Slide 16: Tuition Fees by Student Grade & Educational Stage.
Matches Group 18 reference: 4 progressive educational stage columns with escalation metrics.
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
        category_pill_text="PRICING MATRIX",
        slide_num_str="16 / 20",
        top_tag="GRADE-WISE PRICING",
        assets_dir=assets_dir,
        pill_color=ACCENT_INDIGO,
    )

    add_slide_header(
        slide,
        title="Rates Rise as Students Grow — The 4 Educational Stages",
        subtitle="Tuition pricing nearly triples as students advance from primary school to competitive college entrance exams.",
    )

    col_w = Inches(2.92)
    top_y = Inches(2.10)
    col_h = Inches(4.75)

    stages = [
        (
            "STAGE 1",
            "PRIMARY SCHOOL",
            "Class 1 to 5 (Ages 6–10)",
            "₹420 / hr",
            "Base Entry Fee",
            [
                "Focus: Foundational reading, writing, and elementary arithmetic.",
                "High female tutor share (74.2%) providing patient 1-on-1 homework help.",
                "Commonly booked as 5-day-per-week monthly tuition packages.",
            ],
            ACCENT_EMERALD,
            PILL_MINT,
        ),
        (
            "STAGE 2",
            "MIDDLE SCHOOL",
            "Class 6 to 8 (Ages 11–13)",
            "₹590 / hr",
            "+40.5% vs Primary",
            [
                "Focus: Transition from general science to Physics, Chemistry, and Algebra.",
                "Language grammar mastery (English, Hindi, Sanskrit, Regional).",
                "Shift from broad homework help to subject-specific clarity.",
            ],
            ACCENT_INDIGO,
            PILL_BLUE,
        ),
        (
            "STAGE 3",
            "SENIOR SECONDARY",
            "Class 9 to 12 (Ages 14–17)",
            "₹880 / hr",
            "+49.2% vs Middle",
            [
                "Focus: High-stakes 10th and 12th Board examinations (CBSE, ICSE, State).",
                "Heavy demand for specialized PCM / PCB educators.",
                "Tutors with past board evaluation experience command premium rates.",
            ],
            ACCENT_AMBER,
            PILL_CORAL,
        ),
        (
            "STAGE 4",
            "HIGHER ED & EXAMS",
            "College & Entrances (Ages 18+)",
            "₹1,240 / hr",
            "+40.9% vs Secondary",
            [
                "Focus: National competitive exams (JEE Main/Adv, NEET, CAT, GATE).",
                "University computer science (Data Structures, Algorithms, AI).",
                "Highest hourly rates in India, often taught by IIT/NIT graduates.",
            ],
            ACCENT_CORAL,
            PILL_BLUE,
        ),
    ]

    for i, (stagetag, stitle, ssubtitle, sfee, sdiff, sbullets, saccent, spill) in enumerate(stages):
        left_x = Inches(0.55 + i * (2.92 + 0.20))

        add_card_box(slide, left_x, top_y, col_w, col_h, fill_color=WHITE)
        add_rectangle(slide, left_x, top_y, col_w, Inches(0.08), saccent, shape=MSO_SHAPE.RECTANGLE)

        # Stage Tag
        add_pill(slide, left_x + Inches(0.18), top_y + Inches(0.20), Inches(1.10), Inches(0.26), stagetag, spill, text_color=saccent, font_size=8.5, bold=True)
        add_textbox(slide, left_x + Inches(0.18), top_y + Inches(0.54), col_w - Inches(0.36), Inches(0.28), text=stitle, font_size=13, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, left_x + Inches(0.18), top_y + Inches(0.82), col_w - Inches(0.36), Inches(0.24), text=ssubtitle, font_size=9.5, font_color=TEXT_LIGHT)

        # Fee Callout Box
        add_card_box(slide, left_x + Inches(0.15), top_y + Inches(1.15), col_w - Inches(0.30), Inches(0.75), fill_color=BG_SUBTLE)
        add_textbox(slide, left_x + Inches(0.25), top_y + Inches(1.20), col_w - Inches(0.50), Inches(0.35), text=sfee, font_size=18, font_color=saccent, bold=True)
        add_textbox(slide, left_x + Inches(0.25), top_y + Inches(1.55), col_w - Inches(0.50), Inches(0.30), text=sdiff, font_size=9, font_color=TEXT_MEDIUM)

        # Bullets
        bullet_text = "\n\n".join([f"• {b}" for b in sbullets])
        add_textbox(slide, left_x + Inches(0.18), top_y + Inches(2.05), col_w - Inches(0.36), Inches(2.55), text=bullet_text, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)
