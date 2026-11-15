"""
Slide 14: Linguistic Demographics & Multilingual Reach.
Matches Group 18 reference: Language leaderboard + Bilingual pedagogy & Foreign language premium cards.
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
        category_pill_text="LANGUAGES",
        slide_num_str="14 / 20",
        top_tag="LINGUISTIC DIVERSITY",
        assets_dir=assets_dir,
        pill_color=ACCENT_AMBER,
    )

    add_slide_header(
        slide,
        title="Linguistic Diversity — India's Multilingual Tutoring Advantage",
        subtitle="78% of tutors speak 2 or more languages, creating a bridge between English curricula and regional mother tongues.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT COLUMN: TOP 8 SPOKEN LANGUAGES (Width: 6.2in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(6.20)

    add_card_box(slide, left_x, top_y, left_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, left_w, Inches(0.08), ACCENT_AMBER, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.20), left_w - Inches(0.6), Inches(0.25), text="SPOKEN LANGUAGES IN TUTORING", font_size=11, font_color=ACCENT_AMBER, bold=True)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.48), left_w - Inches(0.6), Inches(0.35), text="National Language Distribution", font_size=15, font_color=TEXT_DARK, bold=True)

    langs = [
        ("1. English", "104,820 Tutors", "97.3%", "Universal language of instruction"),
        ("2. Hindi", "68,410 Tutors", "63.5%", "Dominant in North and Central India"),
        ("3. Tamil", "14,210 Tutors", "13.2%", "Key Southern vernacular hub"),
        ("4. Telugu", "11,850 Tutors", "11.0%", "Telangana & Andhra Pradesh"),
        ("5. Malayalam", "8,420 Tutors", "7.8%", "High NRI Gulf student reach"),
        ("6. Bengali", "7,910 Tutors", "7.3%", "West Bengal & Kolkata hub"),
        ("7. Marathi", "6,850 Tutors", "6.4%", "State Board & Mumbai bilinguals"),
        ("8. Kannada", "6,120 Tutors", "5.7%", "Bengaluru & Karnataka hub"),
    ]

    for j, (lname, lcount, lpct, lnote) in enumerate(langs):
        ly = top_y + Inches(0.92 + j * 0.44)
        add_textbox(slide, left_x + Inches(0.30), ly, Inches(1.8), Inches(0.24), text=lname, font_size=11, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, left_x + Inches(2.15), ly, Inches(1.2), Inches(0.24), text=lcount, font_size=10, font_color=TEXT_LIGHT)
        add_pill(slide, left_x + Inches(3.40), ly, Inches(0.65), Inches(0.22), lpct, PILL_BLUE, text_color=ACCENT_INDIGO, font_size=8.5, bold=True)
        add_textbox(slide, left_x + Inches(4.15), ly, Inches(1.8), Inches(0.24), text=lnote, font_size=9, font_color=TEXT_MEDIUM)

    # -------------------------------------------------------------
    # RIGHT COLUMN: 3 PEDAGOGY & MARKET TAKEAWAYS (Width: 5.8in)
    # -------------------------------------------------------------
    right_x = Inches(6.95)
    right_w = Inches(5.80)

    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, right_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), right_w - Inches(0.6), Inches(0.25), text="PEDAGOGY & VALUE", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), right_w - Inches(0.6), Inches(0.35), text="The Value of Mother-Tongue Teaching", font_size=15, font_color=TEXT_DARK, bold=True)

    points = [
        (
            "The Bilingual Concept Bridge",
            "While CBSE and ICSE exams are in English, students grasp tough Physics and Math concepts much faster when explained in their mother tongue (Hindi, Tamil, Marathi).",
            ACCENT_INDIGO,
        ),
        (
            "The Foreign Language Premium (+35%)",
            "Tutors teaching European languages (French, German, Spanish) command an average of ₹1,150 / hr, catering to international school students and study-abroad test takers.",
            ACCENT_CORAL,
        ),
        (
            "Global NRI Student Reach",
            "Indian tutors fluent in Malayalam, Telugu, and Hindi actively tutor Indian diaspora students in the UAE, Saudi Arabia, the UK, and the USA via Zoom.",
            ACCENT_EMERALD,
        ),
    ]

    for j, (phead, pbody, paccent) in enumerate(points):
        py = top_y + Inches(1.05 + j * 1.15)
        add_pill(slide, right_x + Inches(0.30), py, Inches(0.45), Inches(0.30), str(j + 1), paccent, text_color=WHITE, font_size=11, bold=True)
        add_textbox(slide, right_x + Inches(0.85), py, right_w - Inches(1.1), Inches(0.25), text=phead, font_size=12.5, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, right_x + Inches(0.85), py + Inches(0.28), right_w - Inches(1.1), Inches(0.75), text=pbody, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)
