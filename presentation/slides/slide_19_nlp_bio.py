"""
Slide 19: NLP Studio & Generative AI Bio Optimizer.
Matches Group 18 reference: Dual-audience profile generator cards + Bio quality audit scorecard.
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
        category_pill_text="AI GENERATION",
        slide_num_str="19 / 20",
        top_tag="NLP BIO OPTIMIZER",
        assets_dir=assets_dir,
        pill_color=ACCENT_CORAL,
    )

    add_slide_header(
        slide,
        title="Generative AI — Dual-Audience Profile & Bio Studio",
        subtitle="Using Gemini 2.5 Flash to automatically rewrite raw tutor profiles into compelling, high-converting pitches.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)
    col_w = Inches(5.95)

    # -------------------------------------------------------------
    # LEFT COLUMN: DUAL-AUDIENCE REASONING (Width: 5.95in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    add_card_box(slide, left_x, top_y, col_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, col_w, Inches(0.08), ACCENT_CORAL, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.20), col_w - Inches(0.6), Inches(0.25), text="DUAL-AUDIENCE GENERATION", font_size=11, font_color=ACCENT_CORAL, bold=True)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.48), col_w - Inches(0.6), Inches(0.35), text="Speaking to Two Different Decision Makers", font_size=16, font_color=TEXT_DARK, bold=True)

    # Pitch 1: Parents
    add_card_box(slide, left_x + Inches(0.25), top_y + Inches(0.95), col_w - Inches(0.50), Inches(1.65), fill_color=BG_SUBTLE)
    add_pill(slide, left_x + Inches(0.40), top_y + Inches(1.08), Inches(1.80), Inches(0.28), "PITCH FOR PARENTS", PILL_CORAL, text_color=ACCENT_CORAL, font_size=9, bold=True)

    parent_desc = (
        "• Tone: Professional, disciplined, and results-focused.\n"
        "• Key Elements: Highlights past student board exam scores (90%+), punctuality, weekly progress tests, and regular parent feedback updates."
    )
    add_textbox(slide, left_x + Inches(0.40), top_y + Inches(1.45), col_w - Inches(0.80), Inches(1.00), text=parent_desc, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)

    # Pitch 2: Students
    add_card_box(slide, left_x + Inches(0.25), top_y + Inches(2.75), col_w - Inches(0.50), Inches(1.65), fill_color=BG_SUBTLE)
    add_pill(slide, left_x + Inches(0.40), top_y + Inches(2.88), Inches(1.80), Inches(0.28), "PITCH FOR STUDENTS", PILL_BLUE, text_color=ACCENT_INDIGO, font_size=9, bold=True)

    student_desc = (
        "• Tone: Friendly, patient, approachable, and encouraging.\n"
        "• Key Elements: Focuses on making hard Math/Physics concepts intuitive, removing exam fear, interactive problem-solving, and zero-judgment doubt clearing."
    )
    add_textbox(slide, left_x + Inches(0.40), top_y + Inches(3.25), col_w - Inches(0.80), Inches(1.00), text=student_desc, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)

    # -------------------------------------------------------------
    # RIGHT COLUMN: AUTOMATED BIO AUDIT SCORECARD (Width: 5.95in)
    # -------------------------------------------------------------
    right_x = Inches(6.80)
    add_card_box(slide, right_x, top_y, col_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, col_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), col_w - Inches(0.6), Inches(0.25), text="AUTOMATED PROFILE AUDIT", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), col_w - Inches(0.6), Inches(0.35), text="Real-Time Bio Quality Scorecard", font_size=16, font_color=TEXT_DARK, bold=True)

    scores = [
        ("Clarity & Readability", "4.8 / 5.0", "Checks sentence structure, eliminating grammar errors and awkward phrasing."),
        ("Proof of Qualifications", "4.9 / 5.0", "Verifies explicit mentions of degrees (B.Tech, B.Ed, M.Sc) and years of experience."),
        ("Teaching Methodology", "4.7 / 5.0", "Ensures the bio describes how the tutor teaches (custom notes, weekly mock tests)."),
        ("Call to Action & Conversion", "4.6 / 5.0", "Includes clear booking triggers (free 30-min trial class, instant message link)."),
    ]

    for j, (slabel, sval, sdetail) in enumerate(scores):
        sy = top_y + Inches(0.95 + j * 0.88)
        add_card_box(slide, right_x + Inches(0.25), sy, col_w - Inches(0.50), Inches(0.80), fill_color=BG_SUBTLE)

        add_textbox(slide, right_x + Inches(0.40), sy + Inches(0.10), Inches(3.4), Inches(0.25), text=slabel, font_size=12, font_color=TEXT_DARK, bold=True)
        add_pill(slide, right_x + Inches(4.30), sy + Inches(0.10), Inches(1.10), Inches(0.24), sval, PILL_MINT, text_color=ACCENT_EMERALD, font_size=9.5, bold=True)
        add_textbox(slide, right_x + Inches(0.40), sy + Inches(0.38), col_w - Inches(0.80), Inches(0.38), text=sdetail, font_size=9.5, font_color=TEXT_MEDIUM)

    # Summary insight
    add_card_box(slide, right_x + Inches(0.25), top_y + Inches(3.95), col_w - Inches(0.50), Inches(0.65), fill_color=WHITE)
    add_textbox(slide, right_x + Inches(0.35), top_y + Inches(4.00), col_w - Inches(0.70), Inches(0.55), text="💡 Impact: Helps talented tutors who may lack polished English copywriting skills present their genuine credentials professionally.", font_size=9.5, font_color=TEXT_MEDIUM, line_spacing=1.1)
