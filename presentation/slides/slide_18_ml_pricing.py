"""
Slide 18: Machine Learning — Smart Fee Prediction & Pricing Engine.
Matches Group 18 reference: ML architecture pipeline card + Sample interactive prediction card.
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
        category_pill_text="SMART PRICING",
        slide_num_str="18 / 20",
        top_tag="MACHINE LEARNING ENGINE",
        assets_dir=assets_dir,
        pill_color=ACCENT_INDIGO,
    )

    add_slide_header(
        slide,
        title="Fair Price Prediction — Machine Learning Pricing Engine",
        subtitle="Trained on 72,001 verified outlier-free tutor profiles to instantly give parents and tutors fair market rates.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)
    col_w = Inches(5.95)

    # -------------------------------------------------------------
    # LEFT COLUMN: ML PIPELINE & FEATURES (Width: 5.95in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    add_card_box(slide, left_x, top_y, col_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, col_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.20), col_w - Inches(0.6), Inches(0.25), text="MODEL PIPELINE ARCHITECTURE", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.48), col_w - Inches(0.6), Inches(0.35), text="Trained on 72,001 Verified Profiles", font_size=16, font_color=TEXT_DARK, bold=True)

    steps = [
        ("1. Tukey IQR Outlier Cleansing", "Statistically filtered extreme placeholder entries (<₹50 or >₹5,000) using interquartile fences, eliminating test and dummy data."),
        ("2. Multi-Factor Feature Vector", "The model trains on 5 key parameters: City Tier (Metro/T2/T3), Teaching Experience (Yrs), Subject Domain, Class Level (1–12), and Delivery Mode (Online/Home)."),
        ("3. Instant Client-Side Inference", "Lightweight regression inference runs right in the browser in under 15ms without requiring heavy cloud GPU dependencies."),
    ]

    for j, (shead, sdesc) in enumerate(steps):
        sy = top_y + Inches(0.95 + j * 1.15)
        add_card_box(slide, left_x + Inches(0.25), sy, col_w - Inches(0.50), Inches(1.05), fill_color=BG_SUBTLE)
        add_textbox(slide, left_x + Inches(0.40), sy + Inches(0.10), col_w - Inches(0.8), Inches(0.25), text=shead, font_size=12, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, left_x + Inches(0.40), sy + Inches(0.38), col_w - Inches(0.8), Inches(0.62), text=sdesc, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)

    # -------------------------------------------------------------
    # RIGHT COLUMN: REAL-WORLD PREDICTION EXAMPLE (Width: 5.95in)
    # -------------------------------------------------------------
    right_x = Inches(6.80)
    add_card_box(slide, right_x, top_y, col_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, col_w, Inches(0.08), ACCENT_EMERALD, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), col_w - Inches(0.6), Inches(0.25), text="LIVE ESTIMATION DEMO", font_size=11, font_color=ACCENT_EMERALD, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), col_w - Inches(0.6), Inches(0.35), text="Real-World Prediction Example", font_size=16, font_color=TEXT_DARK, bold=True)

    # Live Scenario Card
    scen_y = top_y + Inches(0.95)
    add_card_box(slide, right_x + Inches(0.25), scen_y, col_w - Inches(0.50), Inches(1.50), fill_color=BG_SUBTLE)

    demo_text = (
        "• Location: Bengaluru (Tier 1 Metro)\n"
        "• Subject: Mathematics (CBSE Board)\n"
        "• Experience: 5 Years Teaching Experience\n"
        "• Student Stage: Senior Secondary (Class 10)"
    )
    add_textbox(slide, right_x + Inches(0.40), scen_y + Inches(0.12), Inches(3.2), Inches(1.25), text=demo_text, font_size=11, font_color=TEXT_MEDIUM, line_spacing=1.2)

    # Big Result Badge
    add_card_box(slide, right_x + Inches(3.70), scen_y + Inches(0.20), Inches(1.80), Inches(1.10), fill_color=WHITE)
    add_textbox(slide, right_x + Inches(3.80), scen_y + Inches(0.28), Inches(1.60), Inches(0.35), text="₹850 / hr", font_size=22, font_color=ACCENT_EMERALD, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, right_x + Inches(3.80), scen_y + Inches(0.68), Inches(1.60), Inches(0.50), text="Fair Market Rate\n(Band: ₹750–₹950)", font_size=9, font_color=TEXT_LIGHT, align=PP_ALIGN.CENTER)

    # Benefits
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(2.65), col_w - Inches(0.6), Inches(0.25), text="HOW THIS PROTECTS USERS", font_size=11, font_color=ACCENT_INDIGO, bold=True)

    bens = [
        "Protects Parents: Prevents agencies from charging inflated ₹1,500/hr rates for standard middle school classes.",
        "Empowers Tutors: Gives new educators confidence to charge fair rates without accidentally pricing themselves out.",
    ]
    for k, b in enumerate(bens):
        by = top_y + Inches(3.00 + k * 0.85)
        add_pill(slide, right_x + Inches(0.30), by, Inches(0.35), Inches(0.26), str(k + 1), ACCENT_INDIGO, text_color=WHITE, font_size=10, bold=True)
        add_textbox(slide, right_x + Inches(0.75), by, col_w - Inches(1.05), Inches(0.75), text=b, font_size=11, font_color=TEXT_MEDIUM, line_spacing=1.15)
