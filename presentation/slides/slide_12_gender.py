"""
Slide 12: Gender Distribution & Domain Representation.
Matches Group 18 reference: Gender split metrics + Domain comparison chips + Key social findings.
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
        category_pill_text="DEMOGRAPHICS",
        slide_num_str="12 / 20",
        top_tag="GENDER REPRESENTATION",
        assets_dir=assets_dir,
        pill_color=ACCENT_CORAL,
    )

    add_slide_header(
        slide,
        title="Gender Distribution — Women Lead in Private Tutoring",
        subtitle="Women represent 58.5% of all private tutors in India, leading heavily in Languages, Biology, and School Education.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT CARD: OVERALL GENDER METRICS & DOMAIN SPLITS (Width: 6.4in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(6.35)

    add_card_box(slide, left_x, top_y, left_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, left_w, Inches(0.08), ACCENT_CORAL, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.20), left_w - Inches(0.6), Inches(0.25), text="NATIONAL GENDER BREAKDOWN", font_size=11, font_color=ACCENT_CORAL, bold=True)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.48), left_w - Inches(0.6), Inches(0.35), text="Women: 58.5%  |  Men: 41.2%", font_size=16, font_color=TEXT_DARK, bold=True)

    # 2 Comparison Blocks
    # Block A: Women Leading
    add_card_box(slide, left_x + Inches(0.25), top_y + Inches(0.95), left_w - Inches(0.50), Inches(1.65), fill_color=BG_SUBTLE)
    add_pill(slide, left_x + Inches(0.40), top_y + Inches(1.08), Inches(2.2), Inches(0.28), "WOMEN PREDOMINANT", PILL_CORAL, text_color=ACCENT_CORAL, font_size=9, bold=True)

    women_domains = (
        "• Languages & Humanities: 72.4% Women (English, Hindi, French)\n"
        "• Biology & Life Sciences: 68.1% Women (Medical Foundation)\n"
        "• Primary & Middle School: 74.2% Women (Foundational Class 1–8)\n"
        "• Arts & Creative Crafts: 78.5% Women (Drawing, Music, Design)"
    )
    add_textbox(slide, left_x + Inches(0.40), top_y + Inches(1.42), left_w - Inches(0.8), Inches(1.10), text=women_domains, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)

    # Block B: Men Leading
    add_card_box(slide, left_x + Inches(0.25), top_y + Inches(2.75), left_w - Inches(0.50), Inches(1.65), fill_color=BG_SUBTLE)
    add_pill(slide, left_x + Inches(0.40), top_y + Inches(2.88), Inches(2.0), Inches(0.28), "MEN PREDOMINANT", PILL_BLUE, text_color=ACCENT_INDIGO, font_size=9, bold=True)

    men_domains = (
        "• Computer Programming & Tech: 64.5% Men (Python, Java, Web Dev)\n"
        "• Physics & Mechanics: 58.2% Men (Class 11–12 Senior STEM)\n"
        "• Competitive JEE / GATE Math: 54.1% Men (Advanced Coaching)\n"
        "• Commerce & Accountancy: 51.8% Men (CA Foundation & Tax)"
    )
    add_textbox(slide, left_x + Inches(0.40), top_y + Inches(3.22), left_w - Inches(0.8), Inches(1.10), text=men_domains, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)

    # -------------------------------------------------------------
    # RIGHT CARD: SOCIAL & ECONOMIC INSIGHTS (Width: 5.7in)
    # -------------------------------------------------------------
    right_x = Inches(7.10)
    right_w = Inches(5.65)

    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, right_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), right_w - Inches(0.6), Inches(0.25), text="KEY GENDER INSIGHTS", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), right_w - Inches(0.6), Inches(0.35), text="Economic Empowerment", font_size=16, font_color=TEXT_DARK, bold=True)

    insights = [
        ("Flexible Self-Employment", "Private home and online tuition provides highly educated women with independent income, flexible working hours, and work-from-home options."),
        ("Strong Fee Parity in Core Fields", "Unlike corporate wage gaps, female tutors in Languages and Biology charge exactly on par with male peers (₹740/hr vs ₹735/hr), showing market pricing equality based on skill."),
        ("The Coding Pipeline Gap", "Male dominance in programming tuition reflects wider university engineering enrollment trends, highlighting the need for more female coding mentors."),
    ]

    for j, (ititle, idesc) in enumerate(insights):
        iy = top_y + Inches(1.05 + j * 1.15)
        add_pill(slide, right_x + Inches(0.30), iy, Inches(0.45), Inches(0.30), str(j + 1), ACCENT_INDIGO, text_color=WHITE, font_size=11, bold=True)
        add_textbox(slide, right_x + Inches(0.85), iy, right_w - Inches(1.1), Inches(0.25), text=ititle, font_size=12.5, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, right_x + Inches(0.85), iy + Inches(0.28), right_w - Inches(1.1), Inches(0.75), text=idesc, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)
