"""
Slide 9: City Tiers & Urban Clustering — Metros vs Tier 2 & Small Towns.
Matches Group 18 reference: 3 detailed vertical comparison cards with metrics & findings.
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
        category_pill_text="URBAN TIERS",
        slide_num_str="09 / 20",
        top_tag="METRO VS TIER 2/3",
        assets_dir=assets_dir,
        pill_color=ACCENT_AMBER,
    )

    add_slide_header(
        slide,
        title="Urban Clustering — Metros vs Semi-Urban Towns",
        subtitle="Metro tutors command a 45% hourly price premium, but online platforms allow semi-urban tutors to reach students nationwide.",
    )

    card_w = Inches(3.85)
    top_y = Inches(2.10)
    card_h = Inches(4.75)

    tiers = [
        (
            "TIER 1 (METROS)",
            "Delhi NCR, Mumbai, Bengaluru, Chennai, Hyderabad, Kolkata",
            "52,140 Tutors",
            "48.4% of India",
            "₹940 / hr",
            "+19.3% vs national avg",
            "82.1%",
            "Online delivery share",
            [
                "Highest concentration of specialized tutors (Coding, Data Science, Advanced AP/IB).",
                "Higher cost of living drives premium hourly tuition rates.",
                "Heavy reliance on online 1-on-1 tutoring to bypass urban traffic congestion.",
            ],
            ACCENT_CORAL,
            PILL_CORAL,
        ),
        (
            "TIER 2 (REGIONAL HUBS)",
            "Pune, Ahmedabad, Jaipur, Lucknow, Chandigarh, Kochi, Indore",
            "34,220 Tutors",
            "31.8% of India",
            "₹715 / hr",
            "-9.3% vs national avg",
            "71.4%",
            "Online delivery share",
            [
                "Balanced hybrid model: tutors teach in-person local batches by day and online students by night.",
                "Strong focus on State Board, CBSE, and regional entrance examinations.",
                "Fastest growing cohort of full-time independent private educators.",
            ],
            ACCENT_INDIGO,
            PILL_BLUE,
        ),
        (
            "TIER 3 / SEMI-URBAN",
            "District headquarters, industrial towns, and semi-rural centers",
            "21,363 Tutors",
            "19.8% of India",
            "₹520 / hr",
            "-34.0% vs national avg",
            "64.2%",
            "Online delivery share",
            [
                "Offers the most affordable hourly rates, serving budget-conscious families.",
                "Skilled university graduates use online tutoring platforms to earn supplementary income.",
                "High bilingual demand: tutors explain complex science concepts in local vernacular languages.",
            ],
            ACCENT_EMERALD,
            PILL_MINT,
        ),
    ]

    for i, (tname, tcities, tcount, tpct, tfee, tfeediff, tonline, tonlinedesc, bullets, accent, pill_bg) in enumerate(tiers):
        left_x = Inches(0.55 + i * (3.85 + 0.32))

        add_card_box(slide, left_x, top_y, card_w, card_h, fill_color=WHITE)
        add_rectangle(slide, left_x, top_y, card_w, Inches(0.08), accent, shape=MSO_SHAPE.RECTANGLE)

        # Tier title & cities
        add_textbox(slide, left_x + Inches(0.22), top_y + Inches(0.18), card_w - Inches(0.44), Inches(0.28), text=tname, font_size=13, font_color=accent, bold=True)
        add_textbox(slide, left_x + Inches(0.22), top_y + Inches(0.46), card_w - Inches(0.44), Inches(0.38), text=tcities, font_size=9.5, font_color=TEXT_LIGHT)

        # 2 Metric Boxes inside card
        m_w = (card_w - Inches(0.55)) / 2
        # Box 1: Fee
        add_card_box(slide, left_x + Inches(0.22), top_y + Inches(0.92), m_w, Inches(0.85), fill_color=BG_SUBTLE)
        add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.98), m_w - Inches(0.15), Inches(0.32), text=tfee, font_size=16, font_color=accent, bold=True)
        add_textbox(slide, left_x + Inches(0.30), top_y + Inches(1.30), m_w - Inches(0.15), Inches(0.40), text=tfeediff, font_size=8.5, font_color=TEXT_MEDIUM)

        # Box 2: Volume
        add_card_box(slide, left_x + Inches(0.30) + m_w, top_y + Inches(0.92), m_w, Inches(0.85), fill_color=BG_SUBTLE)
        add_textbox(slide, left_x + Inches(0.38) + m_w, top_y + Inches(0.98), m_w - Inches(0.15), Inches(0.32), text=tcount, font_size=13, font_color=TEXT_DARK, bold=True)
        add_textbox(slide, left_x + Inches(0.38) + m_w, top_y + Inches(1.30), m_w - Inches(0.15), Inches(0.40), text=tpct, font_size=8.5, font_color=TEXT_MEDIUM)

        # Bullets
        bullet_text = "\n\n".join([f"• {b}" for b in bullets])
        add_textbox(slide, left_x + Inches(0.22), top_y + Inches(1.95), card_w - Inches(0.44), Inches(2.6), text=bullet_text, font_size=10.5, font_color=TEXT_MEDIUM, line_spacing=1.15)
