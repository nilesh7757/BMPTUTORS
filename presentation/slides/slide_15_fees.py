"""
Slide 15: Hourly Tuition Fee Economics & Pricing Benchmarks.
Matches Group 18 reference: Fee distribution chart on left + Benchmark percentile cards on right.
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
        category_pill_text="PRICING MODELS",
        slide_num_str="15 / 20",
        top_tag="FEE ECONOMICS",
        assets_dir=assets_dir,
        pill_color=ACCENT_CORAL,
    )

    add_slide_header(
        slide,
        title="Hourly Tuition Fee Economics — What Do Parents Pay?",
        subtitle="Across 76,416 stating educators, the national average rate is ₹788 per hour, with 50% of tutors charging ₹400 to ₹1,000.",
    )

    top_y = Inches(2.10)
    card_h = Inches(4.75)

    # -------------------------------------------------------------
    # LEFT COLUMN: FEE DISTRIBUTION CHART (Width: 6.4in)
    # -------------------------------------------------------------
    left_x = Inches(0.55)
    left_w = Inches(6.35)

    add_card_box(slide, left_x, top_y, left_w, card_h, fill_color=WHITE)
    add_rectangle(slide, left_x, top_y, left_w, Inches(0.08), ACCENT_CORAL, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.20), left_w - Inches(0.6), Inches(0.25), text="FEE SPREAD BY ACADEMIC DOMAIN", font_size=11, font_color=ACCENT_CORAL, bold=True)
    add_textbox(slide, left_x + Inches(0.30), top_y + Inches(0.48), left_w - Inches(0.6), Inches(0.35), text="Hourly Tuition Band: ₹50 to ₹5,000 / hr", font_size=15, font_color=TEXT_DARK, bold=True)

    chart_path = os.path.join(assets_dir, "charts", "chart_fees_by_category.png")
    if os.path.exists(chart_path):
        add_image_fit(
            slide,
            chart_path,
            left=left_x + Inches(0.20),
            top=top_y + Inches(0.90),
            max_width=left_w - Inches(0.40),
            max_height=Inches(3.40),
            align_center=True,
            valign_center=True,
        )

    # -------------------------------------------------------------
    # RIGHT COLUMN: 4 PERCENTILE BENCHMARKS (Width: 5.7in)
    # -------------------------------------------------------------
    right_x = Inches(7.10)
    right_w = Inches(5.65)

    add_card_box(slide, right_x, top_y, right_w, card_h, fill_color=WHITE)
    add_rectangle(slide, right_x, top_y, right_w, Inches(0.08), ACCENT_INDIGO, shape=MSO_SHAPE.RECTANGLE)

    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.20), right_w - Inches(0.6), Inches(0.25), text="PRICING BENCHMARKS", font_size=11, font_color=ACCENT_INDIGO, bold=True)
    add_textbox(slide, right_x + Inches(0.30), top_y + Inches(0.48), right_w - Inches(0.6), Inches(0.35), text="Empirical Tuition Percentiles", font_size=15, font_color=TEXT_DARK, bold=True)

    percentiles = [
        ("P25: Entry Band", "₹400 / hr", "Primary & middle school homework help, semi-urban online tutors.", ACCENT_EMERALD, PILL_MINT),
        ("P50: Median Rate", "₹600 / hr", "The most common tuition fee in India for Class 9–10 board preparation.", ACCENT_INDIGO, PILL_BLUE),
        ("P75: Senior Band", "₹1,000 / hr", "Senior secondary (Class 11–12) STEM, Commerce, and ICSE coaching.", ACCENT_AMBER, PILL_CORAL),
        ("P90: Elite Band", "₹1,800+ / hr", "Competitive JEE/NEET entrance mentors, coding, and AP/IB curricula.", ACCENT_CORAL, PILL_BLUE),
    ]

    for j, (pname, pval, pdesc, paccent, ppill) in enumerate(percentiles):
        py = top_y + Inches(0.95 + j * 0.88)
        add_card_box(slide, right_x + Inches(0.25), py, right_w - Inches(0.50), Inches(0.80), fill_color=BG_SUBTLE)
        add_rectangle(slide, right_x + Inches(0.25), py, Inches(0.06), Inches(0.80), paccent, shape=MSO_SHAPE.RECTANGLE)

        add_textbox(slide, right_x + Inches(0.40), py + Inches(0.10), Inches(2.5), Inches(0.24), text=pname, font_size=12, font_color=TEXT_DARK, bold=True)
        add_pill(slide, right_x + Inches(3.90), py + Inches(0.10), Inches(1.20), Inches(0.24), pval, ppill, text_color=paccent, font_size=10, bold=True)
        add_textbox(slide, right_x + Inches(0.40), py + Inches(0.38), right_w - Inches(0.90), Inches(0.38), text=pdesc, font_size=9.5, font_color=TEXT_MEDIUM)

    # Key Takeaway Card at bottom
    add_card_box(slide, right_x + Inches(0.25), top_y + Inches(4.00), right_w - Inches(0.50), Inches(0.62), fill_color=WHITE)
    add_textbox(slide, right_x + Inches(0.35), top_y + Inches(4.06), right_w - Inches(0.70), Inches(0.50), text="💡 Note: The mean (₹788) is higher than the median (₹600) because elite specialized tutors pull up the top end of the curve.", font_size=9.5, font_color=TEXT_MEDIUM, line_spacing=1.1)
