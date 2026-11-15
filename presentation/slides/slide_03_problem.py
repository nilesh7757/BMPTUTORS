"""
Slide 3: Project Roadmap — 4 Horizontal Stage Cards.
Matches Group 18 reference Slide 3: 4 distinct numbered columns (01, 02, 03, 04).
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
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
        category_pill_text="OUR ROADMAP",
        slide_num_str="03 / 20",
        top_tag="HOW WE BUILT IT",
        assets_dir=assets_dir,
        pill_color=ACCENT_INDIGO,
    )

    add_slide_header(
        slide,
        title="How We Built BMP Tutors — Four Steps",
        subtitle="From raw TeacherOn listings to a live website with verified tutors and fair pricing.",
    )

    # 4 Columns side-by-side
    col_w = Inches(2.92)
    gap = Inches(0.20)
    top_y = Inches(2.10)
    col_h = Inches(4.75)

    stages = [
        (
            "01",
            "SCRAPE & VERIFY TUTORS",
            "Collecting and checking every profile",
            [
                "Scraped 107,723 tutor profiles from TeacherOn.",
                "Checked login dates — 76,706 tutors are still active.",
                "31,017 accounts were dormant or removed.",
                "Used a headless browser to verify profiles are live.",
            ],
            ACCENT_CORAL,
            PILL_CORAL,
        ),
        (
            "02",
            "CLEANED INTO TWO TABLES",
            "Simple database structure",
            [
                "One table for tutors (107,723 rows, one per educator).",
                "One table for subjects (12,327 distinct subjects).",
                "No duplicate facts — each tutor appears exactly once.",
            ],
            ACCENT_INDIGO,
            PILL_BLUE,
        ),
        (
            "03",
            "MAPS & PRICING TOOLS",
            "Interactive visuals anyone can use",
            [
                "India map showing tutor counts by state.",
                "3-subject Venn diagram showing subject overlaps.",
                "Searchable web directory with filters.",
                "Export to Word (.docx) or PDF reports.",
            ],
            ACCENT_EMERALD,
            PILL_MINT,
        ),
        (
            "04",
            "FAIR PRICING CALCULATOR",
            "ML-based fee estimator",
            [
                "Predicts hourly fees from experience, city, and subject.",
                "Removes extreme outlier fees (e.g., ₹99,999/hr).",
                "AI writes parent and student profile bios.",
                "Bio scorecard checks how strong your profile is.",
            ],
            ACCENT_AMBER,
            PILL_BLUE,
        ),
    ]

    for i, (num, stage_title, subtitle, bullets, accent, pill_bg) in enumerate(stages):
        left_x = Inches(0.55 + i * (2.92 + 0.20))

        # Main Card Box
        add_card_box(slide, left_x, top_y, col_w, col_h, fill_color=WHITE)

        # Top accent bar
        add_rectangle(
            slide, left_x, top_y, col_w, Inches(0.10), accent, shape=MSO_SHAPE.RECTANGLE
        )

        # Step Number Badge
        add_pill(
            slide,
            left_x + Inches(0.20),
            top_y + Inches(0.25),
            Inches(0.65),
            Inches(0.40),
            num,
            accent,
            text_color=WHITE,
            font_size=13,
            bold=True,
        )

        # Stage Header
        add_textbox(
            slide,
            left_x + Inches(0.20),
            top_y + Inches(0.78),
            col_w - Inches(0.4),
            Inches(0.25),
            text=stage_title,
            font_size=11,
            font_color=accent,
            bold=True,
        )

        add_textbox(
            slide,
            left_x + Inches(0.20),
            top_y + Inches(1.05),
            col_w - Inches(0.4),
            Inches(0.35),
            text=subtitle,
            font_size=13,
            font_color=TEXT_DARK,
            bold=True,
        )

        # Divider line
        add_rectangle(
            slide,
            left_x + Inches(0.20),
            top_y + Inches(1.50),
            col_w - Inches(0.4),
            Pt(1),
            RGBColor(0xEB, 0xF0, 0xF5),
            shape=MSO_SHAPE.RECTANGLE,
        )

        # Bullets in clean simple English
        bullet_text = "\n\n".join([f"• {b}" for b in bullets])
        add_textbox(
            slide,
            left_x + Inches(0.20),
            top_y + Inches(1.65),
            col_w - Inches(0.4),
            Inches(2.9),
            text=bullet_text,
            font_size=10.5,
            font_color=TEXT_MEDIUM,
            line_spacing=1.15,
        )
