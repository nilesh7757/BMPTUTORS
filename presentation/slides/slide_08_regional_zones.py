"""
Slide 8: Regional Deep-Dive — 4 Geographic Zones Matrix.
Matches Group 18 reference: 4 visual zone cards with embedded zone maps & regional benchmarks.
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
        category_pill_text="REGIONAL STUDY",
        slide_num_str="08 / 20",
        top_tag="4 ZONAL CLUSTERS",
        assets_dir=assets_dir,
        pill_color=ACCENT_INDIGO,
    )

    add_slide_header(
        slide,
        title="Regional Landscape — How Tutoring Varies Across 4 Zones",
        subtitle="Comparing student demand, hourly rates, and educator volumes across South, North, West, and East India.",
    )

    # 4 Cards side-by-side (Width: 2.92in each)
    col_w = Inches(2.92)
    top_y = Inches(2.10)
    col_h = Inches(4.75)

    zones = [
        (
            "SOUTH ZONE",
            "38,450 Tutors (35.7%)",
            "₹745 / hr Avg Fee",
            "map_zone_south.png",
            "Karnataka, Tamil Nadu, Telangana, Kerala. Tech & coding dominance with heavy B.Tech graduate tutor participation.",
            ACCENT_INDIGO,
            PILL_BLUE,
        ),
        (
            "NORTH ZONE",
            "31,210 Tutors (29.0%)",
            "₹825 / hr Avg Fee",
            "map_zone_north.png",
            "Delhi NCR, UP, Punjab, Rajasthan. Highest concentration of competitive exam prep (JEE, NEET, UPSC, SSC).",
            ACCENT_CORAL,
            PILL_CORAL,
        ),
        (
            "WEST ZONE",
            "22,140 Tutors (20.5%)",
            "₹810 / hr Avg Fee",
            "map_zone_west.png",
            "Maharashtra and Gujarat. Major commercial tuition hub with strong focus on Commerce, Accountancy, and School Boards.",
            ACCENT_AMBER,
            PILL_BLUE,
        ),
        (
            "EAST & CENTRAL",
            "15,923 Tutors (14.8%)",
            "₹680 / hr Avg Fee",
            "map_zone_east_central.png",
            "West Bengal, Bihar, Odisha, MP. Strong focus on foundational Science and Bengali/Hindi vernacular instruction.",
            ACCENT_EMERALD,
            PILL_MINT,
        ),
    ]

    for i, (zname, zcount, zfee, img_name, zdesc, accent, pill_bg) in enumerate(zones):
        left_x = Inches(0.55 + i * (2.92 + 0.20))

        add_card_box(slide, left_x, top_y, col_w, col_h, fill_color=WHITE)
        add_rectangle(slide, left_x, top_y, col_w, Inches(0.08), accent, shape=MSO_SHAPE.RECTANGLE)

        # Header & Badges
        add_textbox(slide, left_x + Inches(0.18), top_y + Inches(0.18), col_w - Inches(0.36), Inches(0.28), text=zname, font_size=13, font_color=accent, bold=True)
        add_textbox(slide, left_x + Inches(0.18), top_y + Inches(0.48), col_w - Inches(0.36), Inches(0.24), text=zcount, font_size=11, font_color=TEXT_DARK, bold=True)
        add_pill(slide, left_x + Inches(0.18), top_y + Inches(0.76), Inches(1.30), Inches(0.26), zfee, pill_bg, text_color=accent, font_size=9, bold=True)

        # Zone Map Image Container
        img_path = os.path.join(assets_dir, "maps", img_name)
        if os.path.exists(img_path):
            add_image_fit(
                slide,
                img_path,
                left=left_x + Inches(0.15),
                top=top_y + Inches(1.10),
                max_width=col_w - Inches(0.30),
                max_height=Inches(2.20),
                align_center=True,
                valign_center=True,
            )

        # Zone Description
        add_textbox(slide, left_x + Inches(0.18), top_y + Inches(3.40), col_w - Inches(0.36), Inches(1.20), text=zdesc, font_size=10, font_color=TEXT_MEDIUM, line_spacing=1.15)
