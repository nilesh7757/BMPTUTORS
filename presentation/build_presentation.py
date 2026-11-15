#!/usr/bin/env python3
"""
Build BMP Tutors Master Presentation PPTX (20 Slides).
Matches Group18_SC452 reference visual style:
- 13.33 x 7.50 in (16:9 Widescreen)
- Textured background pattern (Picture 2) + rounded white card overlay
- Coral & Indigo pill badges + category tags + slide counters
- High-res choropleths, Venn diagrams, demographic bar charts, and UI mockups
- Clear, simple English throughout with zero confusing jargon
"""

import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from pptx import Presentation
import theme

from slides.slide_01_title import add_slide as s01
from slides.slide_02_agenda import add_slide as s02
from slides.slide_03_problem import add_slide as s03
from slides.slide_04_cleaning import add_slide as s04
from slides.slide_05_database import add_slide as s05
from slides.slide_06_webapp import add_slide as s06
from slides.slide_07_national_map import add_slide as s07
from slides.slide_08_regional_zones import add_slide as s08
from slides.slide_09_city_tiers import add_slide as s09
from slides.slide_10_subjects import add_slide as s10
from slides.slide_11_venn_studio import add_slide as s11
from slides.slide_12_gender import add_slide as s12
from slides.slide_13_experience import add_slide as s13
from slides.slide_14_languages import add_slide as s14
from slides.slide_15_fees import add_slide as s15
from slides.slide_16_grades_pricing import add_slide as s16
from slides.slide_17_occupations import add_slide as s17
from slides.slide_18_ml_pricing import add_slide as s18
from slides.slide_19_nlp_bio import add_slide as s19
from slides.slide_20_thankyou import add_slide as s20


def build():
    print("================================================================")
    print("  BMP Tutors — Building Master Presentation (Group 18 Style)")
    print("================================================================")

    prs = Presentation()
    prs.slide_width = theme.SLIDE_WIDTH
    prs.slide_height = theme.SLIDE_HEIGHT

    assets_dir = os.path.join(PROJECT_ROOT, "assets")

    slide_builders = [
        ("01", "Title & Team", s01),
        ("02", "Why This Matters", s02),
        ("03", "Project Roadmap", s03),
        ("04", "Data Cleaning & Verification", s04),
        ("05", "3NF Relational Database", s05),
        ("06", "Interactive Web Platform", s06),
        ("07", "National Map & State Distribution", s07),
        ("08", "4 Regional Geographic Zones", s08),
        ("09", "City Tiers & Online Reach", s09),
        ("10", "Academic Subject Demand", s10),
        ("11", "3-Set Venn Diagram Studio", s11),
        ("12", "Gender Representation", s12),
        ("13", "Teaching Experience Profile", s13),
        ("14", "Linguistic Demographics", s14),
        ("15", "Hourly Fee Economics", s15),
        ("16", "Pricing by Education Level", s16),
        ("17", "Gig Economy & Occupations", s17),
        ("18", "Smart Pricing ML Engine", s18),
        ("19", "AI Profile Studio & Bio Optimizer", s19),
        ("20", "Summary & Thank You", s20),
    ]

    for num, name, builder in slide_builders:
        print(f"  [Slide {num}/20] Building {name}...")
        builder(prs, assets_dir)

    output_dir = os.path.join(PROJECT_ROOT, "output")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "BMP_Tutors_Master_Presentation.pptx")
    prs.save(output_path)

    # Also save a copy directly in project root for easy user access
    root_copy = os.path.join(os.path.dirname(PROJECT_ROOT), "BMP_Tutors_Master_Presentation.pptx")
    prs.save(root_copy)

    print("\n================================================================")
    print(f"✅ Presentation saved successfully!")
    print(f"   Primary Output: {output_path}")
    print(f"   Root Access:    {root_copy}")
    print(f"   Slide Count:    {len(prs.slides)} slides")
    print(f"   Dimensions:     {prs.slide_width / 914400:.2f} x {prs.slide_height / 914400:.2f} inches (16:9)")
    print("================================================================")
    return output_path


if __name__ == "__main__":
    build()
