import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import os
import json

# Import the exact slide data from build_presentation_html.py
from build_presentation_html import SLIDES_DATA

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Color constants
BG_COLOR = RGBColor(11, 17, 32)        # #0B1120
CARD_BG = RGBColor(22, 34, 60)         # #16223C
BORDER_COLOR = RGBColor(56, 189, 248)  # #38BDF8
TEXT_WHITE = RGBColor(248, 250, 252)   # #F8FAFC
TEXT_MUTED = RGBColor(148, 163, 184)   # #94A3B8
ACCENT_BLUE = RGBColor(56, 189, 248)   # #38BDF8
ACCENT_GREEN = RGBColor(52, 211, 153)  # #34D399
ACCENT_AMBER = RGBColor(251, 191, 36)  # #FBBF24

blank_layout = prs.slide_layouts[6] # blank slide layout

def add_background(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_COLOR
    bg.line.fill.background()
    return bg

def add_header(slide, badge_text, slide_num, total_slides=17):
    # Top banner shape
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.5))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p = tf.paragraphs[0]
    r1 = p.add_run()
    r1.text = "TUTORMETRIC · SC452  |  "
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = ACCENT_BLUE
    
    r2 = p.add_run()
    r2.text = f"{badge_text.upper()}"
    r2.font.size = Pt(11)
    r2.font.bold = True
    r2.font.color.rgb = TEXT_WHITE
    
    # Right-aligned slide number
    num_box = slide.shapes.add_textbox(Inches(11.0), Inches(0.4), Inches(1.5), Inches(0.5))
    tf_num = num_box.text_frame
    tf_num.word_wrap = False
    tf_num.margin_left = tf_num.margin_top = tf_num.margin_right = tf_num.margin_bottom = 0
    p_num = tf_num.paragraphs[0]
    p_num.alignment = PP_ALIGN.RIGHT
    r_num = p_num.add_run()
    r_num.text = f"{slide_num:02d} / {total_slides:02d}"
    r_num.font.size = Pt(11)
    r_num.font.bold = True
    r_num.font.color.rgb = ACCENT_BLUE

def add_title_block(slide, title, subtitle):
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.95), Inches(11.733), Inches(0.9))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p1 = tf.paragraphs[0]
    r1 = p1.add_run()
    r1.text = title
    r1.font.size = Pt(22)
    r1.font.bold = True
    r1.font.color.rgb = TEXT_WHITE
    
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = subtitle
    r2.font.size = Pt(12)
    r2.font.color.rgb = TEXT_MUTED

def set_speaker_notes(slide, slide_data):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    notes_text = f"""[SLIDE {slide_data['id']:02d}: {slide_data['badge'].upper()} - {slide_data['title'].upper()}]
Target Pacing: {slide_data['target_time']}

Joint Spoken Script (Nilesh Mori & Pratik Chauhan):
"{slide_data['notes']}"

Anticipated Examiner Viva Defense:
Q: {slide_data['viva_q']}
A: {slide_data['viva_a']}
"""
    tf.text = notes_text.strip()

# BUILD SLIDES
print("Building 17 non-overlapping slides...")

for s in SLIDES_DATA:
    slide = prs.slides.add_slide(blank_layout)
    add_background(slide)
    set_speaker_notes(slide, s)
    
    # 1. Slide 1: Title Slide
    if s['type'] == 'title':
        # Large centered title container
        tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        r0 = p0.add_run()
        r0.text = "DA-IICT SC452  ·  B.TECH MINI PROJECT\n"
        r0.font.size = Pt(13)
        r0.font.bold = True
        r0.font.color.rgb = ACCENT_BLUE
        
        p1 = tf.add_paragraph()
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = s['title'] + "\n"
        r1.font.size = Pt(44)
        r1.font.bold = True
        r1.font.color.rgb = TEXT_WHITE
        
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = s['subtitle'] + "\n\n"
        r2.font.size = Pt(16)
        r2.font.color.rgb = TEXT_MUTED
        
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.add_run()
        r3.text = s['content']['summary'] + "\n\n"
        r3.font.size = Pt(12)
        r3.font.color.rgb = RGBColor(203, 213, 225)
        
        p4 = tf.add_paragraph()
        p4.alignment = PP_ALIGN.CENTER
        r4 = p4.add_run()
        r4.text = f"JOINT EQUAL AUTHORS: Nilesh Mori (202301473) & Pratik Chauhan (202301435)\n"
        r4.font.size = Pt(13)
        r4.font.bold = True
        r4.font.color.rgb = ACCENT_GREEN
        
        r5 = p4.add_run()
        r5.text = f"Mentor: {s['content']['mentor']}  ·  {s['content']['institution']}  ·  Live: {s['content']['live_url']}"
        r5.font.size = Pt(11)
        r5.font.color.rgb = TEXT_MUTED
        continue

    # Standard header for slides 2-17
    add_header(slide, s['badge'], s['id'])
    add_title_block(slide, s['title'], s['subtitle'])
    
    # 2. Slide 2: Cards Grid (Motivation)
    if s['type'] == 'cards_grid':
        coords = [
            (Inches(0.8), Inches(2.0), Inches(5.7), Inches(2.3)),
            (Inches(6.8), Inches(2.0), Inches(5.7), Inches(2.3)),
            (Inches(0.8), Inches(4.5), Inches(5.7), Inches(2.3)),
            (Inches(6.8), Inches(4.5), Inches(5.7), Inches(2.3)),
        ]
        for c, (left, top, width, height) in zip(s['cards'], coords):
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            box.fill.solid()
            box.fill.fore_color.rgb = CARD_BG
            box.line.color.rgb = RGBColor(30, 58, 102)
            
            tf = box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.2)
            p = tf.paragraphs[0]
            r_title = p.add_run()
            r_title.text = f"{c['title']}  [{c['badge']}]\n"
            r_title.font.size = Pt(13)
            r_title.font.bold = True
            r_title.font.color.rgb = ACCENT_BLUE
            
            p2 = tf.add_paragraph()
            r_desc = p2.add_run()
            r_desc.text = c['text']
            r_desc.font.size = Pt(11)
            r_desc.font.color.rgb = RGBColor(226, 232, 240)

    # 3. Slide 3: Comparison (Value & Scope)
    elif s['type'] == 'comparison':
        # Left card (Legacy)
        box_left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.7), Inches(4.8))
        box_left.fill.solid()
        box_left.fill.fore_color.rgb = RGBColor(38, 20, 30)
        box_left.line.color.rgb = RGBColor(159, 18, 57)
        tf_l = box_left.text_frame
        tf_l.word_wrap = True
        tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = Inches(0.25)
        p_l = tf_l.paragraphs[0]
        r_lt = p_l.add_run()
        r_lt.text = f"✕ {s['legacy']['title']}\n\n"
        r_lt.font.size = Pt(15)
        r_lt.font.bold = True
        r_lt.font.color.rgb = RGBColor(251, 113, 133)
        for it in s['legacy']['items']:
            p_item = tf_l.add_paragraph()
            r_it = p_item.add_run()
            r_it.text = f"•  {it}\n"
            r_it.font.size = Pt(12)
            r_it.font.color.rgb = RGBColor(226, 232, 240)
            
        # Right card (TutorMetric)
        box_right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(4.8))
        box_right.fill.solid()
        box_right.fill.fore_color.rgb = RGBColor(16, 44, 40)
        box_right.line.color.rgb = ACCENT_GREEN
        tf_r = box_right.text_frame
        tf_r.word_wrap = True
        tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = Inches(0.25)
        p_r = tf_r.paragraphs[0]
        r_rt = p_r.add_run()
        r_rt.text = f"✓ {s['tutormetric']['title']}\n\n"
        r_rt.font.size = Pt(15)
        r_rt.font.bold = True
        r_rt.font.color.rgb = ACCENT_GREEN
        for it in s['tutormetric']['items']:
            p_item = tf_r.add_paragraph()
            r_it = p_item.add_run()
            r_it.text = f"✓  {it}\n"
            r_it.font.size = Pt(12)
            r_it.font.color.rgb = TEXT_WHITE

    # 4. Slide 4: Roadmap (Project Tracker)
    elif s['type'] == 'roadmap':
        grid = [
            (Inches(0.8), Inches(2.0), Inches(3.7), Inches(2.3)),
            (Inches(4.8), Inches(2.0), Inches(3.7), Inches(2.3)),
            (Inches(8.8), Inches(2.0), Inches(3.7), Inches(2.3)),
            (Inches(0.8), Inches(4.5), Inches(3.7), Inches(2.3)),
            (Inches(4.8), Inches(4.5), Inches(3.7), Inches(2.3)),
            (Inches(8.8), Inches(4.5), Inches(3.7), Inches(2.3)),
        ]
        for st, (left, top, width, height) in zip(s['steps'], grid):
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            box.fill.solid()
            box.fill.fore_color.rgb = CARD_BG
            box.line.color.rgb = RGBColor(30, 58, 102)
            tf = box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.18)
            p = tf.paragraphs[0]
            r_num = p.add_run()
            r_num.text = f"{st['num']}  ·  {st['tasks']}\n"
            r_num.font.size = Pt(11)
            r_num.font.bold = True
            r_num.font.color.rgb = ACCENT_BLUE
            
            p_ph = tf.add_paragraph()
            r_ph = p_ph.add_run()
            r_ph.text = f"{st['phase']}\n"
            r_ph.font.size = Pt(12)
            r_ph.font.bold = True
            r_ph.font.color.rgb = TEXT_WHITE
            
            p_desc = tf.add_paragraph()
            r_desc = p_desc.add_run()
            r_desc.text = st['desc']
            r_desc.font.size = Pt(10)
            r_desc.font.color.rgb = TEXT_MUTED

    # 5. Slide 5: Data Pipeline Detail
    elif s['type'] == 'pipeline_detail':
        # 4 Metric pills across top
        metric_width = Inches(2.8)
        for i, m in enumerate(s['metrics']):
            m_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.98), Inches(2.0), metric_width, Inches(1.0))
            m_box.fill.solid()
            m_box.fill.fore_color.rgb = CARD_BG
            m_box.line.color.rgb = RGBColor(30, 58, 102)
            tf = m_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r_v = p.add_run()
            r_v.text = m['val'] + "\n"
            r_v.font.size = Pt(16)
            r_v.font.bold = True
            r_v.font.color.rgb = ACCENT_BLUE
            r_l = p.add_run()
            r_l.text = m['lbl']
            r_l.font.size = Pt(10)
            r_l.font.color.rgb = TEXT_MUTED

        # 4 Protocol blocks in 2x2 grid below
        proto_coords = [
            (Inches(0.8), Inches(3.2), Inches(5.7), Inches(1.8)),
            (Inches(6.8), Inches(3.2), Inches(5.7), Inches(1.8)),
            (Inches(0.8), Inches(5.1), Inches(5.7), Inches(1.8)),
            (Inches(6.8), Inches(5.1), Inches(5.7), Inches(1.8)),
        ]
        for p_item, (left, top, width, height) in zip(s['protocol'], proto_coords):
            p_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            p_box.fill.solid()
            p_box.fill.fore_color.rgb = RGBColor(16, 26, 46)
            p_box.line.color.rgb = RGBColor(30, 58, 102)
            tf = p_box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.18)
            p = tf.paragraphs[0]
            r_st = p.add_run()
            r_st.text = f"◈ {p_item['step']}\n"
            r_st.font.size = Pt(12)
            r_st.font.bold = True
            r_st.font.color.rgb = ACCENT_GREEN
            p2 = tf.add_paragraph()
            r_d = p2.add_run()
            r_d.text = p_item['desc']
            r_d.font.size = Pt(10)
            r_d.font.color.rgb = RGBColor(226, 232, 240)

    # 6. Slide 6: Database Schema (3NF)
    elif s['type'] == 'database_schema':
        # 2 Table definition boxes
        box_t1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.7), Inches(2.8))
        box_t1.fill.solid()
        box_t1.fill.fore_color.rgb = CARD_BG
        box_t1.line.color.rgb = ACCENT_BLUE
        tf1 = box_t1.text_frame
        tf1.word_wrap = True
        tf1.margin_left = tf1.margin_right = tf1.margin_top = Inches(0.2)
        p1 = tf1.paragraphs[0]
        r1 = p1.add_run()
        r1.text = f"TABLE: teachers  [{s['tables'][0]['rows']}]\n"
        r1.font.size = Pt(13)
        r1.font.bold = True
        r1.font.color.rgb = ACCENT_BLUE
        p1_desc = tf1.add_paragraph()
        r1_d = p1_desc.add_run()
        r1_d.text = f"{s['tables'][0]['desc']}\n\nAttributes:\n"
        r1_d.font.size = Pt(10)
        r1_d.font.color.rgb = TEXT_MUTED
        p1_cols = tf1.add_paragraph()
        r1_c = p1_cols.add_run()
        r1_c.text = " · ".join(s['tables'][0]['columns'])
        r1_c.font.size = Pt(9.5)
        r1_c.font.color.rgb = RGBColor(226, 232, 240)

        box_t2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(2.8))
        box_t2.fill.solid()
        box_t2.fill.fore_color.rgb = CARD_BG
        box_t2.line.color.rgb = ACCENT_GREEN
        tf2 = box_t2.text_frame
        tf2.word_wrap = True
        tf2.margin_left = tf2.margin_right = tf2.margin_top = Inches(0.2)
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = f"TABLE: teacher_subjects  [{s['tables'][1]['rows']}]\n"
        r2.font.size = Pt(13)
        r2.font.bold = True
        r2.font.color.rgb = ACCENT_GREEN
        p2_desc = tf2.add_paragraph()
        r2_d = p2_desc.add_run()
        r2_d.text = f"{s['tables'][1]['desc']}\n\nAttributes:\n"
        r2_d.font.size = Pt(10)
        r2_d.font.color.rgb = TEXT_MUTED
        p2_cols = tf2.add_paragraph()
        r2_c = p2_cols.add_run()
        r2_c.text = " · ".join(s['tables'][1]['columns'])
        r2_c.font.size = Pt(9.5)
        r2_c.font.color.rgb = RGBColor(226, 232, 240)

        # 3 Advantage metrics below
        for i, adv in enumerate(s['advantages']):
            adv_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 4.0), Inches(5.0), Inches(3.7), Inches(1.8))
            adv_box.fill.solid()
            adv_box.fill.fore_color.rgb = RGBColor(16, 26, 46)
            adv_box.line.color.rgb = RGBColor(30, 58, 102)
            tf_a = adv_box.text_frame
            tf_a.word_wrap = True
            p = tf_a.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r_st = p.add_run()
            r_st.text = adv['stat'] + "\n"
            r_st.font.size = Pt(18)
            r_st.font.bold = True
            r_st.font.color.rgb = ACCENT_GREEN
            r_lbl = p.add_run()
            r_lbl.text = adv['label'] + "\n"
            r_lbl.font.size = Pt(11)
            r_lbl.font.bold = True
            r_lbl.font.color.rgb = TEXT_WHITE
            r_dc = p.add_run()
            r_dc.text = adv['desc']
            r_dc.font.size = Pt(9.5)
            r_dc.font.color.rgb = TEXT_MUTED

    # 7. Slide 7: Taxonomy Cleaning
    elif s['type'] == 'taxonomy_cleaning':
        # 4 Quick stats
        for i, st in enumerate(s['stats']):
            stat_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.98), Inches(2.0), Inches(2.8), Inches(1.0))
            stat_box.fill.solid()
            stat_box.fill.fore_color.rgb = CARD_BG
            stat_box.line.color.rgb = RGBColor(30, 58, 102)
            tf = stat_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r_v = p.add_run()
            r_v.text = st['num'] + "\n"
            r_v.font.size = Pt(16)
            r_v.font.bold = True
            r_v.font.color.rgb = ACCENT_BLUE
            r_l = p.add_run()
            r_l.text = st['label']
            r_l.font.size = Pt(10)
            r_l.font.color.rgb = TEXT_MUTED

        # 6 Category cards in 2x3 grid
        cat_grid = [
            (Inches(0.8), Inches(3.2), Inches(3.7), Inches(1.7)),
            (Inches(4.8), Inches(3.2), Inches(3.7), Inches(1.7)),
            (Inches(8.8), Inches(3.2), Inches(3.7), Inches(1.7)),
            (Inches(0.8), Inches(5.1), Inches(3.7), Inches(1.7)),
            (Inches(4.8), Inches(5.1), Inches(3.7), Inches(1.7)),
            (Inches(8.8), Inches(5.1), Inches(3.7), Inches(1.7)),
        ]
        for cat, (left, top, width, height) in zip(s['categories'], cat_grid):
            c_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            c_box.fill.solid()
            c_box.fill.fore_color.rgb = RGBColor(16, 26, 46)
            c_box.line.color.rgb = RGBColor(30, 58, 102)
            tf = c_box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.15)
            p = tf.paragraphs[0]
            r_cn = p.add_run()
            r_cn.text = f"{cat['name']}  ({cat['count']})\n"
            r_cn.font.size = Pt(11)
            r_cn.font.bold = True
            r_cn.font.color.rgb = ACCENT_GREEN
            p2 = tf.add_paragraph()
            r_ex = p2.add_run()
            r_ex.text = f"Canonicals: {cat['examples']}"
            r_ex.font.size = Pt(10)
            r_ex.font.color.rgb = RGBColor(203, 213, 225)

    # 8-12. Split Layout Slides: Image on Left + Structured Data on Right
    elif s['type'] in ['geography_analysis', 'tier_breakdown', 'ladder_chart', 'venn_analysis', 'demographics_grid', 'feature_spotlight']:
        img_path = os.path.join('/home/nilesh7757/BMPTUTORS1', s['image'])
        
        # Place image on left (Inches(0.8), Inches(2.0), width 5.2, height up to 4.8)
        if os.path.exists(img_path):
            try:
                slide.shapes.add_picture(img_path, Inches(0.8), Inches(2.0), width=Inches(5.4))
            except Exception as e:
                print(f"Error adding picture {img_path}: {e}")

        # Place structured right panel (Inches(6.6), Inches(2.0), width 5.9, height 4.8)
        panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.6), Inches(2.0), Inches(5.9), Inches(4.8))
        panel.fill.solid()
        panel.fill.fore_color.rgb = CARD_BG
        panel.line.color.rgb = RGBColor(30, 58, 102)
        tf = panel.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.25)
        
        # Populate right panel depending on type
        if s['type'] == 'geography_analysis':
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = "TOP 5 STATES (54.4% OF ALL TUTORS):\n\n"
            r0.font.size = Pt(13)
            r0.font.bold = True
            r0.font.color.rgb = ACCENT_BLUE
            for st in s['top_states']:
                p_st = tf.add_paragraph()
                r_st = p_st.add_run()
                r_st.text = f"{st['rank']}. {st['state']}  ·  {st['count']} ({st['pct']})  ·  Avg: {st['fee']}\n"
                r_st.font.size = Pt(11)
                r_st.font.color.rgb = TEXT_WHITE
            p_tk = tf.add_paragraph()
            r_tk = p_tk.add_run()
            r_tk.text = f"\nKey Finding: {s['takeaway']}"
            r_tk.font.size = Pt(10)
            r_tk.font.color.rgb = TEXT_MUTED

        elif s['type'] == 'tier_breakdown':
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = "CITY TIER DISTRIBUTION (N = 107,673):\n\n"
            r0.font.size = Pt(13)
            r0.font.bold = True
            r0.font.color.rgb = ACCENT_BLUE
            for t in s['tiers']:
                p_t = tf.add_paragraph()
                r_t = p_t.add_run()
                r_t.text = f"• {t['tier']}: {t['share']} ({t['count']})  ·  Avg: {t['avg_fee']}\n"
                r_t.font.size = Pt(11)
                r_t.font.bold = True
                r_t.font.color.rgb = TEXT_WHITE
                r_td = p_t.add_run()
                r_td.text = f"  {t['desc']}\n"
                r_td.font.size = Pt(9.5)
                r_td.font.color.rgb = TEXT_MUTED
            p_in = tf.add_paragraph()
            r_in = p_in.add_run()
            r_in.text = f"\nTakeaway: {s['insight']}"
            r_in.font.size = Pt(10)
            r_in.font.color.rgb = ACCENT_AMBER

        elif s['type'] == 'ladder_chart':
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = "MONOTONIC LEVEL PRICING LADDER:\n\n"
            r0.font.size = Pt(13)
            r0.font.bold = True
            r0.font.color.rgb = ACCENT_BLUE
            for l in s['levels']:
                p_l = tf.add_paragraph()
                r_l = p_l.add_run()
                r_l.text = f"• {l['level']}:  {l['fee']}  ({l['index']})\n"
                r_l.font.size = Pt(10.5)
                r_l.font.color.rgb = TEXT_WHITE
            p_f = tf.add_paragraph()
            r_f = p_f.add_run()
            r_f.text = "\n" + "\n".join(f"◈ {fact}" for fact in s['distribution_facts'])
            r_f.font.size = Pt(9.5)
            r_f.font.color.rgb = TEXT_MUTED

        elif s['type'] == 'venn_analysis':
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = "STEM TRIO INTERSECTION (MATHS, PHYSICS, CHEM):\n\n"
            r0.font.size = Pt(13)
            r0.font.bold = True
            r0.font.color.rgb = ACCENT_BLUE
            for v in s['venn_stats']:
                p_v = tf.add_paragraph()
                r_v = p_v.add_run()
                r_v.text = f"• {v['lbl']}:  {v['val']}\n"
                r_v.font.size = Pt(11)
                r_v.font.bold = True
                r_v.font.color.rgb = ACCENT_GREEN
                r_vd = p_v.add_run()
                r_vd.text = f"  {v['desc']}\n"
                r_vd.font.size = Pt(9.5)
                r_vd.font.color.rgb = TEXT_MUTED
            p_tk = tf.add_paragraph()
            r_tk = p_tk.add_run()
            r_tk.text = f"\nSet Theory Insight: {s['takeaway']}"
            r_tk.font.size = Pt(10)
            r_tk.font.color.rgb = RGBColor(226, 232, 240)

        elif s['type'] == 'demographics_grid':
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = "GENDER RATIO BY SUBJECT & DEGREE PREMIUM:\n\n"
            r0.font.size = Pt(13)
            r0.font.bold = True
            r0.font.color.rgb = ACCENT_BLUE
            for g in s['gender_breakdown']:
                p_g = tf.add_paragraph()
                r_g = p_g.add_run()
                r_g.text = f"• {g['subject']}: {g['female_pct']}  ·  {g['male_pct']}\n"
                r_g.font.size = Pt(10)
                r_g.font.color.rgb = TEXT_WHITE
            p_dp = tf.add_paragraph()
            r_dp = p_dp.add_run()
            r_dp.text = "\nAcademic Degree & Gig Dynamics:\n"
            r_dp.font.size = Pt(11)
            r_dp.font.bold = True
            r_dp.font.color.rgb = ACCENT_GREEN
            for d in s['degree_premium']:
                p_d = tf.add_paragraph()
                r_d = p_d.add_run()
                r_d.text = f"◈ {d['tier']}: {d['fee']} ({d['gap']})\n"
                r_d.font.size = Pt(9.5)
                r_d.font.color.rgb = TEXT_MUTED

        elif s['type'] == 'feature_spotlight':
            p0 = tf.paragraphs[0]
            r0 = p0.add_run()
            r0.text = "NINE REACTIVE SEARCH & FILTER UTILITIES:\n\n"
            r0.font.size = Pt(13)
            r0.font.bold = True
            r0.font.color.rgb = ACCENT_BLUE
            for f in s['filters']:
                p_f = tf.add_paragraph()
                r_f = p_f.add_run()
                r_f.text = f"✓ {f['name']}: {f['desc']}\n"
                r_f.font.size = Pt(10)
                r_f.font.color.rgb = TEXT_WHITE

    # 13. Slide 13: Econometric Model (Ridge)
    elif s['type'] == 'model_metrics':
        # 4 Metric cards
        for i, m in enumerate(s['model_stats']):
            m_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.98), Inches(2.0), Inches(2.8), Inches(1.1))
            m_box.fill.solid()
            m_box.fill.fore_color.rgb = CARD_BG
            m_box.line.color.rgb = RGBColor(30, 58, 102)
            tf = m_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r_st = p.add_run()
            r_st.text = m['stat'] + "\n"
            r_st.font.size = Pt(18)
            r_st.font.bold = True
            r_st.font.color.rgb = ACCENT_BLUE
            r_lb = p.add_run()
            r_lb.text = m['label'] + "\n"
            r_lb.font.size = Pt(10.5)
            r_lb.font.bold = True
            r_lb.font.color.rgb = TEXT_WHITE
            r_sb = p.add_run()
            r_sb.text = m['sub']
            r_sb.font.size = Pt(8.5)
            r_sb.font.color.rgb = TEXT_MUTED

        # 2 Comparison cards: Observable vs Unobservable
        obs_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(3.3), Inches(5.7), Inches(3.5))
        obs_box.fill.solid()
        obs_box.fill.fore_color.rgb = RGBColor(16, 36, 32)
        obs_box.line.color.rgb = ACCENT_GREEN
        tf_o = obs_box.text_frame
        tf_o.word_wrap = True
        tf_o.margin_left = tf_o.margin_right = tf_o.margin_top = Inches(0.2)
        p_o = tf_o.paragraphs[0]
        r_ot = p_o.add_run()
        r_ot.text = f"✓ {s['econometric_comparison']['observable']['title']}\n\n"
        r_ot.font.size = Pt(12)
        r_ot.font.bold = True
        r_ot.font.color.rgb = ACCENT_GREEN
        for it in s['econometric_comparison']['observable']['items']:
            p_oi = tf_o.add_paragraph()
            r_oi = p_oi.add_run()
            r_oi.text = f"•  {it}\n"
            r_oi.font.size = Pt(10.5)
            r_oi.font.color.rgb = RGBColor(226, 232, 240)

        unobs_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(3.3), Inches(5.7), Inches(3.5))
        unobs_box.fill.solid()
        unobs_box.fill.fore_color.rgb = RGBColor(40, 28, 16)
        unobs_box.line.color.rgb = ACCENT_AMBER
        tf_u = unobs_box.text_frame
        tf_u.word_wrap = True
        tf_u.margin_left = tf_u.margin_right = tf_u.margin_top = Inches(0.2)
        p_u = tf_u.paragraphs[0]
        r_ut = p_u.add_run()
        r_ut.text = f"◈ {s['econometric_comparison']['unobservable']['title']}\n\n"
        r_ut.font.size = Pt(12)
        r_ut.font.bold = True
        r_ut.font.color.rgb = ACCENT_AMBER
        for it in s['econometric_comparison']['unobservable']['items']:
            p_ui = tf_u.add_paragraph()
            r_ui = p_ui.add_run()
            r_ui.text = f"•  {it}\n"
            r_ui.font.size = Pt(10.5)
            r_ui.font.color.rgb = RGBColor(226, 232, 240)

    # 15. Slide 15: Two Interactive Tools
    elif s['type'] == 'two_tools':
        t1_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.7), Inches(4.8))
        t1_box.fill.solid()
        t1_box.fill.fore_color.rgb = CARD_BG
        t1_box.line.color.rgb = ACCENT_BLUE
        tf1 = t1_box.text_frame
        tf1.word_wrap = True
        tf1.margin_left = tf1.margin_right = tf1.margin_top = Inches(0.25)
        p1 = tf1.paragraphs[0]
        r1 = p1.add_run()
        r1.text = f"⚡ {s['tool_a']['name']}\n\n"
        r1.font.size = Pt(15)
        r1.font.bold = True
        r1.font.color.rgb = ACCENT_BLUE
        p1_desc = tf1.add_paragraph()
        r1_d = p1_desc.add_run()
        r1_d.text = f"{s['tool_a']['desc']}\n\nFeatures:\n"
        r1_d.font.size = Pt(11)
        r1_d.font.color.rgb = TEXT_MUTED
        for f in s['tool_a']['features']:
            p_f = tf1.add_paragraph()
            r_f = p_f.add_run()
            r_f.text = f"✓  {f}\n"
            r_f.font.size = Pt(11)
            r_f.font.color.rgb = RGBColor(226, 232, 240)

        t2_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(4.8))
        t2_box.fill.solid()
        t2_box.fill.fore_color.rgb = CARD_BG
        t2_box.line.color.rgb = ACCENT_GREEN
        tf2 = t2_box.text_frame
        tf2.word_wrap = True
        tf2.margin_left = tf2.margin_right = tf2.margin_top = Inches(0.25)
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = f"✨ {s['tool_b']['name']}\n\n"
        r2.font.size = Pt(15)
        r2.font.bold = True
        r2.font.color.rgb = ACCENT_GREEN
        p2_desc = tf2.add_paragraph()
        r2_d = p2_desc.add_run()
        r2_d.text = f"{s['tool_b']['desc']}\n\nFeatures:\n"
        r2_d.font.size = Pt(11)
        r2_d.font.color.rgb = TEXT_MUTED
        for f in s['tool_b']['features']:
            p_f = tf2.add_paragraph()
            r_f = p_f.add_run()
            r_f.text = f"✓  {f}\n"
            r_f.font.size = Pt(11)
            r_f.font.color.rgb = RGBColor(226, 232, 240)

    # 16. Slide 16: Findings Grid & Future Work
    elif s['type'] == 'findings_grid':
        for i, f in enumerate(s['findings']):
            f_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.38), Inches(2.0), Inches(2.25), Inches(3.5))
            f_box.fill.solid()
            f_box.fill.fore_color.rgb = CARD_BG
            f_box.line.color.rgb = RGBColor(30, 58, 102)
            tf = f_box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.18)
            p = tf.paragraphs[0]
            r_fn = p.add_run()
            r_fn.text = f"{f['num']}\n"
            r_fn.font.size = Pt(14)
            r_fn.font.bold = True
            r_fn.font.color.rgb = ACCENT_BLUE
            r_ft = p.add_run()
            r_ft.text = f"{f['title']}\n\n"
            r_ft.font.size = Pt(11.5)
            r_ft.font.bold = True
            r_ft.font.color.rgb = TEXT_WHITE
            r_fd = p.add_run()
            r_fd.text = f['desc']
            r_fd.font.size = Pt(10)
            r_fd.font.color.rgb = TEXT_MUTED

        # Future work banner below
        fw_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.7), Inches(11.733), Inches(1.1))
        fw_box.fill.solid()
        fw_box.fill.fore_color.rgb = RGBColor(16, 26, 46)
        fw_box.line.color.rgb = RGBColor(30, 58, 102)
        tf_fw = fw_box.text_frame
        tf_fw.word_wrap = True
        tf_fw.margin_left = tf_fw.margin_right = tf_fw.margin_top = Inches(0.15)
        p_fw = tf_fw.paragraphs[0]
        r_fwt = p_fw.add_run()
        r_fwt.text = "FUTURE RESEARCH ROADMAP:  "
        r_fwt.font.size = Pt(11)
        r_fwt.font.bold = True
        r_fwt.font.color.rgb = ACCENT_GREEN
        r_fwc = p_fw.add_run()
        r_fwc.text = " · ".join(s['future_work'])
        r_fwc.font.size = Pt(10.5)
        r_fwc.font.color.rgb = RGBColor(203, 213, 225)

    # 17. Slide 17: Conclusion & Demo
    elif s['type'] == 'conclusion_demo':
        left_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.7), Inches(4.8))
        left_box.fill.solid()
        left_box.fill.fore_color.rgb = CARD_BG
        left_box.line.color.rgb = ACCENT_BLUE
        tf_l = left_box.text_frame
        tf_l.word_wrap = True
        tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = Inches(0.25)
        p_l = tf_l.paragraphs[0]
        r_lt = p_l.add_run()
        r_lt.text = "PROJECT CREDENTIALS\n\n"
        r_lt.font.size = Pt(13)
        r_lt.font.bold = True
        r_lt.font.color.rgb = ACCENT_BLUE
        
        p_cr1 = tf_l.add_paragraph()
        r_cr1 = p_cr1.add_run()
        r_cr1.text = f"Title: {s['credits']['title']}\n\n"
        r_cr1.font.size = Pt(12)
        r_cr1.font.bold = True
        r_cr1.font.color.rgb = TEXT_WHITE
        
        p_cr2 = tf_l.add_paragraph()
        r_cr2 = p_cr2.add_run()
        r_cr2.text = f"Authors: {s['credits']['authors']}\n\nMentor: {s['credits']['mentor']}\n\nLive Deployment: {s['credits']['url']}"
        r_cr2.font.size = Pt(11)
        r_cr2.font.color.rgb = RGBColor(226, 232, 240)

        right_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.0), Inches(5.7), Inches(4.8))
        right_box.fill.solid()
        right_box.fill.fore_color.rgb = CARD_BG
        right_box.line.color.rgb = ACCENT_GREEN
        tf_r = right_box.text_frame
        tf_r.word_wrap = True
        tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = Inches(0.25)
        p_r = tf_r.paragraphs[0]
        r_rt = p_r.add_run()
        r_rt.text = "LIVE SYSTEM DEMONSTRATION (5 MINS)\n\n"
        r_rt.font.size = Pt(13)
        r_rt.font.bold = True
        r_rt.font.color.rgb = ACCENT_GREEN
        for d in s['demo_checklist']:
            p_d = tf_r.add_paragraph()
            r_d = p_d.add_run()
            r_d.text = f"✓ {d['step']}: {d['action']}\n\n"
            r_d.font.size = Pt(11)
            r_d.font.color.rgb = TEXT_WHITE

# Save generated master PPTX
output_paths = [
    '/home/nilesh7757/Downloads/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/Downloads/BMP_Tutors_Deck_Full.pptx',
    '/home/nilesh7757/BMPTUTORS1/project_showcase_bundle/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/BMPTUTORS1/project_showcase_bundle/BMP_Tutors_Deck_Full.pptx',
    '/home/nilesh7757/BMPTUTORS1/TutorMetric_Master_Presentation.pptx'
]

for out_path in output_paths:
    prs.save(out_path)
    print(f"Saved: {out_path} ({os.path.getsize(out_path):,} bytes)")

print("Built 17 non-overlapping slides successfully in PowerPoint!")
