#!/usr/bin/env python3
"""
Generate Comprehensive Visual Analytics DOCX Report for BMP Tutors.
Extracts all charts, maps, figures, and tables into an editable Microsoft Word document.
"""

import os
import sys
import sqlite3
sys.path.append('/home/nilesh7757/BMPTUTORS1/app')
from geo_service import get_enriched_states_geojson, get_enriched_districts_geojson
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import re
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DB_PATH = "/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db"
OUTPUT_DIR = "/home/nilesh7757/BMPTUTORS1/extracted_images/charts"
DOCX_PATH = "/home/nilesh7757/BMPTUTORS1/BMP_Tutors_Visual_Analytics_Report.docx"
REF_IMG_DIR = "/home/nilesh7757/BMPTUTORS1/extracted_images"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------------------
# XML Helper Functions for DOCX Table Formatting
# -------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="D3D3D3", sz="4", val="single"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), val)
            border.set(qn('w:sz'), sz)
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), color)
            borders.append(border)
        tblPr[0].append(borders)

# -------------------------------------------------------------
# 1. Matplotlib High-Resolution Chart Generation (300 DPI)
# -------------------------------------------------------------
print("Connecting to database and generating 300 DPI charts...")
conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

# Colors
CLR_PRIMARY = "#4f46e5"
CLR_DARK = "#1e293b"
CLR_GRAY = "#475569"
CLR_LIGHT_GRAY = "#94a3b8"
CLR_EMERALD = "#10b981"
CLR_ROSE = "#e11d48"
CLR_BLUE = "#2563eb"
CLR_AMBER = "#f59e0b"
CLR_PURPLE = "#9333ea"
CLR_TEAL = "#0d9488"

# 1. Figure 1: Macro Subjects (Moscow Format)
macro_rows = cur.execute("""
    SELECT 
        CASE 
            WHEN subject_category = 'Science' THEN 'School Curriculum & STEM'
            WHEN subject_category = 'Languages' THEN 'Languages & Linguistics'
            WHEN subject_category = 'Coding' THEN 'Computer Science & Coding'
            WHEN subject_category = 'Commerce' THEN 'Commerce & Management'
            WHEN subject_category = 'Arts' THEN 'Arts & Humanities'
            ELSE 'Personal Dev. & ECA'
        END as name,
        count(*) as count,
        round(count(*) * 100.0 / 403353, 1) as pct
    FROM teacher_subjects
    GROUP BY name
    ORDER BY count DESC
""").fetchall()

fig, ax = plt.subplots(figsize=(9, 4.8), dpi=300)
cats = [r['name'] for r in macro_rows]
counts = [r['count'] for r in macro_rows]
pcts = [r['pct'] for r in macro_rows]

bars = ax.bar(cats, counts, color=CLR_GRAY, edgecolor="#334155", linewidth=1.2, width=0.55)
ax.set_title("Figure 1: Subjects Offered for Private Tuition in India (N = 403,353)", fontsize=13, fontweight='bold', pad=15, color="#0f172a")
ax.set_ylabel("Number of Offered Courses", fontsize=10, fontweight='semibold', color="#334155")
ax.grid(axis='y', linestyle='--', alpha=0.5)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.xticks(rotation=15, ha='right', fontsize=9.5, fontweight='medium')

for bar, cnt, pct in zip(bars, counts, pcts):
    height = bar.get_height()
    ax.annotate(f"{cnt:,}\n({pct}%)",
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 4), textcoords="offset points",
                ha='center', va='bottom', fontsize=8.5, fontweight='bold', color="#1e293b")

ax.set_ylim(0, max(counts) * 1.15)
plt.tight_layout()
fig1_macro_path = os.path.join(OUTPUT_DIR, "fig1_subjects_macro.png")
plt.savefig(fig1_macro_path)
plt.close()

# 2. Figure 1: Top 12 Specific Subjects
top_subjs = cur.execute("""
    SELECT canonical_subject, count(*) as count, round(count(*) * 100.0 / 403353, 1) as pct
    FROM teacher_subjects
    GROUP BY canonical_subject
    ORDER BY count DESC
    LIMIT 12
""").fetchall()

fig, ax = plt.subplots(figsize=(9, 5.2), dpi=300)
sub_names = [r['canonical_subject'] for r in top_subjs][::-1]
sub_counts = [r['count'] for r in top_subjs][::-1]
sub_pcts = [r['pct'] for r in top_subjs][::-1]

bars = ax.barh(sub_names, sub_counts, color=CLR_PRIMARY, edgecolor="#3730a3", linewidth=1, height=0.65)
ax.set_title("Top 12 Most Frequently Offered Tutoring Subjects in India", fontsize=12.5, fontweight='bold', pad=15, color="#0f172a")
ax.set_xlabel("Number of Registered Tutor Offerings", fontsize=10, fontweight='semibold', color="#334155")
ax.grid(axis='x', linestyle='--', alpha=0.5)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

for bar, cnt, pct in zip(bars, sub_counts, sub_pcts):
    width = bar.get_width()
    ax.annotate(f" {cnt:,} ({pct}%)",
                xy=(width, bar.get_y() + bar.get_height() / 2),
                xytext=(3, 0), textcoords="offset points",
                ha='left', va='center', fontsize=8.5, fontweight='bold', color="#1e293b")

ax.set_xlim(0, max(sub_counts) * 1.18)
plt.tight_layout()
fig1_top_path = os.path.join(OUTPUT_DIR, "fig1_top_subjects.png")
plt.savefig(fig1_top_path)
plt.close()

# 3. Figure 2: Experience Distribution Curve (Overall & Gender)
exp_rows = cur.execute("""
    SELECT 
        cast(round(total_teaching_exp_years) as integer) as exp,
        count(*) as total,
        sum(case when gender = 'Male' then 1 else 0 end) as male,
        sum(case when gender = 'Female' then 1 else 0 end) as female
    FROM teachers
    WHERE total_teaching_exp_years IS NOT NULL AND total_teaching_exp_years >= 0 AND total_teaching_exp_years <= 30
    GROUP BY exp
    ORDER BY exp ASC
""").fetchall()

exp_x = [r['exp'] for r in exp_rows]
exp_tot = [r['total'] for r in exp_rows]
exp_male = [r['male'] for r in exp_rows]
exp_fem = [r['female'] for r in exp_rows]

# 3A. Overall Experience Curve
fig, ax = plt.subplots(figsize=(9, 4.5), dpi=300)
ax.plot(exp_x, exp_tot, color=CLR_PRIMARY, linewidth=2.5, marker='o', markersize=4.5, label='All Tutors (N = 107,723)')
ax.fill_between(exp_x, exp_tot, color=CLR_PRIMARY, alpha=0.12)
ax.set_title("Figure 2: Experience Structure of Private Tutors (Years of Teaching Experience)", fontsize=12.5, fontweight='bold', pad=15, color="#0f172a")
ax.set_xlabel("Years of Teaching Experience", fontsize=10, fontweight='semibold', color="#334155")
ax.set_ylabel("Number of Educators", fontsize=10, fontweight='semibold', color="#334155")
ax.set_xticks(range(0, 31, 3))
ax.grid(True, linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Peak annotation
peak_idx = int(np.argmax(exp_tot))
ax.annotate(f'Peak Entry Mode:\n1-2 Yrs ({exp_tot[peak_idx]:,} tutors)',
            xy=(exp_x[peak_idx], exp_tot[peak_idx]),
            xytext=(exp_x[peak_idx] + 3, exp_tot[peak_idx] - 1500),
            arrowprops=dict(facecolor='#1e293b', shrink=0.08, width=1.5, headwidth=6),
            fontsize=9, fontweight='bold', color="#1e293b",
            bbox=dict(boxstyle="round,pad=0.4", fc="#f8fafc", ec="#cbd5e1", lw=1))

plt.tight_layout()
fig2_exp_path = os.path.join(OUTPUT_DIR, "fig2_experience_overall.png")
plt.savefig(fig2_exp_path)
plt.close()

# 3B. Male vs Female Experience Curves
fig, ax = plt.subplots(figsize=(9, 4.5), dpi=300)
ax.plot(exp_x, exp_fem, color=CLR_ROSE, linewidth=2.2, marker='s', markersize=4, label='Female Tutors (58.7%)')
ax.plot(exp_x, exp_male, color=CLR_BLUE, linewidth=2.2, marker='^', markersize=4, label='Male Tutors (41.3%)')
ax.fill_between(exp_x, exp_fem, color=CLR_ROSE, alpha=0.08)
ax.fill_between(exp_x, exp_male, color=CLR_BLUE, alpha=0.08)
ax.set_title("Teaching Experience Distribution by Gender", fontsize=12.5, fontweight='bold', pad=15, color="#0f172a")
ax.set_xlabel("Years of Teaching Experience", fontsize=10, fontweight='semibold', color="#334155")
ax.set_ylabel("Number of Educators", fontsize=10, fontweight='semibold', color="#334155")
ax.set_xticks(range(0, 31, 3))
ax.grid(True, linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.legend(frameon=True, facecolor='#ffffff', edgecolor='#cbd5e1', fontsize=9.5)
plt.tight_layout()
fig2_gender_path = os.path.join(OUTPUT_DIR, "fig2_experience_gender.png")
plt.savefig(fig2_gender_path)
plt.close()

# 4. Target Education Levels Chart
level_rows = cur.execute("""
    SELECT level_category, count(*) as count
    FROM teacher_subjects
    WHERE level_category IS NOT NULL AND level_category != 'Not Specified'
    GROUP BY level_category
    ORDER BY count ASC
""").fetchall()

fig, ax = plt.subplots(figsize=(9, 4.8), dpi=300)
lvl_names = [r['level_category'] for r in level_rows]
lvl_counts = [r['count'] for r in level_rows]

bars = ax.barh(lvl_names, lvl_counts, color=CLR_PURPLE, edgecolor="#6b21a8", linewidth=1, height=0.6)
ax.set_title("Target Educational Levels & Student Learning Stages (100% Classified)", fontsize=12, fontweight='bold', pad=15, color="#0f172a")
ax.set_xlabel("Number of Course Offerings", fontsize=10, fontweight='semibold', color="#334155")
ax.grid(axis='x', linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

for bar, cnt in zip(bars, lvl_counts):
    width = bar.get_width()
    ax.annotate(f" {cnt:,}",
                xy=(width, bar.get_y() + bar.get_height() / 2),
                xytext=(3, 0), textcoords="offset points",
                ha='left', va='center', fontsize=8.5, fontweight='bold', color="#1e293b")

ax.set_xlim(0, max(lvl_counts) * 1.15)
plt.tight_layout()
chart_levels_path = os.path.join(OUTPUT_DIR, "chart_education_levels.png")
plt.savefig(chart_levels_path)
plt.close()

# 5. Highest Qualifications Chart
qual_rows = cur.execute("""
    SELECT highest_qualification, count(*) as count
    FROM teachers
    GROUP BY highest_qualification
    ORDER BY count DESC
""").fetchall()

fig, ax = plt.subplots(figsize=(8, 4.8), dpi=300)
q_labels = [r['highest_qualification'] for r in qual_rows]
q_counts = [r['count'] for r in qual_rows]
colors_pie = [CLR_BLUE, CLR_EMERALD, CLR_AMBER, CLR_PURPLE, CLR_ROSE, CLR_LIGHT_GRAY]

wedges, texts, autotexts = ax.pie(q_counts, labels=q_labels, autopct='%1.1f%%', startangle=140,
                                  colors=colors_pie, textprops=dict(color="#0f172a", fontsize=8.5, fontweight='medium'),
                                  wedgeprops=dict(width=0.45, edgecolor='#ffffff', linewidth=2))
for autotext in autotexts:
    autotext.set_fontsize(8)
    autotext.set_fontweight('bold')
    autotext.set_color('#ffffff')

ax.set_title("Highest Academic Qualifications of Registered Tutors", fontsize=12.5, fontweight='bold', pad=15, color="#0f172a")
plt.tight_layout()
chart_quals_path = os.path.join(OUTPUT_DIR, "chart_qualifications.png")
plt.savefig(chart_quals_path)
plt.close()

# 6. Average Fee by Subject Category Chart
fee_cat_rows = cur.execute("""
    SELECT s.subject_category, round(avg(t.hourly_fee_avg), 0) as avg_fee
    FROM teacher_subjects s
    JOIN teachers t ON s.tutor_id = t.tutor_id
    WHERE t.hourly_fee_avg BETWEEN 50 AND 5000
    GROUP BY s.subject_category
    ORDER BY avg_fee ASC
""").fetchall()

fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=300)
fc_names = [r['subject_category'] for r in fee_cat_rows]
fc_fees = [int(r['avg_fee']) for r in fee_cat_rows]

bars = ax.barh(fc_names, fc_fees, color=CLR_EMERALD, edgecolor="#047857", linewidth=1, height=0.55)
ax.set_title("Average Hourly Tuition Fee by Subject Discipline (₹/hr)", fontsize=12, fontweight='bold', pad=15, color="#0f172a")
ax.set_xlabel("Average Hourly Rate (INR / Hour)", fontsize=10, fontweight='semibold', color="#334155")
ax.grid(axis='x', linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

for bar, fee in zip(bars, fc_fees):
    width = bar.get_width()
    ax.annotate(f" ₹{fee:,}/hr",
                xy=(width, bar.get_y() + bar.get_height() / 2),
                xytext=(3, 0), textcoords="offset points",
                ha='left', va='center', fontsize=9, fontweight='bold', color="#065f46")

ax.set_xlim(0, max(fc_fees) * 1.15)
plt.tight_layout()
chart_fees_path = os.path.join(OUTPUT_DIR, "chart_fees_by_category.png")
plt.savefig(chart_fees_path)
plt.close()

# 7. Competitive Exams Coaching Frequency Chart
exam_rows = cur.execute("""
    SELECT competitive_exam, count(*) as count
    FROM teacher_subjects
    WHERE competitive_exam IS NOT NULL
    GROUP BY competitive_exam
    ORDER BY count DESC
    LIMIT 10
""").fetchall()

fig, ax = plt.subplots(figsize=(9, 4.5), dpi=300)
ex_names = [r['competitive_exam'] for r in exam_rows]
ex_counts = [r['count'] for r in exam_rows]

bars = ax.bar(ex_names, ex_counts, color=CLR_ROSE, edgecolor="#9f1239", linewidth=1, width=0.55)
ax.set_title("Competitive Exam Coaching Frequency in India (Top 10 Target Exams)", fontsize=12, fontweight='bold', pad=15, color="#0f172a")
ax.set_ylabel("Number of Specialized Tutor Listings", fontsize=10, fontweight='semibold', color="#334155")
ax.grid(axis='y', linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
plt.xticks(rotation=20, ha='right', fontsize=9, fontweight='medium')

for bar, cnt in zip(bars, ex_counts):
    height = bar.get_height()
    ax.annotate(f"{cnt:,}",
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 3), textcoords="offset points",
                ha='center', va='bottom', fontsize=8.5, fontweight='bold', color="#1e293b")

ax.set_ylim(0, max(ex_counts) * 1.15)
plt.tight_layout()
chart_exams_path = os.path.join(OUTPUT_DIR, "chart_competitive_exams.png")
plt.savefig(chart_exams_path)
plt.close()

# 8. Tutors by Zonal Region Chart
region_rows = cur.execute("""
    SELECT region, count(*) as count, round(count(*) * 100.0 / 107723, 1) as pct
    FROM teachers
    WHERE region IS NOT NULL
    GROUP BY region
    ORDER BY count DESC
""").fetchall()

fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=300)
reg_names = [r['region'] for r in region_rows]
reg_counts = [r['count'] for r in region_rows]
reg_pcts = [r['pct'] for r in region_rows]

bars = ax.bar(reg_names, reg_counts, color=CLR_TEAL, edgecolor="#0f766e", linewidth=1, width=0.5)
ax.set_title("Tutor Geographic Distribution Across Indian Zonal Regions", fontsize=12, fontweight='bold', pad=15, color="#0f172a")
ax.set_ylabel("Total Tutor Headcount", fontsize=10, fontweight='semibold', color="#334155")
ax.grid(axis='y', linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

for bar, cnt, pct in zip(bars, reg_counts, reg_pcts):
    height = bar.get_height()
    ax.annotate(f"{cnt:,}\n({pct}%)",
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 3), textcoords="offset points",
                ha='center', va='bottom', fontsize=8.5, fontweight='bold', color="#1e293b")

ax.set_ylim(0, max(reg_counts) * 1.15)
plt.tight_layout()
chart_regions_path = os.path.join(OUTPUT_DIR, "chart_regional_distribution.png")
plt.savefig(chart_regions_path)
plt.close()

# 9. Non-Teaching & Side Occupations Analysis & Charts
print("Analyzing non-teaching side occupations across 107,723 tutors...")
side_occ_defs = {
    'University & College Students / Researchers': {
        'pattern': r'\b(student|pursuing|undergraduate student|b\.?tech student|m\.?tech student|research scholar|phd scholar|postdoc|fellow|intern)\b',
        'color': '#3b82f6',
        'subjects': 'Mathematics, Physics, Computer Science, Engineering Basics'
    },
    'Software Engineers & IT Professionals': {
        'pattern': r'\b(software engineer|software developer|web developer|full stack|frontend|backend|data scientist|data analyst|machine learning|ai engineer|python developer|java developer|devops|cloud architect|systems engineer|it consultant|qa engineer|tester|programmer|coder)\b',
        'color': '#6366f1',
        'subjects': 'Python, Java, Web Development, Data Science, C++, DSA'
    },
    'Medical & Healthcare Professionals': {
        'pattern': r'\b(doctor|mbbs|dentist|bds|physiotherapist|pharmacist|nursing|medical practitioner|veterinary|clinician|homeopath|ayurvedic)\b',
        'color': '#06b6d4',
        'subjects': 'Biology, NEET Prep, Anatomy, Physiology, Chemistry'
    },
    'Corporate, Finance & Chartered Accountants': {
        'pattern': r'\b(chartered accountant|\bca\b|chartered accountancy|accountant|financial analyst|investment banker|bank manager|banker|corporate trainer|consultant|business analyst|management consultant|hr manager|marketing manager|sales executive)\b',
        'color': '#10b981',
        'subjects': 'Accountancy, Financial Management, Economics, Taxation, Commerce'
    },
    'Freelancers & Solopreneurs': {
        'pattern': r'\b(freelancer|freelance consultant|solopreneur|entrepreneur|business owner|founder)\b',
        'color': '#8b5cf6',
        'subjects': 'Digital Marketing, Spoken English, Web Design, Business Studies'
    },
    'Core Industrial Engineers (Mech/Civil/EE)': {
        'pattern': r'\b(mechanical engineer|civil engineer|electrical engineer|electronics engineer|chemical engineer|embedded systems|vlsi|structural engineer|cad designer|aerospace)\b',
        'color': '#f59e0b',
        'subjects': 'Engineering Mechanics, AutoCAD, Mathematics, Thermodynamics'
    },
    'Creative Arts, Design & Media': {
        'pattern': r'\b(graphic designer|ui/ux designer|content writer|copywriter|journalist|animator|photographer|video editor|author|translator|voice artist)\b',
        'color': '#14b8a6',
        'subjects': 'Graphic Design, Creative Writing, Languages, Art, Animation'
    },
    'Legal & Civil Services Aspirants / Officers': {
        'pattern': r'\b(lawyer|advocate|legal consultant|attorney|civil servant|ias|ips|upsc aspirant|government employee|govt officer)\b',
        'color': '#ec4899',
        'subjects': 'Political Science, Legal Studies, History, UPSC GS, Constitution'
    },
    'Homemakers & Home-Based Educators': {
        'pattern': r'\b(homemaker|housewife|stay at home mom)\b',
        'color': '#f43f5e',
        'subjects': 'Primary School All Subjects, Hindi, Phonics, Early Maths'
    }
}

df_raw_teachers = pd.read_sql('''
    SELECT 
        tutor_id, teacher_hookline, experience_raw, description, 
        hourly_fee_avg, teaches_online
    FROM teachers
''', conn)

matched_side_list = []
for idx, row in df_raw_teachers.iterrows():
    text = (str(row['teacher_hookline'] or '') + ' ' + str(row['experience_raw'] or '') + ' ' + str(row['description'] or '')[:250]).lower()
    for cat_name, info in side_occ_defs.items():
        if re.search(info['pattern'], text):
            fee = row['hourly_fee_avg']
            valid_fee = fee if (pd.notnull(fee) and 50 <= fee <= 5000) else np.nan
            online = 1 if row['teaches_online'] in [1, '1', True, 'Yes', 'yes'] else 0
            matched_side_list.append({
                'tutor_id': row['tutor_id'],
                'category': cat_name,
                'fee': valid_fee,
                'online': online,
                'color': info['color'],
                'subjects': info['subjects']
            })
            break

df_side = pd.DataFrame(matched_side_list)
total_side_cohort = len(df_side)
agg_side = df_side.groupby('category').agg(
    count=('tutor_id', 'count'),
    avg_fee=('fee', 'mean'),
    median_fee=('fee', 'median'),
    online_pct=('online', lambda x: x.mean() * 100)
).reset_index()
agg_side['pct'] = agg_side['count'] * 100.0 / total_side_cohort
agg_side['color'] = agg_side['category'].map(lambda c: side_occ_defs[c]['color'])
agg_side['subjects'] = agg_side['category'].map(lambda c: side_occ_defs[c]['subjects'])

# 9A. Figure 6A: Distribution Chart
agg_cnt = agg_side.sort_values(by='count', ascending=True)
fig, ax = plt.subplots(figsize=(9.2, 5.0), dpi=300)
bars = ax.barh(agg_cnt['category'], agg_cnt['count'], color=agg_cnt['color'], edgecolor='#1e293b', linewidth=0.8, height=0.6)
ax.set_title(f"Figure 6A: Distribution of Non-Teaching Side Occupations (N = {total_side_cohort:,})", fontsize=11.5, fontweight='bold', pad=15, color='#0f172a')
ax.set_xlabel("Identified Tutor Headcount", fontsize=10, fontweight='semibold', color='#334155')
ax.grid(axis='x', linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
for bar, cnt, pct in zip(bars, agg_cnt['count'], agg_cnt['pct']):
    width = bar.get_width()
    ax.annotate(f" {cnt:,} ({pct:.1f}%)",
                xy=(width, bar.get_y() + bar.get_height() / 2),
                xytext=(3, 0), textcoords='offset points',
                ha='left', va='center', fontsize=8.5, fontweight='bold', color='#1e293b')
ax.set_xlim(0, max(agg_cnt['count']) * 1.25)
plt.tight_layout()
fig6a_path = os.path.join(OUTPUT_DIR, "fig6a_side_occupations_dist.png")
plt.savefig(fig6a_path)
plt.close()

# 9B. Figure 6B: Average Hourly Fee Chart
agg_fee = agg_side.sort_values(by='avg_fee', ascending=True)
fig, ax = plt.subplots(figsize=(9.2, 5.0), dpi=300)
bars = ax.barh(agg_fee['category'], agg_fee['avg_fee'], color='#10b981', edgecolor='#047857', linewidth=0.8, height=0.6)
ax.set_title("Figure 6B: Hourly Tuition Pricing Benchmark by Side Profession (₹/hr)", fontsize=11.5, fontweight='bold', pad=15, color='#0f172a')
ax.set_xlabel("Average Hourly Fee (INR / Hour)", fontsize=10, fontweight='semibold', color='#334155')
ax.grid(axis='x', linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
for bar, fee, med in zip(bars, agg_fee['avg_fee'], agg_fee['median_fee']):
    width = bar.get_width()
    ax.annotate(f" ₹{int(round(fee)):,}/hr (Median: ₹{int(round(med)):,})",
                xy=(width, bar.get_y() + bar.get_height() / 2),
                xytext=(3, 0), textcoords='offset points',
                ha='left', va='center', fontsize=8.5, fontweight='bold', color='#065f46')
ax.set_xlim(0, 1600)
plt.tight_layout()
fig6b_path = os.path.join(OUTPUT_DIR, "fig6b_side_occupations_fee.png")
plt.savefig(fig6b_path)
plt.close()

# 9C. Figure 6C: Online Adoption Rate Chart
agg_online = agg_side.sort_values(by='online_pct', ascending=True)
fig, ax = plt.subplots(figsize=(9.2, 4.8), dpi=300)
bars = ax.barh(agg_online['category'], agg_online['online_pct'], color='#6366f1', edgecolor='#4338ca', linewidth=0.8, height=0.58)
ax.set_title("Figure 6C: Online Tutoring Delivery Share by Side Profession (%)", fontsize=11.5, fontweight='bold', pad=15, color='#0f172a')
ax.set_xlabel("Online Tutoring Share (%)", fontsize=10, fontweight='semibold', color='#334155')
ax.grid(axis='x', linestyle='--', alpha=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
for bar, opct in zip(bars, agg_online['online_pct']):
    width = bar.get_width()
    ax.annotate(f" {opct:.1f}% Online",
                xy=(width, bar.get_y() + bar.get_height() / 2),
                xytext=(3, 0), textcoords='offset points',
                ha='left', va='center', fontsize=8.5, fontweight='bold', color='#312e81')
ax.set_xlim(85, 103)
plt.tight_layout()
fig6c_path = os.path.join(OUTPUT_DIR, "fig6c_side_occupations_online.png")
plt.savefig(fig6c_path)
plt.close()

print("All 300 DPI charts generated successfully!")

# -------------------------------------------------------------
# 2. Build Publication-Grade Word Document (DOCX)
# -------------------------------------------------------------
print("Building Microsoft Word DOCX Document...")
doc = Document()

# Page Setup: Standard Letter / A4 with 0.8 inch margins
sections = doc.sections
for section in sections:
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

# Color Palette Constants
COLOR_TITLE = RGBColor(30, 41, 59)      # Slate 900
COLOR_SUBTITLE = RGBColor(71, 85, 105)  # Slate 600
COLOR_PRIMARY = RGBColor(79, 70, 229)   # Indigo 600

# Document Title
title_p = doc.add_paragraph()
title_p.paragraph_format.space_before = Pt(0)
title_p.paragraph_format.space_after = Pt(4)
r = title_p.add_run("BMP Tutors: Comprehensive Visual Analytics & Market Intelligence")
r.font.name = "Arial"
r.font.size = Pt(20)
r.font.bold = True
r.font.color.rgb = COLOR_TITLE

# Subtitle
sub_p = doc.add_paragraph()
sub_p.paragraph_format.space_after = Pt(14)
r = sub_p.add_run("Complete Empirical Dossier: 107,723 Verified Educators, 403,353 Subject Offerings, 37 States & UTs, and Academic Pricing Architecture")
r.font.name = "Arial"
r.font.size = Pt(10.5)
r.font.color.rgb = COLOR_SUBTITLE

# Metadata Bar
meta_p = doc.add_paragraph()
meta_p.paragraph_format.space_after = Pt(14)
r = meta_p.add_run("Dataset Source: ")
r.font.bold = True
r.font.size = Pt(9)
meta_p.add_run("TeacherOn India Platform • ").font.size = Pt(9)
r = meta_p.add_run("Date: ")
r.font.bold = True
r.font.size = Pt(9)
meta_p.add_run("September 2026 • ").font.size = Pt(9)
r = meta_p.add_run("Status: ")
r.font.bold = True
r.font.size = Pt(9)
meta_p.add_run("Validated, Cleaned & Normalized (Master 3-Sheet Relational Schema)").font.size = Pt(9)

# Executive KPI Summary Table
kpi_table = doc.add_table(rows=2, cols=6)
kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(kpi_table, color="CBD5E1", sz="4")

kpis = [
    ("TOTAL TUTORS", "107,723", "100% deduplicated"),
    ("ALL-INDIA AVG FEE", "₹657 / hr", "Valid ₹50-₹5,000"),
    ("ONLINE ADOPTION", "74.3%", "80,038 tutors"),
    ("ACTIVE EDUCATORS", "85,618", "79.5% active rate"),
    ("SUBJECT LINKS", "403,353", "3.7 subjects / tutor"),
    ("STATES & UTS", "37 Covered", "Full pan-India scope")
]

for col_idx, (label, val, sub) in enumerate(kpis):
    # Header cell
    c_top = kpi_table.cell(0, col_idx)
    set_cell_background(c_top, "1E293B")
    set_cell_margins(c_top, top=80, bottom=60, left=100, right=100)
    p = c_top.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(label)
    run.font.name = "Arial"
    run.font.size = Pt(7.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

    # Value cell
    c_bot = kpi_table.cell(1, col_idx)
    set_cell_background(c_bot, "F8FAFC")
    set_cell_margins(c_bot, top=100, bottom=80, left=100, right=100)
    p = c_bot.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(val)
    run.font.name = "Arial"
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_TITLE

    p2 = c_bot.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(2)
    run2 = p2.add_run(sub)
    run2.font.name = "Arial"
    run2.font.size = Pt(7)
    run2.font.color.rgb = COLOR_SUBTITLE

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# =========================================================================
# SECTION 1: GEOGRAPHIC MARKET DISTRIBUTION & MAP OF INDIA
# =========================================================================
h1 = doc.add_heading(level=1)
r = h1.add_run("1. Geographic Market Distribution: Map of India & District Spatial Structure")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p = doc.add_paragraph(
    "Private tutoring supply in India displays strong geographic concentration across regional economic hubs and educational belts. "
    "Replicating the clean academic architecture of the Visual Analytics Dashboard, this section presents: "
    "(1) The Map of India showing state-wise administrative boundaries with direct in-polygon verified tutor counts for all 36 states and union territories, "
    "(2) The complete census table covering all 37 Indian states and union territories, and "
    "(3) Individual state administrative diagrams displaying localized district boundaries, headcount labels, and full district data tables for every state."
)
p.paragraph_format.space_after = Pt(12)

# Helper function for adding figures
def add_figure_block(doc, img_path, title_text, caption_text, width=Inches(6.2), space_after=14, page_break=False):
    if page_break:
        doc.add_page_break()
    if os.path.exists(img_path):
        h = doc.add_heading(level=2)
        hr = h.add_run(title_text)
        hr.font.name = "Arial"
        hr.font.color.rgb = COLOR_TITLE
        
        doc.add_picture(img_path, width=width)
        cp = doc.add_paragraph()
        cp.paragraph_format.space_before = Pt(4)
        cp.paragraph_format.space_after = Pt(space_after)
        cpr = cp.add_run(caption_text)
        cpr.font.name = "Arial"
        cpr.font.size = Pt(8.5)
        cpr.font.italic = True
        cpr.font.color.rgb = COLOR_SUBTITLE

MAPS_DIR = os.path.join(REF_IMG_DIR, "maps")

# 1.1 Map of India: State-Wise Boundaries & Data Labels (Dashboard Style)
add_figure_block(
    doc,
    os.path.join(MAPS_DIR, "india_all_states_labeled_dashboard.png"),
    "Figure 1A: Map of India — State-Wise Boundaries & Data Labels (N = 107,723)",
    "State administrative boundaries with direct in-polygon numerical labels and choropleth shading, matching the Visual Analytics Dashboard. "
    "Every single state and union territory in India is labeled with its verified tutor headcount. "
    "Leading markets include Tamil Nadu (16,811), Kerala (14,337), Maharashtra (9,550), Karnataka (9,300), Delhi NCR (8,600), "
    "Uttar Pradesh (8,114), Telangana (7,810), West Bengal (7,026), and Haryana (3,228)."
)

# 1.5 Table 1: All 37 States
doc.add_page_break()
h2 = doc.add_heading(level=2)
r = h2.add_run("Table 1: All-India State & Union Territory Complete Distribution")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

states_data = cur.execute("""
    SELECT 
        state, region, 
        count(*) as tutor_count,
        round(count(*) * 100.0 / 107723, 2) as pct_share,
        round(avg(case when hourly_fee_avg between 50 and 5000 then hourly_fee_avg end), 0) as avg_fee,
        round(avg(teaches_online) * 100, 1) as online_pct,
        sum(case when is_active = 1 then 1 else 0 end) as active_count
    FROM teachers
    WHERE state IS NOT NULL AND state NOT IN ('Other / International', 'Other')
    GROUP BY state
    ORDER BY tutor_count DESC
""").fetchall()

t1 = doc.add_table(rows=len(states_data) + 2, cols=8)
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t1, color="CBD5E1", sz="4")

headers_t1 = ["Rank", "State / Union Territory", "Zone", "Tutors (N)", "Share (%)", "Avg Fee (₹)", "Online %", "Active"]
widths_t1 = [Inches(0.55), Inches(1.8), Inches(0.85), Inches(0.85), Inches(0.75), Inches(0.85), Inches(0.75), Inches(0.8)]

for col_idx, h_text in enumerate(headers_t1):
    c = t1.cell(0, col_idx)
    c.width = widths_t1[col_idx]
    set_cell_background(c, "0F172A")
    set_cell_margins(c, top=80, bottom=60, left=60, right=60)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
    run = p.add_run(h_text)
    run.font.name = "Arial"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

tot_cnt = 0
tot_act = 0
w_fee_sum = 0
w_fee_n = 0
w_onl_sum = 0

for row_idx, s in enumerate(states_data, start=1):
    tot_cnt += s['tutor_count']
    tot_act += s['active_count']
    if s['avg_fee']:
        w_fee_sum += s['avg_fee'] * s['tutor_count']
        w_fee_n += s['tutor_count']
    w_onl_sum += s['online_pct'] * s['tutor_count']

    row_vals = [
        str(row_idx),
        s['state'],
        s['region'] or "Other",
        f"{s['tutor_count']:,}",
        f"{s['pct_share']:.2f}%",
        f"₹{int(s['avg_fee'])}" if s['avg_fee'] else "—",
        f"{s['online_pct']}%",
        f"{s['active_count']:,}"
    ]
    bg = "F8FAFC" if row_idx % 2 == 0 else "FFFFFF"
    for col_idx, val in enumerate(row_vals):
        c = t1.cell(row_idx, col_idx)
        c.width = widths_t1[col_idx]
        set_cell_background(c, bg)
        set_cell_margins(c, top=50, bottom=50, left=60, right=60)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
        run = p.add_run(val)
        run.font.name = "Arial"
        run.font.size = Pt(7.5)
        if col_idx in (1, 3):
            run.font.bold = True

# Summary footer row
foot_idx = len(states_data) + 1
nat_fee = int(round(w_fee_sum / w_fee_n)) if w_fee_n else 657
nat_onl = f"{w_onl_sum / tot_cnt:.1f}%"
foot_vals = ["—", "Total / Pan-India Average", "All Zones", f"{tot_cnt:,}", "100.00%", f"₹{nat_fee}", nat_onl, f"{tot_act:,}"]

for col_idx, val in enumerate(foot_vals):
    c = t1.cell(foot_idx, col_idx)
    c.width = widths_t1[col_idx]
    set_cell_background(c, "E2E8F0")
    set_cell_margins(c, top=70, bottom=70, left=60, right=60)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
    run = p.add_run(val)
    run.font.name = "Arial"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = COLOR_TITLE

foot_p = doc.add_paragraph()
foot_p.paragraph_format.space_before = Pt(4)
foot_p.paragraph_format.space_after = Pt(14)
r = foot_p.add_run("Notes: All 37 states and union territories sorted by total tutor volume. Average hourly rates exclude unverified outliers outside ₹50–₹5,000/hr.")
r.font.name = "Arial"
r.font.size = Pt(7.5)
r.font.italic = True
r.font.color.rgb = COLOR_SUBTITLE

# =========================================================================
# SUBSECTION 1.3: INDIVIDUAL STATES DISTRICT-WISE DATA & PROFILES
# =========================================================================
doc.add_page_break()
h1_dist = doc.add_heading(level=1)
r = h1_dist.add_run("1.3 Individual States & Union Territories: Complete District-Wise Data Portfolio")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p_dist_intro = doc.add_paragraph(
    "This section provides the comprehensive micro-geographic breakdown of private tutoring supply for every individual Indian state and union territory. "
    "Written on the top of each state's section is an executive profile detailing total registered educator volume, national market share, pan-India rank, "
    "average hourly lesson fee, online delivery adoption, and active engagement, followed by the complete census of all constituent administrative districts."
)
p_dist_intro.paragraph_format.space_after = Pt(14)

states_geo_all = get_enriched_states_geojson()
sorted_states_all = sorted(states_geo_all['features'], key=lambda f: f['properties'].get('tutor_count', 0), reverse=True)

for st_rank, s_feat in enumerate(sorted_states_all, 1):
    sp = s_feat['properties']
    s_name = sp['state']
    s_cnt = sp.get('tutor_count', 0)
    s_fee = sp.get('avg_fee', 0)
    s_online = sp.get('online_pct', 0)
    s_active = sp.get('active_count', 0)
    s_region = sp.get('region', 'Other')
    s_pct = round(s_cnt * 100.0 / 107723, 2)
    s_act_pct = round(s_active * 100.0 / s_cnt, 1) if s_cnt > 0 else 0.0

    # Heading for the individual state
    h_st = doc.add_heading(level=2)
    r = h_st.add_run(f"State {st_rank}: {s_name} — District-Wise Tutoring Supply (N = {s_cnt:,})")
    r.font.name = "Arial"
    r.font.color.rgb = COLOR_TITLE

    # Executive Profile on top of table
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(2)
    p_top.paragraph_format.space_after = Pt(6)
    
    r = p_top.add_run(f"EXECUTIVE SUMMARY & STATE PROFILE: {s_name.upper()}\n")
    r.font.name = "Arial"
    r.font.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = COLOR_PRIMARY
    
    prof_text = (
        f"• Total Registered Educators: {s_cnt:,} verified tutors ({s_pct:.2f}% national market share | Pan-India Rank: #{st_rank} of 36)\n"
        f"• Average Hourly Tuition Rate: ₹{int(round(s_fee))} / hr (statewide median pricing)\n"
        f"• Online Delivery Adoption: {s_online:.1f}% of educators teach via digital platforms\n"
        f"• Active Educators: {s_active:,} tutors ({s_act_pct:.1f}% engagement rate)\n"
        f"• Geographic Zone: {s_region} Region"
    )
    r = p_top.add_run(prof_text)
    r.font.name = "Arial"
    r.font.size = Pt(8)
    r.font.color.rgb = COLOR_TITLE

    # Embed that particular state's district boundary diagram with data labels
    s_slug = s_name.lower().replace(" ", "_").replace("&", "and").replace("/", "_").replace("(", "").replace(")", "").replace("-", "_")
    s_map_img = os.path.join(REF_IMG_DIR, "state_maps", f"{s_slug}.png")
    if os.path.exists(s_map_img):
        doc.add_picture(s_map_img, width=Inches(5.6))
        cap_map = doc.add_paragraph()
        cap_map.paragraph_format.space_before = Pt(3)
        cap_map.paragraph_format.space_after = Pt(8)
        rc = cap_map.add_run(f"Figure 1.{st_rank}: {s_name} District Boundary Diagram — Localized Tutor Counts (N = {s_cnt:,})")
        rc.font.name = "Arial"
        rc.font.size = Pt(8)
        rc.font.italic = True
        rc.font.color.rgb = COLOR_SUBTITLE

    # Get district data for this state
    d_res = get_enriched_districts_geojson(s_name)
    d_feats = d_res.get('features', [])
    d_sorted = sorted(d_feats, key=lambda x: x['properties'].get('tutor_count', 0), reverse=True)

    # Create district table
    num_rows = max(len(d_sorted), 1) + 2  # header + rows + summary footer
    t_dist = doc.add_table(rows=num_rows, cols=6)
    t_dist.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_dist, color="CBD5E1", sz="4")

    headers_dst = ["Rank", "District / Administrative Division", "Tutors (N)", "Share of State (%)", "Avg Fee (₹/hr)", "Online %"]
    widths_dst = [Inches(0.6), Inches(2.5), Inches(1.0), Inches(1.0), Inches(1.1), Inches(0.9)]

    for col_idx, h_text in enumerate(headers_dst):
        c = t_dist.cell(0, col_idx)
        c.width = widths_dst[col_idx]
        set_cell_background(c, "1E293B")
        set_cell_margins(c, top=70, bottom=50, left=60, right=60)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx >= 2 else (WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
        run = p.add_run(h_text)
        run.font.name = "Arial"
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)

    if not d_sorted:
        # Single fallback row
        c0 = t_dist.cell(1, 0); c0.paragraphs[0].add_run("1").font.size = Pt(7.5)
        c1 = t_dist.cell(1, 1); c1.paragraphs[0].add_run(s_name).font.size = Pt(7.5)
        c2 = t_dist.cell(1, 2); c2.paragraphs[0].add_run("0").font.size = Pt(7.5)
        c3 = t_dist.cell(1, 3); c3.paragraphs[0].add_run("0.0%").font.size = Pt(7.5)
        c4 = t_dist.cell(1, 4); c4.paragraphs[0].add_run("—").font.size = Pt(7.5)
        c5 = t_dist.cell(1, 5); c5.paragraphs[0].add_run("—").font.size = Pt(7.5)
    else:
        for d_idx, df in enumerate(d_sorted, start=1):
            dp = df['properties']
            d_name = dp.get('district', 'Unknown')
            d_cnt = dp.get('tutor_count', 0)
            d_fee = dp.get('avg_fee', 0)
            d_onl = dp.get('online_pct', 0.0)
            d_share = f"{(d_cnt * 100.0 / s_cnt):.1f}%" if s_cnt > 0 else "0.0%"
            
            row_vals = [
                str(d_idx),
                d_name,
                f"{d_cnt:,}",
                d_share,
                f"₹{int(round(d_fee))}" if d_fee > 0 else "—",
                f"{d_onl:.1f}%"
            ]
            bg = "F8FAFC" if d_idx % 2 == 0 else "FFFFFF"
            
            for col_idx, val in enumerate(row_vals):
                c = t_dist.cell(d_idx, col_idx)
                c.width = widths_dst[col_idx]
                set_cell_background(c, bg)
                set_cell_margins(c, top=45, bottom=45, left=60, right=60)
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx >= 2 else (WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
                run = p.add_run(val)
                run.font.name = "Arial"
                run.font.size = Pt(7.5)
                if col_idx in (1, 2) and d_cnt > 0:
                    run.font.bold = True

    # Footer summary row
    foot_idx = num_rows - 1
    foot_vals = ["—", f"{s_name} Total", f"{s_cnt:,}", "100.0%", f"₹{int(round(s_fee))}" if s_fee > 0 else "—", f"{s_online:.1f}%"]
    for col_idx, val in enumerate(foot_vals):
        c = t_dist.cell(foot_idx, col_idx)
        c.width = widths_dst[col_idx]
        set_cell_background(c, "E2E8F0")
        set_cell_margins(c, top=60, bottom=60, left=60, right=60)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx >= 2 else (WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
        run = p.add_run(val)
        run.font.name = "Arial"
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = COLOR_TITLE

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

# =========================================================================
# SECTION 2: GENDER DISTRIBUTION (TABLE 1)
# =========================================================================
doc.add_page_break()
h1 = doc.add_heading(level=1)
r = h1.add_run("2. Gender Distribution Across Tutoring Disciplines")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p = doc.add_paragraph(
    "In line with academic literature on private tutoring, gender segregation differs across subject domains. "
    "Overall, women comprise 58.7% of registered tutors in India ($N = 62,551$), while men represent 41.3% ($N = 44,001$). "
    "However, STEM sciences and computer programming demonstrate significantly higher male tutor shares compared to languages and primary education."
)
p.paragraph_format.space_after = Pt(8)

h2 = doc.add_heading(level=2)
r = h2.add_run("Table 2: Gender Distribution by Tutoring Subject (Broad Domains & Top Subjects)")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

# Fetch Gender Table Data
gender_cats = cur.execute("""
    SELECT 
        ts.subject_category as subject,
        COUNT(DISTINCT ts.tutor_id) as n,
        COUNT(DISTINCT CASE WHEN t.gender = 'Male' THEN ts.tutor_id END) as male,
        COUNT(DISTINCT CASE WHEN t.gender = 'Female' THEN ts.tutor_id END) as female
    FROM teacher_subjects ts
    JOIN teachers t ON ts.tutor_id = t.tutor_id
    GROUP BY ts.subject_category
    ORDER BY n DESC
""").fetchall()

gender_subs = cur.execute("""
    SELECT 
        ts.canonical_subject as subject,
        COUNT(DISTINCT ts.tutor_id) as n,
        COUNT(DISTINCT CASE WHEN t.gender = 'Male' THEN ts.tutor_id END) as male,
        COUNT(DISTINCT CASE WHEN t.gender = 'Female' THEN ts.tutor_id END) as female
    FROM teacher_subjects ts
    JOIN teachers t ON ts.tutor_id = t.tutor_id
    GROUP BY ts.canonical_subject
    HAVING n >= 4000
    ORDER BY n DESC
""").fetchall()

tot_row = cur.execute("""
    SELECT 
        COUNT(DISTINCT ts.tutor_id) as n,
        COUNT(DISTINCT CASE WHEN t.gender = 'Male' THEN ts.tutor_id END) as male,
        COUNT(DISTINCT CASE WHEN t.gender = 'Female' THEN ts.tutor_id END) as female
    FROM teacher_subjects ts
    JOIN teachers t ON ts.tutor_id = t.tutor_id
""").fetchone()

total_rows_t2 = len(gender_cats) + 1 + len(gender_subs) + 2
t2 = doc.add_table(rows=total_rows_t2, cols=4)
t2.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t2, color="CBD5E1", sz="4")

headers_t2 = ["Subject Domain / Discipline", "Men (%)", "Women (%)", "Sample Size (N)"]
widths_t2 = [Inches(3.4), Inches(1.1), Inches(1.1), Inches(1.4)]

for col_idx, h_text in enumerate(headers_t2):
    c = t2.cell(0, col_idx)
    c.width = widths_t2[col_idx]
    set_cell_background(c, "0F172A")
    set_cell_margins(c, top=80, bottom=60, left=70, right=70)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx >= 1 else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(h_text)
    run.font.name = "Arial"
    run.font.size = Pt(8.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

curr_r = 1
CAT_MAP = {
    'Science': 'Natural sciences & mathematics (physics, chemistry, biology, math)',
    'Languages': 'Foreign & modern languages (English, Hindi, Tamil, French, etc.)',
    'Coding': 'Computer science & programming (Python, Java, web dev, AI/ML)',
    'Arts': 'History, geography, social sciences & humanities',
    'Commerce': 'Commerce, economics & business studies (accounting, finance)',
    'Misc/ECA': 'Creative arts, music, yoga & extracurricular activities',
    'All Subjects': 'General elementary & primary school foundation (all subjects)'
}

for r_data in gender_cats:
    m = r_data['male']
    f = r_data['female']
    tot = m + f
    m_pct = round(m / tot * 100) if tot else 0
    f_pct = round(f / tot * 100) if tot else 0
    
    vals = [CAT_MAP.get(r_data['subject'], r_data['subject']), f"{m_pct}%", f"{f_pct}%", f"{r_data['n']:,}"]
    for c_i, val in enumerate(vals):
        c = t2.cell(curr_r, c_i)
        c.width = widths_t2[c_i]
        set_cell_background(c, "F8FAFC" if curr_r % 2 == 0 else "FFFFFF")
        set_cell_margins(c, top=55, bottom=55, left=70, right=70)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_i >= 1 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(val)
        run.font.name = "Arial"
        run.font.size = Pt(8)
        if c_i == 0: run.font.bold = True
    curr_r += 1

# Section Sub-header
c_sub = t2.cell(curr_r, 0)
set_cell_background(c_sub, "E2E8F0")
for c_i in range(1, 4):
    set_cell_background(t2.cell(curr_r, c_i), "E2E8F0")
p = c_sub.paragraphs[0]
run = p.add_run("Top Specific Individual Tutoring Disciplines (N >= 4,000)")
run.font.name = "Arial"
run.font.size = Pt(8)
run.font.bold = True
run.font.color.rgb = COLOR_TITLE
curr_r += 1

for r_data in gender_subs:
    m = r_data['male']
    f = r_data['female']
    tot = m + f
    m_pct = round(m / tot * 100) if tot else 0
    f_pct = round(f / tot * 100) if tot else 0
    
    vals = [f"  • {r_data['subject']}", f"{m_pct}%", f"{f_pct}%", f"{r_data['n']:,}"]
    for c_i, val in enumerate(vals):
        c = t2.cell(curr_r, c_i)
        c.width = widths_t2[c_i]
        set_cell_background(c, "FFFFFF")
        set_cell_margins(c, top=45, bottom=45, left=70, right=70)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_i >= 1 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(val)
        run.font.name = "Arial"
        run.font.size = Pt(7.5)
    curr_r += 1

# Footer total
tot_m = tot_row['male']
tot_f = tot_row['female']
tot_v = tot_m + tot_f
foot_vals = ["Total Unique Educators", f"{round(tot_m/tot_v*100)}%", f"{round(tot_f/tot_v*100)}%", f"{tot_row['n']:,}"]

for c_i, val in enumerate(foot_vals):
    c = t2.cell(curr_r, c_i)
    c.width = widths_t2[c_i]
    set_cell_background(c, "E2E8F0")
    set_cell_margins(c, top=70, bottom=70, left=70, right=70)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_i >= 1 else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(val)
    run.font.name = "Arial"
    run.font.size = Pt(8.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_TITLE

p_note = doc.add_paragraph()
p_note.paragraph_format.space_before = Pt(4)
p_note.paragraph_format.space_after = Pt(14)
r = p_note.add_run("Notes: If tutors offer tutoring in two or more subjects, they are counted for each subject. Percentages are calculated across tutors with specified gender records.")
r.font.name = "Arial"
r.font.size = Pt(7.5)
r.font.italic = True
r.font.color.rgb = COLOR_SUBTITLE

# =========================================================================
# SECTION 3: PRICING & FEE ARCHITECTURE (TABLE 2 / TABLE 2.8)
# =========================================================================
doc.add_page_break()
h1 = doc.add_heading(level=1)
r = h1.add_run("3. Pricing Architecture & Educational Stage Escalation")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p = doc.add_paragraph(
    "Tuition rates in India demonstrate sharp escalation across educational tiers. "
    "Elementary stages (Primary & Middle School) command lower hourly fees (₹650–₹700/hr), "
    "whereas Senior Secondary (Classes 11–12 board and entrance exam preparation) and College-level tuition "
    "climb significantly, exceeding ₹1,000–₹1,220/hr, particularly in Computer Science/Coding and Mathematics."
)
p.paragraph_format.space_after = Pt(8)

h2 = doc.add_heading(level=2)
r = h2.add_run("Table 3: Hourly Tuition Fees by Educational Level and Geographic Zone (Maths & English)")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

# Fetch Table 2.8 data
raw_fees = cur.execute("""
    SELECT 
        CASE 
            WHEN ts.canonical_subject = 'Mathematics' THEN 'Maths'
            WHEN ts.canonical_subject = 'English' THEN 'English'
            WHEN ts.canonical_subject IN ('Physics', 'Chemistry', 'Biology & Life Sciences', 'Science (General)') THEN 'Sciences'
            WHEN ts.subject_category = 'Coding' OR ts.canonical_subject = 'Computer Science & Programming' THEN 'Coding'
            WHEN ts.subject_category = 'Languages' AND ts.canonical_subject != 'English' THEN 'Languages'
            ELSE 'Other'
        END as subj,
        CASE 
            WHEN ts.level_category IN ('Primary (Class 1–5)', 'Middle School (Class 6–8)', 'Pre-Primary (Early Childhood / KG)') THEN 'Primary'
            WHEN ts.level_category LIKE 'Secondary (Class 9–10%' THEN 'Secondary'
            WHEN ts.level_category LIKE 'Senior Secondary%' THEN 'SeniorSec'
            WHEN ts.level_category LIKE 'Higher Education%' OR ts.level_category LIKE 'Professional%' THEN 'HigherEd'
            ELSE 'Other'
        END as lvl,
        COALESCE(NULLIF(t.region, ''), 'Other') as reg,
        count(*) as n, 
        round(avg(t.hourly_fee_avg), 0) as fee
    FROM teacher_subjects ts
    JOIN teachers t ON ts.tutor_id = t.tutor_id
    WHERE t.hourly_fee_avg BETWEEN 50 AND 5000
    GROUP BY subj, lvl, reg
""").fetchall()

nat_fees = cur.execute("""
    SELECT 
        CASE 
            WHEN ts.canonical_subject = 'Mathematics' THEN 'Maths'
            WHEN ts.canonical_subject = 'English' THEN 'English'
            WHEN ts.canonical_subject IN ('Physics', 'Chemistry', 'Biology & Life Sciences', 'Science (General)') THEN 'Sciences'
            WHEN ts.subject_category = 'Coding' OR ts.canonical_subject = 'Computer Science & Programming' THEN 'Coding'
            WHEN ts.subject_category = 'Languages' AND ts.canonical_subject != 'English' THEN 'Languages'
            ELSE 'Other'
        END as subj,
        CASE 
            WHEN ts.level_category IN ('Primary (Class 1–5)', 'Middle School (Class 6–8)', 'Pre-Primary (Early Childhood / KG)') THEN 'Primary'
            WHEN ts.level_category LIKE 'Secondary (Class 9–10%' THEN 'Secondary'
            WHEN ts.level_category LIKE 'Senior Secondary%' THEN 'SeniorSec'
            WHEN ts.level_category LIKE 'Higher Education%' OR ts.level_category LIKE 'Professional%' THEN 'HigherEd'
            ELSE 'Other'
        END as lvl,
        count(*) as n, 
        round(avg(t.hourly_fee_avg), 0) as fee
    FROM teacher_subjects ts
    JOIN teachers t ON ts.tutor_id = t.tutor_id
    WHERE t.hourly_fee_avg BETWEEN 50 AND 5000
    GROUP BY subj, lvl
""").fetchall()

table_data = {}
for r in nat_fees:
    if r['subj'] != 'Other' and r['lvl'] != 'Other':
        table_data.setdefault('Nationwide', {}).setdefault(r['lvl'], {})[r['subj']] = {'fee': int(r['fee']), 'n': r['n']}

for r in raw_fees:
    if r['subj'] != 'Other' and r['lvl'] != 'Other' and r['reg'] not in ('Other', 'Not Specified'):
        table_data.setdefault(r['reg'], {}).setdefault(r['lvl'], {})[r['subj']] = {'fee': int(r['fee']), 'n': r['n']}

regions_order = ['Nationwide', 'North', 'South', 'West', 'East', 'Central']
levels_def = [
    ('Primary', 'Primary (Class 1-8)'),
    ('Secondary', 'Secondary (Class 9-10)'),
    ('SeniorSec', 'Senior Sec (Class 11-12)'),
    ('HigherEd', 'Higher Ed (College)')
]

t3 = doc.add_table(rows=len(regions_order) + 2, cols=9)
t3.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t3, color="CBD5E1", sz="4")

# Top Header Row (Merged level headers)
c_reg = t3.cell(0, 0)
set_cell_background(c_reg, "0F172A")
p = c_reg.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = p.add_run("Region / Zone")
run.font.name = "Arial"
run.font.size = Pt(8)
run.font.bold = True
run.font.color.rgb = RGBColor(255, 255, 255)

col_offset = 1
for l_key, l_label in levels_def:
    for sub_idx, sub_name in enumerate(["Maths", "English"]):
        c_head = t3.cell(0, col_offset + sub_idx)
        set_cell_background(c_head, "1E293B")
        set_cell_margins(c_head, top=60, bottom=40, left=40, right=40)
        p = c_head.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"{l_label[:11]}\n{sub_name}")
        run.font.name = "Arial"
        run.font.size = Pt(7)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
    col_offset += 2

# Sub-Header Row
c_reg_sub = t3.cell(1, 0)
set_cell_background(c_reg_sub, "1E293B")
p = c_reg_sub.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = p.add_run("Geographic Zone")
run.font.name = "Arial"
run.font.size = Pt(7.5)
run.font.bold = True
run.font.color.rgb = RGBColor(255, 255, 255)

col_offset = 1
for _ in levels_def:
    for sub_name in ["Maths (₹|n)", "English (₹|n)"]:
        c = t3.cell(1, col_offset)
        set_cell_background(c, "334155")
        set_cell_margins(c, top=40, bottom=40, left=40, right=40)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(sub_name)
        run.font.name = "Arial"
        run.font.size = Pt(7)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        col_offset += 1

# Fill Regional Rows
for r_idx, reg in enumerate(regions_order, start=2):
    is_nat = (reg == 'Nationwide')
    bg = "E2E8F0" if is_nat else ("F8FAFC" if r_idx % 2 == 0 else "FFFFFF")
    
    c_lbl = t3.cell(r_idx, 0)
    set_cell_background(c_lbl, bg)
    set_cell_margins(c_lbl, top=50, bottom=50, left=60, right=60)
    p = c_lbl.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Nationwide (All India)" if is_nat else f"{reg} Zone")
    run.font.name = "Arial"
    run.font.size = Pt(7.5)
    run.font.bold = True

    c_offset = 1
    for l_key, _ in levels_def:
        m = table_data.get(reg, {}).get(l_key, {}).get('Maths', {'fee': 0, 'n': 0})
        e = table_data.get(reg, {}).get(l_key, {}).get('English', {'fee': 0, 'n': 0})

        for item in [m, e]:
            c_val = t3.cell(r_idx, c_offset)
            set_cell_background(c_val, bg)
            set_cell_margins(c_val, top=50, bottom=50, left=40, right=40)
            p = c_val.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            txt = f"₹{item['fee']} (n={item['n']})" if item['n'] > 0 else "—"
            run = p.add_run(txt)
            run.font.name = "Arial"
            run.font.size = Pt(7)
            if is_nat: run.font.bold = True
            c_offset += 1

p_note = doc.add_paragraph()
p_note.paragraph_format.space_before = Pt(4)
p_note.paragraph_format.space_after = Pt(12)
r = p_note.add_run("Notes: Values show mean hourly price (₹/hr) and sample size (n) for tutors offering each subject at that educational tier.")
r.font.name = "Arial"
r.font.size = Pt(7.5)
r.font.italic = True
r.font.color.rgb = COLOR_SUBTITLE

# Cross-Tabulation Matrix
h3 = doc.add_heading(level=2)
r = h3.add_run("Cross-Tabulation Matrix: Mean Fees (₹/hr) Across All 5 Core Subjects & Stages")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

matrix_subjs = [
    ('Maths', 'Mathematics'),
    ('English', 'English Language & Literature'),
    ('Sciences', 'Natural Sciences (Physics, Chem, Bio)'),
    ('Coding', 'Computer Science & Programming'),
    ('Languages', 'Other Languages (Hindi, Regional)')
]

t_mat = doc.add_table(rows=len(matrix_subjs) + 1, cols=7)
t_mat.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t_mat, color="CBD5E1", sz="4")

mat_headers = ["Subject Discipline", "Primary (Class 1-8)", "Secondary (Class 9-10)", "Senior Sec (11-12)", "Higher Ed (College)", "Combined Mean", "Total N"]
mat_widths = [Inches(2.2), Inches(1.0), Inches(1.0), Inches(1.0), Inches(1.0), Inches(1.0), Inches(0.8)]

for c_i, h_text in enumerate(mat_headers):
    c = t_mat.cell(0, c_i)
    c.width = mat_widths[c_i]
    set_cell_background(c, "0F172A")
    set_cell_margins(c, top=70, bottom=50, left=50, right=50)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_i >= 1 else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(h_text)
    run.font.name = "Arial"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

for r_i, (s_key, s_label) in enumerate(matrix_subjs, start=1):
    tot_n = 0
    tot_sum = 0
    row_cells = [s_label]

    for l_key, _ in levels_def:
        item = table_data.get('Nationwide', {}).get(l_key, {}).get(s_key, {'fee': 0, 'n': 0})
        if item['n'] > 0:
            tot_n += item['n']
            tot_sum += item['fee'] * item['n']
            row_cells.append(f"₹{item['fee']} (n={item['n']:,})")
        else:
            row_cells.append("—")

    comb_mean = round(tot_sum / tot_n) if tot_n else 0
    row_cells.append(f"₹{comb_mean:,}/hr")
    row_cells.append(f"{tot_n:,}")

    bg = "F8FAFC" if r_i % 2 == 0 else "FFFFFF"
    for c_i, val in enumerate(row_cells):
        c = t_mat.cell(r_i, c_i)
        c.width = mat_widths[c_i]
        set_cell_background(c, bg)
        set_cell_margins(c, top=55, bottom=55, left=50, right=50)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_i >= 1 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(val)
        run.font.name = "Arial"
        run.font.size = Pt(7.5)
        if c_i in (0, 5): run.font.bold = True

p_note = doc.add_paragraph()
p_note.paragraph_format.space_before = Pt(4)
p_note.paragraph_format.space_after = Pt(14)
r = p_note.add_run("Key Finding: Computer Science & Coding commands the highest tuition fee at all levels (₹947 to ₹1,222/hr, combined ₹1,111/hr), followed by Mathematics (₹948/hr) and Natural Sciences (₹879/hr).")
r.font.name = "Arial"
r.font.size = Pt(8)
r.font.bold = True
r.font.color.rgb = COLOR_TITLE

# =========================================================================
# SECTION 4: ACADEMIC FIGURES (FIGURES 1 & 2)
# =========================================================================
doc.add_page_break()
h1 = doc.add_heading(level=1)
r = h1.add_run("4. Academic Figures: Subjects & Teaching Experience")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p = doc.add_paragraph(
    "Visualizations constructed in accordance with empirical specifications from reference tutoring literature. "
    "Figure 1 depicts overall subject volume across broad clusters and top disciplines. "
    "Figure 2 illustrates educator experience structure (serving as proxy for age distribution)."
)
p.paragraph_format.space_after = Pt(12)

# Figure 1: Macro Disciplines
doc.add_picture(fig1_macro_path, width=Inches(6.4))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(8)
r = cap_p.add_run("Figure 1: Subjects Offered for Private Tuition in India (Macro Academic Clusters, N = 403,353).")
r.font.name = "Arial"
r.font.size = Pt(9)
r.font.bold = True

p_desc = doc.add_paragraph(
    "Analytical Commentary: Academic curriculum subjects and STEM represent over 55.1% of all tutoring offerings (222,113 courses), "
    "underscoring the strong school curriculum doubt-clearing and board-prep focus of the Indian market. "
    "Modern languages follow at 18.1% (72,831 offerings), while professional coding commands 11.4% (46,025 offerings)."
)
p_desc.paragraph_format.space_after = Pt(14)

# Figure 1: Top 12 Specific Subjects
doc.add_picture(fig1_top_path, width=Inches(6.4))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(8)
r = cap_p.add_run("Figure 1B: Top 12 Specific Individual Tutoring Disciplines by Offerings Volume.")
r.font.name = "Arial"
r.font.size = Pt(9)
r.font.bold = True

p_desc = doc.add_paragraph(
    "Analytical Commentary: Mathematics is the single largest individual subject with 59,649 offerings (14.8% market share), "
    "nearly double English (32,246 offerings, 8.0%) and General Science (29,616 offerings, 7.3%). "
    "Physics (21,951) and Biology (18,435) also demonstrate substantial dedicated demand."
)
p_desc.paragraph_format.space_after = Pt(14)

# Figure 2: Experience Distribution Curve
doc.add_page_break()
doc.add_picture(fig2_exp_path, width=Inches(6.4))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(8)
r = cap_p.add_run("Figure 2: Experience Structure of Private Tutors (Teaching Experience in Years).")
r.font.name = "Arial"
r.font.size = Pt(9)
r.font.bold = True

p_desc = doc.add_paragraph(
    "Analytical Commentary: The teaching experience distribution shows sharp right-skewness, peaking early at 1–2 years of experience "
    "(16,689 tutors at 1 year; 15,411 tutors at 2 years). "
    "Median experience is 3.0 years, mean experience is 4.6 years, and 47.3% of tutors possess 3 or fewer years of experience. "
    "Local spikes occur at round professional milestones (5, 10, 15, 20, 25 years), reflecting career self-reporting patterns."
)
p_desc.paragraph_format.space_after = Pt(14)

# Figure 2: Gender Experience Comparison
doc.add_picture(fig2_gender_path, width=Inches(6.4))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(8)
r = cap_p.add_run("Figure 2B: Experience Distribution Disaggregated by Gender (Women vs Men).")
r.font.name = "Arial"
r.font.size = Pt(9)
r.font.bold = True

p_desc = doc.add_paragraph(
    "Analytical Commentary: Women maintain higher participation rates across early and mid-career stages (years 0 through 10), "
    "while male tutors display slightly higher persistence in late-career brackets (15+ years), "
    "consistent with broader workforce gender dynamics."
)
p_desc.paragraph_format.space_after = Pt(14)

# =========================================================================
# SECTION 5: CORE MARKET STRUCTURE CHARTS
# =========================================================================
doc.add_page_break()
h1 = doc.add_heading(level=1)
r = h1.add_run("5. Core Market Structure & Pedagogical Profiles")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p = doc.add_paragraph(
    "Comprehensive distribution charts covering educator credentials, pedagogical seniority tiers, "
    "competitive examination coaching frequency, and target learning stages across 107,723 tutors."
)
p.paragraph_format.space_after = Pt(12)

# Grid of charts
doc.add_picture(chart_levels_path, width=Inches(6.2))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(10)
r = cap_p.add_run("Chart 1: Target Educational Levels (Senior Secondary 11-12 Boards & JEE/NEET form the primary market share).")
r.font.name = "Arial"
r.font.size = Pt(8.5)
r.font.bold = True

doc.add_picture(chart_quals_path, width=Inches(5.6))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(10)
r = cap_p.add_run("Chart 2: Highest Academic Qualifications (Post-graduates & Master's degree holders represent over 43% of tutors).")
r.font.name = "Arial"
r.font.size = Pt(8.5)
r.font.bold = True

doc.add_page_break()
doc.add_picture(chart_fees_path, width=Inches(6.2))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(10)
r = cap_p.add_run("Chart 3: Average Hourly Fee Across Subject Specializations (Coding and STEM command premium pricing).")
r.font.name = "Arial"
r.font.size = Pt(8.5)
r.font.bold = True

doc.add_picture(chart_exams_path, width=Inches(6.2))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(10)
r = cap_p.add_run("Chart 4: Competitive Exam Coaching Frequency (JEE, NEET, IELTS, and SAT lead test preparation demand).")
r.font.name = "Arial"
r.font.size = Pt(8.5)
r.font.bold = True

doc.add_picture(chart_regions_path, width=Inches(6.2))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(10)
r = cap_p.add_run("Chart 5: Geographic Supply Disaggregated by Zonal Region (Southern and Northern zones supply over 65% of tutors).")
r.font.name = "Arial"
r.font.size = Pt(8.5)
r.font.bold = True

# =========================================================================
# SECTION 6: PROFESSIONAL LANDSCAPE — NON-TEACHING & SIDE OCCUPATIONS ANALYSIS
# =========================================================================
doc.add_page_break()
h1 = doc.add_heading(level=1)
r = h1.add_run("6. Professional Landscape — Non-Teaching & Side Occupations Analysis")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p = doc.add_paragraph(
    "A defining structural feature of India's private tutoring market is its function as a vibrant gig-economy and moonlighting "
    "platform for highly qualified domain practitioners whose primary career lies outside traditional school education. "
    "Across the dataset of 107,723 educators, detailed parsing of biographical narratives, professional hooklines, and credentials "
    f"identified {total_side_cohort:,} tutors ({total_side_cohort*100.0/107723:.1f}% of the national supply) who actively practice a non-teaching primary or side profession. "
    "This section analyzes the occupational diversity, pricing power differentials, online delivery readiness, and pedagogical "
    "specializations across these diverse professional cohorts."
)
p.paragraph_format.space_after = Pt(12)

# Figure 6A: Distribution Chart
doc.add_picture(fig6a_path, width=Inches(6.4))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(8)
r = cap_p.add_run(f"Figure 6A: Professional Landscape — Distribution of Non-Teaching Side Occupations (N = {total_side_cohort:,}).")
r.font.name = "Arial"
r.font.size = Pt(9)
r.font.bold = True

p_desc = doc.add_paragraph(
    "Analytical Commentary: University and college students along with research scholars represent the single largest cohort "
    "(26,063 tutors, 71.4% of side practitioners), leveraging peer tutoring for supplementary income while pursuing higher degrees. "
    "Software engineers and IT professionals form the second-largest professional group (4,267 tutors, 11.7%), followed by "
    "medical and healthcare professionals (1,862 tutors, 5.1%), and corporate finance practitioners including Chartered Accountants "
    "(1,767 tutors, 4.8%). Specialized freelance consultants, core industrial engineers, civil service aspirants, and homemakers "
    "account for the remainder of this multi-disciplinary teaching supply."
)
p_desc.paragraph_format.space_after = Pt(14)

# Figure 6B: Pricing Benchmark
doc.add_picture(fig6b_path, width=Inches(6.4))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(8)
r = cap_p.add_run("Figure 6B: Hourly Tuition Pricing Benchmark by Side Profession (Mean & Median INR / Hour).")
r.font.name = "Arial"
r.font.size = Pt(9)
r.font.bold = True

p_desc = doc.add_paragraph(
    "Analytical Commentary: Hourly fees demonstrate sharp occupational stratification. Independent freelancers and startup founders "
    "command the highest hourly average rate (₹1,183/hr, median ₹950/hr), closely followed by Corporate, Finance & Chartered Accountants "
    "(₹1,099/hr, median ₹800/hr) and Software Engineers & IT Professionals (₹1,051/hr, median ₹750/hr). In contrast, University and "
    "college students provide accessible high-volume tuition at an average rate of ₹767/hr (median ₹550/hr), while home-based homemakers "
    "average ₹582/hr (median ₹400/hr). This demonstrates that industry practitioners monetizing niche technical expertise command "
    "up to an 80% pricing premium over entry-level peer educators."
)
p_desc.paragraph_format.space_after = Pt(14)

# Figure 6C: Online Adoption
doc.add_picture(fig6c_path, width=Inches(6.4))
cap_p = doc.add_paragraph()
cap_p.paragraph_format.space_after = Pt(8)
r = cap_p.add_run("Figure 6C: Online Tutoring Delivery Share Across Non-Teaching Side Professions (%).")
r.font.name = "Arial"
r.font.size = Pt(9)
r.font.bold = True

p_desc = doc.add_paragraph(
    "Analytical Commentary: Side-profession tutors demonstrate overwhelming digital readiness, with online delivery shares exceeding "
    "96% across all categories. Healthcare professionals (98.7%) and software engineers (98.5%) exhibit the highest online teaching "
    "adoption, utilizing virtual classrooms to overcome geographic constraints while balancing their full-time primary careers."
)
p_desc.paragraph_format.space_after = Pt(14)

# Table 4: Structural Matrix
h2 = doc.add_heading(level=2)
r = h2.add_run("Table 4: Comprehensive Structural Matrix of Non-Teaching & Side Occupations")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p = doc.add_paragraph(
    "Granular benchmark detailing headcount distribution, cohort share, hourly fee metrics (mean & median), "
    "online adoption, and core curricular specializations for all 9 identified non-teaching occupational groups."
)
p.paragraph_format.space_after = Pt(8)

t4_rows = len(agg_side) + 2
t4 = doc.add_table(rows=t4_rows, cols=7)
t4.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t4, color="CBD5E1", sz="4")

headers_t4 = ["Side Occupation Category", "Tutors (N)", "Share (%)", "Avg Fee", "Median Fee", "Online (%)", "Core Specializations"]
widths_t4 = [Inches(1.85), Inches(0.70), Inches(0.60), Inches(0.75), Inches(0.75), Inches(0.65), Inches(1.70)]

for col_idx, h_text in enumerate(headers_t4):
    c = t4.cell(0, col_idx)
    c.width = widths_t4[col_idx]
    set_cell_background(c, "0F172A")
    set_cell_margins(c, top=80, bottom=60, left=50, right=50)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx in [1, 2, 3, 4, 5] else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(h_text)
    run.font.name = "Arial"
    run.font.size = Pt(7.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

agg_sorted_desc = agg_side.sort_values(by='count', ascending=False).reset_index(drop=True)

for row_idx, r_data in agg_sorted_desc.iterrows():
    table_row_idx = row_idx + 1
    bg_color = "FFFFFF" if row_idx % 2 == 0 else "F8FAFC"
    
    cat_str = str(r_data['category'])
    cnt_str = f"{int(r_data['count']):,}"
    pct_str = f"{r_data['pct']:.1f}%"
    avg_f_str = f"₹{int(round(r_data['avg_fee'])):,}"
    med_f_str = f"₹{int(round(r_data['median_fee'])):,}"
    onl_str = f"{r_data['online_pct']:.1f}%"
    subj_str = str(r_data['subjects'])

    row_vals = [cat_str, cnt_str, pct_str, avg_f_str, med_f_str, onl_str, subj_str]

    for col_idx, val_text in enumerate(row_vals):
        c = t4.cell(table_row_idx, col_idx)
        c.width = widths_t4[col_idx]
        set_cell_background(c, bg_color)
        set_cell_margins(c, top=60, bottom=60, left=50, right=50)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx in [1, 2, 3, 4, 5] else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(val_text)
        run.font.name = "Arial"
        run.font.size = Pt(7.5)
        if col_idx == 0:
            run.font.bold = True
            run.font.color.rgb = COLOR_TITLE
        elif col_idx in [3, 4]:
            run.font.bold = True
            run.font.color.rgb = RGBColor(6, 95, 70)
        else:
            run.font.color.rgb = COLOR_SUBTITLE

tot_row_idx = t4_rows - 1
tot_bg = "E2E8F0"
tot_vals = [
    "Total Identified Cohort", 
    f"{total_side_cohort:,}", 
    "100.0%", 
    f"₹{int(round(df_side['fee'].mean())):,}", 
    f"₹{int(round(df_side['fee'].median())):,}", 
    f"{df_side['online'].mean()*100:.1f}%", 
    "Cross-Disciplinary Industry & Academic Base"
]

for col_idx, val_text in enumerate(tot_vals):
    c = t4.cell(tot_row_idx, col_idx)
    c.width = widths_t4[col_idx]
    set_cell_background(c, tot_bg)
    set_cell_margins(c, top=70, bottom=70, left=50, right=50)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col_idx in [1, 2, 3, 4, 5] else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(val_text)
    run.font.name = "Arial"
    run.font.size = Pt(7.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_TITLE

p_note = doc.add_paragraph()
p_note.paragraph_format.space_before = Pt(6)
p_note.paragraph_format.space_after = Pt(14)
r = p_note.add_run(f"Note: Identified through automated computational extraction from self-reported hooklines, bios, and qualification records. "
                   f"Hourly fees exclude outliers outside ₹50–₹5,000/hr range. Represents {total_side_cohort*100.0/107723:.1f}% of total registered tutors (N = 107,723).")
r.font.name = "Arial"
r.font.size = Pt(7)
r.font.italic = True
r.font.color.rgb = COLOR_SUBTITLE

# =========================================================================
# SECTION 7: METHODOLOGY & SCHEMA NOTES
# =========================================================================
doc.add_page_break()
h1 = doc.add_heading(level=1)
r = h1.add_run("7. Methodology, Data Cleaning & Relational Architecture")
r.font.name = "Arial"
r.font.color.rgb = COLOR_TITLE

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(6)
r = p.add_run("Relational Bridge Normalization: ")
r.font.bold = True
p.add_run(
    "To resolve multi-subject profile contamination where tutors teach up to 15 subjects across multiple levels, "
    "the dataset was normalized into a 3-sheet relational schema: 'Cleaned_Tutors' (107,723 unique tutor master rows), "
    "'Tutor_Subjects_Bridge' (403,353 standardized subject-grade-exam linkages), and 'Data_Dictionary' defining all 37 attributes."
)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(6)
r = p.add_run("Experience Role & Seniority Tier Imputation: ")
r.font.bold = True
p.add_run(
    "Experience roles (School Educator, Faculty, Online Educator, Student Tutor) were classified by parsing biographical text "
    "and professional hooklines, eliminating missing values (0% 'Not Specified'). Seniority tiers (Entry-Level, Mid-Level, "
    "Senior Educator, Star Faculty) were inferred from years of experience, degrees, and hourly fee levels."
)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(14)
r = p.add_run("Price Outlier Bounds: ")
r.font.bold = True
p.add_run(
    "Hourly fees were validated within ₹50 to ₹5,000/hr, excluding corrupted entry fees while preserving genuine premium coaching tiers. "
    "All summary tables compute sample-weighted means ($n$) to prevent aggregation bias."
)

# Document Footer
p_foot = doc.add_paragraph()
p_foot.paragraph_format.space_before = Pt(20)
r = p_foot.add_run("BMP Tutors Research Working Paper • All Visuals & Tables Ready for Publication • Generated September 2026")
r.font.name = "Arial"
r.font.size = Pt(8)
r.font.italic = True
r.font.color.rgb = COLOR_SUBTITLE

print(f"Saving final DOCX to {DOCX_PATH}...")
doc.save(DOCX_PATH)
file_size = os.path.getsize(DOCX_PATH)
print(f"DOCX created successfully! File size: {file_size:,} bytes ({file_size / (1024*1024):.2f} MB)")
conn.close()
