#!/usr/bin/env python3
"""
Build Google Docs & Sheets Companion Workbook for BMP Tutors.
Enables 100% editable, live-linked charts in Google Docs via Google Sheets.
"""

import openpyxl
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_companion_workbook():
    wb = openpyxl.Workbook()

    # Style presets
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    subtotal_fill = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
    subtotal_font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    total_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # -------------------------------------------------------------
    # Sheet 1: README & Guide
    # -------------------------------------------------------------
    ws_readme = wb.active
    ws_readme.title = "README_Google_Docs_Guide"
    ws_readme.views.sheetView[0].showGridLines = True

    guide_lines = [
        ("HOW TO USE THESE CHARTS IN GOOGLE DOCS (2-CLICK LINKING)", 16, True, "1E3A8A"),
        ("", 10, False, None),
        ("Why did Word (.docx) charts turn into images in Google Docs?", 12, True, "0F172A"),
        ("• Microsoft Word uses Office OpenXML DrawingML (<c:chart>) with embedded Excel workbooks.", 10, False, "334155"),
        ("• Google Docs DOES NOT support OpenXML charts and automatically converts them to static images.", 10, False, "334155"),
        ("• Instead, Google Docs exclusively supports native editable charts linked via GOOGLE SHEETS.", 10, False, "334155"),
        ("", 10, False, None),
        ("How to get 100% Fully Editable Charts in Google Docs:", 12, True, "0F172A"),
        ("Step 1: Upload this file ('BMP_Tutors_Google_Docs_Charts_Companion.xlsx') to Google Drive.", 10, False, "334155"),
        ("Step 2: Right-click and choose 'Open with > Google Sheets'.", 10, False, "334155"),
        ("Step 3: Open your report in Google Docs.", 10, False, "334155"),
        ("Step 4: In Google Docs menu, click:  Insert  >  Chart  >  From Sheets.", 10, True, "1D4ED8"),
        ("Step 5: Select this spreadsheet, choose the desired Chart tab, and click 'Import' with 'Link to spreadsheet' checked.", 10, False, "334155"),
        ("Step 6: Done! You can now change values, titles, or colors anytime in Google Sheets and click 'Update' in Google Docs.", 10, True, "047857"),
        ("", 10, False, None),
        ("Alternative (Offline Microsoft Word):", 12, True, "0F172A"),
        ("• If you open 'BMP_Tutors_Visual_Analytics_Report.docx' directly in Microsoft Word (Desktop or free Word Online at office.com), all charts are ALREADY 100% native and editable without any extra steps.", 10, False, "334155"),
        ("• Furthermore, all Tables in the DOCX are ALREADY 100% editable inside Google Docs without converting anything.", 10, False, "334155"),
    ]

    for r_idx, (text, sz, is_bold, clr) in enumerate(guide_lines, start=1):
        cell = ws_readme.cell(row=r_idx, column=1, value=text)
        cell.font = Font(name="Calibri", size=sz, bold=is_bold, color=clr or "000000")

    ws_readme.column_dimensions['A'].width = 110

    # -------------------------------------------------------------
    # Helper to style a table
    # -------------------------------------------------------------
    def format_table(ws, start_row, num_cols, align_right_cols=[]):
        for col_idx in range(1, num_cols + 1):
            cell = ws.cell(row=start_row, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_right if col_idx in align_right_cols else align_left
            cell.border = thin_border

        for col in ws.columns:
            max_len = max(len(str(c.value or '')) for c in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # -------------------------------------------------------------
    # Sheet 2: Figure 1 - Macro Academic Clusters
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="Fig1_Macro_Subjects")
    ws1.views.sheetView[0].showGridLines = True
    ws1.append(["Subject Academic Cluster", "Offered Courses", "Market Share (%)"])
    data_fig1 = [
        ("School Curriculum & STEM", 184288, 45.7),
        ("Languages & Linguistics", 58248, 14.4),
        ("Computer Science & Coding", 37778, 9.4),
        ("Commerce & Management", 16311, 4.0),
        ("Arts & Humanities", 15991, 4.0),
        ("Personal Development & ECA", 15078, 3.7)
    ]
    for row in data_fig1:
        ws1.append(list(row))
    format_table(ws1, 1, 3, align_right_cols=[2, 3])

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Figure 1: Subjects Offered for Private Tuition in India"
    chart1.y_axis.title = "Number of Offered Courses"
    chart1.x_axis.title = "Macro Academic Cluster"
    data_ref1 = Reference(ws1, min_col=2, min_row=1, max_row=7)
    cats_ref1 = Reference(ws1, min_col=1, min_row=2, max_row=7)
    chart1.add_data(data_ref1, titles_from_data=True)
    chart1.set_categories(cats_ref1)
    chart1.legend = None
    chart1.width = 16
    chart1.height = 9.5
    ws1.add_chart(chart1, "E2")

    # -------------------------------------------------------------
    # Sheet 3: Figure 1B - Top 12 Specific Disciplines
    # -------------------------------------------------------------
    ws1b = wb.create_sheet(title="Fig1B_Top_Disciplines")
    ws1b.views.sheetView[0].showGridLines = True
    ws1b.append(["Tutoring Subject Discipline", "Offered Courses", "Share of Offerings (%)"])
    data_fig1b = [
        ("Mathematics", 62799, 15.6),
        ("English", 36696, 9.1),
        ("Physics", 29813, 7.4),
        ("Chemistry", 24393, 6.0),
        ("Biology & Life Sciences", 16867, 4.2),
        ("Computer Science & Coding", 13735, 3.4),
        ("Hindi", 12151, 3.0),
        ("Commerce & Accounts", 11210, 2.8),
        ("Economics", 8604, 2.1),
        ("Science (General K-10)", 7842, 1.9),
        ("Social Studies & History", 6520, 1.6),
        ("Spoken English & Communication", 5110, 1.3)
    ]
    for row in data_fig1b:
        ws1b.append(list(row))
    format_table(ws1b, 1, 3, align_right_cols=[2, 3])

    chart1b = BarChart()
    chart1b.type = "bar"
    chart1b.style = 10
    chart1b.title = "Figure 1B: Top 12 Individual Tutoring Disciplines"
    chart1b.x_axis.title = "Number of Offered Courses"
    chart1b.y_axis.title = "Subject Discipline"
    data_ref1b = Reference(ws1b, min_col=2, min_row=1, max_row=13)
    cats_ref1b = Reference(ws1b, min_col=1, min_row=2, max_row=13)
    chart1b.add_data(data_ref1b, titles_from_data=True)
    chart1b.set_categories(cats_ref1b)
    chart1b.legend = None
    chart1b.width = 16
    chart1b.height = 11
    ws1b.add_chart(chart1b, "E2")

    # -------------------------------------------------------------
    # Sheet 4: Figure 2A - National Gender Distribution
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Fig2A_Gender_Overview")
    ws2.views.sheetView[0].showGridLines = True
    ws2.append(["Gender Cohort", "Tutors (N)", "National Share (%)"])
    data_fig2 = [
        ("Female Educators", 62466, 58.0),
        ("Male Educators", 43964, 40.8),
        ("Not Specified", 1286, 1.2)
    ]
    for row in data_fig2:
        ws2.append(list(row))
    format_table(ws2, 1, 3, align_right_cols=[2, 3])

    chart2 = PieChart()
    chart2.title = "Figure 2A: National Gender Breakdown (N = 107,716)"
    data_ref2 = Reference(ws2, min_col=2, min_row=1, max_row=4)
    cats_ref2 = Reference(ws2, min_col=1, min_row=2, max_row=4)
    chart2.add_data(data_ref2, titles_from_data=True)
    chart2.set_categories(cats_ref2)
    chart2.width = 14
    chart2.height = 8.5
    ws2.add_chart(chart2, "E2")

    # -------------------------------------------------------------
    # Sheet 5: Figure 3E - Multi-Subject Breadth
    # -------------------------------------------------------------
    ws3e = wb.create_sheet(title="Fig3E_Subject_Breadth")
    ws3e.views.sheetView[0].showGridLines = True
    ws3e.append(["Number of Subjects Taught", "Tutors Offering", "Share (%)"])
    data_fig3e = [
        ("1 Subject Only", 33998, 52.2),
        ("2 Subjects", 19818, 30.5),
        ("3 Subjects", 9631, 14.8),
        ("4 Subjects", 1451, 2.2),
        ("All 5 Top Subjects", 179, 0.3)
    ]
    for row in data_fig3e:
        ws3e.append(list(row))
    format_table(ws3e, 1, 3, align_right_cols=[2, 3])

    chart3e = BarChart()
    chart3e.type = "col"
    chart3e.style = 10
    chart3e.title = "Figure 3E: Subject Teaching Breadth Across Top 5 Subjects"
    chart3e.y_axis.title = "Number of Tutors"
    chart3e.x_axis.title = "Teaching Breadth"
    data_ref3e = Reference(ws3e, min_col=2, min_row=1, max_row=6)
    cats_ref3e = Reference(ws3e, min_col=1, min_row=2, max_row=6)
    chart3e.add_data(data_ref3e, titles_from_data=True)
    chart3e.set_categories(cats_ref3e)
    chart3e.legend = None
    chart3e.width = 15
    chart3e.height = 9.5
    ws3e.add_chart(chart3e, "E2")

    # -------------------------------------------------------------
    # Sheet 6: Figure 4 - Tutor Experience Distribution
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="Fig4_Experience_Dist")
    ws4.views.sheetView[0].showGridLines = True
    ws4.append(["Experience Tier", "Years Range", "Tutors (N)", "Share (%)", "Avg Fee (₹/hr)"])
    data_fig4 = [
        ("Beginners", "0–1 Years", 26865, 24.9, 620),
        ("Early Career", "2–3 Years", 33692, 31.3, 710),
        ("Mid Career", "4–5 Years", 24947, 23.2, 850),
        ("Experienced", "6–10 Years", 11107, 10.3, 1050),
        ("Senior Mentors", "11–15 Years", 5015, 4.7, 1280),
        ("Veterans", "16+ Years", 6090, 5.7, 1450)
    ]
    for row in data_fig4:
        ws4.append(list(row))
    format_table(ws4, 1, 5, align_right_cols=[3, 4, 5])

    chart4 = BarChart()
    chart4.type = "col"
    chart4.style = 10
    chart4.title = "Figure 4: Teaching Experience Distribution Across India"
    chart4.y_axis.title = "Number of Tutors"
    chart4.x_axis.title = "Experience Bracket"
    data_ref4 = Reference(ws4, min_col=3, min_row=1, max_row=7)
    cats_ref4 = Reference(ws4, min_col=2, min_row=2, max_row=7)
    chart4.add_data(data_ref4, titles_from_data=True)
    chart4.set_categories(cats_ref4)
    chart4.legend = None
    chart4.width = 15
    chart4.height = 9.5
    ws4.add_chart(chart4, "G2")

    # -------------------------------------------------------------
    # Sheet 7: Figure 5 - State Hourly Rates
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title="Fig5_State_Fees")
    ws5.views.sheetView[0].showGridLines = True
    ws5.append(["State / Union Territory", "Average Hourly Fee (₹/hr)", "Tutors (N)"])
    data_fig5 = [
        ("Karnataka", 650, 20387),
        ("Maharashtra", 620, 16390),
        ("Delhi NCR", 640, 10731),
        ("Tamil Nadu", 580, 10452),
        ("Telangana", 610, 8435),
        ("Uttar Pradesh", 520, 8107),
        ("West Bengal", 540, 6808),
        ("Kerala", 560, 4843),
        ("Gujarat", 590, 3721),
        ("Haryana", 600, 3514),
        ("Rajasthan", 510, 2980),
        ("Madhya Pradesh", 500, 2450)
    ]
    for row in data_fig5:
        ws5.append(list(row))
    format_table(ws5, 1, 3, align_right_cols=[2, 3])

    chart5 = BarChart()
    chart5.type = "col"
    chart5.style = 10
    chart5.title = "Figure 5: Average Hourly Tuition Rates in Major States"
    chart5.y_axis.title = "Average Hourly Fee (₹/hr)"
    chart5.x_axis.title = "State / Territory"
    data_ref5 = Reference(ws5, min_col=2, min_row=1, max_row=13)
    cats_ref5 = Reference(ws5, min_col=1, min_row=2, max_row=13)
    chart5.add_data(data_ref5, titles_from_data=True)
    chart5.set_categories(cats_ref5)
    chart5.legend = None
    chart5.width = 16
    chart5.height = 9.5
    ws5.add_chart(chart5, "E2")

    # -------------------------------------------------------------
    # Sheet 8: Figure 6 - Primary Occupations
    # -------------------------------------------------------------
    ws6 = wb.create_sheet(title="Fig6_Occupations")
    ws6.views.sheetView[0].showGridLines = True
    ws6.append(["Occupation / Background", "Tutors (N)", "Share (%)", "Avg Fee (₹/hr)"])
    data_fig6 = [
        ("Online Educator", 29349, 27.2, 943),
        ("College Student / Scholar", 26061, 24.2, 767),
        ("Independent Tutor", 16188, 15.0, 722),
        ("College Faculty", 6672, 6.2, 1044),
        ("Early-Career Teacher", 5821, 5.4, 638),
        ("Software Engineer / IT", 4267, 4.0, 1051),
        ("Senior Private Educator", 4080, 3.8, 798),
        ("School Teacher", 3704, 3.4, 819)
    ]
    for row in data_fig6:
        ws6.append(list(row))
    format_table(ws6, 1, 4, align_right_cols=[2, 3, 4])

    chart6 = BarChart()
    chart6.type = "bar"
    chart6.style = 10
    chart6.title = "Figure 6: Primary Occupations & Backgrounds of Tutors"
    chart6.x_axis.title = "Number of Tutors"
    chart6.y_axis.title = "Primary Occupation"
    data_ref6 = Reference(ws6, min_col=2, min_row=1, max_row=9)
    cats_ref6 = Reference(ws6, min_col=1, min_row=2, max_row=9)
    chart6.add_data(data_ref6, titles_from_data=True)
    chart6.set_categories(cats_ref6)
    chart6.legend = None
    chart6.width = 16
    chart6.height = 10
    ws6.add_chart(chart6, "F2")

    # -------------------------------------------------------------
    # Sheet 9: Table 2.8 - Fees by Student Level
    # -------------------------------------------------------------
    ws28 = wb.create_sheet(title="Table2_8_Fees_By_Level")
    ws28.views.sheetView[0].showGridLines = True
    ws28.append(["Subject Domain", "Elementary (Classes 1–5)", "Upper Primary (Classes 5–8)", "Secondary (Classes 9–10)", "Senior Sec (Classes 11–12)", "College / University", "Competitive Exams / Pro"])
    data_t28 = [
        ("Mathematics", 450, 550, 700, 850, 1000, 1400),
        ("Computer Science & Coding", 600, 700, 850, 950, 1200, 1800),
        ("Physics", 450, 550, 700, 900, 1050, 1500),
        ("English & Languages", 400, 500, 650, 750, 900, 1250),
        ("Chemistry & Biology", 450, 550, 700, 900, 1050, 1450),
        ("Commerce & Accounting", 450, 500, 650, 850, 1050, 1600)
    ]
    for row in data_t28:
        ws28.append(list(row))
    format_table(ws28, 1, 7, align_right_cols=[2, 3, 4, 5, 6, 7])

    # -------------------------------------------------------------
    # Sheet 10: Table 7 - Teaching vs. Non-Teaching Occupations
    # -------------------------------------------------------------
    ws7 = wb.create_sheet(title="Table7_Occupations_Matrix")
    ws7.views.sheetView[0].showGridLines = True
    ws7.append(["Professional Profile", "Career Category", "Tutors (N)", "Share (%)", "Avg Fee (₹)", "Median Fee (₹)", "Online (%)", "Key Specializations & Core Subjects"])

    # Part A
    ws7.append(["PART A: TEACHING & ACADEMIA (Full-Time Educators)", "Career Vocation", 70321, 65.3, 864, 650, 95.2, "Comprehensive School, College & Board Curricula"])
    teaching_rows = [
        ("Professional Online Educator", "Full-Time Online Educator", 29349, 27.2, 943, 725, 98.5, "K-12 STEM, English, Coding & Exam Prep"),
        ("Independent Private Tutor", "Full-Time Independent Tutor", 16188, 15.0, 722, 500, 95.4, "K-10 All Subjects, Maths, Science Foundations"),
        ("College / University Faculty", "Higher Education Faculty", 6672, 6.2, 1044, 750, 98.2, "Advanced Degree STEM, Calculus, University Physics"),
        ("Early-Career Educator", "Trainee / Junior Teacher", 5821, 5.4, 638, 450, 74.7, "Primary & Middle School Basics, Language Fluency"),
        ("Senior Private Educator", "Master Educator (10+ Yrs)", 4080, 3.8, 798, 600, 91.9, "Board Exam Derivations, Senior Physics & Maths"),
        ("School Educator (Teacher)", "Regular School Staff Teacher", 3704, 3.4, 819, 650, 97.6, "CBSE / ICSE School Syllabus, Chemistry & EVS"),
        ("Competitive Exam Specialist", "Entrance Coaching Specialist", 2413, 2.2, 1043, 750, 97.0, "IIT JEE Advanced, NEET, Olympiads, Quant"),
        ("Private / Home Tutor", "Neighborhood Home Tutor", 2094, 1.9, 682, 450, 88.7, "In-Person Mentoring, Middle School Homework Help"),
    ]
    for row in teaching_rows:
        ws7.append(list(row))

    # Part B
    ws7.append(["PART B: NON-TEACHING PROFESSIONS (Side-Gig Tutoring)", "Side-Career / Part-Time", 37395, 34.7, 849, 625, 96.7, "Applied Industry Skills, STEM & Entrance Mentoring"])
    non_teaching_rows = [
        ("University & College Scholar", "Enrolled University Scholar", 26061, 24.2, 767, 550, 96.1, "Mathematics, Physics, Chemistry, Peer Mentoring"),
        ("Software Engineer & IT Pro", "Working Tech Professional", 4267, 4.0, 1051, 750, 98.5, "Python, Java, Web Dev, DSA, AI/ML, Databases"),
        ("Medical & Healthcare Pro", "Doctor / MBBS / Clinician", 1862, 1.7, 928, 750, 98.7, "NEET UG Biology, Human Physiology, Anatomy"),
        ("Corporate Finance, CA & Banking", "Chartered Accountant / Banker", 1767, 1.6, 1099, 800, 97.7, "Accountancy, Financial Mgmt, CA Foundation, Tax"),
        ("Freelancer & Solopreneur", "Independent Consultant / Owner", 994, 0.9, 1183, 950, 98.2, "Spoken English, Digital Skills, Business Consulting"),
        ("Industry Consultant / Manager", "Corporate Consultant / HR", 833, 0.8, 1037, 750, 98.8, "SQL, Data Analytics, Soft Skills, Applied Maths"),
        ("Core Industrial Engineer", "Mech / Civil / Electrical Engg", 642, 0.6, 972, 750, 96.4, "Engineering Mechanics, AutoCAD, Applied Physics"),
        ("Creative Arts, Design & Media", "Designer / Writer / Media Artist", 423, 0.4, 906, 700, 98.1, "Creative Writing, English Fluency, Graphic Design"),
        ("Legal & Civil Services Aspirant", "Lawyer / UPSC Aspirant", 407, 0.4, 809, 600, 96.1, "Political Science, Constitutional Law, History"),
        ("Homemaker & Home Educator", "Educated Homemaker", 100, 0.1, 582, 400, 96.0, "Early Childhood, Classes 1–5 All Subjects, Phonics"),
        ("Student Tutor / Peer Mentor", "Junior / High School Student", 39, 0.04, 601, 438, 94.9, "Peer-to-Peer Foundations, Junior Science & Maths"),
    ]
    for row in non_teaching_rows:
        ws7.append(list(row))

    # Total
    ws7.append(["NATIONAL TOTAL (All Verified Tutors)", "All Verified Profiles", 107723, 100.0, 859, 650, 95.7, "Complete Pan-India Tutoring Ecosystem"])
    format_table(ws7, 1, 8, align_right_cols=[3, 4, 5, 6, 7])

    # Save
    out_path = "/home/nilesh7757/BMPTUTORS1/BMP_Tutors_Google_Docs_Charts_Companion.xlsx"
    wb.save(out_path)
    print(f"Companion workbook successfully saved to {out_path}!")

if __name__ == "__main__":
    create_companion_workbook()
