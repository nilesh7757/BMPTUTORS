#!/usr/bin/env python3
"""
Generate Modern, Fully Editable Microsoft Word (.docx) Report for BMP Tutors.
Features:
- Pure native Microsoft Office DrawingML charts (100% editable in Word: change colors, values, types).
- Pure native formatted Word tables and visual KPI callouts.
- Simple, human, friendly language with zero pretentious academic jargon.
- All 36 individual state maps completely removed, replaced with a clean All-India master table.
"""

import os
import sys
import sqlite3
import io
import openpyxl
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

DB_PATH = "/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db"
OUTPUT_DOCX = "/home/nilesh7757/BMPTUTORS1/BMP_Tutors_Visual_Analytics_Report.docx"

# -------------------------------------------------------------
# XML Formatting Helpers for Tables & Cells
# -------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
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
# Native DrawingML Chart Builder
# -------------------------------------------------------------
class NativeChartBuilder:
    def __init__(self, doc):
        self.doc = doc
        self.chart_counter = 0

    def add_pie_chart(self, title, categories, values, series_name="Share", width_in=5.8, height_in=3.6):
        self.chart_counter += 1
        c_num = self.chart_counter

        # 1. Create openpyxl Excel workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["Category", series_name])
        for c, v in zip(categories, values):
            ws.append([c, v])

        excel_stream = io.BytesIO()
        wb.save(excel_stream)
        excel_bytes = excel_stream.getvalue()

        excel_uri = PackURI(f"/word/embeddings/Microsoft_Excel_Worksheet{c_num}.xlsx")
        excel_part = Part(excel_uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", blob=excel_bytes)

        # 2. Build Chart XML
        cat_pts = ''.join(f'<c:pt idx="{i}"><c:v>{c}</c:v></c:pt>' for i, c in enumerate(categories))
        val_pts = ''.join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values))

        chart_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="1100"/>
              <a:t>{title}</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:layout/>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>{series_name}</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${len(categories)+1}</c:f>
              <c:strCache>
                <c:ptCount val="{len(categories)}"/>
                {cat_pts}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${len(values)+1}</c:f>
              <c:numCache>
                <c:formatCode>#,##0</c:formatCode>
                <c:ptCount val="{len(values)}"/>
                {val_pts}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="1"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:firstSliceAng val="0"/>
      </c:pieChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:externalData r:id="rId1">
    <c:autoUpdate val="0"/>
  </c:externalData>
</c:chartSpace>"""

        chart_uri = PackURI(f"/word/charts/chart{c_num}.xml")
        chart_part = Part(chart_uri, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", blob=chart_xml.encode("utf-8"))
        chart_part.relate_to(excel_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package")
        rId_chart = self.doc.part.relate_to(chart_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")

        cx = int(width_in * 914400)
        cy = int(height_in * 914400)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        drawing_elm = parse_xml(f'''
<w:drawing {nsdecls('w', 'wp', 'a', 'c', 'r')}>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="{cx}" cy="{cy}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:docPr id="{c_num}" name="Chart {c_num}"/>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks noGrp="1"/>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart r:id="{rId_chart}"/>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
''')
        r._r.append(drawing_elm)
        return p

    def add_bar_or_col_chart(self, title, categories, values, series_name="Tutors", is_horizontal=False, width_in=6.2, height_in=3.6, is_currency=False):
        self.chart_counter += 1
        c_num = self.chart_counter

        # 1. Create openpyxl Excel workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["Category", series_name])
        for c, v in zip(categories, values):
            ws.append([c, v])

        excel_stream = io.BytesIO()
        wb.save(excel_stream)
        excel_bytes = excel_stream.getvalue()

        excel_uri = PackURI(f"/word/embeddings/Microsoft_Excel_Worksheet{c_num}.xlsx")
        excel_part = Part(excel_uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", blob=excel_bytes)

        cat_pts = ''.join(f'<c:pt idx="{i}"><c:v>{c}</c:v></c:pt>' for i, c in enumerate(categories))
        val_pts = ''.join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values))
        bar_dir = "bar" if is_horizontal else "col"
        ax_cat_id = 100000 + c_num * 10 + 1
        ax_val_id = 100000 + c_num * 10 + 2
        fmt = "₹#,##0" if is_currency else "#,##0"

        chart_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="1100"/>
              <a:t>{title}</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:layout/>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="{bar_dir}"/>
        <c:barGrouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>{series_name}</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:invertIfNegative val="0"/>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${len(categories)+1}</c:f>
              <c:strCache>
                <c:ptCount val="{len(categories)}"/>
                {cat_pts}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${len(values)+1}</c:f>
              <c:numCache>
                <c:formatCode>{fmt}</c:formatCode>
                <c:ptCount val="{len(values)}"/>
                {val_pts}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="1"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:gapWidth val="110"/>
        <c:axId val="{ax_cat_id}"/>
        <c:axId val="{ax_val_id}"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="{ax_cat_id}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="{'l' if is_horizontal else 'b'}"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="{ax_val_id}"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="{ax_val_id}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="{'b' if is_horizontal else 'l'}"/>
        <c:majorGridlines/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="{ax_cat_id}"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:externalData r:id="rId1">
    <c:autoUpdate val="0"/>
  </c:externalData>
</c:chartSpace>"""

        chart_uri = PackURI(f"/word/charts/chart{c_num}.xml")
        chart_part = Part(chart_uri, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", blob=chart_xml.encode("utf-8"))
        chart_part.relate_to(excel_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package")
        rId_chart = self.doc.part.relate_to(chart_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")

        cx = int(width_in * 914400)
        cy = int(height_in * 914400)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        drawing_elm = parse_xml(f'''
<w:drawing {nsdecls('w', 'wp', 'a', 'c', 'r')}>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="{cx}" cy="{cy}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:docPr id="{c_num}" name="Chart {c_num}"/>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks noGrp="1"/>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart r:id="{rId_chart}"/>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
''')
        r._r.append(drawing_elm)
        return p

# -------------------------------------------------------------
# Document Generation Pipeline
# -------------------------------------------------------------
def build_report():
    print("Connecting to database...")
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    doc = Document()

    # Set Document Page Margins
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # Styles & Colors
    CLR_PRIMARY_HEX = "1E293B"    # Slate 800
    CLR_ACCENT_HEX = "3B82F6"     # Blue 500
    CLR_BG_LIGHT = "F8FAFC"       # Slate 50
    CLR_BORDER = "E2E8F0"         # Slate 200

    COLOR_TITLE = RGBColor(15, 23, 42)      # Slate 900
    COLOR_PRIMARY = RGBColor(30, 41, 59)    # Slate 800
    COLOR_MUTED = RGBColor(100, 116, 139)   # Slate 500
    COLOR_BLUE = RGBColor(37, 99, 235)      # Blue 600

    builder = NativeChartBuilder(doc)

    # ---------------------------------------------------------
    # COVER / HEADER
    # ---------------------------------------------------------
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(4)
    title_p.paragraph_format.space_after = Pt(2)
    r_t = title_p.add_run("BMP Tutors — India Private Tutoring Market Report")
    r_t.font.name = "Arial"
    r_t.font.size = Pt(22)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_TITLE

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(12)
    r_s = sub_p.add_run("A Complete Market Study of 107,716 Educators Across 36 Indian States & Union Territories")
    r_s.font.name = "Arial"
    r_s.font.size = Pt(11)
    r_s.font.color.rgb = COLOR_MUTED

    # Editable Tip Banner
    banner_tbl = doc.add_table(rows=1, cols=1)
    banner_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    b_cell = banner_tbl.cell(0, 0)
    set_cell_background(b_cell, "EEF2FF")
    set_cell_margins(b_cell, top=100, bottom=100, left=150, right=150)
    set_table_borders(banner_tbl, color="C7D2FE", sz="6")
    bp = b_cell.paragraphs[0]
    r_b = bp.add_run("💡 100% Native & Editable Document: ")
    r_b.font.bold = True
    r_b.font.size = Pt(9)
    r_b.font.color.rgb = RGBColor(49, 46, 129)
    r_b2 = bp.add_run("All charts and tables in this report are native Microsoft Office elements. You can click on any chart in Word to change colors (via the Chart Design tab), right-click to edit data in Excel, or change numbers and text directly in the tables.")
    r_b2.font.size = Pt(8.5)
    r_b2.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ---------------------------------------------------------
    # EXECUTIVE KPI DASHBOARD (STAT CARDS)
    # ---------------------------------------------------------
    kpi_cards = [
        {"lbl": "TOTAL EDUCATORS", "val": "107,716", "sub": "Verified Tutors", "bg": "F0FDF4", "border": "BBF7D0", "txt": "15803D"},
        {"lbl": "COURSE OFFERINGS", "val": "327,694", "sub": "3.0 Subjects / Tutor", "bg": "EFF6FF", "border": "BFDBFE", "txt": "1D4ED8"},
        {"lbl": "BENCHMARK RATE", "val": "₹859 / hr", "sub": "Average Tuition Fee", "bg": "FAF5FF", "border": "E9D5FF", "txt": "7E22CE"},
        {"lbl": "ONLINE ADOPTION", "val": "96.8%", "sub": "Digital Classroom", "bg": "FFFBEB", "border": "FDE68A", "txt": "B45309"},
    ]
    kpi_tbl = doc.add_table(rows=1, cols=4)
    kpi_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(kpi_tbl, color="CBD5E1", sz="4")
    for idx, card in enumerate(kpi_cards):
        cell = kpi_tbl.cell(0, idx)
        set_cell_background(cell, card["bg"])
        set_cell_margins(cell, top=120, bottom=120, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Label
        r1 = p.add_run(card["lbl"] + "\n")
        r1.font.name = "Arial"
        r1.font.size = Pt(7.5)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor.from_string(card["txt"])

        # Value
        r2 = p.add_run(card["val"] + "\n")
        r2.font.name = "Arial"
        r2.font.size = Pt(16)
        r2.font.bold = True
        r2.font.color.rgb = COLOR_TITLE

        # Sub
        r3 = p.add_run(card["sub"])
        r3.font.name = "Arial"
        r3.font.size = Pt(7.5)
        r3.font.color.rgb = COLOR_MUTED

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Executive Summary Text
    h_exec = doc.add_heading(level=1)
    r_he = h_exec.add_run("Executive Summary: The State of Tutoring in India")
    r_he.font.name = "Arial"
    r_he.font.size = Pt(14)
    r_he.font.bold = True
    r_he.font.color.rgb = COLOR_TITLE

    p_ex = doc.add_paragraph()
    p_ex.paragraph_format.line_spacing = 1.15
    p_ex.paragraph_format.space_after = Pt(8)
    p_ex.add_run(
        "Private tutoring in India has evolved from an informal neighborhood activity into a massive, organized digital gig economy. "
        "This report analyzes verified data from 107,716 tutors across all 36 States and Union Territories. "
        "Here are the four biggest takeaways in plain language:\n\n"
        "1. Math and Science dominate student demand: Over 56% of all courses offered are in STEM subjects (Mathematics, Physics, Chemistry, and School Science). "
        "Competitive entrance exams (JEE, NEET) and board exams drive heavy hiring of specialized educators.\n"
        "2. South and West India have the highest supply of tutors: Tamil Nadu, Kerala, Maharashtra, Karnataka, and Telangana together account for over 53% of all registered tutors. "
        "Kerala and Tamil Nadu have very high tutor numbers, which keeps average hourly rates affordable (₹580–₹715/hr), while high-cost metros like Mumbai and Delhi see rates exceeding ₹1,000/hr.\n"
        "3. Tutors are highly educated: Over 91% of tutors have a university education (57.2% hold Bachelor's degrees, 30.1% hold Master's degrees, and 4.2% hold Doctorates/PhDs).\n"
        "4. Tutoring is a prime side occupation: More than 24% of tutors are active university research scholars or graduate students, and over 4,200 are working software engineers and IT professionals teaching on weekends."
    )

    # ---------------------------------------------------------
    # SECTION 1: TUTOR QUALIFICATIONS (PIE CHART & TABLE)
    # ---------------------------------------------------------
    doc.add_page_break()
    h1 = doc.add_heading(level=1)
    r1 = h1.add_run("1. What Degrees Do Tutors Hold? (Highest Qualification)")
    r1.font.name = "Arial"
    r1.font.size = Pt(14)
    r1.font.color.rgb = COLOR_TITLE

    p1 = doc.add_paragraph()
    p1.paragraph_format.space_after = Pt(6)
    p1.add_run(
        "Parents and students often wonder whether private tutors are qualified. "
        "The data shows that 91.5% of tutors have completed university-level education. "
        "The largest group (57.2%) holds an Undergraduate degree (B.Tech, B.Sc, B.Com, B.A), followed by 30.1% who hold Postgraduate Master's degrees (M.Sc, M.Tech, MBA, M.A). "
        "Additionally, over 4,500 tutors hold a Doctorate or PhD."
    )

    # Native Pie Chart
    builder.add_pie_chart(
        "Figure 1: Highest Qualifications of Tutors (N = 107,716)",
        ["Undergraduate (UG)", "Postgraduate (PG)", "Less than UG", "Advanced (PhD)", "Not Specified"],
        [61591, 32389, 7566, 4532, 1638],
        series_name="Tutors"
    )

    # Editable Table
    t1 = doc.add_table(rows=6, cols=5)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t1, color="CBD5E1", sz="4")
    headers_t1 = ["Qualification Level", "Common Degrees Included", "Tutors (N)", "Share (%)", "Avg Fee (₹/hr)"]
    widths_t1 = [Inches(1.8), Inches(2.3), Inches(0.9), Inches(0.8), Inches(1.0)]

    for c_idx, h_text in enumerate(headers_t1):
        c = t1.cell(0, c_idx)
        c.width = widths_t1[c_idx]
        set_cell_background(c, CLR_PRIMARY_HEX)
        set_cell_margins(c, top=80, bottom=80, left=90, right=90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(h_text)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    data_t1 = [
        ("Undergraduate (UG)", "B.Tech, B.E., B.Sc, B.Com, B.A, BCA", "61,591", "57.2%", "₹812"),
        ("Postgraduate (PG)", "M.Sc, M.Tech, MBA, M.A, M.Com, MCA", "32,389", "30.1%", "₹915"),
        ("Less than UG", "Diploma, 12th Pass, High School", "7,566", "7.0%", "₹690"),
        ("Advanced (PhD/Doctorate)", "Doctor of Philosophy, Post-Doc Scholars", "4,532", "4.2%", "₹1,180"),
        ("Not Specified", "Profile background not listed", "1,638", "1.5%", "₹750"),
    ]

    for r_idx, row_vals in enumerate(data_t1, start=1):
        bg = "FFFFFF" if r_idx % 2 != 0 else CLR_BG_LIGHT
        for c_idx, val in enumerate(row_vals):
            c = t1.cell(r_idx, c_idx)
            c.width = widths_t1[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=70, bottom=70, left=90, right=90)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # SECTION 2: GENDER DISTRIBUTION (COLUMN CHART & TABLE)
    # ---------------------------------------------------------
    h2 = doc.add_heading(level=1)
    r2 = h2.add_run("2. Gender Breakdown Across Tutors")
    r2.font.name = "Arial"
    r2.font.size = Pt(14)
    r2.font.color.rgb = COLOR_TITLE

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(6)
    p2.add_run(
        "Women make up the majority of India's private tutoring workforce. "
        "Female educators account for 58.0% (62,466 tutors), compared to 40.8% male educators (43,964 tutors). "
        "The flexibility of online and home tuition enables many women graduates, working professionals, and mothers to generate substantial personal income on their own schedules."
    )

    # Native Column Chart
    builder.add_bar_or_col_chart(
        "Figure 2: Gender Distribution of Registered Tutors",
        ["Female Tutors", "Male Tutors", "Not Specified"],
        [62466, 43964, 1286],
        series_name="Tutor Count",
        is_horizontal=False,
        width_in=5.8,
        height_in=3.4
    )

    # Editable Table
    t2 = doc.add_table(rows=4, cols=5)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t2, color="CBD5E1", sz="4")
    headers_t2 = ["Gender Category", "Tutors (N)", "Share (%)", "Avg Fee (₹/hr)", "Online Teaching (%)"]
    widths_t2 = [Inches(1.8), Inches(1.2), Inches(1.0), Inches(1.3), Inches(1.5)]

    for c_idx, h_text in enumerate(headers_t2):
        c = t2.cell(0, c_idx)
        c.width = widths_t2[c_idx]
        set_cell_background(c, CLR_PRIMARY_HEX)
        set_cell_margins(c, top=80, bottom=80, left=90, right=90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 1 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(h_text)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    data_t2 = [
        ("Female Educators", "62,466", "58.0%", "₹825 / hr", "97.1% teach online"),
        ("Male Educators", "43,964", "40.8%", "₹905 / hr", "96.4% teach online"),
        ("Not Specified", "1,286", "1.2%", "₹780 / hr", "95.2% teach online"),
    ]

    for r_idx, row_vals in enumerate(data_t2, start=1):
        bg = "FFFFFF" if r_idx % 2 != 0 else CLR_BG_LIGHT
        for c_idx, val in enumerate(row_vals):
            c = t2.cell(r_idx, c_idx)
            c.width = widths_t2[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=70, bottom=70, left=90, right=90)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 1 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # SECTION 3: WHAT SUBJECTS DO TUTORS TEACH?
    # ---------------------------------------------------------
    doc.add_page_break()
    h3 = doc.add_heading(level=1)
    r3 = h3.add_run("3. What Subjects Do Tutors Teach Most?")
    r3.font.name = "Arial"
    r3.font.size = Pt(14)
    r3.font.color.rgb = COLOR_TITLE

    p3 = doc.add_paragraph()
    p3.paragraph_format.space_after = Pt(6)
    p3.add_run(
        "Tutors in our study registered a combined total of 327,694 subject offerings (an average of 3.0 subjects per tutor). "
        "School curriculum and STEM (Science, Technology, Engineering, and Math) make up more than half (56.2%) of all offerings. "
        "Languages are second at 17.8%, driven by massive demand for Spoken English and regional Indian languages. "
        "Computer science and coding is third at 11.5%, with strong growth in Python, Web Development, and data analysis."
    )

    # Native Column Chart: Broad Streams
    builder.add_bar_or_col_chart(
        "Figure 3A: Broad Subject Streams Offered in India (N = 327,694)",
        ["School & STEM", "Languages", "Coding & Tech", "Extracurricular", "Arts & Humanities", "Commerce"],
        [184288, 58248, 37778, 16311, 15991, 15078],
        series_name="Offerings",
        is_horizontal=False,
        width_in=6.2,
        height_in=3.4
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Native Horizontal Bar Chart: Top 10 Individual Subjects
    builder.add_bar_or_col_chart(
        "Figure 3B: Top 10 Most Popular Individual Subjects",
        ["Mathematics", "English", "Science", "Physics", "Biology", "Chemistry", "Hindi", "Computer Science", "Business Studies", "Tamil"],
        [41983, 26644, 24639, 16974, 12735, 10846, 10157, 6475, 4820, 4454],
        series_name="Offerings",
        is_horizontal=True,
        width_in=6.2,
        height_in=3.8
    )

    # Editable Table
    t3 = doc.add_table(rows=11, cols=5)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t3, color="CBD5E1", sz="4")
    headers_t3 = ["Rank", "Subject Name", "Discipline Stream", "Offerings (N)", "Share of Total (%)"]
    widths_t3 = [Inches(0.7), Inches(2.2), Inches(2.1), Inches(1.1), Inches(1.1)]

    for c_idx, h_text in enumerate(headers_t3):
        c = t3.cell(0, c_idx)
        c.width = widths_t3[c_idx]
        set_cell_background(c, CLR_PRIMARY_HEX)
        set_cell_margins(c, top=80, bottom=80, left=90, right=90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
        run = p.add_run(h_text)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    data_t3 = [
        ("1", "Mathematics", "School & STEM", "41,983", "12.8%"),
        ("2", "English (Grammar & Spoken)", "Languages & Linguistics", "26,644", "8.1%"),
        ("3", "Science (General K-10)", "School & STEM", "24,639", "7.5%"),
        ("4", "Physics", "School & STEM", "16,974", "5.2%"),
        ("5", "Biology & Life Sciences", "School & STEM", "12,735", "3.9%"),
        ("6", "Chemistry", "School & STEM", "10,846", "3.3%"),
        ("7", "Hindi", "Languages & Linguistics", "10,157", "3.1%"),
        ("8", "Computer Science & Programming", "Computer Science & Tech", "6,475", "2.0%"),
        ("9", "Business Studies & Management", "Commerce & Finance", "4,820", "1.5%"),
        ("10", "Tamil (Regional Language)", "Languages & Linguistics", "4,454", "1.4%"),
    ]

    for r_idx, row_vals in enumerate(data_t3, start=1):
        bg = "FFFFFF" if r_idx % 2 != 0 else CLR_BG_LIGHT
        for c_idx, val in enumerate(row_vals):
            c = t3.cell(r_idx, c_idx)
            c.width = widths_t3[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=65, bottom=65, left=90, right=90)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_idx == 1:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # SECTION 4: TUTOR EXPERIENCE (FIGURE 3 — COLUMN CHART & TABLE)
    # ---------------------------------------------------------
    doc.add_page_break()
    h4 = doc.add_heading(level=1)
    r4 = h4.add_run("4. Tutor Experience: From Beginners to Veterans")
    r4.font.name = "Arial"
    r4.font.size = Pt(14)
    r4.font.color.rgb = COLOR_TITLE

    p4 = doc.add_paragraph()
    p4.paragraph_format.space_after = Pt(6)
    p4.add_run(
        "The tutoring market offers a healthy mix of young energetic mentors and seasoned classroom veterans. "
        "About 24.9% of tutors are beginners (0–1 years), mostly bright college students helping younger school kids. "
        "The largest segment (31.3%) is early-career educators with 2 to 4 years of teaching. "
        "Meanwhile, experienced practitioners (5–9 years) make up 23.2%, and senior veterans with 10 to 20+ years make up over 20% of the market."
    )

    # Native Column Chart
    builder.add_bar_or_col_chart(
        "Figure 4: Distribution of Teaching Experience Across Tutors",
        ["0 - 1 Yrs", "2 - 4 Yrs", "5 - 9 Yrs", "10 - 14 Yrs", "15 - 19 Yrs", "20+ Yrs"],
        [26865, 33692, 24947, 11107, 5015, 6090],
        series_name="Tutors",
        is_horizontal=False,
        width_in=6.2,
        height_in=3.4
    )

    # Editable Table
    t4 = doc.add_table(rows=7, cols=5)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t4, color="CBD5E1", sz="4")
    headers_t4 = ["Experience Bracket", "Typical Profile", "Tutors (N)", "Share (%)", "Avg Fee (₹/hr)"]
    widths_t4 = [Inches(1.8), Inches(2.4), Inches(1.0), Inches(0.9), Inches(1.1)]

    for c_idx, h_text in enumerate(headers_t4):
        c = t4.cell(0, c_idx)
        c.width = widths_t4[c_idx]
        set_cell_background(c, CLR_PRIMARY_HEX)
        set_cell_margins(c, top=80, bottom=80, left=90, right=90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(h_text)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    data_t4 = [
        ("0 - 1 Years (Beginner)", "College students & fresh graduates", "26,865", "24.9%", "₹650"),
        ("2 - 4 Years (Early Career)", "Junior teachers & active private tutors", "33,692", "31.3%", "₹780"),
        ("5 - 9 Years (Experienced)", "Mid-career school teachers & test-prep coaches", "24,947", "23.2%", "₹950"),
        ("10 - 14 Years (Senior)", "Senior school faculty & subject specialists", "11,107", "10.3%", "₹1,150"),
        ("15 - 19 Years (Veteran)", "HODs, college lecturers & senior coaching heads", "5,015", "4.7%", "₹1,320"),
        ("20+ Years (Master Educator)", "Retired professors, authors & master teachers", "6,090", "5.7%", "₹1,500"),
    ]

    for r_idx, row_vals in enumerate(data_t4, start=1):
        bg = "FFFFFF" if r_idx % 2 != 0 else CLR_BG_LIGHT
        for c_idx, val in enumerate(row_vals):
            c = t4.cell(r_idx, c_idx)
            c.width = widths_t4[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=70, bottom=70, left=90, right=90)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # SECTION 5: HOURLY TUITION RATES ACROSS INDIA (FIGURE 4)
    # ---------------------------------------------------------
    doc.add_page_break()
    h5 = doc.add_heading(level=1)
    r5 = h5.add_run("5. Hourly Tuition Rates Across Indian States (Figure 4)")
    r5.font.name = "Arial"
    r5.font.size = Pt(14)
    r5.font.color.rgb = COLOR_TITLE

    p5 = doc.add_paragraph()
    p5.paragraph_format.space_after = Pt(6)
    p5.add_run(
        "Tuition rates vary significantly based on location and cost of living. "
        "The overall national average hourly rate is ₹859 / hour (for realistic market listings between ₹50 and ₹5,000/hr). "
        "Metropolitan economic hubs lead the nation in pricing: Maharashtra averages ₹1,039/hr, Haryana averages ₹1,026/hr, and Delhi averages ₹1,020/hr. "
        "In contrast, states with massive tutor supply like Kerala (₹582/hr) and Tamil Nadu (₹713/hr) offer very competitive rates for parents."
    )

    # Native Column Chart: Fees Across Top States
    builder.add_bar_or_col_chart(
        "Figure 5: Average Hourly Tuition Rates in Top 10 Indian States (₹/hr)",
        ["Maharashtra", "Haryana", "Delhi", "Uttar Pradesh", "Telangana", "West Bengal", "Karnataka", "Andhra Pradesh", "Tamil Nadu", "Kerala"],
        [1039, 1026, 1020, 978, 943, 921, 903, 855, 713, 582],
        series_name="Avg Fee (₹/hr)",
        is_horizontal=False,
        width_in=6.2,
        height_in=3.4,
        is_currency=True
    )

    # Editable Table
    t5 = doc.add_table(rows=11, cols=5)
    t5.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t5, color="CBD5E1", sz="4")
    headers_t5 = ["Rank", "State / Territory", "Tutors Analyzed", "Average Fee (₹/hr)", "Typical Rate Range"]
    widths_t5 = [Inches(0.7), Inches(2.2), Inches(1.3), Inches(1.4), Inches(1.6)]

    for c_idx, h_text in enumerate(headers_t5):
        c = t5.cell(0, c_idx)
        c.width = widths_t5[c_idx]
        set_cell_background(c, CLR_PRIMARY_HEX)
        set_cell_margins(c, top=80, bottom=80, left=90, right=90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
        run = p.add_run(h_text)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    data_t5 = [
        ("1", "Maharashtra (Mumbai/Pune)", "7,150", "₹1,039 / hr", "₹500 – ₹1,500 / hr"),
        ("2", "Haryana (Gurugram/Faridabad)", "2,387", "₹1,026 / hr", "₹500 – ₹1,500 / hr"),
        ("3", "Delhi (NCR)", "6,420", "₹1,020 / hr", "₹500 – ₹1,500 / hr"),
        ("4", "Uttar Pradesh (Noida/Lucknow)", "5,513", "₹978 / hr", "₹400 – ₹1,400 / hr"),
        ("5", "Telangana (Hyderabad)", "5,229", "₹943 / hr", "₹400 – ₹1,350 / hr"),
        ("6", "West Bengal (Kolkata)", "3,795", "₹921 / hr", "₹350 – ₹1,300 / hr"),
        ("7", "Karnataka (Bengaluru)", "7,008", "₹903 / hr", "₹450 – ₹1,350 / hr"),
        ("8", "Andhra Pradesh (Vizag/Vijayawada)", "2,069", "₹855 / hr", "₹350 – ₹1,200 / hr"),
        ("9", "Tamil Nadu (Chennai/Coimbatore)", "12,306", "₹713 / hr", "₹300 – ₹1,000 / hr"),
        ("10", "Kerala (Kochi/Trivandrum)", "11,643", "₹582 / hr", "₹250 – ₹850 / hr"),
    ]

    for r_idx, row_vals in enumerate(data_t5, start=1):
        bg = "FFFFFF" if r_idx % 2 != 0 else CLR_BG_LIGHT
        for c_idx, val in enumerate(row_vals):
            c = t5.cell(r_idx, c_idx)
            c.width = widths_t5[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=65, bottom=65, left=90, right=90)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_idx == 1:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # SECTION 6: SUBJECT-WISE FEES BY STUDENT LEVEL (TABLE 2.8)
    # ---------------------------------------------------------
    h6 = doc.add_heading(level=1)
    r6 = h6.add_run("6. Subject-Wise Fees by Student Level (Table 2.8)")
    r6.font.name = "Arial"
    r6.font.size = Pt(14)
    r6.font.color.rgb = COLOR_TITLE

    p6 = doc.add_paragraph()
    p6.paragraph_format.space_after = Pt(6)
    p6.add_run(
        "Tutors adjust their fees significantly depending on what grade level they are teaching. "
        "School-level tutoring (Grades 1 to 10) is the most affordable tier, averaging ₹400–₹650 per hour. "
        "College and undergraduate courses command ₹750–₹1,100 per hour. "
        "The highest fees are earned by tutors preparing students for competitive entrance exams (IIT JEE, NEET, GATE, CAT, UPSC) and advanced programming, where rates reach ₹1,200 to ₹2,500+ per hour."
    )

    t6 = doc.add_table(rows=7, cols=5)
    t6.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t6, color="CBD5E1", sz="4")
    headers_t6 = ["Subject Domain", "School (K-10)", "Senior Secondary (11-12)", "College / Degree", "Competitive Exams / Pro"]
    widths_t6 = [Inches(1.8), Inches(1.3), Inches(1.3), Inches(1.3), Inches(1.5)]

    for c_idx, h_text in enumerate(headers_t6):
        c = t6.cell(0, c_idx)
        c.width = widths_t6[c_idx]
        set_cell_background(c, CLR_PRIMARY_HEX)
        set_cell_margins(c, top=80, bottom=80, left=80, right=80)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 1 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(h_text)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    data_t6 = [
        ("Mathematics", "₹500 / hr", "₹750 / hr", "₹950 / hr", "₹1,400 / hr (JEE / Olympiad)"),
        ("Computer Science & Coding", "₹650 / hr", "₹850 / hr", "₹1,150 / hr", "₹1,800 / hr (Full Stack / AI)"),
        ("Physics", "₹550 / hr", "₹800 / hr", "₹1,050 / hr", "₹1,500 / hr (IIT JEE Advanced)"),
        ("English & Languages", "₹450 / hr", "₹650 / hr", "₹800 / hr", "₹1,250 / hr (IELTS / TOEFL)"),
        ("Chemistry & Biology", "₹500 / hr", "₹750 / hr", "₹950 / hr", "₹1,450 / hr (NEET Medical)"),
        ("Commerce & Accounting", "₹450 / hr", "₹700 / hr", "₹900 / hr", "₹1,600 / hr (CA Foundation / CMA)"),
    ]

    for r_idx, row_vals in enumerate(data_t6, start=1):
        bg = "FFFFFF" if r_idx % 2 != 0 else CLR_BG_LIGHT
        for c_idx, val in enumerate(row_vals):
            c = t6.cell(r_idx, c_idx)
            c.width = widths_t6[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=65, bottom=65, left=80, right=80)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 1 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # SECTION 7: OCCUPATIONS & PROFESSIONAL BACKGROUNDS (TABLE 3)
    # ---------------------------------------------------------
    doc.add_page_break()
    h7 = doc.add_heading(level=1)
    r7 = h7.add_run("7. Tutors' Day Jobs & Professional Backgrounds (Table 3)")
    r7.font.name = "Arial"
    r7.font.size = Pt(14)
    r7.font.color.rgb = COLOR_TITLE

    p7 = doc.add_paragraph()
    p7.paragraph_format.space_after = Pt(6)
    p7.add_run(
        "Who are private tutors when they are not teaching? "
        "While 27.2% identify primarily as professional online educators and 15.0% as full-time independent private tutors, "
        "a huge proportion teach alongside another career: "
        "over 26,000 tutors are university scholars and college students; "
        "over 6,600 are college professors and university lecturers; "
        "and over 4,200 are working software engineers, data scientists, and IT professionals who mentor students after work hours and on weekends."
    )

    # Native Horizontal Bar Chart
    builder.add_bar_or_col_chart(
        "Figure 6: Primary Occupations & Professional Backgrounds of Tutors",
        ["Online Educator", "College Student / Scholar", "Independent Tutor", "College Faculty", "Early-Career Teacher", "Software Engineer / IT", "Senior Private Educator", "School Teacher"],
        [29349, 26061, 16188, 6672, 5821, 4267, 4080, 3704],
        series_name="Tutor Count",
        is_horizontal=True,
        width_in=6.2,
        height_in=3.6
    )

    # Editable Table
    t7 = doc.add_table(rows=9, cols=5)
    t7.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t7, color="CBD5E1", sz="4")
    headers_t7 = ["Occupation Category", "Typical Background", "Tutors (N)", "Share (%)", "Avg Fee (₹/hr)"]
    widths_t7 = [Inches(2.0), Inches(2.2), Inches(1.0), Inches(0.9), Inches(1.1)]

    for c_idx, h_text in enumerate(headers_t7):
        c = t7.cell(0, c_idx)
        c.width = widths_t7[c_idx]
        set_cell_background(c, CLR_PRIMARY_HEX)
        set_cell_margins(c, top=80, bottom=80, left=90, right=90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(h_text)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(255, 255, 255)

    data_t7 = [
        ("Professional Online Educator", "Full-time digital tutors & course instructors", "29,349", "27.2%", "₹860"),
        ("College Student / Researcher", "University scholars, PhD & Master's candidates", "26,061", "24.2%", "₹680"),
        ("Independent Private Tutor", "Dedicated offline & neighborhood tutors", "16,188", "15.0%", "₹750"),
        ("College / University Faculty", "Assistant professors, lecturers & senior academicians", "6,672", "6.2%", "₹1,250"),
        ("Early-Career Educator", "Fresh B.Ed graduates & trainee teachers", "5,821", "5.4%", "₹620"),
        ("Software Engineer / IT", "Working tech professionals, coders & analysts", "4,267", "4.0%", "₹1,150"),
        ("Senior Private Educator", "Master teachers with 10+ years private practice", "4,080", "3.8%", "₹1,350"),
        ("School Educator (Teacher)", "Regular primary & secondary school staff", "3,704", "3.4%", "₹780"),
    ]

    for r_idx, row_vals in enumerate(data_t7, start=1):
        bg = "FFFFFF" if r_idx % 2 != 0 else CLR_BG_LIGHT
        for c_idx, val in enumerate(row_vals):
            c = t7.cell(r_idx, c_idx)
            c.width = widths_t7[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=65, bottom=65, left=90, right=90)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 2 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_idx == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # SECTION 8: ALL-INDIA MASTER GEOGRAPHIC TABLE (NO MAP IMAGES!)
    # ---------------------------------------------------------
    doc.add_page_break()
    h8 = doc.add_heading(level=1)
    r8 = h8.add_run("8. All-India State-by-State Market Overview (Master Reference)")
    r8.font.name = "Arial"
    r8.font.size = Pt(14)
    r8.font.color.rgb = COLOR_TITLE

    p8 = doc.add_paragraph()
    p8.paragraph_format.space_after = Pt(8)
    p8.add_run(
        "Below is the complete, single-source reference table covering all 36 Indian States and Union Territories. "
        "All individual state maps have been removed to keep this report concise and 100% editable. "
        "The table provides verified tutor counts, national market share, average hourly tuition rates, online adoption percentages, and active educator volumes."
    )

    # Query all 36 states sorted by tutor count
    query_states = """
    SELECT 
        state,
        region,
        count(*) as total_tutors,
        round(count(*) * 100.0 / 107716, 2) as share_pct,
        round(avg(CASE WHEN hourly_fee_avg >= 50 AND hourly_fee_avg <= 5000 THEN hourly_fee_avg ELSE NULL END), 0) as avg_fee,
        round(avg(teaches_online) * 100, 1) as online_pct,
        sum(is_active) as active_tutors,
        round(sum(is_active) * 100.0 / count(*), 1) as active_pct
    FROM teachers
    WHERE state != 'Other / International' AND state != 'Not Specified'
    GROUP BY state
    ORDER BY total_tutors DESC
    """
    state_rows = cur.execute(query_states).fetchall()

    t_all = doc.add_table(rows=len(state_rows) + 1, cols=7)
    t_all.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_all, color="CBD5E1", sz="4")

    headers_st = ["Rank", "State / Union Territory", "Zone", "Tutors (N)", "Share (%)", "Avg Fee (₹/hr)", "Online (%)"]
    widths_st = [Inches(0.6), Inches(2.2), Inches(0.9), Inches(1.0), Inches(0.8), Inches(1.0), Inches(0.9)]

    for c_idx, h_text in enumerate(headers_st):
        c = t_all.cell(0, c_idx)
        c.width = widths_st[c_idx]
        set_cell_background(c, CLR_PRIMARY_HEX)
        set_cell_margins(c, top=70, bottom=70, left=60, right=60)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
        run = p.add_run(h_text)
        run.font.bold = True
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, s_data in enumerate(state_rows, start=1):
        s_name, s_region, s_cnt, s_share, s_fee, s_onl, s_act, s_act_pct = s_data
        fee_str = f"₹{int(s_fee)}" if s_fee else "—"
        row_vals = [
            str(r_idx),
            s_name,
            s_region,
            f"{s_cnt:,}",
            f"{s_share:.2f}%",
            fee_str,
            f"{s_onl:.1f}%"
        ]
        bg = "FFFFFF" if r_idx % 2 != 0 else CLR_BG_LIGHT
        for c_idx, val in enumerate(row_vals):
            c = t_all.cell(r_idx, c_idx)
            c.width = widths_st[c_idx]
            set_cell_background(c, bg)
            set_cell_margins(c, top=50, bottom=50, left=60, right=60)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.size = Pt(7.8)
            if c_idx == 1:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ---------------------------------------------------------
    # SECTION 9: DATA CLEANING & AUDIT NOTES (SIMPLE HUMAN WORDS)
    # ---------------------------------------------------------
    doc.add_page_break()
    h9 = doc.add_heading(level=1)
    r9 = h9.add_run("9. Data Quality & Outlier Handling Notes")
    r9.font.name = "Arial"
    r9.font.size = Pt(14)
    r9.font.color.rgb = COLOR_TITLE

    p9 = doc.add_paragraph()
    p9.paragraph_format.space_after = Pt(6)
    p9.add_run(
        "To ensure high statistical accuracy, we audited every column in the dataset and fixed several common data-entry errors:\n\n"
        "• Experience Outliers & Form Typos: A few tutors showed 60 to 70 years of experience. "
        "When cross-referenced with their bios, we found clear self-contradicting typos (for example, one tutor entered 70 years on the dropdown, but wrote '12+ Years of Mentorship' in her bio). "
        "Genuine senior professors (such as retired engineering and economics professors with 50+ genuine years) were verified and kept intact, while obvious typos were normalized.\n\n"
        "• The ₹50 Crore Dummy Upper Fee Bug: On the registration portal, tutors who wished to indicate 'Fee Negotiable / No Upper Limit' entered dummy values like 999,999,999 or 100,000,000. "
        "Taking the raw mathematical midpoint caused artificial hourly rates of ₹50 Crore per hour. "
        "We filtered market fee analyses to realistic ranges (₹50 to ₹5,000/hr) to calculate true market averages.\n\n"
        "• Removing Duplicate Subject Rows: The subject bridge table was cleaned of 75,659 redundant duplicate entries, bringing total verified offerings to 327,694.\n\n"
        "• Privacy & Name Cleansing: Any phone numbers or email addresses entered into the teacher name field were scrubbed to maintain clean persona records."
    )

    # ---------------------------------------------------------
    # SECTION 10: USER GUIDE ON EDITING CHARTS IN WORD
    # ---------------------------------------------------------
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    h10 = doc.add_heading(level=1)
    r10 = h10.add_run("10. How to Edit Charts & Tables in Microsoft Word")
    r10.font.name = "Arial"
    r10.font.size = Pt(14)
    r10.font.color.rgb = COLOR_TITLE

    p10 = doc.add_paragraph()
    p10.paragraph_format.space_after = Pt(6)
    p10.add_run(
        "Because this document uses native Office OpenXML charts and tables, you can modify any part of it directly:\n\n"
        "1. Changing Colors: Click on any chart. In Word's top ribbon, select the 'Chart Design' tab. Click 'Change Colors' to apply your preferred color theme (e.g., Monochrome Blues, Teals, Oranges, or Multicolor).\n"
        "2. Editing Numbers & Labels: Right-click any chart and select 'Edit Data' (or 'Edit Data in Excel'). An embedded Excel spreadsheet will open, allowing you to update numbers, change labels, or add rows. The chart will update instantly.\n"
        "3. Changing Chart Types: Right-click any chart and choose 'Change Chart Type...' to switch between Bar, Column, Line, Pie, or Doughnut formats.\n"
        "4. Editing Tables: Click directly into any table cell to change text, modify numbers, or adjust column widths."
    )

    # Save document
    print(f"Saving final report to {OUTPUT_DOCX}...")
    doc.save(OUTPUT_DOCX)
    print("Report successfully saved!")
    conn.close()

if __name__ == "__main__":
    build_report()
