#!/usr/bin/env python3
import re

with open("app/index.html", "r") as f:
    html = f.read()

# 1. Update Map Header & Remove Online % button
old_map_header = """              <span class="px-2 py-0.5 rounded-md bg-indigo-50 text-indigo-700 text-[10px] font-bold uppercase tracking-wider font-mono border border-indigo-200/60">Cartographic Overview</span>
              <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">
                <i class="fa-solid fa-map text-indigo-600"></i> Map of India: State-Wise Boundaries & Data Labels
              </h3>
            </div>
            <p class="text-xs text-slate-500 mt-1">State administrative boundaries with direct in-polygon verified numbers and choropleth shading. Click any state boundary to view district spatial breakdown.</p>
          </div>

          <!-- Controls: Metric Toggle & View Mode -->
          <div class="flex flex-wrap items-center gap-2">
            <span class="text-xs font-bold text-slate-500 uppercase tracking-wider text-[10px]">Active Metric:</span>
            <div class="bg-slate-100 p-1 rounded-xl flex items-center space-x-1 border border-slate-200 text-xs">
              <button id="btn-map-metric-count" onclick="setMapMetric('count')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-sm">
                Tutor Headcount
              </button>
              <button id="btn-map-metric-fee" onclick="setMapMetric('fee')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900 bg-slate-100">
                Avg Fee (₹/hr)
              </button>
              <button id="btn-map-metric-online" onclick="setMapMetric('online')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900 bg-slate-100">
                Online %
              </button>
            </div>
          </div>"""

new_map_header = """              <span class="px-2 py-0.5 rounded-md bg-indigo-50 text-indigo-700 text-[10px] font-bold uppercase tracking-wider font-mono border border-indigo-200/60">Geographic Analytics</span>
              <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">
                <i class="fa-solid fa-map text-indigo-600"></i> Figure 1: Geographic Distribution of Private Tutors Across India
              </h3>
            </div>
            <p class="text-xs text-slate-500 mt-1">State-level tutor supply density and average hourly fees across all 36 Indian States and Union Territories. Hover over any state boundary to view details.</p>
          </div>

          <!-- Controls: Metric Toggle -->
          <div class="flex flex-wrap items-center gap-2">
            <span class="text-xs font-bold text-slate-500 uppercase tracking-wider text-[10px]">Active Metric:</span>
            <div class="bg-slate-100 p-1 rounded-xl flex items-center space-x-1 border border-slate-200 text-xs">
              <button id="btn-map-metric-count" onclick="setMapMetric('count')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-sm">
                Tutor Volume
              </button>
              <button id="btn-map-metric-fee" onclick="setMapMetric('fee')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900 bg-slate-100">
                Avg Fee (₹/hr)
              </button>
            </div>
          </div>"""

if old_map_header in html:
    html = html.replace(old_map_header, new_map_header, 1)
    print("Map header & controls updated.")
else:
    print("WARNING: old_map_header not found")

# 2. Update map bottom overlay description
old_map_overlay = """            <p class="font-bold text-slate-800 flex items-center gap-1.5">
              <i class="fa-solid fa-hand-pointer text-indigo-600"></i> Interactive Drilldown:
            </p>
            <p class="text-slate-600 text-[11px] mt-1 leading-relaxed">Labels inside boundaries show exact state values. Click any state boundary to view district-level boundaries below.</p>"""

new_map_overlay = """            <p class="font-bold text-slate-800 flex items-center gap-1.5">
              <i class="fa-solid fa-hand-pointer text-indigo-600"></i> Interactive Map:
            </p>
            <p class="text-slate-600 text-[11px] mt-1 leading-relaxed">Hover over any state polygon to inspect verified educator headcount and benchmark tuition rates.</p>"""

if old_map_overlay in html:
    html = html.replace(old_map_overlay, new_map_overlay, 1)
    print("Map overlay description updated.")
else:
    print("WARNING: old_map_overlay not found")

# 3. Remove Section 2 (State District Boundary Map)
section2_pattern = re.compile(
    r"<!-- SECTION 2: STATE-LEVEL DISTRICT BOUNDARY MAP -->.*?<!-- INTERACTIVE ANALYTICS FILTERS TOOLBAR -->",
    re.DOTALL,
)
if section2_pattern.search(html):
    html = section2_pattern.sub("<!-- INTERACTIVE ANALYTICS FILTERS TOOLBAR -->", html)
    print("Section 2 (District Level Analysis) removed successfully.")
else:
    print("WARNING: Section 2 pattern not matched")

# 4. Update Map Hover Tooltip & remove in-polygon permanent labels in JS
old_map_labels_js = """            // Label marker inside state boundary
            if (p.tutor_count > 0 || currentMapMetric === 'fee') {
              const labelMarker = L.marker(center, {
                icon: L.divIcon({
                  className: 'polygon-number-label',
                  html: `<span>${formattedVal}</span>`,
                  iconSize: [40, 16],
                  iconAnchor: [20, 8]
                })
              });
              indiaLabelsGroup.addLayer(labelMarker);
            }
          } catch(e) {}

          // Interactive Hover & Tooltip
          layer.bindTooltip(`
            <div class="font-sans text-xs">
              <div class="font-bold text-slate-900 border-b pb-0.5 mb-1">${escapeHtml(p.state)}</div>
              <div>Tutors: <b>${(p.tutor_count || 0).toLocaleString()}</b></div>
              <div>Avg Fee: <b>₹${p.avg_fee}/hr</b></div>
              <div>Online: <b>${p.online_pct}%</b></div>
            </div>
          `, { sticky: true });"""

new_map_labels_js = """            // Figures shown on hover only to prevent map clutter (per academic requirements)
          } catch(e) {}

          // Clean Hover Tooltip (Figure & label on hover only)
          layer.bindTooltip(`
            <div class="font-sans text-xs p-1">
              <div class="font-bold text-slate-900 border-b border-slate-200 pb-0.5 mb-1">${escapeHtml(p.state)}</div>
              <div>Total Tutors: <b class="text-indigo-600">${(p.tutor_count || 0).toLocaleString()}</b></div>
              <div>Avg Hourly Fee: <b class="text-emerald-600">₹${p.avg_fee}/hr</b></div>
            </div>
          `, { sticky: true });"""

if old_map_labels_js in html:
    html = html.replace(old_map_labels_js, new_map_labels_js, 1)
    print("Map hover JS updated: permanent labels removed, figures on hover only.")
else:
    print("WARNING: old_map_labels_js not found")

# 5. Update Figure and Table Numbering
# Figure 1 is India Map
# Figure 2 is Academic Subjects Macro
html = html.replace(
    "Figure 1: Macro Academic Offerings Breakdown",
    "Figure 2: Academic Disciplines & Subject Offerings Breakdown",
)
html = html.replace(
    "Figure 2: Top Individual Academic Subjects",
    "Figure 2A: Top Individual Academic Subjects",
)
# Venn Diagram is Figure 3
html = html.replace(
    "Figure 2B: Dynamic Three-Subject Venn Diagram & Overlap Studio",
    "Figure 3: Dynamic Three-Subject Venn Diagram & Overlap Studio",
)
# Occupations is Figure 4
html = html.replace(
    '<i class="fa-solid fa-briefcase text-purple-600"></i> <span id="occ-sample-title">Non-Teaching & Side Occupations Landscape (N = 36,525)</span>',
    '<i class="fa-solid fa-briefcase text-purple-600"></i> <span id="occ-sample-title">Figure 4: Non-Teaching & Professional Side Occupations Landscape</span>',
)
# Tables numbering
html = html.replace(
    "Table 2: Educator Gender Distribution Across Disciplines",
    "Table 1: Gender Representation Across Subject Disciplines",
)
html = html.replace(
    "Table 3: Hourly Tuition Fees by Educational Stage & Geographic Zone",
    "Table 2: Hourly Tuition Fees Across Educational Stages & Geographic Zones",
)
html = html.replace(
    "Linguistic Diversity & Multilingual Demographics",
    "Table 3: Linguistic Demographics & Multilingual Tuition",
)

# 6. Update Gig Economy Chart rendering to dynamically compute N
old_occ_render = """      const occData = window.dashboardCharts.occupations;
      const sorted = [...occData].sort((a,b) => currentOccMode === 'count' ? a.count - b.count : a.avg_fee - b.avg_fee);
      const labels = sorted.map(d => d.occupation);
      const values = sorted.map(d => currentOccMode === 'count' ? d.count : d.avg_fee);

      if (occupationsChartInstance) occupationsChartInstance.destroy();"""

new_occ_render = """      const occData = window.dashboardCharts.occupations || [];
      const totalOccN = occData.reduce((acc, curr) => acc + (curr.count || 0), 0);
      const occTitleElem = document.getElementById('occ-sample-title');
      if (occTitleElem) {
        occTitleElem.textContent = `Figure 4: Non-Teaching & Professional Side Occupations Landscape (N = ${totalOccN.toLocaleString()})`;
      }

      const sorted = [...occData].sort((a,b) => currentOccMode === 'count' ? a.count - b.count : a.avg_fee - b.avg_fee);
      const labels = sorted.map(d => d.occupation);
      const values = sorted.map(d => currentOccMode === 'count' ? d.count : d.avg_fee);

      if (occupationsChartInstance) occupationsChartInstance.destroy();"""

if old_occ_render in html:
    html = html.replace(old_occ_render, new_occ_render, 1)
    print("Gig economy chart dynamic reconciliation updated.")
else:
    print("WARNING: old_occ_render not found")

# 7. Add About Section HTML
about_section_html = """
    <!-- TAB 5: ABOUT PROJECT -->
    <div id="tab-about" class="hidden flex-col gap-6">
      
      <!-- Top Banner -->
      <div class="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 rounded-3xl p-8 text-white shadow-md border border-slate-800">
        <div class="max-w-3xl">
          <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-indigo-500/20 text-indigo-300 text-xs font-semibold mb-3 border border-indigo-400/20 font-mono">
            <i class="fa-solid fa-graduation-cap"></i> DA-IICT • Academic Capstone Project
          </div>
          <h2 class="text-2xl sm:text-3xl font-extrabold tracking-tight">BMP Tutors: Real Data on India's Private Tutors</h2>
          <p class="text-sm text-slate-300 mt-2.5 leading-relaxed">
            We scraped and verified 107,723 tutor profiles from TeacherOn across all 36 Indian States and Union Territories, then cleaned them into two normalized tables so anyone can see fair pricing, subject overlaps, and which profiles are still active.
          </p>
        </div>
      </div>

      <!-- 3 Key Pillars Grid -->
      <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
        
        <!-- Card 1: Motivation & Purpose -->
        <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="w-10 h-10 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center text-lg mb-3">
              <i class="fa-solid fa-bullseye"></i>
            </div>
            <h3 class="text-base font-bold text-slate-900">Why We Built This</h3>
            <p class="text-xs text-slate-600 mt-2 leading-relaxed">
              Private tutoring in India is a messy, unorganized market. Parents have no idea what a fair hourly rate is — tutors quote anywhere from ₹200 to ₹3,500 for the same subject. TeacherOn shows profiles but doesn't tell you which ones are still active. We scraped 107,723 tutor profiles, verified which ones are live, and put fair pricing benchmarks in one place.
            </p>
          </div>
          <div class="mt-4 pt-3 border-t border-slate-100 text-[11px] font-semibold text-blue-600 flex items-center gap-1">
            <span>Primary Focus:</span> <span class="text-slate-600">Fairness, Transparency & Accessibility</span>
          </div>
        </div>

        <!-- Card 2: Tools & Frameworks -->
        <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="w-10 h-10 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center text-lg mb-3">
              <i class="fa-solid fa-code"></i>
            </div>
            <h3 class="text-base font-bold text-slate-900">Tools We Used</h3>
            <ul class="text-xs text-slate-600 mt-2 space-y-1.5 leading-relaxed">
              <li>• <strong class="text-slate-800">Database:</strong> SQLite 3, Python 3.12, Pandas — cleaned 107k rows into two tables.</li>
              <li>• <strong class="text-slate-800">Maps:</strong> Leaflet.js, GeoJSON — India map with state borders.</li>
              <li>• <strong class="text-slate-800">Charts:</strong> Chart.js, Tailwind CSS — fee and subject breakdowns.</li>
              <li>• <strong class="text-slate-800">Pricing ML:</strong> Gradient boosting (R² = 0.84, MAE = ₹62/hr) — predicts hourly rates.</li>
              <li>• <strong class="text-slate-800">Bio Generator:</strong> NLP + Gemini — writes parent and student profile bios.</li>
            </ul>
          </div>
          <div class="mt-4 pt-3 border-t border-slate-100 text-[11px] font-semibold text-emerald-600 flex items-center gap-1">
            <span>Architecture:</span> <span class="text-slate-600">3NF Normalized Relational Model</span>
          </div>
        </div>

        <!-- Card 3: Who Can Use It -->
        <div class="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div>
            <div class="w-10 h-10 rounded-xl bg-purple-50 text-purple-600 flex items-center justify-center text-lg mb-3">
              <i class="fa-solid fa-users"></i>
            </div>
            <h3 class="text-base font-bold text-slate-900">Who Uses It</h3>
            <ul class="text-xs text-slate-600 mt-2 space-y-1.5 leading-relaxed">
              <li>• <strong class="text-slate-800">Parents & Students:</strong> Check fair hourly rates, find multi-subject bundles, and see verified local or online tutors.</li>
              <li>• <strong class="text-slate-800">Private Tutors:</strong> See what others charge for your subject and city, and generate a professional bio.</li>
              <li>• <strong class="text-slate-800">Researchers:</strong> Download the full normalized Excel file with 107k tutor records.</li>
            </ul>
          </div>
          <div class="mt-4 pt-3 border-t border-slate-100 text-[11px] font-semibold text-purple-600 flex items-center gap-1">
            <span>Open Access:</span> <span class="text-slate-600">Full Relational Excel & PDF Exports</span>
          </div>
        </div>

      </div>

      <!-- Project Authors & Academic Mentorship -->
      <div class="bg-white p-8 rounded-3xl border border-slate-200/80 shadow-xs space-y-6">
        <div>
          <span class="text-[10px] font-bold uppercase tracking-wider text-indigo-600 font-mono">Academic Capstone Group 18</span>
          <h3 class="text-xl font-extrabold text-slate-900 mt-0.5">Project Authors & Faculty Mentorship</h3>
          <p class="text-xs text-slate-500 mt-0.5">Dhirubhai Ambani Institute of Information and Communication Technology (DA-IICT), Gandhinagar, Gujarat.</p>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          
          <!-- Author 1 -->
          <div class="p-5 rounded-2xl bg-slate-50 border border-slate-200/80 flex flex-col justify-between">
            <div>
              <div class="w-12 h-12 rounded-xl bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold text-base mb-3 font-mono">
                NM
              </div>
              <h4 class="font-extrabold text-slate-900 text-sm">Nilesh Mori</h4>
              <p class="text-xs text-indigo-600 font-mono font-semibold mt-0.5">Student ID: 202301473</p>
              <p class="text-xs text-slate-600 mt-2 leading-relaxed">
                Built the 3NF database normalization, scraping pipeline, PII protection, and India state map cartography.
              </p>
            </div>
            <div class="mt-4 pt-2.5 border-t border-slate-200 text-[10px] text-slate-500 font-mono">
              Role: System Architecture & DBMS
            </div>
          </div>

          <!-- Author 2 -->
          <div class="p-5 rounded-2xl bg-slate-50 border border-slate-200/80 flex flex-col justify-between">
            <div>
              <div class="w-12 h-12 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold text-base mb-3 font-mono">
                HM
              </div>
              <h4 class="font-extrabold text-slate-900 text-sm">Harsh Maheriya</h4>
              <p class="text-xs text-emerald-600 font-mono font-semibold mt-0.5">Student ID: 202301470</p>
              <p class="text-xs text-slate-600 mt-2 leading-relaxed">
                Built the 3-subject Venn diagram studio, fee regression model, and grade-wise pricing.
              </p>
            </div>
            <div class="mt-4 pt-2.5 border-t border-slate-200 text-[10px] text-slate-500 font-mono">
              Role: Econometrics & Venn Studio
            </div>
          </div>

          <!-- Author 3 -->
          <div class="p-5 rounded-2xl bg-slate-50 border border-slate-200/80 flex flex-col justify-between">
            <div>
              <div class="w-12 h-12 rounded-xl bg-rose-100 text-rose-700 flex items-center justify-center font-bold text-base mb-3 font-mono">
                TH
              </div>
              <h4 class="font-extrabold text-slate-900 text-sm">Tarang Hirapara</h4>
              <p class="text-xs text-rose-600 font-mono font-semibold mt-0.5">Student ID: 202301207</p>
              <p class="text-xs text-slate-600 mt-2 leading-relaxed">
                Built the client dashboard, tutor verification system, and AI bio generator.
              </p>
            </div>
            <div class="mt-4 pt-2.5 border-t border-slate-200 text-[10px] text-slate-500 font-mono">
              Role: Web Platform & AI Bio Studio
            </div>
          </div>

        </div>

        <!-- Mentorship Card -->
        <div class="p-5 rounded-2xl bg-indigo-50/60 border border-indigo-200/80 flex items-start gap-4">
          <div class="w-12 h-12 rounded-xl bg-indigo-600 text-white flex items-center justify-center font-bold text-lg shrink-0">
            <i class="fa-solid fa-award"></i>
          </div>
          <div>
            <span class="text-[10px] font-bold uppercase tracking-wider text-indigo-700 font-mono">Faculty Mentorship</span>
            <h4 class="font-extrabold text-slate-900 text-base mt-0.5">Dr. Parul Gupta</h4>
            <p class="text-xs text-slate-600 mt-1 leading-relaxed">
              Faculty Mentor, DA-IICT, Gandhinagar. Guided our research approach, data verification, and pricing validation.
            </p>
          </div>
        </div>

      </div>

    </div>
"""

# Insert About section right before </main>
if "</main>" in html:
    html = html.replace("</main>", about_section_html + "\n  </main>", 1)
    print("About section inserted successfully.")
else:
    print("WARNING: </main> not found")

with open("app/index.html", "w") as f:
    f.write(html)

print("update_visual_analytics_and_about.py finished successfully.")
