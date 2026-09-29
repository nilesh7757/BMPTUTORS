#!/usr/bin/env python3
"""
Generate Modern UI/UX Redesigned Dashboard for BMP Tutors.
Applies modern SaaS design engineering, Tailwind CSS, Inter & JetBrains Mono typography,
shimmer skeleton loaders, micro-interactions, faceted filters, and responsive layout.
"""

import os
import re

BACKUP_PATH = "/home/nilesh7757/BMPTUTORS1/app/index.html.bak"
TARGET_PATH = "/home/nilesh7757/BMPTUTORS1/app/index.html"

with open(BACKUP_PATH, 'r') as f:
    orig = f.read()

# Extract script block
script_start = orig.find('<script>\n    // State')
if script_start == -1:
    script_start = orig.find('<script>\n    // State')
    if script_start == -1:
        script_start = orig.rfind('<script>')

script_end = orig.rfind('</script>')
orig_script = orig[script_start + len('<script>'):script_end]

# Now let's enhance the script to support:
# 1. Occupation filter in params and populateSelect
# 2. Table row rendering with avatar initials, glowing status dot, occupation badge
# 3. Tutor detail modal with occupation display
# 4. Occupations chart in Visual Analytics (Figure 6)
# 5. Skeleton loader on table load

script_mod = orig_script

# 1. Add populateSelect for occupation in loadFilters
old_pop = "populateSelect('filter-role', data.experience_roles);"
new_pop = """populateSelect('filter-role', data.experience_roles);
        if (data.occupations && document.getElementById('filter-occupation')) {
          populateSelect('filter-occupation', data.occupations);
        }"""
script_mod = script_mod.replace(old_pop, new_pop, 1)

# 2. Add occupation param in loadTutors
old_load_params = "experience_role: document.getElementById('filter-role').value,"
new_load_params = """experience_role: document.getElementById('filter-role').value,
        occupation: document.getElementById('filter-occupation') ? document.getElementById('filter-occupation').value : '',"""
script_mod = script_mod.replace(old_load_params, new_load_params, 1)

# 3. Add reset for filter-occupation in resetFilters
old_reset = "document.getElementById('filter-role').value = '';"
new_reset = """document.getElementById('filter-role').value = '';
      if (document.getElementById('filter-occupation')) document.getElementById('filter-occupation').value = '';"""
script_mod = script_mod.replace(old_reset, new_reset, 1)

# 4. Add occupation in viewTutorDetails modal
old_modal_role = "document.getElementById('modal-role').textContent = t.experience_role || 'Tutor';"
new_modal_role = """document.getElementById('modal-role').textContent = t.experience_role || 'Tutor';
        if (document.getElementById('modal-occupation')) {
          document.getElementById('modal-occupation').textContent = t.occupation || t.experience_role || 'Educator';
        }"""
script_mod = script_mod.replace(old_modal_role, new_modal_role, 1)

# 5. Enhanced table row rendering in loadTutors
# Let's see how tbody.innerHTML is rendered for skeleton and rows
skeleton_html = """
      tbody.innerHTML = Array(8).fill(0).map(() => `
        <tr class="animate-pulse">
          <td class="py-4 px-4"><div class="h-4 w-12 bg-slate-200 rounded"></div></td>
          <td class="py-4 px-4">
            <div class="flex items-center gap-3">
              <div class="w-9 h-9 rounded-full bg-slate-200 shrink-0"></div>
              <div class="space-y-1.5 flex-1">
                <div class="h-3.5 w-32 bg-slate-200 rounded"></div>
                <div class="h-2.5 w-48 bg-slate-100 rounded"></div>
              </div>
            </div>
          </td>
          <td class="py-4 px-4"><div class="h-5 w-24 bg-slate-200 rounded-full"></div></td>
          <td class="py-4 px-4"><div class="h-3.5 w-24 bg-slate-200 rounded"></div></td>
          <td class="py-4 px-4"><div class="h-3.5 w-20 bg-slate-200 rounded"></div></td>
          <td class="py-4 px-4"><div class="h-5 w-40 bg-slate-200 rounded"></div></td>
          <td class="py-4 px-4"><div class="h-5 w-16 bg-slate-200 rounded-full"></div></td>
          <td class="py-4 px-4 text-center"><div class="h-7 w-16 bg-slate-200 rounded-lg mx-auto"></div></td>
        </tr>
      `).join('');
"""

script_mod = re.sub(r'tbody\.innerHTML\s*=\s*`\s*<tr>\s*<td colspan="7".*?</td>\s*</tr>\s*`;', skeleton_html.strip(), script_mod, flags=re.DOTALL)

# Update row HTML generation in loadTutors
old_row_gen_start = "html += `\n            <tr class=\"hover:bg-slate-50/80 transition-colors\">"
# We'll replace the row HTML generator with the modern card/table row
new_row_gen = """
          // Helper for initials
          const names = (t.teacher_name || 'Anonymous').trim().split(' ');
          const initials = (names[0][0] + (names.length > 1 ? names[names.length - 1][0] : '')).toUpperCase();
          const avatarColors = [
            'bg-indigo-100 text-indigo-700 ring-indigo-200',
            'bg-blue-100 text-blue-700 ring-blue-200',
            'bg-emerald-100 text-emerald-700 ring-emerald-200',
            'bg-purple-100 text-purple-700 ring-purple-200',
            'bg-rose-100 text-rose-700 ring-rose-200',
            'bg-amber-100 text-amber-700 ring-amber-200'
          ];
          const avatarColor = avatarColors[Math.abs(t.tutor_id.split('').reduce((a,c)=>a+c.charCodeAt(0),0)) % avatarColors.length];

          // Occupation badge style
          let occBadge = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-medium bg-slate-100 text-slate-700 border border-slate-200">
            ${escapeHtml(t.occupation || t.experience_role || 'Tutor')}
          </span>`;
          if (t.occupation) {
            if (t.occupation.includes('Software') || t.occupation.includes('IT')) {
              occBadge = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200/60">
                <i class="fa-solid fa-code text-[9px] mr-1"></i> Software & IT
              </span>`;
            } else if (t.occupation.includes('Student') || t.occupation.includes('Researcher')) {
              occBadge = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-blue-50 text-blue-700 border border-blue-200/60">
                <i class="fa-solid fa-graduation-cap text-[9px] mr-1"></i> Student / Scholar
              </span>`;
            } else if (t.occupation.includes('Medical') || t.occupation.includes('Healthcare')) {
              occBadge = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-cyan-50 text-cyan-700 border border-cyan-200/60">
                <i class="fa-solid fa-user-doctor text-[9px] mr-1"></i> Medical / Doctor
              </span>`;
            } else if (t.occupation.includes('Chartered Accountant') || t.occupation.includes('Finance')) {
              occBadge = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200/60">
                <i class="fa-solid fa-chart-line text-[9px] mr-1"></i> Finance & CA
              </span>`;
            } else if (t.occupation.includes('Engineer')) {
              occBadge = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-amber-50 text-amber-700 border border-amber-200/60">
                <i class="fa-solid fa-gears text-[9px] mr-1"></i> Core Engineer
              </span>`;
            } else if (t.occupation.includes('Faculty') || t.occupation.includes('School')) {
              occBadge = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-purple-50 text-purple-700 border border-purple-200/60">
                <i class="fa-solid fa-chalkboard-user text-[9px] mr-1"></i> Academic Faculty
              </span>`;
            }
          }

          // Active indicator dot
          const statusDot = t.is_active == 1 
            ? `<span class="inline-flex items-center gap-1 text-[10px] font-semibold text-emerald-600 bg-emerald-50/80 px-1.5 py-0.5 rounded-md border border-emerald-200/50"><span class="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>Active</span>`
            : `<span class="inline-flex items-center text-[10px] text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded-md">Inactive</span>`;

          html += `
            <tr class="hover:bg-indigo-50/30 transition-colors group cursor-pointer" onclick="viewTutorDetails('${t.tutor_id}')">
              <td class="py-3.5 px-4 font-mono font-bold text-xs text-indigo-600">
                <span class="bg-indigo-50/80 group-hover:bg-indigo-100 text-indigo-700 px-2 py-1 rounded-md border border-indigo-200/60 transition-colors">
                  ${t.tutor_id}
                </span>
              </td>
              <td class="py-3.5 px-4">
                <div class="flex items-center gap-3">
                  <div class="w-8 h-8 rounded-full ${avatarColor} ring-1 flex items-center justify-center text-xs font-bold shrink-0">
                    ${initials}
                  </div>
                  <div class="min-w-0">
                    <div class="flex items-center gap-2">
                      <span class="font-semibold text-slate-900 text-sm group-hover:text-indigo-600 transition-colors truncate">${escapeHtml(t.teacher_name || 'Anonymous')}</span>
                      ${statusDot}
                    </div>
                    <div class="text-xs text-slate-500 truncate max-w-xs mt-0.5">${escapeHtml(t.teacher_hookline || 'No headline available')}</div>
                  </div>
                </div>
              </td>
              <td class="py-3.5 px-4">
                <div class="flex flex-col gap-1 items-start">
                  ${occBadge}
                  <span class="text-[10px] text-slate-400 font-medium">${escapeHtml(t.seniority_tier || 'Practitioner')}</span>
                </div>
              </td>
              <td class="py-3.5 px-4">
                <div class="font-medium text-slate-800 text-xs flex items-center gap-1">
                  <i class="fa-solid fa-location-dot text-slate-400 text-[10px]"></i>
                  <span>${escapeHtml(t.city || t.state || 'India')}</span>
                </div>
                <div class="text-[11px] text-slate-500 ml-3.5">${escapeHtml(t.state || '')} <span class="text-slate-400">(${escapeHtml(t.region || '—')})</span></div>
              </td>
              <td class="py-3.5 px-4">
                <div class="text-xs font-semibold text-slate-800">${escapeHtml(t.highest_qualification || 'Degree')}</div>
                <div class="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
                  <i class="fa-regular fa-clock text-slate-400 text-[10px]"></i>
                  <span>${t.total_teaching_exp_years != null ? t.total_teaching_exp_years + ' yrs exp' : '0 yrs exp'}</span>
                </div>
              </td>
              <td class="py-3.5 px-4">
                <div class="flex flex-wrap gap-1 items-center max-w-sm">
                  ${subjsBadges}
                </div>
              </td>
              <td class="py-3.5 px-4">
                <span class="inline-flex items-center px-2.5 py-1 rounded-lg text-xs font-bold font-mono bg-emerald-50 text-emerald-700 border border-emerald-200/80">
                  ${t.hourly_fee_avg ? `₹${Math.round(t.hourly_fee_avg)}<span class="text-[10px] text-emerald-600/70 font-sans ml-0.5">/hr</span>` : 'N/A'}
                </span>
              </td>
              <td class="py-3.5 px-4 text-center" onclick="event.stopPropagation()">
                <button onclick="viewTutorDetails('${t.tutor_id}')" class="px-3 py-1.5 bg-slate-100 hover:bg-indigo-600 hover:text-white text-slate-700 rounded-lg text-xs font-semibold transition-all shadow-2xs flex items-center gap-1.5 mx-auto">
                  <span>Profile</span>
                  <i class="fa-solid fa-arrow-right text-[10px]"></i>
                </button>
              </td>
            </tr>`;
"""

# Replace the row rendering
old_row_match = re.search(r'data\.items\.forEach\(t\s*=>\s*\{(.*?)html\s*\+=\s*`\s*<tr class="hover:bg-slate-50/80.*?</tr>\s*`;\s*\}\);', script_mod, re.DOTALL)
if old_row_match:
    full_block = old_row_match.group(0)
    # extract subjsBadges logic
    subjs_logic = """
        data.items.forEach(t => {
          let subjsBadges = '';
          if (t.subjects && t.subjects.length) {
            subjsBadges = t.subjects.slice(0, 3).map(s => {
              let badgeColor = "bg-slate-100 text-slate-700 border-slate-200";
              if (s.subject_category === 'Science') badgeColor = "bg-blue-50 text-blue-700 border-blue-200/70";
              if (s.subject_category === 'Coding') badgeColor = "bg-purple-50 text-purple-700 border-purple-200/70";
              if (s.subject_category === 'Languages') badgeColor = "bg-amber-50 text-amber-700 border-amber-200/70";
              if (s.subject_category === 'Commerce') badgeColor = "bg-emerald-50 text-emerald-700 border-emerald-200/70";
              return `<span class="inline-flex items-center px-2 py-0.5 rounded-md text-[10.5px] font-medium border ${badgeColor}">
                ${escapeHtml(s.canonical_subject)}
                ${s.level_category ? `<span class="ml-1 opacity-75 text-[9px] font-mono">(${escapeHtml(s.level_category)})</span>` : ''}
              </span>`;
            }).join(' ');

            if (t.subjects.length > 3) {
              subjsBadges += ` <span class="text-[10px] text-slate-500 font-semibold px-1 py-0.5 rounded bg-slate-100 border border-slate-200">+${t.subjects.length - 3}</span>`;
            }
          } else {
            subjsBadges = '<span class="text-xs text-slate-400 italic">None listed</span>';
          }
    """ + new_row_gen + "\n        });"
    script_mod = script_mod.replace(full_block, subjs_logic)

# 6. Add Occupations Chart logic
occ_chart_func = """
    // Occupations Chart (Figure 6)
    let occupationsChartInstance = null;
    let currentOccMode = 'count';

    function switchOccupationsChart(mode) {
      currentOccMode = mode;
      document.getElementById('btn-occ-count').className = mode === 'count'
        ? "px-3 py-1 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-2xs"
        : "px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900";
      document.getElementById('btn-occ-fee').className = mode === 'fee'
        ? "px-3 py-1 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-2xs"
        : "px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900";
      renderOccupationsChart();
    }

    function renderOccupationsChart() {
      const canvas = document.getElementById('chart-occupations');
      if (!canvas || !window.dashboardCharts || !window.dashboardCharts.occupations) return;
      
      const occData = window.dashboardCharts.occupations;
      const sorted = [...occData].sort((a,b) => currentOccMode === 'count' ? a.count - b.count : a.avg_fee - b.avg_fee);
      const labels = sorted.map(d => d.occupation);
      const values = sorted.map(d => currentOccMode === 'count' ? d.count : d.avg_fee);

      if (occupationsChartInstance) occupationsChartInstance.destroy();

      occupationsChartInstance = new Chart(canvas, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [{
            label: currentOccMode === 'count' ? 'Tutor Headcount' : 'Average Hourly Fee (₹/hr)',
            data: values,
            backgroundColor: currentOccMode === 'count' ? '#4f46e5' : '#10b981',
            borderRadius: 6,
            borderSkipped: false
          }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: '#0f172a',
              titleFont: { size: 12, weight: 'bold' },
              padding: 10,
              cornerRadius: 8,
              callbacks: {
                label: function(ctx) {
                  return currentOccMode === 'count' 
                    ? ` ${ctx.raw.toLocaleString()} Registered Tutors`
                    : ` ₹${ctx.raw.toLocaleString()} / hour`;
                }
              }
            }
          },
          scales: {
            x: {
              grid: { color: '#f1f5f9' },
              ticks: {
                callback: function(v) {
                  return currentOccMode === 'count' ? v.toLocaleString() : '₹' + v.toLocaleString();
                },
                font: { family: 'JetBrains Mono', size: 10 }
              }
            },
            y: {
              grid: { display: false },
              ticks: { font: { size: 10.5, weight: '500' } }
            }
          }
        }
      });
    }
"""

# Insert occ_chart_func right before window.addEventListener('DOMContentLoaded'
dom_idx = script_mod.find("window.addEventListener('DOMContentLoaded'")
if dom_idx != -1:
    script_mod = script_mod[:dom_idx] + occ_chart_func + "\n" + script_mod[dom_idx:]

# Call renderOccupationsChart in loadCharts
load_charts_idx = script_mod.find("window.genderAnalyticsData = data;")
if load_charts_idx != -1:
    script_mod = script_mod[:load_charts_idx] + "window.dashboardCharts = data;\n        if (data.occupations) { renderOccupationsChart(); }\n        " + script_mod[load_charts_idx:]

# -------------------------------------------------------------
# Construct the Complete Modern HTML
# -------------------------------------------------------------

full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>BMP Tutors - National Market Intelligence Platform</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" />
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
    body {{ font-family: 'Inter', system-ui, -apple-system, sans-serif; }}
    .font-mono {{ font-family: 'JetBrains Mono', monospace; }}
    .custom-scrollbar::-webkit-scrollbar {{ width: 5px; height: 5px; }}
    .custom-scrollbar::-webkit-scrollbar-thumb {{ background: #cbd5e1; border-radius: 4px; }}
    .custom-scrollbar::-webkit-scrollbar-track {{ background: #f8fafc; }}
    .leaflet-popup-content-wrapper {{ border-radius: 14px; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1); border: 1px solid #e2e8f0; }}
    .white-map-container {{ background: #ffffff !important; }}
    .polygon-number-label {{ background: transparent; border: none; text-align: center; pointer-events: none; }}
    .polygon-number-label span {{
      font-family: 'JetBrains Mono', monospace;
      font-size: 10px;
      font-weight: 700;
      color: #0f172a;
      background: rgba(255, 255, 255, 0.92);
      border: 1px solid rgba(15, 23, 42, 0.15);
      padding: 1px 4px;
      border-radius: 4px;
      box-shadow: 0 1px 2px rgba(0,0,0,0.06);
      display: inline-block;
      white-space: nowrap;
    }}
    .shadow-2xs {{ box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05); }}
    .shadow-xs {{ box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.07), 0 1px 2px -1px rgba(0, 0, 0, 0.07); }}
  </style>
</head>
<body class="bg-slate-50/70 text-slate-800 min-h-screen flex flex-col antialiased selection:bg-indigo-100 selection:text-indigo-900">

  <!-- TOP EXECUTIVE NAVIGATION BAR -->
  <header class="bg-slate-900/95 backdrop-blur-md text-white border-b border-slate-800 sticky top-0 z-40 transition-all">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4">
      
      <!-- Brand & Title -->
      <div class="flex items-center gap-3 shrink-0">
        <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-600 to-violet-500 shadow-md flex items-center justify-center font-bold text-lg text-white ring-1 ring-white/20">
          <i class="fa-solid fa-graduation-cap"></i>
        </div>
        <div>
          <div class="flex items-center gap-2">
            <h1 class="text-base font-extrabold tracking-tight text-white">BMP Tutors</h1>
            <span class="px-2 py-0.5 rounded-full bg-slate-800 border border-slate-700/80 text-[10px] font-semibold text-emerald-400 flex items-center gap-1.5 font-mono">
              <span class="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
              Live Data
            </span>
          </div>
          <p class="text-[11px] text-slate-400 font-medium">National Market Intelligence & Analytics Platform</p>
        </div>
      </div>

      <!-- Segmented Navigation Controller -->
      <nav class="hidden md:flex items-center bg-slate-800/80 p-1 rounded-xl ring-1 ring-white/10 shadow-inner">
        <button id="tab-btn-explorer" onclick="switchTab('explorer')" class="px-4 py-1.5 rounded-lg text-xs font-semibold transition-all bg-indigo-600 text-white shadow-xs flex items-center gap-2">
          <i class="fa-solid fa-table-cells-large"></i>
          <span>Tutors Explorer</span>
        </button>
        <button id="tab-btn-analytics" onclick="switchTab('analytics')" class="px-4 py-1.5 rounded-lg text-xs font-semibold transition-all text-slate-300 hover:text-white hover:bg-slate-700/60 flex items-center gap-2">
          <i class="fa-solid fa-chart-column"></i>
          <span>Visual Analytics</span>
        </button>
        <button id="tab-btn-schema" onclick="switchTab('schema')" class="px-4 py-1.5 rounded-lg text-xs font-semibold transition-all text-slate-300 hover:text-white hover:bg-slate-700/60 flex items-center gap-2">
          <i class="fa-solid fa-layer-group"></i>
          <span>Schema & Exports</span>
        </button>
      </nav>

      <!-- Action Buttons -->
      <div class="flex items-center gap-2 shrink-0">
        <a href="/report" target="_blank" class="hidden sm:inline-flex px-3 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 hover:border-slate-600 transition-all items-center gap-1.5 shadow-2xs">
          <i class="fa-solid fa-file-invoice text-indigo-400"></i>
          <span>Academic Report</span>
          <i class="fa-solid fa-arrow-up-right-from-square text-[9px] opacity-60"></i>
        </a>

        <a href="/api/download/master-excel" download="BMP_Tutors_Cleaned_Master.xlsx" class="px-3.5 py-1.5 rounded-lg text-xs font-bold bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white shadow-xs transition-all flex items-center gap-2 ring-1 ring-emerald-400/30" title="Download Master Cleaned Dataset (107,723 Tutors + 403k Bridge Links + Occupation)">
          <i class="fa-solid fa-file-excel text-sm"></i>
          <span>Master Excel</span>
          <span class="px-1.5 py-0.2 rounded bg-black/25 text-[10px] font-mono">82 MB</span>
        </a>
      </div>
    </div>

    <!-- Mobile Segmented Controller -->
    <div class="flex md:hidden px-4 pb-3 border-t border-slate-800/80 pt-2">
      <nav class="flex w-full bg-slate-800/80 p-1 rounded-xl ring-1 ring-white/10">
        <button onclick="switchTab('explorer')" class="flex-1 py-1.5 rounded-lg text-xs font-semibold text-center bg-indigo-600 text-white">Explorer</button>
        <button onclick="switchTab('analytics')" class="flex-1 py-1.5 rounded-lg text-xs font-semibold text-center text-slate-300">Analytics</button>
        <button onclick="switchTab('schema')" class="flex-1 py-1.5 rounded-lg text-xs font-semibold text-center text-slate-300">Schema</button>
      </nav>
    </div>
  </header>

  <!-- METRIC SUMMARY KPI CARDS -->
  <section class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-6 w-full">
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
      
      <!-- Card 1: Unique Tutors -->
      <div class="bg-white rounded-2xl p-5 shadow-xs border border-slate-200/80 hover:shadow-md hover:border-slate-300 transition-all group relative overflow-hidden">
        <div class="h-1 bg-indigo-500 absolute top-0 left-0 right-0"></div>
        <div class="flex items-center justify-between">
          <div>
            <p class="text-[11px] font-bold uppercase tracking-wider text-slate-500">Verified Unique Tutors</p>
            <h3 id="stat-tutors" class="text-3xl font-extrabold text-slate-900 mt-1 font-mono tracking-tight">107,723</h3>
            <p class="text-xs text-emerald-600 mt-1.5 font-medium flex items-center gap-1">
              <i class="fa-solid fa-circle-check text-[11px]"></i>
              <span>100% Normalized Star Schema</span>
            </p>
          </div>
          <div class="w-12 h-12 rounded-2xl bg-indigo-50 text-indigo-600 flex items-center justify-center text-xl shrink-0 group-hover:scale-105 transition-transform">
            <i class="fa-solid fa-users"></i>
          </div>
        </div>
      </div>

      <!-- Card 2: Bridge Offerings -->
      <div class="bg-white rounded-2xl p-5 shadow-xs border border-slate-200/80 hover:shadow-md hover:border-slate-300 transition-all group relative overflow-hidden">
        <div class="h-1 bg-violet-500 absolute top-0 left-0 right-0"></div>
        <div class="flex items-center justify-between">
          <div>
            <p class="text-[11px] font-bold uppercase tracking-wider text-slate-500">Subject-Grade Linkages</p>
            <h3 id="stat-bridge" class="text-3xl font-extrabold text-slate-900 mt-1 font-mono tracking-tight">403,353</h3>
            <p class="text-xs text-violet-600 mt-1.5 font-medium flex items-center gap-1">
              <i class="fa-solid fa-diagram-project text-[11px]"></i>
              <span>3.7 Subjects per Educator</span>
            </p>
          </div>
          <div class="w-12 h-12 rounded-2xl bg-violet-50 text-violet-600 flex items-center justify-center text-xl shrink-0 group-hover:scale-105 transition-transform">
            <i class="fa-solid fa-book-open"></i>
          </div>
        </div>
      </div>

      <!-- Card 3: Avg Hourly Fee -->
      <div class="bg-white rounded-2xl p-5 shadow-xs border border-slate-200/80 hover:shadow-md hover:border-slate-300 transition-all group relative overflow-hidden">
        <div class="h-1 bg-emerald-500 absolute top-0 left-0 right-0"></div>
        <div class="flex items-center justify-between">
          <div>
            <p class="text-[11px] font-bold uppercase tracking-wider text-slate-500">Benchmark Hourly Rate</p>
            <h3 id="stat-fee" class="text-3xl font-extrabold text-slate-900 mt-1 font-mono tracking-tight">₹416.4 <span class="text-base font-sans font-medium text-slate-500">/hr</span></h3>
            <p class="text-xs text-slate-500 mt-1.5 font-medium flex items-center gap-1">
              <i class="fa-solid fa-shield-halved text-emerald-600 text-[11px]"></i>
              <span>Outlier-Trimmed Mean (₹50-₹5k)</span>
            </p>
          </div>
          <div class="w-12 h-12 rounded-2xl bg-emerald-50 text-emerald-600 flex items-center justify-center text-xl shrink-0 group-hover:scale-105 transition-transform">
            <i class="fa-solid fa-indian-rupee-sign"></i>
          </div>
        </div>
      </div>

      <!-- Card 4: Active Profiles -->
      <div class="bg-white rounded-2xl p-5 shadow-xs border border-slate-200/80 hover:shadow-md hover:border-slate-300 transition-all group relative overflow-hidden">
        <div class="h-1 bg-amber-500 absolute top-0 left-0 right-0"></div>
        <div class="flex items-center justify-between">
          <div>
            <p class="text-[11px] font-bold uppercase tracking-wider text-slate-500">Live Active Profiles</p>
            <h3 id="stat-active" class="text-3xl font-extrabold text-slate-900 mt-1 font-mono tracking-tight">76,706</h3>
            <p class="text-xs text-amber-600 mt-1.5 font-medium flex items-center gap-1">
              <i class="fa-solid fa-bolt text-[11px]"></i>
              <span>71.2% URL Verified Active Rate</span>
            </p>
          </div>
          <div class="w-12 h-12 rounded-2xl bg-amber-50 text-amber-600 flex items-center justify-center text-xl shrink-0 group-hover:scale-105 transition-transform">
            <i class="fa-solid fa-user-check"></i>
          </div>
        </div>
      </div>

    </div>
  </section>

  <!-- MAIN TAB CONTENT AREA -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 w-full flex-1 flex flex-col">

    <!-- TAB 1: EXPLORER -->
    <div id="tab-explorer" class="flex flex-col lg:flex-row gap-6 flex-1">
      
      <!-- SIDEBAR FACETED FILTERS -->
      <aside class="w-full lg:w-80 bg-white p-5 rounded-2xl shadow-xs border border-slate-200/80 flex-shrink-0 flex flex-col gap-4">
        <div class="flex items-center justify-between border-b border-slate-100 pb-3">
          <div class="flex items-center gap-2">
            <div class="w-6 h-6 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center text-xs">
              <i class="fa-solid fa-sliders"></i>
            </div>
            <h2 class="font-bold text-slate-900 text-sm">Filter Dataset</h2>
          </div>
          <button onclick="resetFilters()" class="text-xs text-slate-500 hover:text-indigo-600 font-semibold flex items-center gap-1 transition-colors">
            <i class="fa-solid fa-rotate-left text-[10px]"></i> Reset
          </button>
        </div>

        <!-- Search input -->
        <div>
          <label class="block text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Keyword Search</label>
          <div class="relative">
            <input type="text" id="filter-search" placeholder="ID (T1), name, city, hookline..." class="w-full pl-9 pr-3 py-2 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs" />
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
          </div>
        </div>

        <!-- Facet Group 1: Academics & Subject -->
        <div class="space-y-3 pt-1">
          <div class="text-[10px] font-bold uppercase tracking-wider text-indigo-700 bg-indigo-50/60 px-2 py-1 rounded-md">
            Academics & Curriculum
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Subject Category</label>
            <select id="filter-subject-category" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Categories</option>
            </select>
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Level / Grade</label>
            <select id="filter-level" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Levels</option>
            </select>
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Competitive Exam</label>
            <select id="filter-exam" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Exams (JEE, NEET, IELTS...)</option>
            </select>
          </div>
        </div>

        <!-- Facet Group 2: Geography -->
        <div class="space-y-3 pt-2 border-t border-slate-100">
          <div class="text-[10px] font-bold uppercase tracking-wider text-blue-700 bg-blue-50/60 px-2 py-1 rounded-md">
            Geographic Territory
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Region</label>
            <select id="filter-region" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Regions</option>
            </select>
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">State / UT</label>
            <select id="filter-state" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All States & UTs</option>
            </select>
          </div>
        </div>

        <!-- Facet Group 3: Professional & Career -->
        <div class="space-y-3 pt-2 border-t border-slate-100">
          <div class="text-[10px] font-bold uppercase tracking-wider text-emerald-700 bg-emerald-50/60 px-2 py-1 rounded-md">
            Professional & Credentials
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Occupation (Side / Primary)</label>
            <select id="filter-occupation" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Occupations (IT, CA, Doctors...)</option>
            </select>
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Seniority Tier</label>
            <select id="filter-seniority" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Seniority Tiers</option>
            </select>
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Experience Role</label>
            <select id="filter-role" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Roles</option>
            </select>
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Highest Qualification</label>
            <select id="filter-qualification" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Qualifications</option>
            </select>
          </div>
        </div>

        <!-- Facet Group 4: Preferences -->
        <div class="space-y-3 pt-2 border-t border-slate-100">
          <div class="text-[10px] font-bold uppercase tracking-wider text-amber-700 bg-amber-50/60 px-2 py-1 rounded-md">
            Delivery & Demographics
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Gender</label>
            <select id="filter-gender" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">All Genders</option>
              <option value="Female">Female</option>
              <option value="Male">Male</option>
            </select>
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Teaches Online</label>
            <select id="filter-online" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
              <option value="">Any Mode</option>
              <option value="1">Yes (Online)</option>
              <option value="0">No (Offline only)</option>
            </select>
          </div>
        </div>

        <button onclick="applyFilters()" class="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 active:scale-[0.98] text-white rounded-xl text-xs font-bold shadow-xs transition-all mt-2 flex items-center justify-center gap-2">
          <i class="fa-solid fa-magnifying-glass text-[11px]"></i>
          <span>Apply Filters</span>
        </button>
      </aside>

      <!-- TABLE & PAGINATION CONTAINER -->
      <section class="flex-1 bg-white rounded-2xl shadow-xs border border-slate-200/80 flex flex-col overflow-hidden">
        
        <!-- Results Header Toolbar -->
        <div class="px-6 py-4 border-b border-slate-100 flex flex-col md:flex-row items-start md:items-center justify-between gap-3 bg-slate-50/40">
          <div class="flex items-center gap-2.5">
            <h3 class="font-bold text-slate-900 text-sm">Tutor Profiles</h3>
            <span id="results-count-badge" class="text-xs font-semibold px-2.5 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200/70 font-mono">
              Loading...
            </span>
          </div>
          
          <!-- Download Buttons & Page size selector -->
          <div class="flex flex-wrap items-center gap-2.5 text-xs text-slate-600 w-full md:w-auto justify-between md:justify-end">
            <!-- Downloads -->
            <div class="flex items-center gap-2">
              <select onchange="if(this.value){{window.location.href='/api/download/category-excel?category='+this.value; this.value='';}}" class="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg text-xs font-semibold text-slate-700 transition-all cursor-pointer" title="Download fast category-specific thematic Excel files">
                <option value="">📁 By Domain...</option>
                <option value="coding">💻 Coding & Tech (~15 MB)</option>
                <option value="science">🔬 Science & Math (~62 MB)</option>
                <option value="languages">🗣️ Languages (~27 MB)</option>
                <option value="commerce">📈 Commerce & Finance (~7 MB)</option>
                <option value="arts">🎨 Arts & Humanities (~8 MB)</option>
                <option value="eca">♟️ Extracurricular & Skills (~6 MB)</option>
              </select>

              <button onclick="downloadFilteredExcel()" class="px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 rounded-lg text-xs font-semibold transition-all flex items-center gap-1.5 shadow-2xs" title="Download matching rows based on current filters">
                <i class="fa-solid fa-filter text-[10px]"></i>
                <span>Filtered Export</span>
              </button>
            </div>

            <div class="flex items-center gap-1.5 pl-2 border-l border-slate-200">
              <span class="text-slate-500 font-medium text-[11px]">Rows:</span>
              <select id="page-size" onchange="changePageSize()" class="bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs focus:outline-none focus:border-indigo-500 shadow-2xs">
                <option value="15">15</option>
                <option value="25" selected>25</option>
                <option value="50">50</option>
                <option value="100">100</option>
              </select>
            </div>
          </div>
        </div>

        <!-- Table Container -->
        <div class="flex-1 overflow-x-auto custom-scrollbar">
          <table class="w-full text-left border-collapse text-sm">
            <thead>
              <tr class="bg-slate-50/80 border-b border-slate-200/80 text-[11px] font-bold text-slate-500 uppercase tracking-wider sticky top-0 backdrop-blur-sm z-10">
                <th class="py-3 px-4">Tutor ID</th>
                <th class="py-3 px-4">Educator Profile</th>
                <th class="py-3 px-4">Occupation & Role</th>
                <th class="py-3 px-4">Location</th>
                <th class="py-3 px-4">Qualification & Exp</th>
                <th class="py-3 px-4">Subjects Taught (Bridge)</th>
                <th class="py-3 px-4">Fee / hr</th>
                <th class="py-3 px-4 text-center">Action</th>
              </tr>
            </thead>
            <tbody id="tutors-table-body" class="divide-y divide-slate-100 text-slate-700">
              <tr>
                <td colspan="8" class="py-16 text-center text-slate-400">
                  <i class="fa-solid fa-circle-notch fa-spin text-2xl text-indigo-500 mb-2 block mx-auto"></i>
                  <p class="text-xs font-medium text-slate-500">Loading verified educator records...</p>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- Pagination Controls -->
        <div class="px-6 py-3.5 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3 bg-slate-50/50">
          <div id="page-info" class="text-xs text-slate-500 font-medium">
            Showing 1 to 25 of 107,723 entries
          </div>
          <div class="flex items-center space-x-1">
            <button onclick="goToPage(1)" class="px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white text-xs hover:bg-slate-50 disabled:opacity-40 shadow-2xs transition-colors" id="btn-first" title="First Page">
              <i class="fa-solid fa-angles-left text-[10px]"></i>
            </button>
            <button onclick="changePage(-1)" class="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-xs hover:bg-slate-50 disabled:opacity-40 shadow-2xs transition-colors flex items-center gap-1" id="btn-prev">
              <i class="fa-solid fa-chevron-left text-[10px]"></i> <span>Prev</span>
            </button>
            <span id="current-page-display" class="px-3 py-1.5 text-xs font-bold text-indigo-700 bg-indigo-50 border border-indigo-200/80 rounded-lg font-mono">1</span>
            <button onclick="changePage(1)" class="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-xs hover:bg-slate-50 disabled:opacity-40 shadow-2xs transition-colors flex items-center gap-1" id="btn-next">
              <span>Next</span> <i class="fa-solid fa-chevron-right text-[10px]"></i>
            </button>
            <button onclick="goToLastPage()" class="px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white text-xs hover:bg-slate-50 disabled:opacity-40 shadow-2xs transition-colors" id="btn-last" title="Last Page">
              <i class="fa-solid fa-angles-right text-[10px]"></i>
            </button>
          </div>
        </div>
      </section>
    </div>

    <!-- TAB 2: VISUAL ANALYTICS & GEOGRAPHIC MAPS -->
    <div id="tab-analytics" class="hidden flex-col gap-6">

      <!-- VISUAL ANALYTICS HEADER & DOWNLOAD BAR -->
      <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
        <div>
          <div class="flex items-center gap-2">
            <span class="px-2.5 py-0.5 rounded-md bg-indigo-50 text-indigo-700 text-[10px] font-bold uppercase tracking-wider font-mono border border-indigo-200/60">Empirical Research</span>
            <span class="text-xs text-slate-400 font-mono">107,723 Tutors &bull; 403,353 Offerings &bull; 37 States & UTs</span>
          </div>
          <h2 class="text-xl font-extrabold text-slate-900 mt-1 flex items-center gap-2 tracking-tight">
            <i class="fa-solid fa-chart-line text-indigo-600"></i> Visual Analytics & National Market Dossier
          </h2>
          <p class="text-xs text-slate-500 mt-0.5">Empirical breakdown of private tutoring demand, hourly fee benchmarks, subject distributions, and district-level cartography.</p>
        </div>

        <!-- Download & Export Actions -->
        <div class="flex items-center gap-2 flex-wrap shrink-0">
          <a href="/report" target="_blank" class="px-3.5 py-2 rounded-xl text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 transition-all flex items-center gap-1.5 shadow-2xs">
            <i class="fa-solid fa-arrow-up-right-from-square text-slate-500"></i> Full Printable Dossier
          </a>
          <a href="/api/download/visual-analytics-docx" download="BMP_Tutors_Visual_Analytics_Report.docx" class="px-4 py-2 rounded-xl text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white transition-all flex items-center gap-2 shadow-xs active:scale-[0.98]">
            <i class="fa-solid fa-file-word text-sm"></i> Download Word (.docx)
          </a>
          <a href="/api/export/visual-analytics-pdf" download="BMP_Tutors_Visual_Analytics_Complete_Report.pdf" class="px-4 py-2 rounded-xl text-xs font-bold bg-rose-600 hover:bg-rose-700 text-white transition-all flex items-center gap-2 shadow-xs active:scale-[0.98]">
            <i class="fa-solid fa-file-pdf text-sm"></i> Download PDF
          </a>
        </div>
      </div>

      <!-- SECTION 1: ALL-INDIA ACADEMIC BOUNDARY MAP -->
      <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
        <div class="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4 mb-4 pb-3 border-b border-slate-100">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2 py-0.5 rounded-md bg-indigo-50 text-indigo-700 text-[10px] font-bold uppercase tracking-wider font-mono border border-indigo-200/60">Cartographic Overview</span>
              <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">
                <i class="fa-solid fa-map text-indigo-600"></i> Map of India: State-Wise Boundaries & Data Labels
              </h3>
            </div>
            <p class="text-xs text-slate-500 mt-1">State administrative boundaries with direct in-polygon verified numbers and choropleth shading. Click any state boundary to view district spatial breakdown.</p>
          </div>

          <!-- Controls: Metric Toggle & View Mode -->
          <div class="flex flex-wrap items-center gap-2">
            <span class="text-xs font-bold text-slate-500 uppercase tracking-wider text-[10px]">Active Metric:</span>
            <div class="bg-slate-100 p-1 rounded-xl flex items-center space-x-1 border border-slate-200">
              <button id="btn-metric-count" onclick="setMapMetric('count')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-2xs">
                Tutor Count
              </button>
              <button id="btn-metric-fee" onclick="setMapMetric('fee')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
                Avg Fee (₹/hr)
              </button>
              <button id="btn-metric-online" onclick="setMapMetric('online')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">
                Online %
              </button>
            </div>
          </div>
        </div>

        <!-- Dynamic Color Scale Legend -->
        <div class="flex flex-wrap items-center justify-between gap-3 mb-3 p-2.5 bg-slate-50 rounded-xl border border-slate-200 text-xs">
          <div class="flex items-center gap-2">
            <span class="font-bold text-slate-700 text-[11px]" id="legend-title">Tutor Count Range:</span>
            <div id="legend-scale-bar" class="flex items-center space-x-1"></div>
          </div>
          <div class="text-[11px] text-slate-500 flex items-center gap-2 font-medium">
            <span><i class="fa-solid fa-square text-indigo-800 mr-1"></i>High Concentration</span>
            <span><i class="fa-solid fa-square text-purple-200 mr-1"></i>Low / Emerging</span>
          </div>
        </div>

        <!-- Academic White India Map Container -->
        <div class="relative w-full rounded-2xl overflow-hidden border border-slate-300 shadow-xs bg-white" style="height: 520px;">
          <div id="map-india" class="w-full h-full z-10 white-map-container"></div>
          <div class="absolute bottom-4 left-4 z-20 bg-white/95 backdrop-blur-md p-3.5 rounded-xl shadow-md border border-slate-200 text-xs max-w-xs">
            <p class="font-bold text-slate-800 flex items-center gap-1.5">
              <i class="fa-solid fa-hand-pointer text-indigo-600"></i> Interactive Drilldown:
            </p>
            <p class="text-slate-600 text-[11px] mt-1 leading-relaxed">Labels inside boundaries show exact state values. Click any state boundary to view district-level boundaries below.</p>
          </div>
        </div>
      </div>

      <!-- SECTION 2: STATE-LEVEL DISTRICT BOUNDARY MAP -->
      <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
        <div class="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4 mb-4 pb-3 border-b border-slate-100">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-700 text-[10px] font-bold uppercase tracking-wider font-mono border border-emerald-200/60">District Spatial Deep-Dive</span>
              <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">
                <i class="fa-solid fa-draw-polygon text-emerald-600"></i> State District Boundaries & Spatial Distribution
              </h3>
            </div>
            <p class="text-xs text-slate-500 mt-1">Detailed administrative district polygons for the selected state with local tutor numbers and pricing metrics.</p>
          </div>

          <!-- State Selector Controls -->
          <div class="flex flex-wrap items-center gap-2 w-full sm:w-auto">
            <label class="text-xs font-bold text-slate-500 uppercase tracking-wider text-[10px]">Select State:</label>
            <select id="map-state-select" onchange="onStateSelectChange(this.value)" class="bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-semibold text-slate-800 focus:outline-none focus:border-indigo-500 shadow-2xs">
              <option value="Karnataka">Karnataka</option>
            </select>
          </div>
        </div>

        <!-- State Quick Metric Highlights -->
        <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4">
          <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200/80">
            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">State Analyzed</span>
            <span id="state-metric-name" class="font-extrabold text-indigo-700 text-sm mt-0.5 block">Karnataka</span>
          </div>
          <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200/80">
            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Total Tutors</span>
            <span id="state-metric-tutors" class="font-extrabold text-slate-900 text-sm mt-0.5 block font-mono">—</span>
          </div>
          <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200/80">
            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Avg Hourly Fee</span>
            <span id="state-metric-fee" class="font-extrabold text-emerald-600 text-sm mt-0.5 block font-mono">—</span>
          </div>
          <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200/80">
            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Online Delivery</span>
            <span id="state-metric-online" class="font-extrabold text-violet-600 text-sm mt-0.5 block font-mono">—</span>
          </div>
        </div>

        <!-- Split Map + District Table -->
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
          <div class="lg:col-span-7 rounded-2xl overflow-hidden border border-slate-300 shadow-xs relative bg-white" style="height: 480px;">
            <div id="map-state-cities" class="w-full h-full z-10 white-map-container"></div>
          </div>

          <div class="lg:col-span-5 bg-slate-50/70 rounded-2xl border border-slate-200/80 p-4 h-[480px] flex flex-col shadow-xs">
            <div class="flex items-center justify-between pb-3 border-b border-slate-200">
              <h4 class="font-bold text-xs text-slate-800 flex items-center gap-1.5">
                <i class="fa-solid fa-list-ol text-emerald-600"></i>
                <span id="city-table-state-title">Districts in Karnataka</span>
              </h4>
              <span id="city-count-badge" class="text-[10px] font-bold bg-white px-2 py-0.5 rounded-full border border-slate-200 text-slate-600 font-mono">0 Districts</span>
            </div>
            <div class="flex-1 overflow-y-auto custom-scrollbar mt-2 pr-1">
              <table class="w-full text-left text-xs">
                <thead>
                  <tr class="text-[10px] font-bold text-slate-500 uppercase border-b border-slate-200">
                    <th class="py-2">District</th>
                    <th class="py-2 text-right">Tutors</th>
                    <th class="py-2 text-right">Avg Fee</th>
                  </tr>
                </thead>
                <tbody id="city-table-body" class="divide-y divide-slate-100 font-medium">
                  <tr><td colspan="3" class="py-8 text-center text-slate-400">Loading districts...</td></tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>

      <!-- NEW SECTION: PROFESSIONAL LANDSCAPE — NON-TEACHING & SIDE OCCUPATIONS -->
      <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
        <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-4 pb-3 border-b border-slate-100">
          <div>
            <div class="flex items-center gap-2">
              <span class="px-2 py-0.5 rounded-md bg-purple-50 text-purple-700 text-[10px] font-bold uppercase tracking-wider font-mono border border-purple-200/60">Professional Gig Economy</span>
              <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">
                <i class="fa-solid fa-briefcase text-purple-600"></i> Non-Teaching & Side Occupations Landscape (N = 36,525)
              </h3>
            </div>
            <p class="text-xs text-slate-500 mt-1">Cross-disciplinary breakdown of 36,500+ tutors practicing an outside profession (IT, CA, Doctors, Research Scholars, Engineers).</p>
          </div>

          <div class="bg-slate-100 p-1 rounded-xl flex items-center space-x-1 border border-slate-200 text-xs">
            <button id="btn-occ-count" onclick="switchOccupationsChart('count')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-2xs">Tutor Volume</button>
            <button id="btn-occ-fee" onclick="switchOccupationsChart('fee')" class="px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900">Avg Fee (₹/hr)</button>
          </div>
        </div>

        <div class="h-80 my-2 relative">
          <canvas id="chart-occupations"></canvas>
        </div>

        <p class="text-[11px] text-slate-500 italic pt-2 border-t border-slate-100">
          <strong>Key Insight:</strong> Independent consultants, Chartered Accountants, and Software Engineers command an 80% fee premium (₹1,050–₹1,180/hr) over entry-level university peer tutors (₹767/hr).
        </p>
      </div>

      <!-- SECTION 3: ACADEMIC FIGURES (FIGURE 1 & FIGURE 2) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <!-- Figure 1 -->
        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80 flex flex-col justify-between">
          <div>
            <div class="flex items-center justify-between gap-2 mb-2">
              <h3 class="text-base font-bold text-slate-900 flex items-center gap-2">
                <i class="fa-solid fa-chart-column text-indigo-600"></i> Figure 1: Subjects Offered for Tuition
              </h3>
              <div class="bg-slate-100 p-0.5 rounded-lg flex items-center space-x-0.5 border border-slate-200 text-[11px]">
                <button id="btn-fig1-macro" onclick="switchFigure1('macro')" class="px-2.5 py-0.5 rounded-md font-bold bg-indigo-600 text-white shadow-2xs">Macro Clusters</button>
                <button id="btn-fig1-top" onclick="switchFigure1('top')" class="px-2.5 py-0.5 rounded-md font-medium text-slate-600 hover:text-slate-900">Top 12 Disciplines</button>
              </div>
            </div>
            <p class="text-xs text-slate-500 mt-0.5">Categorical volume of private tutoring offerings across India (N = 403,353).</p>
          </div>

          <div class="h-72 my-4 relative">
            <canvas id="chart-figure-1"></canvas>
          </div>

          <p class="text-[11px] text-slate-500 italic pt-2 border-t border-slate-100">
            <strong>Notes:</strong> Values indicate absolute frequency (N) and percentage share across all 403,353 subject links.
          </p>
        </div>

        <!-- Figure 2: Experience Structure -->
        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80 flex flex-col justify-between">
          <div>
            <div class="flex items-center justify-between gap-2 mb-2">
              <h3 class="text-base font-bold text-slate-900 flex items-center gap-2">
                <i class="fa-solid fa-chart-line text-indigo-600"></i> Figure 2: Teaching Experience Structure
              </h3>
              <div class="bg-slate-100 p-0.5 rounded-lg flex items-center space-x-0.5 border border-slate-200 text-[11px]">
                <button id="btn-fig3-total" onclick="switchFigure3('total')" class="px-2.5 py-0.5 rounded-md font-bold bg-indigo-600 text-white shadow-2xs">Total Curve</button>
                <button id="btn-fig3-gender" onclick="switchFigure3('gender')" class="px-2.5 py-0.5 rounded-md font-medium text-slate-600 hover:text-slate-900">Men vs Women</button>
              </div>
            </div>
            <p class="text-xs text-slate-500 mt-0.5">Empirical distribution of teaching experience across tutors (0 to 30 years).</p>
          </div>

          <!-- Summary Badges -->
          <div class="grid grid-cols-4 gap-2 mt-2 text-center text-xs">
            <div class="bg-slate-50 p-2 rounded-xl border border-slate-100">
              <div class="text-[10px] text-slate-400 uppercase font-bold">Peak Mode</div>
              <div class="font-bold font-mono text-indigo-600 text-sm">1–2 Yrs</div>
            </div>
            <div class="bg-slate-50 p-2 rounded-xl border border-slate-100">
              <div class="text-[10px] text-slate-400 uppercase font-bold">Median Exp</div>
              <div class="font-bold font-mono text-slate-800 text-sm">3.0 Yrs</div>
            </div>
            <div class="bg-slate-50 p-2 rounded-xl border border-slate-100">
              <div class="text-[10px] text-slate-400 uppercase font-bold">Mean Exp</div>
              <div class="font-bold font-mono text-slate-800 text-sm">4.6 Yrs</div>
            </div>
            <div class="bg-slate-50 p-2 rounded-xl border border-slate-100">
              <div class="text-[10px] text-slate-400 uppercase font-bold">&le; 3 Yrs Share</div>
              <div class="font-bold font-mono text-slate-800 text-sm">47.3%</div>
            </div>
          </div>

          <div class="h-60 my-2 relative">
            <canvas id="chart-figure-3"></canvas>
          </div>

          <p class="text-[11px] text-slate-500 italic pt-2 border-t border-slate-100">
            <strong>Notes:</strong> Shows early-career entry peaking at 1–2 years with a long professional tail up to 30 years.
          </p>
        </div>
      </div>

      <!-- SECTION 4: ACADEMIC TABLES (TABLE 2 & TABLE 3) -->
      <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
        <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-4 pb-3 border-b border-slate-100">
          <div>
            <span class="px-2 py-0.5 rounded-md bg-rose-50 text-rose-700 text-[10px] font-bold uppercase tracking-wider font-mono border border-rose-200/60">Demographic Breakdown</span>
            <h3 class="text-base font-bold text-slate-900 mt-1">Table 2: Gender Distribution by Tutoring Subject</h3>
            <p class="text-xs text-slate-500 mt-0.5">Disaggregated participation rates for women vs men across academic domains.</p>
          </div>
          <div class="flex items-center gap-2">
            <div class="bg-slate-100 p-0.5 rounded-lg flex items-center space-x-0.5 border border-slate-200 text-xs">
              <button id="btn-gender-cats" onclick="switchGenderTable('cats')" class="px-3 py-1 rounded-md font-bold bg-indigo-600 text-white shadow-2xs">Broad Domains</button>
              <button id="btn-gender-subs" onclick="switchGenderTable('subs')" class="px-3 py-1 rounded-md font-medium text-slate-600 hover:text-slate-900">Top Subjects</button>
            </div>
            <button onclick="copyGenderTable()" class="px-3 py-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-xs font-semibold text-slate-700 transition-colors flex items-center gap-1 shadow-2xs">
              <i class="fa-solid fa-copy text-[10px]"></i> <span id="copy-table-btn-text">Copy Table</span>
            </button>
          </div>
        </div>

        <div class="overflow-x-auto custom-scrollbar">
          <table id="academic-gender-table" class="w-full text-left border-collapse text-xs">
            <thead>
              <tr class="bg-slate-900 text-white font-semibold">
                <th class="py-2.5 px-4 rounded-tl-lg">Subject Domain / Discipline</th>
                <th class="py-2.5 px-4 text-right">Men (%)</th>
                <th class="py-2.5 px-4 text-right">Women (%)</th>
                <th class="py-2.5 px-4 text-right rounded-tr-lg">Sample Size (N)</th>
              </tr>
            </thead>
            <tbody id="gender-table-body" class="divide-y divide-slate-100 font-medium"></tbody>
            <tfoot id="gender-table-foot" class="border-t-2 border-slate-300 font-bold bg-slate-100"></tfoot>
          </table>
        </div>
      </div>

      <!-- SECTION 5: HOURLY TUITION FEES TABLE (TABLE 3) -->
      <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
        <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-4 pb-3 border-b border-slate-100">
          <div>
            <span class="px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-700 text-[10px] font-bold uppercase tracking-wider font-mono border border-emerald-200/60">Pricing Architecture</span>
            <h3 class="text-base font-bold text-slate-900 mt-1">Table 3: Hourly Tuition Fees by Educational Stage & Geographic Zone</h3>
            <p class="text-xs text-slate-500 mt-0.5">Empirical pricing escalations across primary, secondary, senior secondary, and higher education.</p>
          </div>
          <div class="flex items-center gap-2">
            <div class="bg-slate-100 p-0.5 rounded-lg flex items-center space-x-0.5 border border-slate-200 text-xs">
              <button id="btn-fee-me" onclick="switchFeeTable('me')" class="px-2.5 py-1 rounded-md font-bold bg-indigo-600 text-white shadow-2xs">Maths & English</button>
              <button id="btn-fee-all" onclick="switchFeeTable('all')" class="px-2.5 py-1 rounded-md font-medium text-slate-600 hover:text-slate-900">All 5 Subjects</button>
              <button id="btn-fee-matrix" onclick="switchFeeTable('matrix')" class="px-2.5 py-1 rounded-md font-medium text-slate-600 hover:text-slate-900">Cross-Matrix</button>
            </div>
            <button onclick="copyFeeTable()" class="px-3 py-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-xs font-semibold text-slate-700 transition-colors flex items-center gap-1 shadow-2xs">
              <i class="fa-solid fa-copy text-[10px]"></i> <span id="copy-fee-table-btn-text">Copy Table</span>
            </button>
          </div>
        </div>

        <div class="overflow-x-auto custom-scrollbar">
          <table id="academic-fee-table" class="w-full text-left border-collapse text-xs">
            <thead id="fee-table-head"></thead>
            <tbody id="fee-table-body" class="divide-y divide-slate-100 font-medium"></tbody>
          </table>
        </div>
      </div>

      <!-- SECTION 6: CORE MARKET STRUCTURE CHARTS GRID -->
      <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
          <h3 class="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
            <i class="fa-solid fa-shapes text-indigo-600"></i> Subject Offerings Distribution
          </h3>
          <p class="text-xs text-slate-500 mb-4">Total breakdown across 403,353 subject-tutor links</p>
          <div class="h-64"><canvas id="chart-subjects"></canvas></div>
        </div>

        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
          <h3 class="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
            <i class="fa-solid fa-graduation-cap text-purple-600"></i> Target Educational Levels
          </h3>
          <p class="text-xs text-slate-500 mb-4">100% concrete educational stages across offerings</p>
          <div class="h-64"><canvas id="chart-levels"></canvas></div>
        </div>

        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
          <h3 class="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
            <i class="fa-solid fa-user-graduate text-blue-600"></i> Highest Academic Qualifications
          </h3>
          <p class="text-xs text-slate-500 mb-4">Undergraduate, Postgraduate, Doctorate credentials</p>
          <div class="h-64"><canvas id="chart-qualifications"></canvas></div>
        </div>

        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
          <h3 class="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
            <i class="fa-solid fa-award text-amber-600"></i> Pedagogical Seniority Tiers
          </h3>
          <p class="text-xs text-slate-500 mb-4">Mastery tiers inferred from experience, degrees, fee rates</p>
          <div class="h-64"><canvas id="chart-seniority"></canvas></div>
        </div>

        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
          <h3 class="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
            <i class="fa-solid fa-indian-rupee-sign text-emerald-600"></i> Hourly Fees by Subject Domain
          </h3>
          <p class="text-xs text-slate-500 mb-4">Standardized mean hourly tuition rates (INR / hour)</p>
          <div class="h-64"><canvas id="chart-fees"></canvas></div>
        </div>

        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
          <h3 class="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
            <i class="fa-solid fa-bullseye text-rose-600"></i> Competitive Exam Coaching Frequency
          </h3>
          <p class="text-xs text-slate-500 mb-4">Top 10 specialized target examinations (JEE, NEET, IELTS...)</p>
          <div class="h-64"><canvas id="chart-exams"></canvas></div>
        </div>

        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
          <h3 class="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
            <i class="fa-solid fa-compass text-teal-600"></i> Geographic Supply by Zonal Region
          </h3>
          <p class="text-xs text-slate-500 mb-4">Distribution across South, North, West, East, Central</p>
          <div class="h-64"><canvas id="chart-regions"></canvas></div>
        </div>

        <div class="bg-white p-6 rounded-2xl shadow-xs border border-slate-200/80">
          <h3 class="text-base font-bold text-slate-900 mb-1 flex items-center gap-2">
            <i class="fa-solid fa-city text-cyan-600"></i> City Tier Distribution
          </h3>
          <p class="text-xs text-slate-500 mb-4">Metros (Tier 1) vs Urban Hubs (Tier 2) vs Semi-Urban (Tier 3)</p>
          <div class="h-64"><canvas id="chart-city-tiers"></canvas></div>
        </div>
      </div>
    </div>

    <!-- TAB 3: SCHEMA & EXPORTS -->
    <div id="tab-schema" class="hidden flex-col gap-6">
      <div class="bg-white p-8 rounded-2xl shadow-xs border border-slate-200/80">
        <div class="flex items-center gap-3 mb-4">
          <div class="w-10 h-10 rounded-xl bg-indigo-50 text-indigo-600 flex items-center justify-center text-xl">
            <i class="fa-solid fa-network-wired"></i>
          </div>
          <div>
            <h3 class="text-lg font-bold text-slate-900 tracking-tight">Relational Star Schema & Bridge Architecture</h3>
            <p class="text-xs text-slate-500">Zero data redundancy, zero wide sparsity, relational 3NF production standard</p>
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mt-6">
          <div class="p-5 rounded-xl bg-slate-50 border border-slate-200/80">
            <div class="flex items-center justify-between mb-3">
              <span class="font-bold text-sm text-slate-800 flex items-center gap-2">
                <i class="fa-solid fa-table-list text-blue-600"></i> Table: `teachers` (Master Dimension)
              </span>
              <span class="text-xs bg-blue-100 text-blue-800 px-2.5 py-0.5 rounded-full font-mono font-bold">107,723 rows</span>
            </div>
            <p class="text-xs text-slate-600 mb-3 leading-relaxed">Stores educator profile attributes exactly once, preventing repeating text contamination across subjects.</p>
            <ul class="text-xs space-y-1 font-mono text-slate-700 bg-white p-3 rounded-lg border border-slate-200/80">
              <li>• <span class="text-indigo-600 font-bold">tutor_id</span> (PK: 'T1'..'T107723')</li>
              <li>• teacher_name, hookline, description</li>
              <li>• state, region, city, city_tier, location_raw</li>
              <li>• gender, works_as</li>
              <li>• <span class="text-emerald-700 font-bold">occupation</span> (100% classified domain)</li>
              <li>• experience_role, seniority_tier, total_teaching_exp_years</li>
              <li>• highest_qualification, has_education_degree</li>
              <li>• hourly_fee_min, hourly_fee_max, hourly_fee_avg</li>
              <li>• teaches_online, teaches_at_student_home, can_travel</li>
            </ul>
          </div>

          <div class="p-5 rounded-xl bg-slate-50 border border-slate-200/80">
            <div class="flex items-center justify-between mb-3">
              <span class="font-bold text-sm text-slate-800 flex items-center gap-2">
                <i class="fa-solid fa-diagram-project text-indigo-600"></i> Table: `teacher_subjects` (Relational Bridge)
              </span>
              <span class="text-xs bg-indigo-100 text-indigo-800 px-2.5 py-0.5 rounded-full font-mono font-bold">403,353 rows</span>
            </div>
            <p class="text-xs text-slate-600 mb-3 leading-relaxed">Narrow junction table mapping each tutor to their specific subject disciplines, grade levels, and competitive tests.</p>
            <ul class="text-xs space-y-1 font-mono text-slate-700 bg-white p-3 rounded-lg border border-slate-200/80">
              <li>• bridge_id (PK: Auto increment)</li>
              <li>• <span class="text-indigo-600 font-bold">tutor_id</span> (FK: References teachers.tutor_id)</li>
              <li>• canonical_subject (Standardized subject name)</li>
              <li>• subject_category (Science, Coding, Languages...)</li>
              <li>• level_category (Secondary, Sr Secondary, College...)</li>
              <li>• competitive_exam (JEE, NEET, IELTS, UPSC...)</li>
              <li>• raw_subject_string (Original uncleaned text token)</li>
            </ul>
          </div>
        </div>

        <div class="mt-8 border-t border-slate-200 pt-6">
          <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-4">
            <div>
              <h4 class="font-bold text-slate-900 text-sm">Direct Deliverable Downloads & Export Formats</h4>
              <p class="text-xs text-slate-500">All data exported in open standard formats with 100% original and cleaned attributes</p>
            </div>
            <a href="/api/download/master-excel" download="BMP_Tutors_Cleaned_Master.xlsx" class="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold shadow-xs transition-all flex items-center gap-2">
              <i class="fa-solid fa-file-excel text-sm"></i> Download Complete Master Excel (.xlsx)
            </a>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
            <div class="p-3.5 bg-emerald-50/70 border border-emerald-200/80 rounded-xl text-xs">
              <span class="font-bold text-emerald-900 block flex items-center justify-between">
                <span><i class="fa-solid fa-file-excel mr-1 text-emerald-600"></i> Master Excel File:</span>
                <span class="px-1.5 py-0.5 rounded bg-emerald-200/70 text-emerald-800 text-[10px] font-mono font-bold">82 MB</span>
              </span>
              <span class="font-mono text-slate-700 select-all text-[11px] block mt-1">BMP_Tutors_Cleaned_Master.xlsx</span>
              <p class="text-[10px] text-emerald-700 mt-1">107k Tutors, Cleaned Roles, 19 Occupations</p>
            </div>
            <div class="p-3.5 bg-slate-100 rounded-xl text-xs border border-slate-200">
              <span class="font-semibold text-slate-800 block">SQLite Database:</span>
              <span class="font-mono text-slate-600 select-all text-[11px] block mt-1">cleaned_tutors_data.db (302 MB)</span>
              <p class="text-[10px] text-slate-500 mt-1">Indexed 3NF relational schema</p>
            </div>
            <div class="p-3.5 bg-slate-100 rounded-xl text-xs border border-slate-200">
              <span class="font-semibold text-slate-800 block">Cleaned CSVs:</span>
              <span class="font-mono text-slate-600 select-all text-[11px] block mt-1">cleaned_teachers.csv (128 MB)<br/>cleaned_teacher_subjects.csv (26 MB)</span>
            </div>
            <div class="p-3.5 bg-slate-100 rounded-xl text-xs border border-slate-200">
              <span class="font-semibold text-slate-800 block">Fast Parquet Files:</span>
              <span class="font-mono text-slate-600 select-all text-[11px] block mt-1">cleaned_teachers.parquet (57 MB)<br/>cleaned_teacher_subjects.parquet (3.3 MB)</span>
            </div>
          </div>
        </div>
      </div>
    </div>

  </main>

  <!-- TUTOR DETAIL SLIDE-OVER DRAWER MODAL -->
  <div id="detail-modal" class="fixed inset-0 z-50 hidden bg-slate-950/40 backdrop-blur-sm flex justify-end transition-opacity">
    <div class="bg-white w-full max-w-2xl h-full shadow-2xl flex flex-col transform transition-transform duration-300 translate-x-0 overflow-hidden">
      
      <!-- Drawer Header -->
      <div class="p-6 bg-slate-900 text-white flex items-start justify-between border-b border-slate-800">
        <div class="flex items-center gap-3">
          <div class="w-12 h-12 rounded-2xl bg-gradient-to-tr from-indigo-500 to-violet-500 text-white flex items-center justify-center font-bold text-lg ring-2 ring-white/10 shrink-0">
            <span id="modal-initials">T</span>
          </div>
          <div>
            <div class="flex items-center gap-2">
              <h3 id="modal-name" class="text-lg font-extrabold tracking-tight">Tutor Details</h3>
              <span id="modal-active-badge" class="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider font-mono"></span>
            </div>
            <p id="modal-tutor-id" class="text-xs text-indigo-300 font-mono mt-0.5 font-semibold">T1</p>
          </div>
        </div>
        
        <div class="flex items-center gap-2">
          <a id="modal-url" href="#" target="_blank" class="px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs text-slate-200 border border-slate-700 transition-colors flex items-center gap-1.5" title="Open TeacherOn Profile in new tab">
            <i class="fa-solid fa-arrow-up-right-from-square text-[10px]"></i>
            <span>TeacherOn Profile</span>
          </a>
          <button onclick="closeModal()" class="w-8 h-8 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center text-sm transition-colors">
            <i class="fa-solid fa-xmark"></i>
          </button>
        </div>
      </div>

      <!-- Drawer Scrollable Content -->
      <div class="flex-1 overflow-y-auto p-6 space-y-6 custom-scrollbar text-sm">
        
        <!-- Headline / Hookline -->
        <div class="bg-indigo-50/50 p-4 rounded-xl border border-indigo-100 text-slate-800">
          <span class="text-[10px] font-bold text-indigo-600 uppercase tracking-wider block mb-1">Professional Headline</span>
          <p id="modal-hookline" class="text-sm font-semibold italic text-slate-900 leading-relaxed"></p>
        </div>

        <!-- 4 Metric Cards -->
        <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 text-center">
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200/80">
            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Hourly Fee</span>
            <span id="modal-fee" class="text-base font-extrabold text-emerald-600 block mt-0.5 font-mono">₹500 / hr</span>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200/80">
            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Total Experience</span>
            <span id="modal-exp" class="text-base font-extrabold text-slate-900 block mt-0.5 font-mono">5 yrs</span>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200/80">
            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Seniority Tier</span>
            <span id="modal-seniority" class="text-xs font-bold text-amber-700 block mt-1">Experienced</span>
          </div>
          <div class="p-3 bg-slate-50 rounded-xl border border-slate-200/80">
            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Gender</span>
            <span id="modal-gender" class="text-xs font-bold text-slate-800 block mt-1">Female</span>
          </div>
        </div>

        <!-- Occupation & Role -->
        <div class="p-4 rounded-xl bg-slate-50 border border-slate-200/80">
          <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-1">Career & Professional Domain</span>
          <div class="flex flex-wrap items-center gap-2">
            <span id="modal-occupation" class="px-3 py-1 rounded-full text-xs font-bold bg-indigo-100 text-indigo-800 font-mono">Software Engineer & IT Professional</span>
            <span id="modal-role" class="px-2.5 py-1 rounded-md text-xs font-medium bg-slate-200 text-slate-700">Online Educator</span>
            <span id="modal-works-as" class="px-2.5 py-1 rounded-md text-xs font-medium bg-slate-200 text-slate-700">Individual Teacher</span>
          </div>
        </div>

        <!-- Delivery Modes & Flags -->
        <div>
          <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-2">Teaching Preferences & Delivery</span>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 text-center text-xs">
            <div id="flag-online" class="p-2 rounded-lg border font-semibold">Online</div>
            <div id="flag-home" class="p-2 rounded-lg border font-semibold">Home Visit</div>
            <div id="flag-travel" class="p-2 rounded-lg border font-semibold">Can Travel</div>
            <div id="flag-hw" class="p-2 rounded-lg border font-semibold">Homework Help</div>
          </div>
        </div>

        <!-- Location Details -->
        <div class="p-4 rounded-xl bg-slate-50 border border-slate-200/80">
          <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-1">Geographic Classification</span>
          <div class="flex items-center gap-2 text-sm text-slate-900 font-semibold">
            <i class="fa-solid fa-map-pin text-rose-500"></i>
            <span id="modal-location">Bangalore, Karnataka</span>
          </div>
          <p class="text-xs text-slate-500 mt-1">
            State: <span id="modal-state" class="font-semibold text-slate-700">Karnataka</span> &bull; 
            City Tier: <span id="modal-city-tier" class="font-semibold text-slate-700">Tier 1 (Metro)</span>
          </p>
        </div>

        <!-- Academic Credentials -->
        <div class="p-4 rounded-xl bg-slate-50 border border-slate-200/80">
          <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-1">Academic Credentials</span>
          <p id="modal-qualification" class="font-bold text-slate-900 text-sm">Postgraduate (Master's)</p>
          <p id="modal-edu-deg" class="text-xs text-slate-600 mt-0.5">Education Degree (B.Ed/M.Ed): No</p>
          <p id="modal-speaks" class="text-xs text-slate-500 mt-1">Languages: English, Hindi</p>
        </div>

        <!-- Subjects List (Bridge Links) -->
        <div>
          <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-2">Subject Offerings (Bridge Links)</span>
          <div id="modal-subjects-container" class="space-y-2 max-h-60 overflow-y-auto custom-scrollbar"></div>
        </div>

        <!-- Bio Description -->
        <div>
          <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-2">Full Educator Bio</span>
          <div id="modal-desc" class="p-4 rounded-xl bg-slate-50 border border-slate-200 text-xs text-slate-700 leading-relaxed max-h-60 overflow-y-auto custom-scrollbar whitespace-pre-line font-serif italic"></div>
        </div>

        <!-- System Metadata -->
        <div class="pt-4 border-t border-slate-100 text-[11px] text-slate-400 flex items-center justify-between font-mono">
          <div><span id="modal-source-tab">Tab: scraped</span> &bull; Registered: <span id="modal-registered">—</span></div>
          <div>Last Login: <span id="modal-last-login">—</span></div>
        </div>

      </div>
    </div>
  </div>

  <footer class="bg-white border-t border-slate-200 py-4 text-center text-xs text-slate-500 mt-auto">
    <div class="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-2">
      <p>BMP Tutors &bull; Empirical Market Intelligence & Geospatial Analytics &bull; September 2026</p>
      <div class="flex items-center space-x-4">
        <a href="/report" target="_blank" class="hover:text-indigo-600 transition-colors">Academic Report</a>
        <a href="/api/download/master-excel" class="hover:text-emerald-600 transition-colors">Download Master Dataset</a>
      </div>
    </div>
  </footer>

  <script>
{script_mod}
  </script>
</body>
</html>
"""

with open(TARGET_PATH, 'w') as f:
    f.write(full_html)

print("Redesigned modern dashboard generated successfully at:", TARGET_PATH)
file_size = os.path.getsize(TARGET_PATH)
print(f"File size: {file_size:,} bytes")
