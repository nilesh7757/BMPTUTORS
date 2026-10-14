#!/usr/bin/env python3
"""
Build Redesigned Dashboard for BMP Tutors.
Streamlines buttons, adds comprehensive multi-criteria sorting (Name, Fees, Education, Experience),
and introduces clean collapsible filters, active filter chips, and interactive table headers.
"""

import os
import re

ORIG_PATH = "/home/nilesh7757/BMPTUTORS1/app/index.html.orig"
TARGET_PATH = "/home/nilesh7757/BMPTUTORS1/app/index.html"

with open(ORIG_PATH, "r", encoding="utf-8") as f:
    content = f.read()

# ---------------------------------------------------------------------------
# 1. REDESIGN HEADER (lines ~41 to ~108)
# Replace cluttered multiple download buttons with unified Export dropdown
# ---------------------------------------------------------------------------
old_header_pattern = r'<!-- TOP EXECUTIVE NAVIGATION BAR -->.*?</header>'
new_header = """<!-- TOP EXECUTIVE NAVIGATION BAR -->
  <header class="bg-slate-900 text-white border-b border-slate-800 sticky top-0 z-40 transition-all shadow-sm">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4">
      
      <!-- Brand & Title -->
      <div class="flex items-center gap-3 shrink-0">
        <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-600 to-indigo-500 shadow-md flex items-center justify-center font-bold text-lg text-white ring-1 ring-white/20">
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
          <span>Tutors Directory</span>
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

      <!-- Consolidated Action: Export Menu + Academic Report Link -->
      <div class="flex items-center gap-2 shrink-0">
        <a href="/report" target="_blank" class="hidden sm:inline-flex px-3 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 hover:border-slate-600 transition-all items-center gap-1.5 shadow-2xs">
          <i class="fa-solid fa-file-invoice text-indigo-400"></i>
          <span>Academic Report</span>
          <i class="fa-solid fa-arrow-up-right-from-square text-[9px] opacity-60"></i>
        </a>

        <!-- Consolidated Export Dropdown Button -->
        <div class="relative" id="export-dropdown-container">
          <button id="btn-export-dropdown" onclick="toggleExportMenu(event)" class="px-3.5 py-1.5 rounded-lg text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white shadow-xs transition-all flex items-center gap-1.5 ring-1 ring-indigo-400/30">
            <i class="fa-solid fa-download text-xs text-indigo-200"></i>
            <span>Export Data</span>
            <i class="fa-solid fa-chevron-down text-[9px] transition-transform duration-200" id="export-dropdown-chevron"></i>
          </button>
          
          <div id="export-dropdown-menu" class="hidden absolute right-0 mt-2 w-72 bg-white rounded-xl shadow-xl border border-slate-200 py-1.5 z-50 text-slate-800 divide-y divide-slate-100">
            <div class="p-1">
              <span class="block px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">Master Datasets</span>
              <a href="/api/download/active-master-excel" download="BMP_Tutors_Verified_Active_Master.xlsx" class="flex items-center justify-between px-3 py-2 rounded-lg hover:bg-slate-50 transition-colors group">
                <div class="flex items-center gap-2.5">
                  <div class="w-7 h-7 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center text-xs">
                    <i class="fa-solid fa-user-check"></i>
                  </div>
                  <div>
                    <span class="text-xs font-bold text-slate-900 group-hover:text-indigo-600 block">Active Master Excel</span>
                    <span class="text-[10px] text-slate-500">76,705 URL-Verified Educators</span>
                  </div>
                </div>
                <span class="text-[10px] font-mono text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded">58 MB</span>
              </a>
              <a href="/api/download/master-excel" download="BMP_Tutors_Cleaned_Master.xlsx" class="flex items-center justify-between px-3 py-2 rounded-lg hover:bg-slate-50 transition-colors group">
                <div class="flex items-center gap-2.5">
                  <div class="w-7 h-7 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center text-xs">
                    <i class="fa-solid fa-file-excel"></i>
                  </div>
                  <div>
                    <span class="text-xs font-bold text-slate-900 group-hover:text-indigo-600 block">Complete Cleaned Master</span>
                    <span class="text-[10px] text-slate-500">All 107,723 Records</span>
                  </div>
                </div>
                <span class="text-[10px] font-mono text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded">79 MB</span>
              </a>
            </div>
            <div class="p-1">
              <span class="block px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">Thematic Bundles</span>
              <div class="grid grid-cols-2 gap-1 px-1">
                <a href="/api/download/category-excel?category=coding" class="px-2 py-1.5 rounded-md text-[11px] font-medium text-slate-600 hover:text-indigo-600 hover:bg-slate-50 transition-colors truncate">
                  💻 Coding (~15M)
                </a>
                <a href="/api/download/category-excel?category=science" class="px-2 py-1.5 rounded-md text-[11px] font-medium text-slate-600 hover:text-indigo-600 hover:bg-slate-50 transition-colors truncate">
                  🔬 Science (~62M)
                </a>
                <a href="/api/download/category-excel?category=languages" class="px-2 py-1.5 rounded-md text-[11px] font-medium text-slate-600 hover:text-indigo-600 hover:bg-slate-50 transition-colors truncate">
                  🗣️ Languages (~27M)
                </a>
                <a href="/api/download/category-excel?category=commerce" class="px-2 py-1.5 rounded-md text-[11px] font-medium text-slate-600 hover:text-indigo-600 hover:bg-slate-50 transition-colors truncate">
                  📈 Commerce (~7M)
                </a>
              </div>
            </div>
            <div class="p-1">
              <span class="block px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">Reports</span>
              <a href="/api/export/visual-analytics-pdf" download="BMP_Tutors_Visual_Analytics_Complete_Report.pdf" class="flex items-center gap-2 px-3 py-2 rounded-lg text-xs font-semibold text-rose-600 hover:bg-rose-50 transition-colors">
                <i class="fa-solid fa-file-pdf"></i>
                <span>Download Printable PDF Report</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Mobile Segmented Controller -->
    <div class="flex md:hidden px-4 pb-3 border-t border-slate-800/80 pt-2">
      <nav class="flex w-full bg-slate-800/80 p-1 rounded-xl ring-1 ring-white/10">
        <button onclick="switchTab('explorer')" class="flex-1 py-1.5 rounded-lg text-xs font-semibold text-center bg-indigo-600 text-white">Directory</button>
        <button onclick="switchTab('analytics')" class="flex-1 py-1.5 rounded-lg text-xs font-semibold text-center text-slate-300">Analytics</button>
        <button onclick="switchTab('schema')" class="flex-1 py-1.5 rounded-lg text-xs font-semibold text-center text-slate-300">Schema</button>
      </nav>
    </div>
  </header>"""

content = re.sub(old_header_pattern, new_header, content, flags=re.DOTALL)

# ---------------------------------------------------------------------------
# 2. REDESIGN EXPLORER TAB (Sidebar filters, toolbar, table headers, table body)
# ---------------------------------------------------------------------------
old_explorer_pattern = r'<!-- TAB 1: EXPLORER -->.*?<!-- TAB 2: VISUAL ANALYTICS'
new_explorer = """<!-- TAB 1: EXPLORER -->
    <div id="tab-explorer" class="flex flex-col lg:flex-row gap-6 flex-1">
      
      <!-- SIDEBAR FACETED FILTERS -->
      <aside class="w-full lg:w-72 bg-white p-5 rounded-2xl shadow-xs border border-slate-200/80 flex-shrink-0 flex flex-col gap-3.5">
        <div class="flex items-center justify-between border-b border-slate-100 pb-3">
          <div class="flex items-center gap-2">
            <div class="w-6 h-6 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center text-xs">
              <i class="fa-solid fa-filter"></i>
            </div>
            <h2 class="font-bold text-slate-900 text-sm">Filter Tutors</h2>
          </div>
          <button id="btn-reset-filters" onclick="resetFilters()" class="text-xs text-slate-400 hover:text-indigo-600 font-semibold flex items-center gap-1 transition-colors cursor-pointer">
            <i class="fa-solid fa-rotate-left text-[10px]"></i> Reset
          </button>
        </div>

        <!-- Search input with Instant Debounce and Clear Button -->
        <div>
          <label for="filter-search" class="block text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Search</label>
          <div class="relative">
            <input type="text" id="filter-search" oninput="onSearchInput(this.value)" onkeydown="if(event.key==='Enter') applyFilters()" placeholder="Name, ID (T1), city, headline..." class="w-full pl-8 pr-7 py-2 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs" />
            <i class="fa-solid fa-magnifying-glass absolute left-2.5 top-2.5 text-slate-400 text-xs"></i>
            <button id="btn-clear-search" onclick="clearSearch()" class="hidden absolute right-2.5 top-2 text-slate-400 hover:text-slate-600 text-xs">
              <i class="fa-solid fa-circle-xmark"></i>
            </button>
          </div>
        </div>

        <!-- Primary Core Filter 1: Subject Category -->
        <div>
          <label for="filter-subject-category" class="block text-xs font-semibold text-slate-700 mb-1">Subject Domain</label>
          <select id="filter-subject-category" onchange="applyFilters()" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs cursor-pointer">
            <option value="">All Disciplines</option>
          </select>
        </div>

        <!-- Primary Core Filter 2: State / UT -->
        <div>
          <label for="filter-state" class="block text-xs font-semibold text-slate-700 mb-1">State / Territory</label>
          <select id="filter-state" onchange="applyFilters()" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs cursor-pointer">
            <option value="">All States & UTs</option>
          </select>
        </div>

        <!-- Primary Core Filter 3: Education Level -->
        <div>
          <label for="filter-qualification" class="block text-xs font-semibold text-slate-700 mb-1">Education Level</label>
          <select id="filter-qualification" onchange="applyFilters()" class="w-full py-2 px-3 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs cursor-pointer">
            <option value="">All Qualifications</option>
          </select>
        </div>

        <!-- Secondary Collapsible Accordion -->
        <div class="pt-2 border-t border-slate-100">
          <button id="btn-toggle-more-filters" type="button" onclick="toggleMoreFilters()" class="w-full py-2 px-3 text-xs font-semibold text-slate-600 hover:text-indigo-600 bg-slate-50 hover:bg-slate-100/70 border border-slate-200/80 rounded-xl flex items-center justify-between transition-colors shadow-2xs">
            <span class="flex items-center gap-1.5">
              <i class="fa-solid fa-sliders text-indigo-500 text-[10px]"></i>
              <span>More Filters</span>
            </span>
            <span class="flex items-center gap-1.5">
              <span id="more-filters-badge" class="hidden px-1.5 py-0.2 rounded-full bg-indigo-100 text-indigo-700 text-[10px] font-bold font-mono">0</span>
              <i id="more-filters-chevron" class="fa-solid fa-chevron-down text-[10px] text-slate-400 transition-transform duration-200"></i>
            </span>
          </button>

          <div id="more-filters-container" class="hidden space-y-3 pt-3">
            <div>
              <label for="filter-occupation" class="block text-[11px] font-semibold text-slate-700 mb-1">Career / Occupation</label>
              <select id="filter-occupation" onchange="applyFilters()" class="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
                <option value="">All Occupations</option>
              </select>
            </div>

            <div>
              <label for="filter-level" class="block text-[11px] font-semibold text-slate-700 mb-1">Grade / Level</label>
              <select id="filter-level" onchange="applyFilters()" class="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
                <option value="">All Levels</option>
              </select>
            </div>

            <div>
              <label for="filter-online" class="block text-[11px] font-semibold text-slate-700 mb-1">Delivery Mode</label>
              <select id="filter-online" onchange="applyFilters()" class="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
                <option value="">Any Mode</option>
                <option value="1">Online Only</option>
                <option value="0">Offline / In-person</option>
              </select>
            </div>

            <div>
              <label for="filter-seniority" class="block text-[11px] font-semibold text-slate-700 mb-1">Seniority Tier</label>
              <select id="filter-seniority" onchange="applyFilters()" class="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
                <option value="">All Seniority Tiers</option>
              </select>
            </div>

            <div>
              <label for="filter-role" class="block text-[11px] font-semibold text-slate-700 mb-1">Experience Role</label>
              <select id="filter-role" onchange="applyFilters()" class="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
                <option value="">All Roles</option>
              </select>
            </div>

            <div>
              <label for="filter-exam" class="block text-[11px] font-semibold text-slate-700 mb-1">Competitive Exam</label>
              <select id="filter-exam" onchange="applyFilters()" class="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
                <option value="">All Exams</option>
              </select>
            </div>

            <div>
              <label for="filter-region" class="block text-[11px] font-semibold text-slate-700 mb-1">Geographic Zone</label>
              <select id="filter-region" onchange="applyFilters()" class="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
                <option value="">All Zones</option>
              </select>
            </div>

            <div>
              <label for="filter-gender" class="block text-[11px] font-semibold text-slate-700 mb-1">Gender</label>
              <select id="filter-gender" onchange="applyFilters()" class="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:border-indigo-500 focus:bg-white transition-all shadow-2xs">
                <option value="">All Genders</option>
                <option value="Female">Female</option>
                <option value="Male">Male</option>
              </select>
            </div>
          </div>
        </div>
      </aside>

      <!-- TABLE & PAGINATION CONTAINER -->
      <section class="flex-1 bg-white rounded-2xl shadow-xs border border-slate-200/80 flex flex-col overflow-hidden">
        
        <!-- Results Header Toolbar -->
        <div class="px-6 py-3.5 border-b border-slate-100 flex flex-col md:flex-row items-start md:items-center justify-between gap-3 bg-slate-50/40">
          <div class="flex items-center gap-2.5">
            <h3 class="font-bold text-slate-900 text-sm">Educator Profiles</h3>
            <span id="results-count-badge" class="text-xs font-semibold px-2.5 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200/70 font-mono">
              Loading...
            </span>
          </div>
          
          <!-- Controls: Sort Selector + Page Size + Export -->
          <div class="flex flex-wrap items-center gap-2.5 text-xs w-full md:w-auto justify-between md:justify-end">
            <!-- Sort Selector Dropdown -->
            <div class="flex items-center gap-1.5">
              <label for="table-sort-select" class="text-slate-500 font-semibold text-[11px] flex items-center gap-1">
                <i class="fa-solid fa-arrow-down-short-wide text-indigo-500 text-[10px]"></i>
                <span>Sort:</span>
              </label>
              <select id="table-sort-select" onchange="onSortDropdownChange(this.value)" class="bg-white border border-slate-200 rounded-lg px-2.5 py-1 text-xs font-medium text-slate-700 focus:outline-none focus:border-indigo-500 shadow-2xs cursor-pointer">
                <option value="default_asc">Default (Tutor ID)</option>
                <option value="name_asc">Name: A → Z</option>
                <option value="name_desc">Name: Z → A</option>
                <option value="fee_asc">Fee: Low to High</option>
                <option value="fee_desc">Fee: High to Low</option>
                <option value="education_desc">Education: Highest Degree</option>
                <option value="education_asc">Education: Lowest Degree</option>
                <option value="experience_desc">Experience: Most Experienced</option>
                <option value="experience_asc">Experience: Least Experienced</option>
              </select>
            </div>

            <!-- Page Size -->
            <div class="flex items-center gap-1.5 pl-2 border-l border-slate-200">
              <span class="text-slate-500 font-medium text-[11px]">Rows:</span>
              <select id="page-size" onchange="changePageSize()" class="bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs focus:outline-none focus:border-indigo-500 shadow-2xs cursor-pointer">
                <option value="15">15</option>
                <option value="25" selected>25</option>
                <option value="50">50</option>
                <option value="100">100</option>
              </select>
            </div>

            <!-- Export Filtered Button -->
            <button onclick="downloadFilteredExcel()" class="px-2.5 py-1 bg-slate-100 hover:bg-indigo-50 hover:text-indigo-700 hover:border-indigo-200 text-slate-700 border border-slate-200 rounded-lg text-xs font-semibold transition-all flex items-center gap-1.5 shadow-2xs ml-1" title="Export matching rows to Excel">
              <i class="fa-solid fa-file-excel text-emerald-600 text-[11px]"></i>
              <span>Export</span>
            </button>
          </div>
        </div>

        <!-- Active Filters Bar -->
        <div id="active-filters-bar" class="hidden px-6 py-2 bg-indigo-50/40 border-b border-indigo-100/60 flex flex-wrap items-center gap-1.5 text-xs">
          <!-- Dynamically populated chips -->
        </div>

        <!-- Table Container -->
        <div class="flex-1 overflow-x-auto custom-scrollbar">
          <table class="w-full text-left border-collapse text-sm">
            <thead>
              <tr class="bg-slate-50/80 border-b border-slate-200/80 text-[11px] font-bold text-slate-500 uppercase tracking-wider sticky top-0 backdrop-blur-sm z-10 select-none">
                <th class="py-3 px-4">Tutor ID</th>
                <th class="py-3 px-4 cursor-pointer hover:text-indigo-600 transition-colors group" onclick="toggleSort('name')" title="Click to sort by Educator Name">
                  <div class="flex items-center gap-1.5">
                    <span>Educator Profile</span>
                    <i id="sort-icon-name" class="fa-solid fa-sort text-slate-300 group-hover:text-indigo-500 text-[10px]"></i>
                  </div>
                </th>
                <th class="py-3 px-4">Occupation & Role</th>
                <th class="py-3 px-4">Location</th>
                <th class="py-3 px-4 cursor-pointer hover:text-indigo-600 transition-colors group" onclick="toggleSort('education')" title="Click to sort by Education Qualification">
                  <div class="flex items-center gap-1.5">
                    <span>Qualification & Exp</span>
                    <i id="sort-icon-education" class="fa-solid fa-sort text-slate-300 group-hover:text-indigo-500 text-[10px]"></i>
                  </div>
                </th>
                <th class="py-3 px-4">Subjects Taught (Bridge)</th>
                <th class="py-3 px-4 cursor-pointer hover:text-indigo-600 transition-colors group" onclick="toggleSort('fee')" title="Click to sort by Hourly Fee">
                  <div class="flex items-center gap-1.5">
                    <span>Fee / hr</span>
                    <i id="sort-icon-fee" class="fa-solid fa-sort text-slate-300 group-hover:text-indigo-500 text-[10px]"></i>
                  </div>
                </th>
                <th class="py-3 px-3 text-center w-10"></th>
              </tr>
            </thead>
            <tbody id="tutors-table-body" class="divide-y divide-slate-100 text-slate-700">
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
                <td class="py-4 px-3 text-center"><div class="h-5 w-5 bg-slate-200 rounded-full mx-auto"></div></td>
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
            <button onclick="goToPage(1)" class="px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white text-xs hover:bg-slate-50 disabled:opacity-40 shadow-2xs transition-colors cursor-pointer" id="btn-first" title="First Page">
              <i class="fa-solid fa-angles-left text-[10px]"></i>
            </button>
            <button onclick="changePage(-1)" class="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-xs hover:bg-slate-50 disabled:opacity-40 shadow-2xs transition-colors flex items-center gap-1 cursor-pointer" id="btn-prev">
              <i class="fa-solid fa-chevron-left text-[10px]"></i> <span>Prev</span>
            </button>
            <span id="current-page-display" class="px-3 py-1.5 text-xs font-bold text-indigo-700 bg-indigo-50 border border-indigo-200/80 rounded-lg font-mono">1</span>
            <button onclick="changePage(1)" class="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-xs hover:bg-slate-50 disabled:opacity-40 shadow-2xs transition-colors flex items-center gap-1 cursor-pointer" id="btn-next">
              <span>Next</span> <i class="fa-solid fa-chevron-right text-[10px]"></i>
            </button>
            <button onclick="goToLastPage()" class="px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white text-xs hover:bg-slate-50 disabled:opacity-40 shadow-2xs transition-colors cursor-pointer" id="btn-last" title="Last Page">
              <i class="fa-solid fa-angles-right text-[10px]"></i>
            </button>
          </div>
        </div>
      </section>
    </div>

    <!-- TAB 2: VISUAL ANALYTICS"""

content = re.sub(old_explorer_pattern, new_explorer, content, flags=re.DOTALL)

# ---------------------------------------------------------------------------
# 3. SCRIPT ENHANCEMENTS
# Add sort state, debounce, filter chips, dropdown toggle, and updated row rendering
# ---------------------------------------------------------------------------

# Insert global state variables after `let chartsInitialized = false;`
old_state_init = "let chartsInitialized = false;"
new_state_init = """let chartsInitialized = false;
    let currentSortBy = 'default';
    let currentSortOrder = 'asc';
    let searchDebounceTimer = null;

    // Toggle Export Menu
    function toggleExportMenu(e) {
      if (e) e.stopPropagation();
      const menu = document.getElementById('export-dropdown-menu');
      const chevron = document.getElementById('export-dropdown-chevron');
      if (!menu) return;
      const isHidden = menu.classList.contains('hidden');
      if (isHidden) {
        menu.classList.remove('hidden');
        if (chevron) chevron.style.transform = 'rotate(180deg)';
      } else {
        menu.classList.add('hidden');
        if (chevron) chevron.style.transform = 'rotate(0deg)';
      }
    }

    // Close export menu when clicking outside
    document.addEventListener('click', (e) => {
      const container = document.getElementById('export-dropdown-container');
      const menu = document.getElementById('export-dropdown-menu');
      const chevron = document.getElementById('export-dropdown-chevron');
      if (container && !container.contains(e.target)) {
        if (menu && !menu.classList.contains('hidden')) {
          menu.classList.add('hidden');
          if (chevron) chevron.style.transform = 'rotate(0deg)';
        }
      }
    });

    // Toggle More Filters Collapsible
    function toggleMoreFilters() {
      const container = document.getElementById('more-filters-container');
      const chevron = document.getElementById('more-filters-chevron');
      if (!container) return;
      const isHidden = container.classList.contains('hidden');
      if (isHidden) {
        container.classList.remove('hidden');
        if (chevron) chevron.style.transform = 'rotate(180deg)';
      } else {
        container.classList.add('hidden');
        if (chevron) chevron.style.transform = 'rotate(0deg)';
      }
    }

    // Search Input with Debounce & Clear Icon
    function onSearchInput(val) {
      const clearBtn = document.getElementById('btn-clear-search');
      if (clearBtn) {
        if (val && val.trim().length > 0) {
          clearBtn.classList.remove('hidden');
        } else {
          clearBtn.classList.add('hidden');
        }
      }
      clearTimeout(searchDebounceTimer);
      searchDebounceTimer = setTimeout(() => {
        applyFilters();
      }, 350);
    }

    function clearSearch() {
      const searchInput = document.getElementById('filter-search');
      if (searchInput) searchInput.value = '';
      const clearBtn = document.getElementById('btn-clear-search');
      if (clearBtn) clearBtn.classList.add('hidden');
      applyFilters();
    }

    // Interactive Sorting Controls
    function setSort(sortBy, sortOrder) {
      currentSortBy = sortBy;
      currentSortOrder = sortOrder;
      
      // Sync dropdown
      const sortSelect = document.getElementById('table-sort-select');
      if (sortSelect) {
        const targetVal = `${sortBy}_${sortOrder}`;
        sortSelect.value = targetVal;
      }

      updateSortHeaderUI();
      currentPage = 1;
      loadTutors();
    }

    function toggleSort(column) {
      if (currentSortBy === column) {
        // Toggle direction
        currentSortOrder = currentSortOrder === 'asc' ? 'desc' : 'asc';
      } else {
        currentSortBy = column;
        // Default direction per metric
        if (column === 'fee') {
          currentSortOrder = 'asc'; // lowest fee first by default
        } else if (column === 'education' || column === 'experience') {
          currentSortOrder = 'desc'; // highest credential or experience first
        } else {
          currentSortOrder = 'asc';
        }
      }
      setSort(currentSortBy, currentSortOrder);
    }

    function onSortDropdownChange(val) {
      if (!val) return;
      const parts = val.split('_');
      const sortBy = parts[0];
      const sortOrder = parts[1] || 'asc';
      setSort(sortBy, sortOrder);
    }

    function updateSortHeaderUI() {
      const columns = ['name', 'education', 'fee'];
      columns.forEach(col => {
        const icon = document.getElementById(`sort-icon-${col}`);
        if (!icon) return;
        if (currentSortBy === col) {
          icon.className = currentSortOrder === 'asc' 
            ? 'fa-solid fa-arrow-up-short-wide text-indigo-600 text-[11px]' 
            : 'fa-solid fa-arrow-down-wide-short text-indigo-600 text-[11px]';
        } else {
          icon.className = 'fa-solid fa-sort text-slate-300 group-hover:text-indigo-400 text-[10px]';
        }
      });
    }

    // Active Filters Chips Bar
    function updateActiveFilterChips() {
      const bar = document.getElementById('active-filters-bar');
      const moreBadge = document.getElementById('more-filters-badge');
      if (!bar) return;

      const chips = [];
      let moreCount = 0;

      const searchVal = document.getElementById('filter-search')?.value.trim();
      if (searchVal) chips.push({ label: `Search: "${searchVal}"`, clear: () => clearSearch() });

      const subjVal = document.getElementById('filter-subject-category')?.value;
      if (subjVal) chips.push({ label: `Subject: ${subjVal}`, clear: () => { document.getElementById('filter-subject-category').value = ''; applyFilters(); } });

      const stateVal = document.getElementById('filter-state')?.value;
      if (stateVal) chips.push({ label: `State: ${stateVal}`, clear: () => { document.getElementById('filter-state').value = ''; applyFilters(); } });

      const qualVal = document.getElementById('filter-qualification')?.value;
      if (qualVal) chips.push({ label: `Degree: ${qualVal}`, clear: () => { document.getElementById('filter-qualification').value = ''; applyFilters(); } });

      // Secondary filters
      const occVal = document.getElementById('filter-occupation')?.value;
      if (occVal) { chips.push({ label: `Career: ${occVal}`, clear: () => { document.getElementById('filter-occupation').value = ''; applyFilters(); } }); moreCount++; }

      const levelVal = document.getElementById('filter-level')?.value;
      if (levelVal) { chips.push({ label: `Level: ${levelVal}`, clear: () => { document.getElementById('filter-level').value = ''; applyFilters(); } }); moreCount++; }

      const onlineVal = document.getElementById('filter-online')?.value;
      if (onlineVal !== '') { chips.push({ label: `Mode: ${onlineVal === '1' ? 'Online' : 'Offline'}`, clear: () => { document.getElementById('filter-online').value = ''; applyFilters(); } }); moreCount++; }

      const seniorityVal = document.getElementById('filter-seniority')?.value;
      if (seniorityVal) { chips.push({ label: `Tier: ${seniorityVal}`, clear: () => { document.getElementById('filter-seniority').value = ''; applyFilters(); } }); moreCount++; }

      const roleVal = document.getElementById('filter-role')?.value;
      if (roleVal) { chips.push({ label: `Role: ${roleVal}`, clear: () => { document.getElementById('filter-role').value = ''; applyFilters(); } }); moreCount++; }

      const examVal = document.getElementById('filter-exam')?.value;
      if (examVal) { chips.push({ label: `Exam: ${examVal}`, clear: () => { document.getElementById('filter-exam').value = ''; applyFilters(); } }); moreCount++; }

      const regionVal = document.getElementById('filter-region')?.value;
      if (regionVal) { chips.push({ label: `Zone: ${regionVal}`, clear: () => { document.getElementById('filter-region').value = ''; applyFilters(); } }); moreCount++; }

      const genderVal = document.getElementById('filter-gender')?.value;
      if (genderVal) { chips.push({ label: `Gender: ${genderVal}`, clear: () => { document.getElementById('filter-gender').value = ''; applyFilters(); } }); moreCount++; }

      // Custom Sort chip
      if (currentSortBy !== 'default') {
        const sortLabels = {
          'name': currentSortOrder === 'asc' ? 'Name (A→Z)' : 'Name (Z→A)',
          'fee': currentSortOrder === 'asc' ? 'Fee (Low→High)' : 'Fee (High→Low)',
          'education': currentSortOrder === 'desc' ? 'Education (Highest)' : 'Education (Lowest)',
          'experience': currentSortOrder === 'desc' ? 'Exp (Most)' : 'Exp (Least)'
        };
        chips.push({ label: `Sort: ${sortLabels[currentSortBy] || currentSortBy}`, clear: () => setSort('default', 'asc') });
      }

      // Update More Filters count badge
      if (moreBadge) {
        if (moreCount > 0) {
          moreBadge.textContent = moreCount;
          moreBadge.classList.remove('hidden');
        } else {
          moreBadge.classList.add('hidden');
        }
      }

      // Render Bar
      if (chips.length > 0) {
        bar.classList.remove('hidden');
        bar.innerHTML = `
          <span class="text-[11px] font-bold text-slate-500 uppercase tracking-wider mr-1">Active:</span>
          ${chips.map((c, i) => `
            <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-semibold bg-white text-indigo-700 border border-indigo-200/80 shadow-2xs">
              <span>${escapeHtml(c.label)}</span>
              <button onclick="window.__clearFilterChip(${i})" class="text-indigo-400 hover:text-indigo-800 transition-colors ml-0.5 cursor-pointer">
                <i class="fa-solid fa-xmark text-[10px]"></i>
              </button>
            </span>
          `).join('')}
          <button onclick="resetFilters()" class="text-[11px] text-slate-500 hover:text-rose-600 font-semibold underline decoration-slate-300 ml-2 cursor-pointer transition-colors">
            Clear all
          </button>
        `;
        window.__filterChipsCallbacks = chips.map(c => c.clear);
      } else {
        bar.classList.add('hidden');
        bar.innerHTML = '';
        window.__filterChipsCallbacks = [];
      }
    }

    window.__clearFilterChip = function(index) {
      if (window.__filterChipsCallbacks && window.__filterChipsCallbacks[index]) {
        window.__filterChipsCallbacks[index]();
      }
    };"""

content = content.replace(old_state_init, new_state_init, 1)

# ---------------------------------------------------------------------------
# 4. UPDATE loadTutors()
# Pass sort params, render active filter chips, and render clean rows without redundant action button
# ---------------------------------------------------------------------------

# Update URLSearchParams in loadTutors
old_params_code = """      const params = new URLSearchParams({
        page: currentPage,
        page_size: pageSize,
        search: document.getElementById('filter-search').value,
        subject_category: document.getElementById('filter-subject-category').value,
        level: document.getElementById('filter-level').value,
        exam: document.getElementById('filter-exam').value,
        region: document.getElementById('filter-region').value,
        state: document.getElementById('filter-state').value,
        qualification: document.getElementById('filter-qualification').value,
        seniority_tier: document.getElementById('filter-seniority').value,
        experience_role: document.getElementById('filter-role').value,
        occupation: document.getElementById('filter-occupation') ? document.getElementById('filter-occupation').value : '',
        gender: document.getElementById('filter-gender').value,
        teaches_online: document.getElementById('filter-online').value
      });"""

new_params_code = """      updateActiveFilterChips();
      updateSortHeaderUI();

      const params = new URLSearchParams({
        page: currentPage,
        page_size: pageSize,
        sort_by: currentSortBy,
        sort_order: currentSortOrder,
        search: document.getElementById('filter-search') ? document.getElementById('filter-search').value.trim() : '',
        subject_category: document.getElementById('filter-subject-category') ? document.getElementById('filter-subject-category').value : '',
        level: document.getElementById('filter-level') ? document.getElementById('filter-level').value : '',
        exam: document.getElementById('filter-exam') ? document.getElementById('filter-exam').value : '',
        region: document.getElementById('filter-region') ? document.getElementById('filter-region').value : '',
        state: document.getElementById('filter-state') ? document.getElementById('filter-state').value : '',
        qualification: document.getElementById('filter-qualification') ? document.getElementById('filter-qualification').value : '',
        seniority_tier: document.getElementById('filter-seniority') ? document.getElementById('filter-seniority').value : '',
        experience_role: document.getElementById('filter-role') ? document.getElementById('filter-role').value : '',
        occupation: document.getElementById('filter-occupation') ? document.getElementById('filter-occupation').value : '',
        gender: document.getElementById('filter-gender') ? document.getElementById('filter-gender').value : '',
        teaches_online: document.getElementById('filter-online') ? document.getElementById('filter-online').value : ''
      });"""

content = content.replace(old_params_code, new_params_code, 1)

# Update row HTML generation in loadTutors to remove clunky action button and add clean chevron
old_row_html = """          html += `
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
            </tr>`;"""

new_row_html = """          html += `
            <tr class="hover:bg-indigo-50/40 transition-colors group cursor-pointer" onclick="viewTutorDetails('${t.tutor_id}')" title="Click to view full educator dossier">
              <td class="py-3.5 px-4 font-mono font-bold text-xs text-indigo-600">
                <span class="bg-indigo-50 group-hover:bg-indigo-100 text-indigo-700 px-2 py-1 rounded-md border border-indigo-200/60 transition-colors">
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
              <td class="py-3.5 px-3 text-center text-slate-300 group-hover:text-indigo-600 transition-colors">
                <i class="fa-solid fa-chevron-right text-xs"></i>
              </td>
            </tr>`;"""

content = content.replace(old_row_html, new_row_html, 1)

# Update resetFilters() to also reset sort and search
old_reset_func = """    function resetFilters() {
      document.getElementById('filter-search').value = '';
      document.getElementById('filter-subject-category').value = '';
      document.getElementById('filter-level').value = '';
      document.getElementById('filter-exam').value = '';
      document.getElementById('filter-region').value = '';
      document.getElementById('filter-state').value = '';
      document.getElementById('filter-qualification').value = '';
      document.getElementById('filter-seniority').value = '';
      document.getElementById('filter-role').value = '';
      if (document.getElementById('filter-occupation')) document.getElementById('filter-occupation').value = '';
      document.getElementById('filter-gender').value = '';
      document.getElementById('filter-online').value = '';
      currentPage = 1;
      loadTutors();
    }"""

new_reset_func = """    function resetFilters() {
      if (document.getElementById('filter-search')) document.getElementById('filter-search').value = '';
      const clearBtn = document.getElementById('btn-clear-search');
      if (clearBtn) clearBtn.classList.add('hidden');
      if (document.getElementById('filter-subject-category')) document.getElementById('filter-subject-category').value = '';
      if (document.getElementById('filter-level')) document.getElementById('filter-level').value = '';
      if (document.getElementById('filter-exam')) document.getElementById('filter-exam').value = '';
      if (document.getElementById('filter-region')) document.getElementById('filter-region').value = '';
      if (document.getElementById('filter-state')) document.getElementById('filter-state').value = '';
      if (document.getElementById('filter-qualification')) document.getElementById('filter-qualification').value = '';
      if (document.getElementById('filter-seniority')) document.getElementById('filter-seniority').value = '';
      if (document.getElementById('filter-role')) document.getElementById('filter-role').value = '';
      if (document.getElementById('filter-occupation')) document.getElementById('filter-occupation').value = '';
      if (document.getElementById('filter-gender')) document.getElementById('filter-gender').value = '';
      if (document.getElementById('filter-online')) document.getElementById('filter-online').value = '';
      
      currentSortBy = 'default';
      currentSortOrder = 'asc';
      const sortSelect = document.getElementById('table-sort-select');
      if (sortSelect) sortSelect.value = 'default_asc';

      currentPage = 1;
      loadTutors();
    }"""

content = content.replace(old_reset_func, new_reset_func, 1)

# Update downloadFilteredExcel() to include sorting and occupation
old_download_func = """    function downloadFilteredExcel() {
      const params = new URLSearchParams({
        search: document.getElementById('filter-search').value,
        subject_category: document.getElementById('filter-subject-category').value,
        level: document.getElementById('filter-level').value,
        exam: document.getElementById('filter-exam').value,
        region: document.getElementById('filter-region').value,
        state: document.getElementById('filter-state').value,
        qualification: document.getElementById('filter-qualification').value,
        seniority_tier: document.getElementById('filter-seniority').value,
        experience_role: document.getElementById('filter-role').value,
        gender: document.getElementById('filter-gender').value,
        teaches_online: document.getElementById('filter-online').value
      });
      window.location.href = `/api/download/filtered-excel?${params.toString()}`;
    }"""

new_download_func = """    function downloadFilteredExcel() {
      const params = new URLSearchParams({
        sort_by: currentSortBy,
        sort_order: currentSortOrder,
        search: document.getElementById('filter-search') ? document.getElementById('filter-search').value.trim() : '',
        subject_category: document.getElementById('filter-subject-category') ? document.getElementById('filter-subject-category').value : '',
        level: document.getElementById('filter-level') ? document.getElementById('filter-level').value : '',
        exam: document.getElementById('filter-exam') ? document.getElementById('filter-exam').value : '',
        region: document.getElementById('filter-region') ? document.getElementById('filter-region').value : '',
        state: document.getElementById('filter-state') ? document.getElementById('filter-state').value : '',
        qualification: document.getElementById('filter-qualification') ? document.getElementById('filter-qualification').value : '',
        seniority_tier: document.getElementById('filter-seniority') ? document.getElementById('filter-seniority').value : '',
        experience_role: document.getElementById('filter-role') ? document.getElementById('filter-role').value : '',
        occupation: document.getElementById('filter-occupation') ? document.getElementById('filter-occupation').value : '',
        gender: document.getElementById('filter-gender') ? document.getElementById('filter-gender').value : '',
        teaches_online: document.getElementById('filter-online') ? document.getElementById('filter-online').value : ''
      });
      window.location.href = `/api/download/filtered-excel?${params.toString()}`;
    }"""

content = content.replace(old_download_func, new_download_func, 1)

with open(TARGET_PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("Successfully written redesigned dashboard to", TARGET_PATH)
