#!/usr/bin/env python3
import re

with open('app/index.html', 'r') as f:
    html = f.read()

# 1. Update Table Header
old_thead = '''          <table class="w-full text-left border-collapse text-sm">
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
            </thead>'''

new_thead = '''          <table class="w-full text-left border-collapse text-sm">
            <thead>
              <tr class="bg-slate-50/80 border-b border-slate-200/80 text-[11px] font-bold text-slate-500 uppercase tracking-wider sticky top-0 backdrop-blur-sm z-10 select-none">
                <th class="py-3 px-3 w-20 text-center">Tutor ID</th>
                <th class="py-3 px-4 cursor-pointer hover:text-indigo-600 transition-colors group min-w-[190px]" onclick="toggleSort('name')" title="Click to sort by Educator Name">
                  <div class="flex items-center gap-1.5">
                    <span>Educator</span>
                    <i id="sort-icon-name" class="fa-solid fa-sort text-slate-300 group-hover:text-indigo-500 text-[10px]"></i>
                  </div>
                </th>
                <th class="py-3 px-4 min-w-[240px]">Subjects Taught</th>
                <th class="py-3 px-4 cursor-pointer hover:text-indigo-600 transition-colors group min-w-[125px]" onclick="toggleSort('fee')" title="Click to sort by Tuition Fee">
                  <div class="flex items-center gap-1.5">
                    <span>Tuition Fee</span>
                    <i id="sort-icon-fee" class="fa-solid fa-sort text-slate-300 group-hover:text-indigo-500 text-[10px]"></i>
                  </div>
                </th>
                <th class="py-3 px-4 min-w-[150px]">Location</th>
                <th class="py-3 px-4 cursor-pointer hover:text-indigo-600 transition-colors group min-w-[150px]" onclick="toggleSort('education')" title="Click to sort by Qualification">
                  <div class="flex items-center gap-1.5">
                    <span>Qualification & Exp</span>
                    <i id="sort-icon-education" class="fa-solid fa-sort text-slate-300 group-hover:text-indigo-500 text-[10px]"></i>
                  </div>
                </th>
                <th class="py-3 px-3 text-center w-10"></th>
              </tr>
            </thead>'''

if old_thead in html:
    html = html.replace(old_thead, new_thead, 1)
    print("Table header successfully reordered.")
else:
    print("WARNING: old_thead not found")

# 2. Update renderTutors row structure
old_row_block = '''          html += `
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
            </tr>`;'''

new_row_block = '''          // Fee display logic
          let feeDisplay = '';
          if (t.hourly_fee_avg) {
            feeDisplay = `<span class="inline-flex items-center px-2.5 py-1 rounded-lg text-xs font-bold font-mono bg-emerald-50 text-emerald-700 border border-emerald-200/80">₹${Math.round(t.hourly_fee_avg)}<span class="text-[10px] text-emerald-600/70 font-sans ml-0.5">/hr</span></span>`;
          } else if (t.monthly_fee_avg) {
            feeDisplay = `<span class="inline-flex items-center px-2.5 py-1 rounded-lg text-xs font-bold font-mono bg-blue-50 text-blue-700 border border-blue-200/80">₹${Math.round(t.monthly_fee_avg).toLocaleString()}<span class="text-[10px] text-blue-600/70 font-sans ml-0.5">/mo</span></span>`;
          } else if (t.fee_raw) {
            const cleanRaw = t.fee_raw.split('(')[0].trim();
            feeDisplay = `<span class="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-medium bg-slate-100 text-slate-700" title="${escapeHtml(t.fee_raw)}">${escapeHtml(cleanRaw)}</span>`;
          } else {
            feeDisplay = `<span class="text-xs text-slate-400 font-medium">Negotiable</span>`;
          }

          html += `
            <tr class="hover:bg-indigo-50/40 transition-colors group cursor-pointer" onclick="viewTutorDetails('${t.tutor_id}')" title="Click to view educator dossier">
              <td class="py-3.5 px-3 text-center font-mono font-bold text-xs text-indigo-600">
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
                      <span class="font-semibold text-slate-900 text-sm group-hover:text-indigo-600 transition-colors truncate">${escapeHtml(t.teacher_name || 'Tutor')}</span>
                      ${statusDot}
                    </div>
                    <div class="text-xs text-slate-500 truncate max-w-xs mt-0.5">${escapeHtml(t.teacher_hookline || t.occupation || 'Private Educator')}</div>
                  </div>
                </div>
              </td>
              <td class="py-3.5 px-4">
                <div class="flex flex-wrap gap-1 items-center max-w-sm">
                  ${subjsBadges}
                </div>
              </td>
              <td class="py-3.5 px-4">
                ${feeDisplay}
              </td>
              <td class="py-3.5 px-4">
                <div class="font-medium text-slate-800 text-xs flex items-center gap-1">
                  <i class="fa-solid fa-location-dot text-slate-400 text-[10px]"></i>
                  <span>${escapeHtml(t.city || t.state || 'India')}</span>
                </div>
                <div class="text-[11px] text-slate-500 ml-3.5">${escapeHtml(t.state || '')} <span class="text-slate-400">(${escapeHtml(t.region || '—')})</span></div>
              </td>
              <td class="py-3.5 px-4">
                <div class="text-xs font-semibold text-slate-800">${escapeHtml(t.highest_qualification || 'Graduate')}</div>
                <div class="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
                  <i class="fa-regular fa-clock text-slate-400 text-[10px]"></i>
                  <span>${t.total_teaching_exp_years != null ? t.total_teaching_exp_years + ' yrs exp' : '1+ yrs exp'}</span>
                </div>
              </td>
              <td class="py-3.5 px-3 text-center text-slate-300 group-hover:text-indigo-600 transition-colors">
                <i class="fa-solid fa-chevron-right text-xs"></i>
              </td>
            </tr>`;'''

if old_row_block in html:
    html = html.replace(old_row_block, new_row_block, 1)
    print("Table row rendering successfully updated.")
else:
    print("WARNING: old_row_block not found")

# 3. Update Tutor Detail Modal
# Fee label
html = html.replace(
    '<span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Hourly Fee</span>',
    '<span id="modal-fee-label" class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Tuition Rate</span>'
)

# Subjects header
html = html.replace(
    '<span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-2">Subject Offerings (Bridge Links)</span>',
    '<span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider block mb-2">Subject Offerings</span>'
)

# Footer
old_footer = '''        <!-- System Metadata -->
        <div class="pt-4 border-t border-slate-100 text-[11px] text-slate-400 flex items-center justify-between font-mono">
          <div><span id="modal-source-tab">Tab: scraped</span> &bull; Registered: <span id="modal-registered">—</span></div>
          <div>Last Login: <span id="modal-last-login">—</span></div>
        </div>'''

new_footer = '''        <!-- Verified Metadata -->
        <div class="pt-4 border-t border-slate-100 text-[11px] text-slate-500 flex items-center justify-between font-mono">
          <div>Registered: <span id="modal-registered" class="font-semibold text-slate-700">—</span></div>
          <div>Last Activity: <span id="modal-last-login" class="font-semibold text-slate-700">—</span></div>
        </div>'''

if old_footer in html:
    html = html.replace(old_footer, new_footer, 1)
    print("Modal footer successfully cleaned.")
else:
    print("WARNING: old_footer not found")

# 4. Update viewTutorDetails fee logic in JS
old_modal_fee = "document.getElementById('modal-fee').textContent = t.hourly_fee_avg ? `₹${Math.round(t.hourly_fee_avg)} / hr` : (t.fee_raw || 'Negotiable / Unstated');"
new_modal_fee = '''        if (t.hourly_fee_avg) {
          document.getElementById('modal-fee').textContent = `₹${Math.round(t.hourly_fee_avg)} / hr`;
          const lbl = document.getElementById('modal-fee-label');
          if (lbl) lbl.textContent = 'Hourly Rate';
        } else if (t.monthly_fee_avg) {
          document.getElementById('modal-fee').textContent = `₹${Math.round(t.monthly_fee_avg).toLocaleString()} / mo`;
          const lbl = document.getElementById('modal-fee-label');
          if (lbl) lbl.textContent = 'Monthly Rate';
        } else {
          document.getElementById('modal-fee').textContent = t.fee_raw ? t.fee_raw.split('(')[0].trim() : 'Negotiable';
          const lbl = document.getElementById('modal-fee-label');
          if (lbl) lbl.textContent = 'Tuition Fee';
        }'''

if old_modal_fee in html:
    html = html.replace(old_modal_fee, new_modal_fee, 1)
    print("Modal fee JS logic updated.")
else:
    print("WARNING: old_modal_fee not found")

# 5. Fix discrepancy text in docx modal
html = html.replace('(87.4k active + 19.0k dormant + 1.2k unstated)', '(76.7k verified active + 31.0k inactive/removed)')

with open('app/index.html', 'w') as f:
    f.write(html)

print("update_index_table_and_modal.py finished successfully.")
