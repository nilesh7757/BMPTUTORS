with open('app/report.html', 'r') as f:
    html = f.read()

# 1. Remove Section 1C (District-wise data)
import re
p = re.compile(r'<!-- SECTION 1C: INDIVIDUAL STATE DISTRICT-WISE DATA & PROFILES -->.*?</section>', re.DOTALL)
if p.search(html):
    html = p.sub('', html, count=1)
    print('Removed Section 1C (District level analysis) from report.html')
else:
    print('WARNING: Section 1C not found')

# 2. Remove Online % from Table 1 Header
html = html.replace('<th class="py-2 px-3 text-right">Online %</th>', '')

# 3. Remove Online % from Table 1 JS rendering
html = html.replace('<td class="py-2 px-3 text-right">${r.online_pct}%</td>', '')
html = html.replace('<th class="py-2 px-3 text-right">${d.total_online_pct}%</th>', '')

# 4. Remove district initialization in JS
html = html.replace('initDistrictReport(statesData);', '// District analysis removed per mentor instructions')

# 5. Clean notes
html = html.replace('Online percentage indicates availability for remote/video instruction.', '')
html = html.replace('All 37 States, Tables 2 & 2.8, Figures 1 & 3', 'All 36 States/UTs, Tables 1-3, Figures 1-4')

with open('app/report.html', 'w') as f:
    f.write(html)

print('clean_report_html.py finished successfully.')
