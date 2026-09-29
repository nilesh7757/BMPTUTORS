#!/usr/bin/env python3
"""
Generate Complete Map Portfolio for BMP Tutors:
- All-India Tutor Headcount Choropleth
- All-India Hourly Tuition Fee Choropleth
- All-36 States Comprehensive District Atlas Grid
- 5 Regional Zone District Maps (covering all 36 States & UTs)
- Top 4 Anchor State District Focus Maps (Tamil Nadu, Kerala, Maharashtra, Karnataka)
"""

import os
import json
import sqlite3
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
import matplotlib.colors as mcolors

MAPS_DIR = "/home/nilesh7757/BMPTUTORS1/extracted_images/maps"
os.makedirs(MAPS_DIR, exist_ok=True)

# Load GeoJSON
with open('/home/nilesh7757/BMPTUTORS1/app/geojson/india_states.geojson') as f:
    states_geo = json.load(f)

with open('/home/nilesh7757/BMPTUTORS1/app/geojson/india_districts.geojson') as f:
    dist_geo = json.load(f)

# Group districts by state
dist_by_state = {}
for f in dist_geo['features']:
    st = f['properties'].get('NAME_1')
    if st not in dist_by_state:
        dist_by_state[st] = []
    dist_by_state[st].append(f)

# Load DB stats
conn = sqlite3.connect('/home/nilesh7757/BMPTUTORS1/cleaned_tutors_data.db')
c = conn.cursor()
db_stats = {}
for r in c.execute("""
    SELECT 
        state, 
        count(*) as cnt, 
        round(avg(case when hourly_fee_avg between 50 and 5000 then hourly_fee_avg end), 0) as fee, 
        round(avg(teaches_online)*100, 1) as online,
        sum(case when is_active = 1 then 1 else 0 end) as active
    FROM teachers 
    WHERE state IS NOT NULL 
    GROUP BY state
"""):
    db_stats[r[0]] = {
        'count': r[1], 
        'fee': r[2] or 650, 
        'online': r[3] or 95.0, 
        'active': r[4] or 0
    }

alias_map = {
    'Orissa': 'Odisha',
    'Uttaranchal': 'Uttarakhand',
    'Andaman & Nicobar': 'Andaman and Nicobar Islands',
    'Andaman and Nicobar Islands': 'Andaman and Nicobar Islands',
    'Dadra and Nagar Haveli and Daman and Diu': ['Dadra and Nagar Haveli', 'Daman and Diu'],
    'Jammu & Kashmir': 'Jammu and Kashmir',
    'Jammu and Kashmir': 'Jammu and Kashmir'
}

def get_stats(st):
    k = alias_map.get(st, st)
    if isinstance(k, list):
        cnt = sum(db_stats.get(x, {}).get('count', 0) for x in k)
        fees = [db_stats.get(x, {}).get('fee', 650) for x in k if x in db_stats]
        fee = round(sum(fees)/len(fees), 0) if fees else 650
        online = 95.0
        active = sum(db_stats.get(x, {}).get('active', 0) for x in k)
    else:
        info = db_stats.get(k, {})
        cnt = info.get('count', 0)
        fee = info.get('fee', 650)
        online = info.get('online', 95.0)
        active = info.get('active', 0)
    return cnt, fee, online, active

def extract_rings(geom):
    coords = geom['coordinates']
    rings = []
    if geom['type'] == 'Polygon':
        for r in coords:
            rings.append(np.array(r))
    elif geom['type'] == 'MultiPolygon':
        for p in coords:
            for r in p:
                rings.append(np.array(r))
    return rings

print("1. Generating All-India Tutor Headcount Choropleth...")
fig, ax = plt.subplots(figsize=(10, 11), dpi=300)
patches = []
counts = []
for feat in states_geo['features']:
    st_name = feat['properties'].get('ST_NM')
    cnt, fee, _, _ = get_stats(st_name)
    rings = extract_rings(feat['geometry'])
    for r in rings:
        patches.append(Polygon(r, closed=True))
        counts.append(cnt)

cmap = plt.cm.YlGnBu
norm = mcolors.Normalize(vmin=0, vmax=17000)
pcol = PatchCollection(patches, cmap=cmap, norm=norm, edgecolor='#334155', linewidth=0.75, alpha=0.95)
pcol.set_array(np.array(counts))
ax.add_collection(pcol)
ax.set_xlim(67, 98)
ax.set_ylim(6, 38)
ax.set_aspect('equal')
ax.axis('off')

top_annot = {
    'Tamil Nadu': (78.6, 11.1, 'Tamil Nadu\n16,811 (15.6%)'),
    'Kerala': (76.2, 10.2, 'Kerala\n14,337 (13.3%)'),
    'Maharashtra': (75.5, 19.5, 'Maharashtra\n9,550 (8.9%)'),
    'Karnataka': (75.7, 14.5, 'Karnataka\n9,300 (8.6%)'),
    'Delhi': (77.2, 28.6, 'Delhi NCR\n8,600 (8.0%)'),
    'Uttar Pradesh': (80.5, 27.0, 'Uttar Pradesh\n8,114 (7.5%)'),
    'Telangana': (79.0, 17.8, 'Telangana\n7,810 (7.2%)'),
    'West Bengal': (87.8, 23.5, 'West Bengal\n7,026 (6.5%)'),
    'Rajasthan': (73.5, 26.5, 'Rajasthan\n2,479 (2.3%)'),
    'Gujarat': (71.5, 22.5, 'Gujarat\n2,122 (2.0%)'),
    'Punjab': (75.3, 31.0, 'Punjab\n2,412 (2.2%)'),
    'Haryana': (76.0, 29.0, 'Haryana\n3,228 (3.0%)'),
    'Andhra Pradesh': (80.0, 15.5, 'Andhra Pradesh\n2,894 (2.7%)'),
    'Madhya Pradesh': (77.5, 23.5, 'Madhya Pradesh\n1,839 (1.7%)'),
    'Bihar': (85.5, 25.5, 'Bihar\n1,557 (1.4%)'),
    'Assam': (92.5, 26.2, 'Assam\n795 (0.7%)')
}

for st, (x, y, txt) in top_annot.items():
    ax.annotate(txt, xy=(x, y), fontsize=7.2, fontweight='bold', color='#0f172a',
                bbox=dict(boxstyle='round,pad=0.25', facecolor='white', alpha=0.9, edgecolor='#94a3b8', linewidth=0.6),
                ha='center', va='center')

cbar = plt.colorbar(pcol, ax=ax, orientation='horizontal', fraction=0.035, pad=0.02, shrink=0.6)
cbar.set_label('Total Registered Tutors (Headcount)', fontsize=10, fontweight='bold', color='#1e293b')
cbar.ax.tick_params(labelsize=8.5)
ax.set_title('Figure: All-India Tutoring Supply Distribution by State (N = 107,723)', 
             fontsize=13, fontweight='bold', pad=15, color='#0f172a')
plt.tight_layout()
out_map1 = os.path.join(MAPS_DIR, "all_india_tutor_choropleth.png")
plt.savefig(out_map1)
plt.close()
print("Saved:", out_map1)

print("2. Generating All-India Hourly Tuition Fee Choropleth...")
fig, ax = plt.subplots(figsize=(10, 11), dpi=300)
patches = []
fees = []
for feat in states_geo['features']:
    st_name = feat['properties'].get('ST_NM')
    _, fee, _, _ = get_stats(st_name)
    rings = extract_rings(feat['geometry'])
    for r in rings:
        patches.append(Polygon(r, closed=True))
        fees.append(fee if fee > 0 else 650)

cmap_fee = plt.cm.magma_r
norm_fee = mcolors.Normalize(vmin=550, vmax=1150)
pcol_fee = PatchCollection(patches, cmap=cmap_fee, norm=norm_fee, edgecolor='#334155', linewidth=0.75, alpha=0.95)
pcol_fee.set_array(np.array(fees))
ax.add_collection(pcol_fee)
ax.set_xlim(67, 98)
ax.set_ylim(6, 38)
ax.set_aspect('equal')
ax.axis('off')

fee_annot = {
    'Maharashtra': (75.5, 19.5, 'Maharashtra\n₹1,039 / hr'),
    'Delhi': (77.2, 28.6, 'Delhi NCR\n₹1,020 / hr'),
    'Haryana': (76.0, 29.0, 'Haryana\n₹1,026 / hr'),
    'Karnataka': (75.7, 14.5, 'Karnataka\n₹903 / hr'),
    'Uttar Pradesh': (80.5, 27.0, 'Uttar Pradesh\n₹978 / hr'),
    'Telangana': (79.0, 17.8, 'Telangana\n₹943 / hr'),
    'West Bengal': (87.8, 23.5, 'West Bengal\n₹921 / hr'),
    'Gujarat': (71.5, 22.5, 'Gujarat\n₹981 / hr'),
    'Punjab': (75.3, 31.0, 'Punjab\n₹890 / hr'),
    'Tamil Nadu': (78.6, 11.1, 'Tamil Nadu\n₹713 / hr'),
    'Kerala': (76.2, 10.2, 'Kerala\n₹582 / hr'),
    'Assam': (92.5, 26.2, 'Assam\n₹774 / hr')
}

for st, (x, y, txt) in fee_annot.items():
    ax.annotate(txt, xy=(x, y), fontsize=7.2, fontweight='bold', color='#0f172a',
                bbox=dict(boxstyle='round,pad=0.25', facecolor='white', alpha=0.9, edgecolor='#94a3b8', linewidth=0.6),
                ha='center', va='center')

cbar_fee = plt.colorbar(pcol_fee, ax=ax, orientation='horizontal', fraction=0.035, pad=0.02, shrink=0.6)
cbar_fee.set_label('Average Hourly Tuition Fee (₹ / hr)', fontsize=10, fontweight='bold', color='#1e293b')
cbar_fee.ax.tick_params(labelsize=8.5)
ax.set_title('Figure: All-India Hourly Tuition Fee Distribution by State (₹500 - ₹1,100 Range)', 
             fontsize=13, fontweight='bold', pad=15, color='#0f172a')
plt.tight_layout()
out_map2 = os.path.join(MAPS_DIR, "all_india_fee_choropleth.png")
plt.savefig(out_map2)
plt.close()
print("Saved:", out_map2)

print("3. Generating All-36 States Comprehensive District Atlas Grid...")
sorted_states = sorted(dist_by_state.keys(), key=lambda s: get_stats(s)[0], reverse=True)
fig, axes = plt.subplots(6, 6, figsize=(18, 20), dpi=250)
axes = axes.flatten()

for idx, st_name in enumerate(sorted_states):
    ax = axes[idx]
    feats = dist_by_state[st_name]
    cnt, fee, onl, _ = get_stats(st_name)
    patches = []
    for f in feats:
        rings = extract_rings(f['geometry'])
        for r in rings:
            patches.append(Polygon(r, closed=True))
    
    if cnt >= 5000:
        fc, ec = '#c7d2fe', '#312e81' # Deep Indigo
    elif cnt >= 1500:
        fc, ec = '#bae6fd', '#0369a1' # Blue
    elif cnt >= 300:
        fc, ec = '#ccfbf1', '#0f766e' # Teal
    else:
        fc, ec = '#f1f5f9', '#64748b' # Gray
        
    pcol = PatchCollection(patches, facecolor=fc, edgecolor=ec, linewidth=0.5, alpha=0.9)
    ax.add_collection(pcol)
    ax.autoscale()
    ax.set_aspect('equal')
    ax.axis('off')
    
    disp_name = st_name
    if disp_name == 'Uttaranchal': disp_name = 'Uttarakhand'
    elif disp_name == 'Orissa': disp_name = 'Odisha'
    elif disp_name == 'Dadra and Nagar Haveli and Daman and Diu': disp_name = 'DNH & Daman Diu'
    elif disp_name == 'Andaman and Nicobar Islands': disp_name = 'A & N Islands'
    
    title_txt = f"{disp_name}\n({cnt:,} tutors | ₹{fee:.0f}/h)"
    ax.set_title(title_txt, fontsize=7.5, fontweight='bold', color='#0f172a', pad=3)

plt.suptitle('All 36 States & Union Territories of India: Administrative District Maps & Tutor Density Portfolio', 
             fontsize=14, fontweight='bold', y=0.99, color='#0f172a')
plt.tight_layout()
out_map3 = os.path.join(MAPS_DIR, "all_36_states_district_atlas.png")
plt.savefig(out_map3, bbox_inches='tight')
plt.close()
print("Saved:", out_map3)

print("4. Generating Regional State District Maps (6 Zones)...")

zones = {
    "south": {
        "title": "Southern Zone States (Tamil Nadu, Kerala, Karnataka, Telangana, Andhra Pradesh)",
        "states": ["Tamil Nadu", "Kerala", "Karnataka", "Telangana", "Andhra Pradesh", "Puducherry"],
        "filename": "map_zone_south.png",
        "bbox": (73.5, 8.0, 84.5, 20.0)
    },
    "west": {
        "title": "Western Zone States (Maharashtra, Gujarat, Goa, DNH & Daman Diu)",
        "states": ["Maharashtra", "Gujarat", "Goa", "Dadra and Nagar Haveli and Daman and Diu"],
        "filename": "map_zone_west.png",
        "bbox": (68.0, 14.5, 81.0, 25.0)
    },
    "north": {
        "title": "Northern Zone States (Delhi NCR, UP, Haryana, Punjab, Rajasthan, HP, Uttarakhand, J&K, Ladakh)",
        "states": ["Delhi", "Uttar Pradesh", "Haryana", "Punjab", "Rajasthan", "Himachal Pradesh", "Uttaranchal", "Jammu and Kashmir", "Ladakh", "Chandigarh"],
        "filename": "map_zone_north.png",
        "bbox": (69.0, 23.5, 85.0, 37.5)
    },
    "east_central": {
        "title": "Eastern & Central Zone States (West Bengal, Bihar, Odisha, Jharkhand, MP, Chhattisgarh)",
        "states": ["West Bengal", "Bihar", "Orissa", "Jharkhand", "Madhya Pradesh", "Chhattisgarh"],
        "filename": "map_zone_east_central.png",
        "bbox": (74.0, 17.5, 90.0, 28.5)
    },
    "northeast": {
        "title": "North-Eastern Zone States (Assam, Meghalaya, Tripura, Manipur, Nagaland, Mizoram, Arunachal, Sikkim)",
        "states": ["Assam", "Meghalaya", "Tripura", "Manipur", "Nagaland", "Mizoram", "Arunachal Pradesh", "Sikkim"],
        "filename": "map_zone_northeast.png",
        "bbox": (87.5, 21.5, 97.5, 29.8)
    }
}

for zone_key, zinfo in zones.items():
    fig, ax = plt.subplots(figsize=(9, 7.5), dpi=300)
    state_patches = {}
    
    for st_name in zinfo['states']:
        feats = dist_by_state.get(st_name, [])
        cnt, fee, _, _ = get_stats(st_name)
        patches = []
        for f in feats:
            rings = extract_rings(f['geometry'])
            for r in rings:
                patches.append(Polygon(r, closed=True))
        
        # Determine shade based on count
        if cnt > 10000:
            fc, ec = '#818cf8', '#312e81'
        elif cnt > 5000:
            fc, ec = '#a5b4fc', '#3730a3'
        elif cnt > 2000:
            fc, ec = '#bae6fd', '#0284c7'
        elif cnt > 500:
            fc, ec = '#ccfbf1', '#0d9488'
        else:
            fc, ec = '#f1f5f9', '#64748b'
            
        pcol = PatchCollection(patches, facecolor=fc, edgecolor=ec, linewidth=0.6, alpha=0.9)
        ax.add_collection(pcol)
        
        # Calculate state center
        if feats:
            all_pts = []
            for f in feats:
                for r in extract_rings(f['geometry']):
                    all_pts.append(r)
            if all_pts:
                pts_stacked = np.vstack(all_pts)
                cx, cy = (pts_stacked[:, 0].min() + pts_stacked[:, 0].max()) / 2, (pts_stacked[:, 1].min() + pts_stacked[:, 1].max()) / 2
                clean_name = st_name
                if clean_name == 'Uttaranchal': clean_name = 'Uttarakhand'
                elif clean_name == 'Orissa': clean_name = 'Odisha'
                elif clean_name == 'Dadra and Nagar Haveli and Daman and Diu': clean_name = 'DNH & DD'
                ax.text(cx, cy, f"{clean_name}\n({cnt:,})", fontsize=8, fontweight='bold',
                        color='#0f172a', ha='center', va='center',
                        bbox=dict(boxstyle='round,pad=0.2', facecolor='white', alpha=0.85, edgecolor='#cbd5e1', linewidth=0.5))

    minx, miny, maxx, maxy = zinfo['bbox']
    ax.set_xlim(minx, maxx)
    ax.set_ylim(miny, maxy)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title(f"Figure: {zinfo['title']}", fontsize=11.5, fontweight='bold', pad=12, color='#0f172a')
    plt.tight_layout()
    z_path = os.path.join(MAPS_DIR, zinfo['filename'])
    plt.savefig(z_path)
    plt.close()
    print("Saved:", z_path)

print("5. Generating Top 4 Anchor State District Focus Maps...")
focus_states = [
    ("Tamil Nadu", "map_focus_tamil_nadu.png", "Tamil Nadu District Map (Leading National Market: 16,811 Tutors)", ["Chennai", "Coimbatore", "Madurai", "Tiruchirappalli", "Salem"]),
    ("Kerala", "map_focus_kerala.png", "Kerala District Map (Highest Tutor Density: 14,337 Tutors)", ["Ernakulam", "Thiruvananthapuram", "Kozhikode", "Thrissur", "Kollam"]),
    ("Maharashtra", "map_focus_maharashtra.png", "Maharashtra District Map (Top Economic Market: 9,550 Tutors | ₹1,039/hr)", ["Mumbai Suburban", "Pune", "Thane", "Nagpur", "Nashik"]),
    ("Karnataka", "map_focus_karnataka.png", "Karnataka District Map (Premier Tech Cluster: 9,300 Tutors | ₹903/hr)", ["Bangalore Urban", "Mysore", "Dakshin Kannad", "Dharwad", "Belgaum"])
]

for st_name, fname, ftitle, key_districts in focus_states:
    feats = dist_by_state.get(st_name, [])
    fig, ax = plt.subplots(figsize=(8, 8.5), dpi=300)
    patches = []
    
    for f in feats:
        dname = f['properties'].get('NAME_2')
        rings = extract_rings(f['geometry'])
        for r in rings:
            patches.append(Polygon(r, closed=True))
        
        # District center label
        all_pts = np.vstack(rings)
        cx = (all_pts[:, 0].min() + all_pts[:, 0].max()) / 2
        cy = (all_pts[:, 1].min() + all_pts[:, 1].max()) / 2
        
        is_key = any(kd.lower() in dname.lower() for kd in key_districts)
        ax.text(cx, cy, dname, fontsize=6 if not is_key else 7.5,
                fontweight='bold' if is_key else 'normal',
                color='#1e1b4b' if is_key else '#475569',
                ha='center', va='center',
                bbox=dict(boxstyle='round,pad=0.15', facecolor='#fef08a' if is_key else 'white', 
                          alpha=0.85 if is_key else 0.6, edgecolor='#ca8a04' if is_key else '#cbd5e1', linewidth=0.5))

    cnt, fee, onl, act = get_stats(st_name)
    pcol = PatchCollection(patches, facecolor='#e0e7ff', edgecolor='#4338ca', linewidth=0.7, alpha=0.9)
    ax.add_collection(pcol)
    ax.autoscale()
    ax.set_aspect('equal')
    ax.axis('off')
    
    subtitle = f"Census: {cnt:,} Registered Educators | Avg Rate: ₹{fee:.0f}/hr | Online Delivery: {onl}% | Active: {act:,}"
    ax.set_title(f"{ftitle}\n{subtitle}", fontsize=10.5, fontweight='bold', pad=12, color='#0f172a')
    plt.tight_layout()
    f_path = os.path.join(MAPS_DIR, fname)
    plt.savefig(f_path)
    plt.close()
    print("Saved:", f_path)

print("All Cartographic Maps and Atlases Generated Successfully!")
