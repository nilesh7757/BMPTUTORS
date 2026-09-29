#!/usr/bin/env python3
"""
Generate individual district diagram maps for all 36 states and union territories in India.
Each diagram displays:
- State district boundaries
- District names and verified tutor counts labeled directly inside/on the district polygon
- Choropleth fill based on local tutor headcount
"""

import os
import sys
import json
import sqlite3
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
import matplotlib.colors as mcolors

sys.path.append('/home/nilesh7757/BMPTUTORS1/app')
from geo_service import get_enriched_states_geojson, get_enriched_districts_geojson

OUTPUT_DIR = "/home/nilesh7757/BMPTUTORS1/extracted_images/state_maps"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def slugify(text):
    return text.lower().replace(" ", "_").replace("&", "and").replace("/", "_").replace("(", "").replace(")", "").replace("-", "_")

def render_state_map(st_name, feats, out_path):
    fig, ax = plt.subplots(figsize=(8, 7.5), dpi=220, facecolor='white')
    ax.set_facecolor('white')
    
    patches = []
    counts = []
    labels = []
    
    for f in feats:
        p = f['properties']
        dname = p.get('district', '')
        cnt = p.get('tutor_count', 0)
        geom = f['geometry']
        coords = geom['coordinates']
        rings = []
        if geom['type'] == 'Polygon':
            for r in coords: rings.append(np.array(r))
        elif geom['type'] == 'MultiPolygon':
            for poly in coords:
                for r in poly: rings.append(np.array(r))
        for r in rings:
            patches.append(Polygon(r, closed=True))
            counts.append(cnt)
        if rings:
            all_pts = np.vstack(rings)
            cx = (all_pts[:, 0].min() + all_pts[:, 0].max()) / 2
            cy = (all_pts[:, 1].min() + all_pts[:, 1].max()) / 2
            labels.append((dname, cnt, cx, cy))
            
    max_c = max(counts) if counts and max(counts) > 0 else 10
    norm = mcolors.PowerNorm(gamma=0.35, vmin=0, vmax=max_c)
    cmap = plt.cm.YlGnBu
    
    pcol = PatchCollection(patches, cmap=cmap, norm=norm, edgecolor='#1e293b', linewidth=0.8, alpha=0.92)
    pcol.set_array(np.array(counts))
    ax.add_collection(pcol)
    
    # Sort labels so districts with higher counts stand out
    labels.sort(key=lambda x: x[1], reverse=True)
    
    # If a state has many districts (e.g. UP 70, MP 48), adjust label font size
    num_dist = len(labels)
    base_fsize = 7.0 if num_dist <= 20 else (6.2 if num_dist <= 35 else 5.2)
    
    for dname, cnt, cx, cy in labels:
        if num_dist > 40 and cnt == 0:
            continue # Don't clutter if zero tutors in crowded 70-district states
            
        lbl = f"{dname}\n{cnt:,}" if cnt > 0 else dname
        is_high = (cnt >= 100) or (num_dist <= 15 and cnt > 0)
        
        ax.annotate(lbl, xy=(cx, cy), 
                    fontsize=base_fsize if is_high else (base_fsize - 0.8),
                    fontweight='bold' if is_high else 'normal',
                    color='#0f172a' if is_high else '#475569',
                    bbox=dict(boxstyle='round,pad=0.18', facecolor='white', alpha=0.85 if is_high else 0.65, 
                              edgecolor='#cbd5e1' if not is_high else '#94a3b8', linewidth=0.5),
                    ha='center', va='center')
                    
    ax.autoscale()
    ax.set_aspect('equal')
    ax.axis('off')
    
    tot_cnt = sum(f['properties'].get('tutor_count', 0) for f in feats)
    ax.set_title(f"{st_name}: District-Wise Tutoring Supply (Total N = {tot_cnt:,})", 
                 fontsize=11.5, fontweight='bold', pad=12, color='#0f172a')
                 
    plt.tight_layout()
    plt.savefig(out_path, facecolor='white', bbox_inches='tight')
    plt.close()

def render_puducherry_map(feats, out_path):
    fig, axes = plt.subplots(2, 2, figsize=(9, 8.5), dpi=220, facecolor='white')
    axes = axes.flatten()

    enclave_meta = {
        'Puducherry': ('Puducherry District (Main Enclave, Coromandel Coast, TN)', '#818cf8'),
        'Karaikal': ('Karaikal District (Enclave within Tamil Nadu, 130 km South)', '#bae6fd'),
        'Mahe': ('Mahe District (Enclave within Kerala, Malabar Coast)', '#ccfbf1'),
        'Yanam': ('Yanam District (Enclave within Andhra Pradesh, Godavari Delta)', '#ccfbf1')
    }

    # Sort feats by custom order
    order = ['Puducherry', 'Karaikal', 'Mahe', 'Yanam']
    sorted_f = sorted(feats, key=lambda f: order.index(f['properties'].get('district', '')) if f['properties'].get('district', '') in order else 99)

    for idx, f in enumerate(sorted_f):
        ax = axes[idx]
        ax.set_facecolor('#f8fafc')
        p = f['properties']
        dname = p.get('district', '')
        cnt = p.get('tutor_count', 0)
        fee = p.get('avg_fee', 0)
        
        geom = f['geometry']
        coords = geom['coordinates']
        rings = []
        if geom['type'] == 'Polygon':
            for r in coords: rings.append(np.array(r))
        elif geom['type'] == 'MultiPolygon':
            for poly in coords:
                for r in poly: rings.append(np.array(r))
                
        patches = [Polygon(r, closed=True) for r in rings]
        all_pts = np.vstack(rings)
        cx = (all_pts[:, 0].min() + all_pts[:, 0].max()) / 2
        cy = (all_pts[:, 1].min() + all_pts[:, 1].max()) / 2
        
        meta_title, fc = enclave_meta.get(dname, (f'{dname} Enclave', '#bae6fd'))
        pcol = PatchCollection(patches, facecolor=fc, edgecolor='#1e293b', linewidth=1.2, alpha=0.95)
        ax.add_collection(pcol)
        
        lbl = f"{dname}\n{cnt:,} Tutors\n(Avg: ₹{fee:.0f}/hr)" if fee > 0 else f"{dname}\n{cnt:,} Tutors"
        ax.annotate(lbl, xy=(cx, cy), fontsize=9, fontweight='bold', color='#0f172a',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.92, edgecolor='#94a3b8', linewidth=0.8),
                    ha='center', va='center')
                    
        ax.autoscale()
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        dx = (xlim[1] - xlim[0]) * 0.18
        dy = (ylim[1] - ylim[0]) * 0.18
        ax.set_xlim(xlim[0] - dx, xlim[1] + dx)
        ax.set_ylim(ylim[0] - dy, ylim[1] + dy)
        ax.set_aspect('equal')
        ax.axis('off')
        ax.set_title(meta_title, fontsize=9, fontweight='bold', pad=8, color='#1e293b')

    plt.suptitle('Puducherry UT: Complete 4-Enclave District Cartographic Breakdown\n(Total Registered Tutors: 360 | Statewide Rate: ₹700/hr)',
                 fontsize=11.5, fontweight='bold', color='#0f172a', y=0.98)
    plt.tight_layout()
    plt.savefig(out_path, facecolor='white', bbox_inches='tight')
    plt.close()

def render_dnh_daman_diu_map(feats, out_path):
    fig, axes = plt.subplots(1, len(feats), figsize=(10, 4.5), dpi=220, facecolor='white')
    if len(feats) == 1: axes = [axes]
    
    for idx, f in enumerate(feats):
        ax = axes[idx]
        ax.set_facecolor('#f8fafc')
        p = f['properties']
        dname = p.get('district', '')
        cnt = p.get('tutor_count', 0)
        fee = p.get('avg_fee', 0)
        
        geom = f['geometry']
        coords = geom['coordinates']
        rings = []
        if geom['type'] == 'Polygon':
            for r in coords: rings.append(np.array(r))
        elif geom['type'] == 'MultiPolygon':
            for poly in coords:
                for r in poly: rings.append(np.array(r))
        patches = [Polygon(r, closed=True) for r in rings]
        all_pts = np.vstack(rings)
        cx = (all_pts[:, 0].min() + all_pts[:, 0].max()) / 2
        cy = (all_pts[:, 1].min() + all_pts[:, 1].max()) / 2
        
        fc = '#818cf8' if cnt > 10 else '#bae6fd'
        pcol = PatchCollection(patches, facecolor=fc, edgecolor='#1e293b', linewidth=1.2, alpha=0.95)
        ax.add_collection(pcol)
        
        lbl = f"{dname}\n{cnt:,} Tutors"
        ax.annotate(lbl, xy=(cx, cy), fontsize=8.5, fontweight='bold', color='#0f172a',
                    bbox=dict(boxstyle='round,pad=0.25', facecolor='white', alpha=0.92, edgecolor='#94a3b8', linewidth=0.8),
                    ha='center', va='center')
        ax.autoscale()
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        dx = (xlim[1] - xlim[0]) * 0.15
        dy = (ylim[1] - ylim[0]) * 0.15
        ax.set_xlim(xlim[0] - dx, xlim[1] + dx)
        ax.set_ylim(ylim[0] - dy, ylim[1] + dy)
        ax.set_aspect('equal')
        ax.axis('off')
        ax.set_title(f"{dname} Division", fontsize=9, fontweight='bold', pad=8, color='#1e293b')

    plt.suptitle('Dadra and Nagar Haveli and Daman and Diu: District Breakdown (Total N = 22)',
                 fontsize=11, fontweight='bold', color='#0f172a', y=0.98)
    plt.tight_layout()
    plt.savefig(out_path, facecolor='white', bbox_inches='tight')
    plt.close()

def main():
    print("Fetching states and generating individual district maps...")
    states_geo = get_enriched_states_geojson()
    sorted_states = sorted(states_geo['features'], key=lambda f: f['properties'].get('tutor_count', 0), reverse=True)
    
    for idx, feat in enumerate(sorted_states, 1):
        st_name = feat['properties']['state']
        slug = slugify(st_name)
        out_path = os.path.join(OUTPUT_DIR, f"{slug}.png")
        
        d_res = get_enriched_districts_geojson(st_name)
        feats = d_res.get('features', [])
        
        if st_name == 'Puducherry':
            render_puducherry_map(feats, out_path)
        elif st_name == 'Dadra and Nagar Haveli and Daman and Diu':
            render_dnh_daman_diu_map(feats, out_path)
        else:
            render_state_map(st_name, feats, out_path)
        print(f"[{idx}/36] Generated: {st_name} -> {out_path}")

if __name__ == "__main__":
    main()
