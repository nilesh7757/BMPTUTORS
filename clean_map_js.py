with open('app/index.html', 'r') as f:
    html = f.read()

# Update setMapMetric
old_set_map = """    function setMapMetric(metric) {
      currentMapMetric = metric;

      // Update button styles
      ['count', 'fee', 'online'].forEach(m => {
        const btn = document.getElementById(`btn-metric-${m}`);
        if (m === metric) {
          btn.className = "px-3 py-1 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-sm";
        } else {
          btn.className = "px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900 bg-slate-100";
        }
      });

      updateLegendUI();

      if (statesGeoDataCache) {
        renderIndiaChoropleth(statesGeoDataCache);
      }
      if (currentDistrictGeoCache) {
        renderDistrictChoropleth(currentDistrictGeoCache);
      }
    }"""

new_set_map = """    function setMapMetric(metric) {
      currentMapMetric = metric;

      // Update button styles
      ['count', 'fee'].forEach(m => {
        const btn = document.getElementById(`btn-metric-${m}`);
        if (!btn) return;
        if (m === metric) {
          btn.className = "px-3 py-1 rounded-lg text-xs font-bold transition-all bg-indigo-600 text-white shadow-sm";
        } else {
          btn.className = "px-3 py-1 rounded-lg text-xs font-bold transition-all text-slate-600 hover:text-slate-900 bg-slate-100";
        }
      });

      updateLegendUI();

      if (statesGeoDataCache) {
        renderIndiaChoropleth(statesGeoDataCache);
      }
    }"""

if old_set_map in html:
    html = html.replace(old_set_map, new_set_map, 1)
    print('Updated setMapMetric')
else:
    print('old_set_map not found')

# Update loadGeoJSONData
old_load_geo = """        // Populate State Dropdown
        const stateSelect = document.getElementById('map-state-select');
        stateSelect.innerHTML = '';
        
        // Sort states by tutor count descending
        const sortedFeatures = [...statesGeoDataCache.features].sort((a, b) => (b.properties.tutor_count || 0) - (a.properties.tutor_count || 0));
        
        sortedFeatures.forEach(feat => {
          const p = feat.properties;
          const opt = document.createElement('option');
          opt.value = p.state;
          opt.textContent = `${p.state} (${(p.tutor_count || 0).toLocaleString()} tutors)`;
          stateSelect.appendChild(opt);
        });

        renderIndiaChoropleth(statesGeoDataCache);

        // Load default state
        selectState(currentSelectedState);"""

new_load_geo = """        renderIndiaChoropleth(statesGeoDataCache);"""

if old_load_geo in html:
    html = html.replace(old_load_geo, new_load_geo, 1)
    print('Updated loadGeoJSONData')
else:
    print('old_load_geo not found')

with open('app/index.html', 'w') as f:
    f.write(html)
