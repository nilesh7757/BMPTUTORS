import re
import os

with open('/home/nilesh7757/BMPTUTORS1/app/index.html.bak', 'r') as f:
    orig = f.read()

# Locate script
script_start = orig.find('<script>\n    // State')
if script_start == -1:
    script_start = orig.find('<script>\n    // State')
    if script_start == -1:
        # find any script after main
        script_start = orig.rfind('<script>')

script_end = orig.rfind('</script>')
orig_script = orig[script_start + len('<script>'):script_end]

print("Script length:", len(orig_script))
