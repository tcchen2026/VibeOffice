#!/usr/bin/env python3
"""Writes tests/shapes/powerpoint.json: where PowerPoint 2016 paints each preset shape (fill or 1 pt outline) in the wide
cells of tmp/shapes/shapes.pptx (tools/shapes/deck.mjs), from its PDF of the deck, at half a pixel per point, for
tests/shapes.test.mjs.
  python3 -I tools/shapes/masks.py office.pdf"""
import base64, json, os, pathlib, re, subprocess, sys, tempfile
from PIL import Image

ROOT = pathlib.Path(__file__).resolve().parents[2]
pdf = sys.argv[1]
names = sorted(k for k in re.findall(r'^  (\w+): \{', (ROOT / 'public/common/presets.js').read_text(), re.M)
               if not re.match(r'^(line|lineInv|straightConnector1|bentConnector\d|curvedConnector\d)$', k))
COLS, PER, M = 5, 15, 4          # cells per row and slide; the margin around each shape, in points
W, H = 110, 70                   # the wide cells' shapes, in points
tmp = tempfile.mkdtemp()
subprocess.run(['pdftoppm', '-scale-to-x', '480', '-scale-to-y', '-1', '-gray', pdf, os.path.join(tmp, 'p')], check=True)
pages = sorted(os.listdir(tmp), key=lambda f: int(f.split('-')[-1].split('.')[0]))
out = {'scale': 0.5, 'margin': M, 'w': W, 'h': H, 'shapes': {}}
for k, name in enumerate(names):
    page = Image.open(os.path.join(tmp, pages[k // PER]))
    i = k % PER
    cx, cy = 20 + (i % COLS) * 188, 14 + (i // COLS) * 176
    x, y = cx + (184 - W) / 2 - M, cy + 6 + (150 - H) / 2 - M
    x0, y0 = int(x // 2), int(y // 2)   # the crop's corner, in pixels; the shape's margin starts x / 2 - x0 into it
    crop = page.crop((x0, y0, x0 + (W + 2 * M) // 2, y0 + (H + 2 * M) // 2))
    bits = bytearray((crop.width * crop.height + 7) // 8)
    for j, v in enumerate(crop.get_flattened_data()):
        if v < 235: bits[j >> 3] |= 0x80 >> (j & 7)
    out['shapes'][name] = [base64.b64encode(bytes(bits)).decode(), x / 2 - x0, y / 2 - y0]
out['cols'], out['rows'] = crop.width, crop.height
dest = ROOT / 'tests/shapes/powerpoint.json'
dest.parent.mkdir(exist_ok=True)
dest.write_text(json.dumps(out, indent=0) + '\n')
print(len(out['shapes']), 'shapes,', dest.stat().st_size, 'bytes:', dest.relative_to(ROOT))
