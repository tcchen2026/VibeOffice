#!/usr/bin/env python3
"""Compares each preset's drawing in PowerPoint's PDF of tmp/shapes/shapes.pptx (tools/shapes/deck.mjs) with ours: the
cells' painted pixels (fill or outline), as a share of the two together that only one of them paints.
  python3 -I tools/shapes/compare.py office.pdf oursDir [outDir [deck]]     oursDir: <deck>__<slide>.png, 1920 px wide
Prints the worst shapes first and writes side-by-side crops of each (PowerPoint left) to outDir. deck: shapes (the
default), or warps for tools/shapes/warp-deck.py's text warps (cell names in tmp/shapes/warps.txt)."""
import os, subprocess, sys, tempfile
from PIL import Image, ImageChops, ImageFilter

pdf, ours = sys.argv[1], sys.argv[2]
out = sys.argv[3] if len(sys.argv) > 3 and sys.argv[3] != '-' else None
deck = sys.argv[4] if len(sys.argv) > 4 else 'shapes'
S = 2   # pixels per point
COLS, ROWS = 5, 3
tmp = tempfile.mkdtemp()
subprocess.run(['pdftoppm', '-scale-to-x', str(960 * S), '-scale-to-y', '-1', '-png', pdf, os.path.join(tmp, 'p')], check=True)
pages = sorted(os.listdir(tmp), key=lambda f: int(f.split('-')[-1].split('.')[0]))


def painted(im):
    """pixels that aren't (nearly) white, grown by a pixel to forgive anti-aliasing"""
    g = im.convert('L').point(lambda v: 255 if v < 235 else 0)
    return g


rows = []
per = COLS * ROWS
for si, f in enumerate(pages):
    a = Image.open(os.path.join(tmp, f)).convert('RGB')
    mine = os.path.join(ours, '%s__%d.png' % (deck, si + 1))
    if not os.path.exists(mine):
        continue
    b = Image.open(mine).convert('RGB').resize(a.size)
    for i in range(per):
        cx, cy = 20 + (i % COLS) * 188, 14 + (i // COLS) * 176
        box = tuple(v * S for v in (cx, cy, cx + 184, cy + 156))
        ca, cb = a.crop(box), b.crop(box)
        ma, mb = painted(ca), painted(cb)
        union = ImageChops.lighter(ma, mb)
        nu = sum(1 for v in union.get_flattened_data() if v)
        if not nu:
            continue
        # one-sided pixels, ignoring those within a pixel of the other side's paint
        da = ImageChops.subtract(ma, mb.filter(ImageFilter.MaxFilter(3)))
        db = ImageChops.subtract(mb, ma.filter(ImageFilter.MaxFilter(3)))
        nd = sum(1 for v in da.get_flattened_data() if v) + sum(1 for v in db.get_flattened_data() if v)
        rows.append((round(100 * nd / nu, 2), si + 1, i, ca, cb))
rows.sort(key=lambda r: -r[0])
names_file = os.path.join(ours, 'names.txt') if deck == 'shapes' else os.path.join(os.path.dirname(__file__), '../../tmp/shapes/%s.txt' % deck)
names = [l.strip() for l in open(names_file)] if os.path.exists(names_file) else None
half = len(pages) // 2 if deck == 'shapes' else len(pages)
for score, slide, i, ca, cb in rows:
    idx = ((slide - 1) % half) * per + i
    name = names[idx] if names and idx < len(names) else '?'
    print('%6.2f  slide %2d cell %2d  %s%s' % (score, slide, i, name, ' (tall)' if slide > half else ''))
    if out:
        os.makedirs(out, exist_ok=True)
        m = Image.new('RGB', (ca.width * 2 + 6, ca.height), (255, 0, 0))
        m.paste(ca, (0, 0)); m.paste(cb, (ca.width + 6, 0))
        m.save(os.path.join(out, '%06.2f-%s%s.png' % (score, name, '-tall' if slide > half else '')))
