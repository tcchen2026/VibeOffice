#!/usr/bin/env python3
"""Slides as PowerPoint draws them against Lectern's pictures of them (tools/lectern-pictures.mjs): a score per slide
(the mean difference of small blurred copies, plus the share of pixels far apart: 0 is the same picture) and side-by-side
pictures, PowerPoint's on the left, worst first by name; outDir/scores.json has every score.
  python3 -I tools/lectern-vs-powerpoint.py officeDir oursDir outDir
officeDir holds <name>/file/office.pdf, as the Office test API's client writes them (docs/testing.md)."""
import json, os, re, subprocess, sys
from PIL import Image, ImageChops, ImageFilter, ImageStat

pp, lec, out = sys.argv[1:4]
os.makedirs(out, exist_ok=True)
rows = []
for deck in sorted(os.listdir(pp)):
    pdf = os.path.join(pp, deck, 'file', 'office.pdf')
    if not os.path.exists(pdf) or not os.path.exists(os.path.join(lec, '%s__1.png' % deck)):
        continue
    pdir = os.path.join(pp, deck, 'file', 'p960')
    if not os.path.isdir(pdir):
        os.makedirs(pdir)
        subprocess.run(['pdftoppm', '-scale-to-x', '960', '-scale-to-y', '-1', '-png', pdf, os.path.join(pdir, 'p')], check=True)
    pages = sorted(os.listdir(pdir), key=lambda f: int(re.search(r'(\d+)', f).group(1)))
    for i, f in enumerate(pages):
        mine = os.path.join(lec, '%s__%d.png' % (deck, i + 1))
        if not os.path.exists(mine):
            rows.append({'deck': deck, 'slide': i + 1, 'score': 99, 'note': 'no Lectern slide'}); continue
        a = Image.open(os.path.join(pdir, f)).convert('RGB'); b = Image.open(mine).convert('RGB')
        if abs(a.height - b.height * a.width / b.width) > 2:
            rows.append({'deck': deck, 'slide': i + 1, 'score': 98, 'note': 'height %d vs %d' % (a.height, round(b.height * a.width / b.width))})
        b = b.resize(a.size)
        size = (192, round(192 * a.height / a.width))
        pa, pb = (x.resize(size, Image.BILINEAR).filter(ImageFilter.GaussianBlur(1.2)) for x in (a, b))
        d = ImageChops.difference(pa, pb)
        mad = sum(ImageStat.Stat(d).mean) / 3
        big = ImageStat.Stat(d.convert('L').point(lambda v: 255 if v > 50 else 0)).mean[0] / 2.55
        score = round(mad + big, 2)
        rows.append({'deck': deck, 'slide': i + 1, 'score': score, 'mad': round(mad, 2), 'big': round(big, 2)})
        m = Image.new('RGB', (a.width * 2 + 12, a.height), (255, 0, 0))
        m.paste(a, (0, 0)); m.paste(b, (a.width + 12, 0))
        m.save(os.path.join(out, '%06.2f__%s__%d.png' % (score, deck[:70], i + 1)))
rows.sort(key=lambda r: -r['score'])
json.dump(rows, open(os.path.join(out, 'scores.json'), 'w'), indent=1)
for r in rows:
    print('%6.2f  %s #%d %s' % (r['score'], r['deck'][:90], r['slide'], r.get('note', '')))
