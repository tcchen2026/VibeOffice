"""Compare Lectern's slide renderings with LibreOffice's (both 480 px wide PNGs).
For every slide that has both, report a difference score (0 = identical) computed on small blurred images,
which is tolerant to anti-aliasing and font differences but sensitive to missing or misplaced content,
wrong colours and wrong backgrounds. Writes a JSONL of scores and side-by-side montages of the worst ones.
Usage: python3 -I viscompare.py lecternPngDir loPngDir out.jsonl [montageDir] [N]"""
import json, os, re, sys
from PIL import Image, ImageFilter, ImageChops, ImageDraw, ImageStat

lec, lo, outp = sys.argv[1], sys.argv[2], sys.argv[3]
mdir = sys.argv[4] if len(sys.argv) > 4 else None
N = int(sys.argv[5]) if len(sys.argv) > 5 else 60

lo_pages = {}
for f in os.listdir(lo):
    m = re.match(r'(.*)-0*(\d+)\.png$', f)
    if m: lo_pages[(m.group(1), int(m.group(2)))] = os.path.join(lo, f)

def prep(path, size):
    im = Image.open(path).convert('RGB').resize(size, Image.BILINEAR)
    return im.filter(ImageFilter.GaussianBlur(1.5))

rows = []
for f in sorted(os.listdir(lec)):
    m = re.match(r'(.*)__(\d+)\.png$', f)
    if not m: continue
    key = (m.group(1), int(m.group(2)))
    if key not in lo_pages: continue
    a = Image.open(os.path.join(lec, f)); b = Image.open(lo_pages[key])
    ar, br = a.height / a.width, b.height / b.width
    size = (96, max(1, round(96 * ar)))
    if abs(ar - br) > 0.02:
        rows.append({'deck': key[0], 'slide': key[1], 'score': 99.0, 'note': 'aspect %.3f vs %.3f' % (ar, br)})
        continue
    pa, pb = prep(os.path.join(lec, f), size), prep(lo_pages[key], size)
    diff = ImageChops.difference(pa, pb)
    st = ImageStat.Stat(diff)
    mad = sum(st.mean) / 3
    # share of pixels that differ a lot (content present in one and not the other)
    g = diff.convert('L').point(lambda v: 255 if v > 60 else 0)
    big = ImageStat.Stat(g).mean[0] / 255 * 100
    rows.append({'deck': key[0], 'slide': key[1], 'score': round(mad, 2), 'big': round(big, 2), 'lec': os.path.join(lec, f), 'lo': lo_pages[key]})

with open(outp, 'w') as o:
    for r in rows: o.write(json.dumps(r) + '\n')
scores = sorted(r['score'] for r in rows)
print('slides compared', len(rows), 'decks', len({r['deck'] for r in rows}))
if scores:
    for q in (0.5, 0.75, 0.9, 0.95):
        print('  p%d score %.2f' % (q * 100, scores[int(len(scores) * q) - 1]))
    print('  slides with score > 12:', sum(1 for s in scores if s > 12), ' > 20:', sum(1 for s in scores if s > 20))

if mdir:
    os.makedirs(mdir, exist_ok=True)
    worst = sorted([r for r in rows if 'lec' in r], key=lambda r: -r['score'])[:N]
    for i, r in enumerate(worst):
        a = Image.open(r['lec']).convert('RGB'); b = Image.open(r['lo']).convert('RGB')
        b = b.resize((a.width, round(a.width * b.height / b.width)))
        W = a.width * 2 + 12; H = max(a.height, b.height) + 22
        c = Image.new('RGB', (W, H), (90, 90, 90))
        c.paste(a, (0, 22)); c.paste(b, (a.width + 12, 22))
        d = ImageDraw.Draw(c)
        d.text((4, 4), ('Lectern  |  LibreOffice   %s s%d  score %.1f' % (r['deck'][:60], r['slide'], r['score'])), fill=(255, 255, 255))
        c.save(os.path.join(mdir, '%03d_%s_s%d.png' % (i, r['deck'][:50], r['slide'])))
    print('montages in', mdir)
