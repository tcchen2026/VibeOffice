"""Side-by-side montages for chosen slides: before | after | LibreOffice.
Usage: python3 -I triptych.py slides.json beforeDir afterDir loDir outDir
slides.json: [[deckFile, slideNo], ...]"""
import json, os, sys
from PIL import Image, ImageDraw
sl, before, after, lo, out = sys.argv[1:6]
os.makedirs(out, exist_ok=True)
n = 0
for f, k in json.load(open(sl)):
    b = os.path.splitext(f)[0]
    paths = [os.path.join(before, '%s__%d.png' % (b, k)), os.path.join(after, '%s__%d.png' % (b, k))]
    lop = None
    for cand in ('%s-%d.png' % (b, k), '%s-%02d.png' % (b, k), '%s-%03d.png' % (b, k)):
        if os.path.exists(os.path.join(lo, cand)): lop = os.path.join(lo, cand); break
    paths.append(lop)
    ims = []
    for p in paths:
        if p and os.path.exists(p): ims.append(Image.open(p).convert('RGB'))
        else: ims.append(None)
    if not ims[1]: continue
    W = ims[1].width
    ims = [im.resize((W, round(W * im.height / im.width))) if im else Image.new('RGB', (W, ims[1].height), (60, 60, 60)) for im in ims]
    H = max(im.height for im in ims) + 20
    c = Image.new('RGB', (W * 3 + 16, H), (90, 90, 90))
    for i, im in enumerate(ims): c.paste(im, (i * (W + 8), 20))
    ImageDraw.Draw(c).text((4, 4), 'before | after | LibreOffice   %s s%d' % (b[:70], k), fill=(255, 255, 255))
    c.save(os.path.join(out, '%03d.png' % n)); n += 1
print('montages', n)
