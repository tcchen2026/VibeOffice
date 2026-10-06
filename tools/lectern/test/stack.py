"""Stack montage images vertically, k per sheet. Usage: python3 -I stack.py inDir outDir k"""
import os, sys
from PIL import Image
src, out, k = sys.argv[1], sys.argv[2], int(sys.argv[3])
os.makedirs(out, exist_ok=True)
fs = sorted(f for f in os.listdir(src) if f.endswith('.png'))
for i in range(0, len(fs), k):
    ims = [Image.open(os.path.join(src, f)).convert('RGB') for f in fs[i:i + k]]
    W = max(im.width for im in ims); H = sum(im.height for im in ims) + 6 * (len(ims) - 1)
    c = Image.new('RGB', (W, H), (30, 30, 30)); y = 0
    for im in ims: c.paste(im, (0, y)); y += im.height + 6
    c.save(os.path.join(out, 'sheet%03d.png' % (i // k)))
print(len(range(0, len(fs), k)), 'sheets')
