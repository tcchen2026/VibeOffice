"""Derive a two-member fallback from a public chartEx sample; keep it outside git."""
from pathlib import Path
import argparse
import re
import zipfile

ap = argparse.ArgumentParser()
ap.add_argument('output', type=Path)
ap.add_argument('--corpus', type=Path, default=Path.home() / 'corpora/powerpoint')
a = ap.parse_args()
a.output.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(a.corpus / 'openxml-sdk__254c9ac4ef06__Of16-03.pptx') as z:
    parts = {n: z.read(n) for n in z.namelist()}
name = 'ppt/slides/slide1.xml'
xml = parts[name].decode('utf-8-sig')
match = re.search(r'<mc:Fallback>(.*?)</mc:Fallback>', xml, re.S)
picture = match[1]
width = int(re.search(r'<a:ext cx="(\d+)"', picture)[1])
x = int(re.search(r'<a:off x="(\d+)"', picture)[1])
first = picture.replace(f'cx="{width}"', f'cx="{width // 2}"')
first = re.sub(r'<a:xfrm\b[^>]*>', '<a:xfrm rot="1800000" flipH="1">', first, count=1)
second = re.sub(r'(<p:cNvPr\b[^>]*\bid=")\d+', r'\g<1>100001', first)
second = second.replace('rot="1800000" flipH="1"', 'rot="4200000" flipV="1"', 1)
second = second.replace(f'x="{x}"', f'x="{x + width // 2}"')
xml = xml[:match.start(1)] + first + second + xml[match.end(1):]
parts[name] = xml.encode()
with zipfile.ZipFile(a.output / 'multiple-fallback.pptx', 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in parts.items(): z.writestr(name, data)
print(a.output / 'multiple-fallback.pptx')
