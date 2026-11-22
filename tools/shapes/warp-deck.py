#!/usr/bin/env python3
"""Every DrawingML text warp (presetTextWarpDefinitions.xml, common/presets.js) in a labelled cell, in variants (one,
two and three lines, small and long text; insets, anchoring and alignment), written over the slides of
tmp/shapes/shapes.pptx (tools/shapes/deck.mjs) as tmp/shapes/warps.pptx: open it in PowerPoint and save it as PDF,
then compare its drawing with ours (tools/shapes/compare.py with the cells' names in warps.txt).
  python3 -I tools/shapes/warp-deck.py"""
import pathlib, re, zipfile
from xml.sax.saxutils import escape

ROOT = pathlib.Path(__file__).resolve().parents[2]
BASE, OUT = ROOT / 'tmp/shapes/shapes.pptx', ROOT / 'tmp/shapes/warps.pptx'
EMU = 12700
COLS, ROWS = 5, 3


def names():
    src = (ROOT / 'refs/openoffice/main/oox/source/export/presetTextWarpDefinitions.xml').read_text()
    return sorted(set(re.findall(r'^\s*<(text\w+)>', src, re.M)))


def run(t, sz, extra=''):
    return f'<a:r><a:rPr lang="en-US" sz="{sz * 100}" b="1" dirty="0"{extra}><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>{escape(t)}</a:t></a:r>'


def shape(sid, name, x, y, w, h, warp, lines, sz, body='', algn='ctr', outline=True):
    paras = ''.join(f'<a:p><a:pPr algn="{algn}"/>{run(t, sz)}</a:p>' for t in lines)
    ln = '<a:ln w="3175"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>' if outline else '<a:ln><a:noFill/></a:ln>'
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{escape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{round(x * EMU)}" y="{round(y * EMU)}"/><a:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>{ln}</p:spPr>'
            f'<p:txBody><a:bodyPr wrap="none" {body or "anchor=\"ctr\""} rtlCol="0"><a:prstTxWarp prst="{warp}"><a:avLst/></a:prstTxWarp><a:noAutofit/></a:bodyPr><a:lstStyle/>{paras}</p:txBody></p:sp>')


def label(sid, text, x, y):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="label {sid}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{round(x * EMU)}" y="{round(y * EMU)}"/><a:ext cx="{184 * EMU}" cy="{16 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="en-US" sz="800" dirty="0"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:rPr><a:t>{escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>')


def main():
    ws = names()
    cells = []   # (label, warp, lines, size, bodyPr attributes, alignment)
    for tag, lines, sz in [('1', ['WordArt'], 28), ('2', ['Word', 'Art Text'], 28), ('3', ['One', 'Two', 'Three'], 28), ('small', ['WordArt'], 12), ('long', ['The quick brown fox'], 28)]:
        cells += [(f'{w} {tag}', w, lines, sz, '', 'ctr') for w in ws]
    for w in ('textPlain', 'textArchUp', 'textCircle', 'textInflate', 'textWave1', 'textDeflateInflate', 'textButton'):
        cells += [(f'{w} ins0', w, ['WordArt'], 28, 'lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"', 'ctr'),
                  (f'{w} top', w, ['WordArt'], 28, 'anchor="t"', 'ctr'),
                  (f'{w} left', w, ['Word', 'Art Text'], 28, 'anchor="ctr"', 'l'),
                  (f'{w} 2para4', w, ['Word', 'Art', 'Text', 'Four'], 20, '', 'ctr')]
    per = COLS * ROWS
    slides = [cells[i:i + per] for i in range(0, len(cells), per)]
    zin = zipfile.ZipFile(BASE)
    have = sorted(int(m.group(1)) for n in zin.namelist() for m in [re.match(r'ppt/slides/slide(\d+)\.xml$', n)] if m)
    assert len(slides) <= len(have), (len(slides), len(have))
    head = zin.read('ppt/slides/slide1.xml').decode().split('<p:sp>')[0]
    tail = '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    out = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
    for info in zin.infolist():
        m = re.match(r'ppt/slides/slide(\d+)\.xml$', info.filename)
        if not m:
            out.writestr(info, zin.read(info.filename)); continue
        k = int(m.group(1)) - 1
        xml, sid = '', 2
        for i, (lab, w, lines, sz, body, algn) in enumerate(slides[k] if k < len(slides) else []):
            cx, cy = 20 + (i % COLS) * 188, 14 + (i // COLS) * 176
            xml += shape(sid, lab, cx + 12, cy + 26, 160, 110, w, lines, sz, body, algn) + label(sid + 1, lab, cx, cy + 158)
            sid += 2
        out.writestr(info, head + xml + tail)
    out.close()
    (ROOT / 'tmp/shapes/warps.txt').write_text('\n'.join(c[0] for c in cells) + '\n')
    print(len(cells), 'cells on', len(slides), 'of', len(have), 'slides:', OUT.relative_to(ROOT))


main()
