"""Add interleaved readable and unread decorations to a pinned public design deck."""
import argparse
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as E

ap = argparse.ArgumentParser(description=__doc__)
ap.add_argument('output', type=Path)
ap.add_argument('--corpus', type=Path, default=Path.home() / 'corpora/powerpoint')
args = ap.parse_args()
args.output.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(args.corpus / 'libreoffice__07eb8b76b2cf__master-slides.pptx') as z:
    parts = {n: z.read(n) for n in z.namelist()}
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
for name, data in list(parts.items()):
    if not re.fullmatch(r'ppt/slide(?:Masters|Layouts)/[^/]+\.xml', name):
        continue
    tree = E.fromstring(data)
    ident = max(int(e.get('id')) for e in tree.iter('{' + P + '}cNvPr')) + 1
    def shape(n, x):
        return (f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{n}" name="Review decoration {n - ident + 1}"/>'
                '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
                f'<a:xfrm><a:off x="{x}" y="381000"/><a:ext cx="381000" cy="381000"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="225588"/></a:solidFill></p:spPr></p:sp>')
    unknown = (f'<p:graphicFrame xmlns:p="{P}" xmlns:a="{A}"><p:nvGraphicFramePr><p:cNvPr id="{ident + 1}" name="Unread decoration"/>'
               '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="762000" y="381000"/>'
               '<a:ext cx="381000" cy="381000"/></p:xfrm><a:graphic><a:graphicData uri="urn:vibeoffice:test:unread">'
               '<test:object xmlns:test="urn:vibeoffice:test:unread" value="retain this object"/>'
               '</a:graphicData></a:graphic></p:graphicFrame>')
    # Office repairs the invented graphic type when it is active content. Keep
    # this preservation probe in an unsupported choice with an empty fallback.
    unknown = (f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:test="urn:vibeoffice:test:unread">'
               f'<mc:Choice Requires="test">{unknown}</mc:Choice><mc:Fallback/></mc:AlternateContent>')
    xml = data.decode('utf-8-sig')
    end = re.search(r'</(?:\w+:)?spTree>', xml).start()
    parts[name] = (xml[:end] + shape(ident, 381000) + unknown + shape(ident + 2, 1143000) + xml[end:]).encode()
with zipfile.ZipFile(args.output / 'unread-decorations.pptx', 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in parts.items():
        z.writestr(name, data)
