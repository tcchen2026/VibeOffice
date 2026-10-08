"""Add an AlternateContent/effects regression to an authored media picture.

python3 tools/lectern/test/media-variant.py EmbeddedVideo.pptx OUT.pptx [--strict]
The input/output fixtures remain outside the repository. Validate both before use.
"""
from pathlib import Path
import re
import sys
from zipfile import ZipFile, ZIP_DEFLATED

source, output = map(Path, sys.argv[1:3])
strict = '--strict' in sys.argv
with ZipFile(source) as z:
    parts = {name: z.read(name) for name in z.namelist()}
for name, data in parts.items():
    if not re.match(r'ppt/slides/slide\d+\.xml$', name): continue
    text = data.decode('utf-8-sig')
    found = next((m for m in re.finditer(r'<p:pic\b[\s\S]*?</p:pic>', text) if re.search(r'<a:(?:videoFile|audioFile)\b', m[0])), None)
    if not found: continue
    pic = found[0]
    # Known effects are ignored by Lectern's model today. They must survive a
    # new outer shadow, including in the unselected compatibility branch.
    effects = '<a:effectLst><a:glow rad="127000"><a:srgbClr val="EEDD11"/></a:glow><a:reflection blurRad="6350" dist="127000" dir="5400000"/></a:effectLst>'
    if strict: effects = effects.replace('<a:reflection', '<a:outerShdw blurRad="12700" dist="12700" dir="5400000"><a:srgbClr val="000000"/></a:outerShdw><a:reflection')
    scene = '<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d><a:bevelT w="76200" h="76200"/></a:sp3d>'
    assert '<a:effectLst' not in pic and '<a:scene3d' not in pic, 'Use an authored picture without effects as the base'
    pic = pic.replace('<p:pic>', '<p:pic probe:role="probe:Video">').replace('</p:spPr>', effects + scene + '</p:spPr>')
    fallback = re.sub(r'<a:(?:videoFile|audioFile)\b[^>]*/>', '', pic)
    fallback = re.sub(r'<p:extLst>\s*<p:ext\b[^>]*>\s*<p14:media\b[\s\S]*?</p:extLst>', '', fallback)
    wrapper = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:probe="urn:vibeoffice:preservation-test" mc:Ignorable="probe"><mc:Choice Requires="p14">' + pic + '</mc:Choice><mc:Fallback>' + fallback + '</mc:Fallback></mc:AlternateContent>'
    parts[name] = (text[:found.start()] + wrapper + text[found.end():]).encode()
    break
else: raise ValueError('No media picture was found')
if strict:
    for name, data in parts.items():
        if not name.endswith(('.xml', '.rels')): continue
        text = data.decode('utf-8-sig')
        for old, new in [('drawingml/2006/main', 'drawingml/main'), ('presentationml/2006/main', 'presentationml/main'), ('officeDocument/2006/relationships', 'officeDocument/relationships')]:
            text = text.replace('http://schemas.openxmlformats.org/' + old, 'http://purl.oclc.org/ooxml/' + new)
        parts[name] = text.encode()
output.parent.mkdir(parents=True, exist_ok=True)
with ZipFile(output, 'w', ZIP_DEFLATED) as z:
    for name, data in parts.items(): z.writestr(name, data)
print(output)
