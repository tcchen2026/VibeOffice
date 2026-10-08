"""Schema-authored modern comments on a public legacy-comment deck (not Office-authored).

MS-PPTX 2.16 and the SDK's modern-comment example define the part, author list,
slide extension and anchors. The pinned corpus has no modern comment threads.
https://learn.microsoft.com/en-us/office/open-xml/presentation/how-to-add-a-comment-to-a-slide-in-a-presentation
"""
import argparse
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as E

ap = argparse.ArgumentParser()
ap.add_argument('output', type=Path)
ap.add_argument('--corpus', type=Path, default=Path.home() / 'corpora/powerpoint')
a = ap.parse_args()
a.output.mkdir(parents=True, exist_ok=True)
source = a.corpus / 'poi__788ee515a5cd__45545_Comment.pptx'
with zipfile.ZipFile(source) as z:
    parts = {n: z.read(n) for n in z.namelist()}
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
M = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main'
PC = 'http://schemas.microsoft.com/office/powerpoint/2013/main/command'
AC = 'http://schemas.microsoft.com/office/drawing/2013/main/command'
REL = 'http://schemas.microsoft.com/office/2018/10/relationships/'
author = '{10000000-0000-4000-8000-000000000001}'
second = '{10000000-0000-4000-8000-000000000002}'
shape_creation = '{20000000-0000-4000-8000-000000000001}'
slide = E.fromstring(parts['ppt/slides/slide1.xml'])
shape = slide.find(f'.//{{{P}}}sp/{{{P}}}nvSpPr/{{{P}}}cNvPr')
shape_id = shape.attrib['id']
sld_id = E.fromstring(parts['ppt/presentation.xml']).find(f'{{{P}}}sldIdLst')[0].attrib['id']

def patch(name, fn):
    parts[name] = fn(parts[name].decode('utf-8-sig')).encode()

def body(text):
    return f'<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p188:txBody>'

def anchor(text=False):
    return (f'<ac:{"tx" if text else "de"}MkLst><pc:docMk/><pc:sldMk cId="123456" sldId="{sld_id}"/>'
            f'<ac:spMk id="{shape_id}" creationId="{shape_creation}"/>'
            + ('<ac:txMk cp="0" len="4"/>' if text else '') + f'</ac:{"tx" if text else "de"}MkLst>')

parts['ppt/authors.xml'] = (f'<p188:authorLst xmlns:p188="{M}">'
    f'<p188:author id="{author}" name="First author" initials="FA" userId="first@example.test" providerId="AD"/>'
    f'<p188:author id="{second}" name="Reply author" initials="RA" userId="reply@example.test" providerId="AD"/>'
    '</p188:authorLst>').encode()
parts['ppt/comments/modern.xml'] = (f'<p188:cmLst xmlns:p188="{M}" xmlns:a="{A}" xmlns:pc="{PC}" xmlns:ac="{AC}">'
    f'<p188:cm id="{{30000000-0000-4000-8000-000000000001}}" authorId="{author}" created="2024-06-01T10:00:00Z"'
    f' status="resolved" title="Check title" assignedTo="{second}" startDate="2024-06-01T10:00:00Z" dueDate="2024-06-02T10:00:00Z" complete="50000">'
    + anchor() + '<p188:pos x="1000" y="2000"/><p188:replyLst>'
    f'<p188:reply id="{{30000000-0000-4000-8000-000000000002}}" authorId="{second}" created="2024-06-01T11:00:00Z">'
    + body('Reply remains attached') + '</p188:reply></p188:replyLst>' + body('Shape comment task') + '</p188:cm>'
    f'<p188:cm id="{{30000000-0000-4000-8000-000000000003}}" authorId="{author}" created="2024-06-01T12:00:00Z">'
    + anchor(True) + body('Text range comment') + '</p188:cm></p188:cmLst>').encode()
patch('ppt/_rels/presentation.xml.rels', lambda s: s.replace('</Relationships>', f'<Relationship Id="rIdModernAuthors" Type="{REL}authors" Target="authors.xml"/></Relationships>'))
patch('ppt/slides/_rels/slide1.xml.rels', lambda s: s.replace('</Relationships>', f'<Relationship Id="rIdModernComments" Type="{REL}comments" Target="../comments/modern.xml"/></Relationships>'))
extension = (f'<p:extLst><p:ext uri="{{6950BFC3-D8DA-4A85-94F7-54DA5524770B}}">'
             f'<p188:commentRel xmlns:p188="{M}" r:id="rIdModernComments"/></p:ext></p:extLst>')
creation = '<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="123456"/></p:ext></p:extLst>'
shape_ext = '<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="' + shape_creation + '"/></a:ext></a:extLst>'
def edit_slide(s):
    pat = r'(<p:cNvPr\b[^>]*\bid="' + re.escape(shape_id) + r'"[^>]*)(/>)'
    s, count = re.subn(pat, lambda m: m[1] + '>' + shape_ext + '</p:cNvPr>', s, count=1)
    assert count == 1, 'Expected an empty authored cNvPr'
    return s.replace('</p:cSld>', creation + '</p:cSld>').replace('</p:sld>', extension + '</p:sld>')
patch('ppt/slides/slide1.xml', edit_slide)
patch('[Content_Types].xml', lambda s: s.replace('</Types>', '<Override PartName="/ppt/authors.xml" ContentType="application/vnd.ms-powerpoint.authors+xml"/><Override PartName="/ppt/comments/modern.xml" ContentType="application/vnd.ms-powerpoint.comments+xml"/></Types>'))
with zipfile.ZipFile(a.output / 'modern-comments.pptx', 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in parts.items():
        z.writestr(name, data)
print(a.output / 'modern-comments.pptx')
