"""SmartArt spike: decks that answer how PowerPoint treats SmartArt written by Lectern, before the writer
is built (docs/lectern.md, Diagrams). Our layout, quick-style and colour definitions are our own, under
urn:vibeoffice.work URNs; nothing is copied from Microsoft's built-in definitions.

    python3 tools/lectern/test/smartart-spike.py BASE.pptx POWERPOINT_SMARTART.pptx OUTDIR

BASE: a one-slide deck saved by Lectern. POWERPOINT_SMARTART: a PowerPoint-authored deck with one SmartArt
(for e/f). Writes:
  a-own-full.pptx        our definitions, full data model (presentation points) and drawing cache
  b-own-no-cache.pptx    as a, without the drawing cache: does PowerPoint lay out from our definition?
  c-own-no-pres.pptx     only document/node/transition points: does PowerPoint rebuild the rest?
  d-own-colorful.pptx    as a, with a multi-colour colour definition: are our colours honoured?
  e-ppt-text-edited.pptx PowerPoint's SmartArt with node text changed in data and drawing
  f-ppt-node-added.pptx  PowerPoint's SmartArt with a node added (no presentation points, no cache)
"""
import os
import re
import sys
import uuid
import zipfile

NS_DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
REL_DRAWING = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'
CT = 'application/vnd.openxmlformats-officedocument.drawingml.'
URN = 'urn:vibeoffice.work/diagram/'
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
EMU = 12700  # per point


def gid(*parts):
    return '{' + str(uuid.uuid5(uuid.NAMESPACE_URL, 'vibeoffice:' + ':'.join(map(str, parts)))).upper() + '}'


def esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


# ---------------------------------------------------------------- our definitions
def layout_process():
    """Basic Process: items in a row, arrows between (our own constraint set)"""
    margins = ''.join(f'<dgm:constr type="{m}" refType="primFontSz" fact="0.25"/>' for m in ('lMarg', 'rMarg', 'tMarg', 'bMarg'))
    return HEAD + f'''<dgm:layoutDef xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}" xmlns:r="{NS_R}" uniqueId="{URN}layout/process">
<dgm:title val="Basic Process"/><dgm:desc val="Steps in a row, joined by arrows."/>
<dgm:catLst><dgm:cat type="process" pri="100"/></dgm:catLst>
<dgm:layoutNode name="diagram">
 <dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>
 <dgm:choose name="direction">
  <dgm:if name="leftToRight" func="var" arg="dir" op="equ" val="norm"><dgm:alg type="lin"/></dgm:if>
  <dgm:else name="rightToLeft"><dgm:alg type="lin"><dgm:param type="linDir" val="fromR"/></dgm:alg></dgm:else>
 </dgm:choose>
 <dgm:shape r:blip=""><dgm:adjLst/></dgm:shape>
 <dgm:presOf/>
 <dgm:constrLst>
  <dgm:constr type="w" for="ch" forName="item" refType="w"/>
  <dgm:constr type="h" for="ch" forName="item" refType="w" refFor="ch" refForName="item" fact="0.6"/>
  <dgm:constr type="w" for="ch" forName="arrow" refType="w" refFor="ch" refForName="item" fact="0.32"/>
  <dgm:constr type="h" for="ch" forName="arrow" op="equ"/>
  <dgm:constr type="primFontSz" for="ch" forName="item" op="equ" val="60"/>
  <dgm:constr type="primFontSz" for="des" forName="arrowText" op="equ"/>
 </dgm:constrLst>
 <dgm:ruleLst/>
 <dgm:forEach name="items" axis="ch" ptType="node">
  <dgm:layoutNode name="item" styleLbl="node1">
   <dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst>
   <dgm:alg type="tx"/>
   <dgm:shape type="roundRect" r:blip=""><dgm:adjLst><dgm:adj idx="1" val="0.12"/></dgm:adjLst></dgm:shape>
   <dgm:presOf axis="desOrSelf" ptType="node"/>
   <dgm:constrLst>{margins}</dgm:constrLst>
   <dgm:ruleLst><dgm:rule type="primFontSz" val="8" fact="NaN" max="NaN"/></dgm:ruleLst>
  </dgm:layoutNode>
  <dgm:forEach name="arrows" axis="followSib" ptType="sibTrans" cnt="1">
   <dgm:layoutNode name="arrow" styleLbl="sibTrans2D1">
    <dgm:alg type="conn"><dgm:param type="begPts" val="auto"/><dgm:param type="endPts" val="auto"/></dgm:alg>
    <dgm:shape type="conn" r:blip=""><dgm:adjLst/></dgm:shape>
    <dgm:constrLst>
     <dgm:constr type="w" val="20"/>
     <dgm:constr type="h" refType="w" fact="0.6"/>
     <dgm:constr type="connDist"/>
     <dgm:constr type="begPad" refType="connDist" fact="0.2"/>
     <dgm:constr type="endPad" refType="connDist" fact="0.2"/>
    </dgm:constrLst>
    <dgm:ruleLst/>
    <dgm:layoutNode name="arrowText" styleLbl="sibTrans2D1">
     <dgm:alg type="tx"><dgm:param type="autoTxRot" val="grav"/></dgm:alg>
     <dgm:shape type="conn" r:blip="" hideGeom="1"><dgm:adjLst/></dgm:shape>
     <dgm:constrLst><dgm:constr type="primFontSz" val="16"/><dgm:constr type="lMarg"/><dgm:constr type="rMarg"/><dgm:constr type="tMarg"/><dgm:constr type="bMarg"/></dgm:constrLst>
     <dgm:ruleLst><dgm:rule type="primFontSz" val="8" fact="NaN" max="NaN"/></dgm:ruleLst>
    </dgm:layoutNode>
   </dgm:layoutNode>
  </dgm:forEach>
 </dgm:forEach>
</dgm:layoutNode>
</dgm:layoutDef>'''


def style_label(name, line, fill, effect, font):
    ref = lambda tag, idx, clr: f'<a:{tag} idx="{idx}"><a:scrgbClr r="0" g="0" b="0"/></a:{tag}>' if clr is None else f'<a:{tag} idx="{idx}">{clr}</a:{tag}>'
    return (f'<dgm:styleLbl name="{name}"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>'
            f'<dgm:sp3d/><dgm:txPr/><dgm:style>{ref("lnRef", line, None)}{ref("fillRef", fill, None)}{ref("effectRef", effect, None)}'
            f'<a:fontRef idx="minor">{font}</a:fontRef></dgm:style></dgm:styleLbl>')


def quick_style():
    """Subtle: theme line 2, theme fill 1, no effect; our own definition"""
    lt1 = '<a:schemeClr val="lt1"/>'
    labels = ''.join(style_label(n, 2, 1, 0, lt1) for n in ('node0', 'node1', 'lnNode1', 'vennNode1', 'alignNode1', 'trAlignAcc1'))
    labels += style_label('sibTrans2D1', 0, 1, 0, lt1) + style_label('sibTrans1D1', 1, 0, 0, '')
    labels += style_label('bgShp', 0, 1, 0, '<a:schemeClr val="tx1"/>')
    return HEAD + (f'<dgm:styleDef xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}" uniqueId="{URN}style/subtle">'
                   f'<dgm:title val="Subtle"/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="100"/></dgm:catLst>'
                   f'<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>{labels}</dgm:styleDef>')


def colors(name, fills):
    """fills: scheme colours cycled across nodes (meth=repeat for one, cycle for several)"""
    clr = lambda vals: ''.join(f'<a:schemeClr val="{v}"/>' for v in vals)
    meth = 'repeat' if len(fills) == 1 else 'cycle'

    def lbl(n, fill, line, text):
        return (f'<dgm:styleLbl name="{n}"><dgm:fillClrLst meth="{meth if n.startswith("node") else "repeat"}">{fill}</dgm:fillClrLst>'
                f'<dgm:linClrLst meth="repeat">{line}</dgm:linClrLst><dgm:effectClrLst/><dgm:txLinClrLst/>'
                f'<dgm:txFillClrLst meth="repeat">{text}</dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl>')
    labels = lbl('node0', clr(fills), clr(['lt1']), clr(['lt1'])) + lbl('node1', clr(fills), clr(['lt1']), clr(['lt1']))
    labels += lbl('sibTrans2D1', f'<a:schemeClr val="{fills[0]}"><a:tint val="60000"/></a:schemeClr>', f'<a:schemeClr val="{fills[0]}"><a:tint val="60000"/></a:schemeClr>', clr(['lt1']))
    labels += lbl('bgShp', f'<a:schemeClr val="{fills[0]}"><a:tint val="40000"/></a:schemeClr>', clr([fills[0]]), clr(['dk1']))
    return HEAD + (f'<dgm:colorsDef xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}" uniqueId="{URN}colors/{name}">'
                   f'<dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="{"accent1" if len(fills) == 1 else "colorful"}" pri="100"/></dgm:catLst>{labels}</dgm:colorsDef>')


# ---------------------------------------------------------------- data model and drawing
def tx(text):
    return f'<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{esc(text)}</a:t></a:r></a:p></dgm:t>' if text else '<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t>'


def data_model(texts, drawing_rid, pres=True, colors_name='accent1'):
    doc = gid('doc')
    n = len(texts)
    nodes = [gid('node', i) for i in range(n)]
    par = [gid('par', i) for i in range(n)]
    sib = [gid('sib', i) for i in range(n)]
    cx = [gid('cxn', i) for i in range(n)]
    pts = [f'<dgm:pt modelId="{doc}" type="doc"><dgm:prSet loTypeId="{URN}layout/process" loCatId="process" qsTypeId="{URN}style/subtle" qsCatId="simple" csTypeId="{URN}colors/{colors_name}" csCatId="{"accent1" if colors_name == "accent1" else "colorful"}" phldr="1"/><dgm:spPr/>{tx("")}</dgm:pt>']
    cxns = []
    for i, t in enumerate(texts):
        pts.append(f'<dgm:pt modelId="{nodes[i]}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/>{tx(t)}</dgm:pt>')
        pts.append(f'<dgm:pt modelId="{par[i]}" type="parTrans" cxnId="{cx[i]}"><dgm:prSet/><dgm:spPr/>{tx("")}</dgm:pt>')
        pts.append(f'<dgm:pt modelId="{sib[i]}" type="sibTrans" cxnId="{cx[i]}"><dgm:prSet/><dgm:spPr/>{tx("")}</dgm:pt>')
        cxns.append(f'<dgm:cxn modelId="{cx[i]}" srcId="{doc}" destId="{nodes[i]}" srcOrd="{i}" destOrd="0" parTransId="{par[i]}" sibTransId="{sib[i]}"/>')
    pres_ids = {}
    if pres:
        p_doc = gid('pres', 'diagram')
        pres_ids['diagram'] = p_doc
        pts.append(f'<dgm:pt modelId="{p_doc}" type="pres"><dgm:prSet presAssocID="{doc}" presName="diagram" presStyleCnt="0"><dgm:presLayoutVars><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:presLayoutVars></dgm:prSet><dgm:spPr/></dgm:pt>')
        cxns.append(f'<dgm:cxn modelId="{gid("cx", "doc")}" type="presOf" srcId="{doc}" destId="{p_doc}" srcOrd="0" destOrd="0" presId="{URN}layout/process"/>')
        order = 0
        for i in range(n):
            p_item = gid('pres', 'item', i)
            pres_ids[('item', i)] = p_item
            pts.append(f'<dgm:pt modelId="{p_item}" type="pres"><dgm:prSet presAssocID="{nodes[i]}" presName="item" presStyleLbl="node1" presStyleIdx="{i}" presStyleCnt="{n}"><dgm:presLayoutVars><dgm:bulletEnabled val="1"/></dgm:presLayoutVars></dgm:prSet><dgm:spPr/></dgm:pt>')
            cxns.append(f'<dgm:cxn modelId="{gid("cx", "item", i)}" type="presOf" srcId="{nodes[i]}" destId="{p_item}" srcOrd="0" destOrd="0" presId="{URN}layout/process"/>')
            cxns.append(f'<dgm:cxn modelId="{gid("cx", "itempar", i)}" type="presParOf" srcId="{p_doc}" destId="{p_item}" srcOrd="{order}" destOrd="0" presId="{URN}layout/process"/>')
            order += 1
            if i < n - 1:
                p_arrow, p_text = gid('pres', 'arrow', i), gid('pres', 'arrowText', i)
                pres_ids[('arrow', i)] = p_arrow
                for pid, name in ((p_arrow, 'arrow'), (p_text, 'arrowText')):
                    pts.append(f'<dgm:pt modelId="{pid}" type="pres"><dgm:prSet presAssocID="{sib[i]}" presName="{name}" presStyleLbl="sibTrans2D1" presStyleIdx="{i}" presStyleCnt="{n - 1}"/><dgm:spPr/></dgm:pt>')
                cxns.append(f'<dgm:cxn modelId="{gid("cx", "arrow", i)}" type="presOf" srcId="{sib[i]}" destId="{p_arrow}" srcOrd="0" destOrd="0" presId="{URN}layout/process"/>')
                cxns.append(f'<dgm:cxn modelId="{gid("cx", "arrowText", i)}" type="presOf" srcId="{sib[i]}" destId="{p_text}" srcOrd="1" destOrd="0" presId="{URN}layout/process"/>')
                cxns.append(f'<dgm:cxn modelId="{gid("cx", "arrowpar", i)}" type="presParOf" srcId="{p_doc}" destId="{p_arrow}" srcOrd="{order}" destOrd="0" presId="{URN}layout/process"/>')
                cxns.append(f'<dgm:cxn modelId="{gid("cx", "arrowTextpar", i)}" type="presParOf" srcId="{p_arrow}" destId="{p_text}" srcOrd="0" destOrd="0" presId="{URN}layout/process"/>')
                order += 1
    ext = (f'<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="{NS_DSP}" relId="{drawing_rid}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst>' if drawing_rid else '')
    xml = HEAD + (f'<dgm:dataModel xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}" xmlns:r="{NS_R}"><dgm:ptLst>{"".join(pts)}</dgm:ptLst>'
                  f'<dgm:cxnLst>{"".join(cxns)}</dgm:cxnLst><dgm:bg/><dgm:whole/>{ext}</dgm:dataModel>')
    return xml, pres_ids


def drawing(texts, pres_ids, w, h, fills):
    """the drawing cache, laid out as our engine will: equal items, arrows a third of an item wide"""
    n = len(texts)
    unit = w / (n + (n - 1) * 0.32)
    iw, ih = unit, min(h, unit * 0.6)
    y = (h - ih) / 2
    sps = []
    style = lambda lnidx, fill: (f'<dsp:style><a:lnRef idx="{lnidx}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>'
                                 f'<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>')
    sz = max(8, min(28, int(iw / EMU / max(len(t) for t in texts) * 1.6)))
    for i, t in enumerate(texts):
        x = i * unit * 1.32
        fill = fills[i % len(fills)]
        sps.append(f'<dsp:sp modelId="{pres_ids[("item", i)]}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>'
                   f'<dsp:spPr><a:xfrm><a:off x="{int(x)}" y="{int(y)}"/><a:ext cx="{int(iw)}" cy="{int(ih)}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 12000"/></a:avLst></a:prstGeom>'
                   f'<a:solidFill><a:schemeClr val="{fill}"/></a:solidFill><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:prstDash val="solid"/></a:ln></dsp:spPr>'
                   f'{style(2, fill)}<dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="68580" tIns="68580" rIns="68580" bIns="68580" numCol="1" spcCol="1270" anchor="ctr" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
                   f'<a:p><a:pPr marL="0" lvl="0" indent="0" algn="ctr" defTabSz="1066800"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPct val="35000"/></a:spcAft></a:pPr>'
                   f'<a:r><a:rPr lang="en-US" sz="{sz * 100}" kern="1200"/><a:t>{esc(t)}</a:t></a:r></a:p></dsp:txBody>'
                   f'<dsp:txXfrm><a:off x="{int(x)}" y="{int(y)}"/><a:ext cx="{int(iw)}" cy="{int(ih)}"/></dsp:txXfrm></dsp:sp>')
        if i < n - 1:
            ax, aw, ah = x + iw + unit * 0.06, unit * 0.2, unit * 0.2 * 1.2
            sps.append(f'<dsp:sp modelId="{pres_ids[("arrow", i)]}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>'
                       f'<dsp:spPr><a:xfrm><a:off x="{int(ax)}" y="{int((h - ah) / 2)}"/><a:ext cx="{int(aw)}" cy="{int(ah)}"/></a:xfrm><a:prstGeom prst="rightArrow"><a:avLst><a:gd name="adj1" fmla="val 60000"/><a:gd name="adj2" fmla="val 50000"/></a:avLst></a:prstGeom>'
                       f'<a:solidFill><a:schemeClr val="{fills[0]}"><a:tint val="60000"/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln></dsp:spPr>{style(0, fills[0])}'
                       f'<dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" spcCol="1270" anchor="ctr" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dsp:txBody>'
                       f'<dsp:txXfrm><a:off x="{int(ax)}" y="{int((h - ah) / 2)}"/><a:ext cx="{int(aw)}" cy="{int(ah)}"/></dsp:txXfrm></dsp:sp>')
    return HEAD + (f'<dsp:drawing xmlns:dgm="{NS_DGM}" xmlns:dsp="{NS_DSP}" xmlns:a="{NS_A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>'
                   f'{"".join(sps)}</dsp:spTree></dsp:drawing>')


# ---------------------------------------------------------------- packaging
def frame(rids, x, y, w, h, shape_id=4):
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="Diagram {shape_id - 1}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm><a:graphic><a:graphicData uri="{NS_DGM}">'
            f'<dgm:relIds xmlns:dgm="{NS_DGM}" r:dm="{rids["dm"]}" r:lo="{rids["lo"]}" r:qs="{rids["qs"]}" r:cs="{rids["cs"]}"/></a:graphicData></a:graphic></p:graphicFrame>')


def build_own(base, out, title, texts, cache=True, pres=True, fills=('accent1',), colors_name='accent1'):
    zin = zipfile.ZipFile(base)
    files = {n: zin.read(n) for n in zin.namelist()}
    slide = files['ppt/slides/slide1.xml'].decode()
    rels = files['ppt/slides/_rels/slide1.xml.rels'].decode()
    used = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    nxt = iter(range(max(used) + 1, max(used) + 20))
    rids = {k: f'rId{next(nxt)}' for k in ('dm', 'lo', 'qs', 'cs', 'dr')}
    w, h, x, y = 7315200, 3200400, 914400, 2286000   # 8 in × 3.5 in at (1 in, 2.5 in)
    data, pres_ids = data_model(texts, rids['dr'] if cache else None, pres=pres, colors_name=colors_name)
    parts = {'ppt/diagrams/data1.xml': (data, 'diagramData+xml', 'diagramData', 'dm'),
             'ppt/diagrams/layout1.xml': (layout_process(), 'diagramLayout+xml', 'diagramLayout', 'lo'),
             'ppt/diagrams/quickStyle1.xml': (quick_style(), 'diagramStyle+xml', 'diagramQuickStyle', 'qs'),
             'ppt/diagrams/colors1.xml': (colors(colors_name, list(fills)), 'diagramColors+xml', 'diagramColors', 'cs')}
    if cache:
        parts['ppt/diagrams/drawing1.xml'] = (drawing(texts, pres_ids, w, h, list(fills)), None, None, 'dr')
    add_rels, add_ct = '', ''
    for name, (xml, ct, rel, key) in parts.items():
        files[name] = xml.encode()
        typ = REL + rel if rel else REL_DRAWING
        add_rels += f'<Relationship Id="{rids[key]}" Type="{typ}" Target="../diagrams/{os.path.basename(name)}"/>'
        add_ct += f'<Override PartName="/{name}" ContentType="{CT + ct if ct else "application/vnd.ms-office.drawingml.diagramDrawing+xml"}"/>'
    files['ppt/slides/_rels/slide1.xml.rels'] = rels.replace('</Relationships>', add_rels + '</Relationships>').encode()
    files['[Content_Types].xml'] = files['[Content_Types].xml'].decode().replace('</Types>', add_ct + '</Types>').encode()
    # the title says which spike this is; the subtitle placeholder (if any) is left empty
    slide = re.sub(r'(<p:ph type="(?:ctrTitle|title)"[^>]*/>.*?<a:p>)(.*?)(</a:p>)', lambda m: m.group(1) + f'<a:r><a:rPr lang="en-US"/><a:t>{esc(title)}</a:t></a:r>' + m.group(3), slide, count=1, flags=re.S)
    slide = re.sub(r'<p:sp>(?:(?!</p:sp>).)*<p:ph type="subTitle".*?</p:sp>', '', slide, count=1, flags=re.S)
    files['ppt/slides/slide1.xml'] = slide.replace('</p:spTree>', frame(rids, x, y, w, h) + '</p:spTree>').encode()
    write(out, files)


def write(out, files):
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
        for n, b in files.items(): z.writestr(n, b)


def edit_powerpoint(src, out, add_node):
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}
    data_name = next(n for n in files if re.match(r'ppt/diagrams/data\d+\.xml$', n))
    data = files[data_name].decode()
    # change the first node's text (in data, and in the drawing cache when it is kept)
    m = re.search(r'<dgm:pt modelId="([^"]+)"(?: type="node")?>(?:(?!</dgm:pt>).)*?<a:t>([^<]+)</a:t>', data, re.S)
    old = m.group(2)
    new = 'Edited in Lectern'
    data = data.replace(f'<a:t>{old}</a:t>', f'<a:t>{new}</a:t>', 1)
    drawing_name = next((n for n in files if re.match(r'ppt/diagrams/drawing\d+\.xml$', n)), None)
    if not add_node:
        if drawing_name: files[drawing_name] = files[drawing_name].decode().replace(f'<a:t>{old}</a:t>', f'<a:t>{new}</a:t>', 1).encode()
    else:
        # a new last child of the document point: node + transition pair + parOf; no presentation points,
        # and no drawing cache, so PowerPoint has to lay the diagram out again
        doc = re.search(r'<dgm:pt modelId="([^"]+)" type="doc"', data).group(1)
        last = max([int(x) for x in re.findall(rf'<dgm:cxn [^>]*srcId="{re.escape(doc)}"[^>]*srcOrd="(\d+)"[^>]*parTransId', data)] or [-1])
        node, par, sib, cx = gid('add', 'node'), gid('add', 'par'), gid('add', 'sib'), gid('add', 'cxn')
        pts = (f'<dgm:pt modelId="{node}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/>{tx("Added in Lectern")}</dgm:pt>'
               f'<dgm:pt modelId="{par}" type="parTrans" cxnId="{cx}"><dgm:prSet/><dgm:spPr/>{tx("")}</dgm:pt>'
               f'<dgm:pt modelId="{sib}" type="sibTrans" cxnId="{cx}"><dgm:prSet/><dgm:spPr/>{tx("")}</dgm:pt>')
        data = data.replace('</dgm:ptLst>', pts + '</dgm:ptLst>', 1)
        data = data.replace('</dgm:cxnLst>', f'<dgm:cxn modelId="{cx}" srcId="{doc}" destId="{node}" srcOrd="{last + 1}" destOrd="0" parTransId="{par}" sibTransId="{sib}"/></dgm:cxnLst>', 1)
        data = re.sub(r'<dgm:extLst>.*?</dgm:extLst>', '', data, flags=re.S)
        if drawing_name:
            del files[drawing_name]
            slide_rels = next(n for n in files if n.startswith('ppt/slides/_rels/'))
            files[slide_rels] = re.sub(r'<Relationship [^>]*diagramDrawing"[^>]*/>', '', files[slide_rels].decode()).encode()
            files['[Content_Types].xml'] = re.sub(rf'<Override PartName="/{re.escape(drawing_name)}"[^>]*/>', '', files['[Content_Types].xml'].decode()).encode()
    files[data_name] = data.encode()
    write(out, files)


def main():
    base, ppt, outdir = sys.argv[1:4]
    os.makedirs(outdir, exist_ok=True)
    steps = ['Plan', 'Build', 'Launch']
    o = lambda n: os.path.join(outdir, n)
    build_own(base, o('a-own-full.pptx'), 'a: our definitions, full', steps)
    build_own(base, o('b-own-no-cache.pptx'), 'b: no drawing cache', steps, cache=False)
    build_own(base, o('c-own-no-pres.pptx'), 'c: no presentation points', steps, cache=False, pres=False)
    build_own(base, o('d-own-colorful.pptx'), 'd: our colourful colours', steps + ['Review'], fills=('accent2', 'accent3', 'accent4', 'accent5'), colors_name='colorful')
    edit_powerpoint(ppt, o('e-ppt-text-edited.pptx'), add_node=False)
    edit_powerpoint(ppt, o('f-ppt-node-added.pptx'), add_node=True)
    print('wrote 6 decks to', outdir)


if __name__ == '__main__':
    main()
