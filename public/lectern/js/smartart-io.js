/* Lectern — SmartArt parts for diagrams (js/diagram.js): the data model written from the items, and our own
 * layout, quick-style and colour definitions. The definitions are written here from our layouts (common/
 * smartart.js) under urn:vibeoffice.work/diagram/… ids; nothing is copied from Office's built-in ones.
 * pptx-write.js adds the drawing (the laid-out shapes, which other applications show) and the slide frame;
 * pptx-read.js reads diagrams back through L.saIO.read.
 */
(function () {
  'use strict';
  const L = window.L;
  const IO = (L.saIO = {});
  const NS_DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
  const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const URN = 'urn:vibeoffice.work/diagram/';
  IO.URN = URN;
  IO.NS = { dgm: NS_DGM, dsp: NS_DSP };
  const X = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /* stable model ids: the same item gets the same GUID on every save */
  function fnv(s, seed) { let h = seed >>> 0; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; } return h.toString(16).padStart(8, '0'); }
  IO.guid = (key) => {
    const x = fnv(key, 2166136261) + fnv(key, 0x9e3779b9) + fnv(key, 0x85ebca6b) + fnv(key, 0xc2b2ae35);
    return `{${x.slice(0, 8)}-${x.slice(8, 12)}-4${x.slice(13, 16)}-${'89ab'[parseInt(x[16], 16) % 4]}${x.slice(17, 20)}-${x.slice(20, 32)}}`.toUpperCase();
  };

  const CAT = { list: 'list', process: 'process', cycle: 'cycle', hierarchy: 'hierarchy', relationship: 'relationship', pyramid: 'pyramid', matrix: 'matrix' };
  const tBody = (text) => `<dgm:t><a:bodyPr/><a:lstStyle/><a:p>${text ? `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${X(text)}</a:t></a:r>` : '<a:endParaRPr lang="en-US" dirty="0"/>'}</a:p></dgm:t>`;

  /** dgm:dataModel: the document point, one point per item, a parent/sibling transition pair per connection */
  IO.dataXML = (sa, key, drawingRelId) => {
    const lay = L.smartart.get(sa.layout) || L.smartart.LAYOUTS[0];
    const doc = IO.guid(key + ':doc');
    const pts = [`<dgm:pt modelId="${doc}" type="doc"><dgm:prSet loTypeId="${URN}layout/${lay.id}" loCatId="${CAT[lay.cat] || 'list'}" qsTypeId="${URN}style/${sa.style || 'simple'}" qsCatId="simple" csTypeId="${URN}colors/${sa.colors || 'accent1'}" csCatId="${sa.colors === 'colorful' ? 'colorful' : 'accent1'}" phldr="1"/><dgm:spPr/>${tBody('')}</dgm:pt>`];
    const cxns = [];
    const walk = (list, parent) => list.forEach((it, i) => {
      const id = IO.guid(key + ':' + it.id), cx = IO.guid(key + ':' + it.id + ':cxn'), par = IO.guid(key + ':' + it.id + ':par'), sib = IO.guid(key + ':' + it.id + ':sib');
      pts.push(`<dgm:pt modelId="${id}"><dgm:prSet phldrT="[Text]"${it.text ? '' : ' phldr="1"'}/><dgm:spPr/>${tBody(it.text)}</dgm:pt>`);
      pts.push(`<dgm:pt modelId="${par}" type="parTrans" cxnId="${cx}"><dgm:prSet/><dgm:spPr/>${tBody('')}</dgm:pt>`);
      pts.push(`<dgm:pt modelId="${sib}" type="sibTrans" cxnId="${cx}"><dgm:prSet/><dgm:spPr/>${tBody('')}</dgm:pt>`);
      cxns.push(`<dgm:cxn modelId="${cx}" srcId="${parent}" destId="${id}" srcOrd="${i}" destOrd="0" parTransId="${par}" sibTransId="${sib}"/>`);
      walk(it.kids || [], id);
    });
    walk(sa.items || [], doc);
    const ext = drawingRelId ? `<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="${NS_DSP}" relId="${drawingRelId}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst>` : '';
    return HEAD + `<dgm:dataModel xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><dgm:ptLst>${pts.join('')}</dgm:ptLst><dgm:cxnLst>${cxns.join('')}</dgm:cxnLst><dgm:bg/><dgm:whole/>${ext}</dgm:dataModel>`;
  };

  /* ---------------------------------------------------------------- layout definitions */
  const margins = (f) => ['lMarg', 'rMarg', 'tMarg', 'bMarg'].map((m) => `<dgm:constr type="${m}" refType="primFontSz" refFor="ch" refForName="node" fact="${f}"/>`).join('');
  const SAMPLE = '<dgm:sampData useDef="1"><dgm:dataModel><dgm:ptLst/><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:sampData><dgm:styleData useDef="1"><dgm:dataModel><dgm:ptLst/><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:styleData><dgm:clrData useDef="1"><dgm:dataModel><dgm:ptLst/><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:clrData>';
  const shape = (type, adj) => `<dgm:shape${type ? ` type="${type}"` : ''} xmlns:r="${NS_R}" r:blip=""><dgm:adjLst>${(adj || []).map(([i, v]) => `<dgm:adj idx="${i}" val="${v}"/>`).join('')}</dgm:adjLst></dgm:shape>`;
  const textNode = (name, styleLbl, geom, adj, extra) => `<dgm:layoutNode name="${name}" styleLbl="${styleLbl}"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/>${shape(geom, adj)}<dgm:presOf axis="desOrSelf" ptType="node"/><dgm:constrLst>${['lMarg', 'rMarg', 'tMarg', 'bMarg'].map((m) => `<dgm:constr type="${m}" refType="primFontSz" fact="0.25"/>`).join('')}</dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst>${extra || ''}</dgm:layoutNode>`;
  const arrowNode = (styleLbl) => `<dgm:forEach name="arrows" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="arrow" styleLbl="${styleLbl}"><dgm:alg type="conn"><dgm:param type="begPts" val="auto"/><dgm:param type="endPts" val="auto"/></dgm:alg>${shape('conn')}<dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="h" refType="w" fact="0.6"/><dgm:constr type="connDist"/><dgm:constr type="begPad" refType="connDist" fact="0.25"/><dgm:constr type="endPad" refType="connDist" fact="0.22"/></dgm:constrLst><dgm:ruleLst/></dgm:layoutNode></dgm:forEach>`;
  const root = (alg, constr, body, vars) => `<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/>${vars || ''}</dgm:varLst>${alg}${shape('')}<dgm:presOf/><dgm:constrLst>${constr}</dgm:constrLst><dgm:ruleLst/>${body}</dgm:layoutNode>`;
  const nodeSize = (w, hf, sp) => `<dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="${hf}"/>${sp != null ? `<dgm:constr type="w" for="ch" forName="space" refType="w" refFor="ch" refForName="node" fact="${sp}"/>` : ''}<dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>`;
  const LAYOUT = {
    blockList: () => root('<dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg>',
      nodeSize(1, 0.6) + '<dgm:constr type="sp" refType="w" refFor="ch" refForName="node" fact="0.1"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'rect')}</dgm:forEach>`),
    verticalBullet: () => root('<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/><dgm:param type="vertAlign" val="mid"/></dgm:alg>',
      '<dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="primFontSz" refFor="ch" refForName="node" fact="0.8"/><dgm:constr type="w" for="ch" forName="childText" refType="w"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/><dgm:constr type="primFontSz" for="ch" forName="childText" refType="primFontSz" refFor="ch" refForName="node" op="equ" fact="0.8"/><dgm:constr type="h" for="ch" forName="spacer" refType="primFontSz" refFor="ch" refForName="node" fact="0.15"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'roundRect', [[1, 0.16667]])}<dgm:choose name="hasKids"><dgm:if name="kids" axis="ch" ptType="node" func="cnt" op="gte" val="1"><dgm:layoutNode name="childText" styleLbl="revTx"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"><dgm:param type="stBulletLvl" val="1"/></dgm:alg>${shape('rect')}<dgm:presOf axis="des" ptType="node"/><dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.2"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.2"/><dgm:constr type="lMarg" refType="w" fact="0.08"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode></dgm:if><dgm:else name="noKids"/></dgm:choose><dgm:forEach name="gap" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="spacer"><dgm:alg type="sp"/>${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:forEach>`),
    process: () => root('<dgm:choose name="direction"><dgm:if name="ltr" func="var" arg="dir" op="equ" val="norm"><dgm:alg type="lin"/></dgm:if><dgm:else name="rtl"><dgm:alg type="lin"><dgm:param type="linDir" val="fromR"/></dgm:alg></dgm:else></dgm:choose>',
      nodeSize(1, 0.6) + '<dgm:constr type="w" for="ch" forName="arrow" refType="w" refFor="ch" refForName="node" fact="0.22"/><dgm:constr type="h" for="ch" forName="arrow" op="equ"/><dgm:constr type="primFontSz" for="des" forName="arrowText" op="equ"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'roundRect', [[1, 0.1]])}${arrowNode('sibTrans2D1')}</dgm:forEach>`),
    chevron: () => root('<dgm:alg type="lin"/>',
      nodeSize(1, 0.4) + '<dgm:constr type="sp" refType="w" refFor="ch" refForName="node" fact="-0.16"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'chevron')}</dgm:forEach>`),
    cycle: () => root('<dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/></dgm:alg>',
      '<dgm:constr type="w" for="ch" forName="node" refType="w" fact="0.24"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node"/><dgm:constr type="w" for="ch" forName="arrow" refType="w" refFor="ch" refForName="node" fact="0.3"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/><dgm:constr type="sp" refType="w" refFor="ch" refForName="node" fact="0.3"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'ellipse')}${arrowNode('sibTrans2D1')}</dgm:forEach>`),
    radial: () => root('<dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/><dgm:param type="ctrShpMap" val="fNode"/></dgm:alg>',
      '<dgm:constr type="w" for="ch" forName="centre" refType="w" fact="0.3"/><dgm:constr type="h" for="ch" forName="centre" refType="w" refFor="ch" refForName="centre"/><dgm:constr type="w" for="ch" forName="node" refType="w" refFor="ch" refForName="centre" fact="0.6"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node"/><dgm:constr type="primFontSz" for="ch" forName="centre" op="equ" val="65"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>',
      `<dgm:forEach name="centres" axis="ch" ptType="node" cnt="1">${textNode('centre', 'node0', 'ellipse')}<dgm:forEach name="kids" axis="ch" ptType="node"><dgm:forEach name="lines" axis="self" ptType="parTrans"><dgm:layoutNode name="line" styleLbl="parChTrans1D2"><dgm:alg type="conn"><dgm:param type="dim" val="1D"/><dgm:param type="begPts" val="auto"/><dgm:param type="endPts" val="auto"/></dgm:alg>${shape('conn')}<dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="begPad"/><dgm:constr type="endPad"/></dgm:constrLst><dgm:ruleLst/></dgm:layoutNode></dgm:forEach>${textNode('node', 'node1', 'ellipse')}</dgm:forEach></dgm:forEach>`),
    orgChart: () => root('<dgm:alg type="hierRoot"/>',
      '<dgm:constr type="w" for="des" forName="node" refType="w" fact="0.24"/><dgm:constr type="h" for="des" forName="node" refType="w" refFor="des" refForName="node" fact="0.55"/><dgm:constr type="primFontSz" for="des" forName="node" op="equ" val="65"/><dgm:constr type="w" for="des" forName="node2" refType="w" refFor="des" refForName="node"/><dgm:constr type="h" for="des" forName="node2" refType="h" refFor="des" refForName="node"/><dgm:constr type="primFontSz" for="des" forName="node2" refType="primFontSz" refFor="des" refForName="node" op="equ"/><dgm:constr type="sibSp" refType="w" refFor="des" refForName="node" fact="0.25"/><dgm:constr type="sp" refType="h" refFor="des" refForName="node" fact="0.7"/>',
      `<dgm:forEach name="tops" axis="ch" ptType="node"><dgm:layoutNode name="branch"><dgm:alg type="hierRoot"/>${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>${textNode('node', 'node1', 'rect')}<dgm:layoutNode name="kids"><dgm:alg type="hierChild"/>${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/><dgm:forEach name="rep" axis="ch" ptType="node"><dgm:forEach name="lines" axis="self" ptType="parTrans"><dgm:layoutNode name="line" styleLbl="parChTrans1D2"><dgm:alg type="conn"><dgm:param type="dim" val="1D"/><dgm:param type="connRout" val="bend"/><dgm:param type="begPts" val="bCtr"/><dgm:param type="endPts" val="tCtr"/></dgm:alg>${shape('conn')}<dgm:presOf axis="self"/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach><dgm:layoutNode name="branch2"><dgm:alg type="hierRoot"/>${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>${textNode('node2', 'node1', 'rect')}<dgm:layoutNode name="kids2"><dgm:alg type="hierChild"/>${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/><dgm:forEach name="more" ref="rep"/></dgm:layoutNode></dgm:layoutNode></dgm:forEach></dgm:layoutNode></dgm:layoutNode></dgm:forEach>`),
    pyramid: () => root('<dgm:alg type="pyra"><dgm:param type="linDir" val="fromT"/><dgm:param type="pyraAcctPos" val="aft"/></dgm:alg>',
      '<dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="h"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'trapezoid')}</dgm:forEach>`),
    venn: () => root('<dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/></dgm:alg>',
      '<dgm:constr type="w" for="ch" forName="node" refType="w" fact="0.5"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node"/><dgm:constr type="sp" refType="w" refFor="ch" refForName="node" fact="-0.4"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'vennNode1', 'ellipse')}</dgm:forEach>`),
    target: () => root('<dgm:alg type="composite"/>',
      '<dgm:constr type="w" for="ch" forName="node" refType="h"/><dgm:constr type="h" for="ch" forName="node" refType="h"/><dgm:constr type="ctrX" for="ch" forName="node" refType="w" fact="0.5"/><dgm:constr type="ctrY" for="ch" forName="node" refType="h" fact="0.5"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'ellipse')}</dgm:forEach>`),
  };
  /* items in a line (any direction), optionally joined by arrows or overlapping */
  const linear = (geom, dir, o) => () => root(`<dgm:alg type="lin"><dgm:param type="linDir" val="${dir}"/></dgm:alg>`,
    nodeSize(1, (o && o.aspect) || 0.6) + ((o && o.sp) != null ? `<dgm:constr type="sp" refType="w" refFor="ch" refForName="node" fact="${o.sp}"/>` : '') + (o && o.arrows ? '<dgm:constr type="w" for="ch" forName="arrow" refType="w" refFor="ch" refForName="node" fact="0.22"/><dgm:constr type="h" for="ch" forName="arrow" op="equ"/>' : ''),
    `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', (o && o.styleLbl) || 'node1', geom, (o && o.adj) || [])}${o && o.arrows ? arrowNode('sibTrans2D1') : ''}</dgm:forEach>`);
  Object.assign(LAYOUT, {
    stackedList: LAYOUT.verticalBullet, vBox: LAYOUT.verticalBullet, linedList: LAYOUT.verticalBullet,
    hBullet: linear('rect', 'fromL', { aspect: 0.3, sp: 0.06 }),
    stepUp: linear('roundRect', 'fromL', { sp: 0.1, adj: [[1, 0.1]] }),
    continuousArrow: linear('roundRect', 'fromL', { aspect: 0.5, sp: 0.12, adj: [[1, 0.16667]] }),
    vProcess: linear('roundRect', 'fromT', { aspect: 0.3, arrows: true, adj: [[1, 0.1]] }),
    timeline: linear('ellipse', 'fromL', { aspect: 1, sp: 1.5 }),
    continuousCycle: LAYOUT.cycle,
    hierarchy: LAYOUT.orgChart, hHierarchy: LAYOUT.orgChart,
    linearVenn: linear('ellipse', 'fromL', { aspect: 1, sp: -0.25, styleLbl: 'vennNode1' }),
    funnel: LAYOUT.pyramid, pyramidList: LAYOUT.pyramid,
    invertedPyramid: () => LAYOUT.pyramid().replace('<dgm:param type="linDir" val="fromT"/>', '<dgm:param type="linDir" val="fromB"/>'),
    matrix: LAYOUT.blockList,
  });
  IO.layoutXML = (id) => {
    const lay = L.smartart.get(id) || L.smartart.LAYOUTS[0];
    return HEAD + `<dgm:layoutDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${URN}layout/${lay.id}"><dgm:title val="${X(lay.name)}"/><dgm:desc val="${X(lay.desc)}"/><dgm:catLst><dgm:cat type="${CAT[lay.cat] || 'list'}" pri="${10000 + L.smartart.LAYOUTS.indexOf(lay)}"/></dgm:catLst>${SAMPLE}${(LAYOUT[lay.id] || LAYOUT.blockList)()}</dgm:layoutDef>`;
  };

  /* ---------------------------------------------------------------- quick styles */
  const LABELS = ['node0', 'node1', 'lnNode1', 'vennNode1', 'alignNode1', 'trAlignAcc1', 'revTx', 'sibTrans2D1', 'parChTrans1D2', 'bgShp'];
  const STYLE = { simple: [2, 0], white: [3, 0], subtle: [2, 1], moderate: [0, 2], intense: [0, 3] };   // [line ref, effect ref]
  IO.styleXML = (id) => {
    const [ln, fx] = STYLE[id] || STYLE.simple;
    const lbl = (n) => {
      const line = n === 'revTx' || n === 'sibTrans2D1' ? 0 : n === 'parChTrans1D2' ? 2 : ln;
      const fill = n === 'revTx' || n === 'parChTrans1D2' ? 0 : 1;
      const font = `<a:fontRef idx="minor"><a:schemeClr val="${n === 'revTx' || n === 'vennNode1' ? 'tx1' : 'lt1'}"/></a:fontRef>`;
      return `<dgm:styleLbl name="${n}"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="${line}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="${fill}"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="${n === 'revTx' ? 0 : fx}"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>${font}</dgm:style></dgm:styleLbl>`;
    };
    const name = (window.L.diagram && L.diagram.STYLES.find((s) => s[0] === id) || [0, 'Simple Fill'])[1];
    return HEAD + `<dgm:styleDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${URN}style/${X(id || 'simple')}"><dgm:title val="${X(name)}"/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10000"/></dgm:catLst><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:styleLbl name="node0"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="${ln}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="${fx}"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl>${LABELS.slice(1).map(lbl).join('')}</dgm:styleDef>`;
  };

  /* ---------------------------------------------------------------- colours */
  /** pal: the theme's usable colours (L.diagram.palette), so PowerPoint colours the diagram as it was drawn here */
  IO.colorsXML = (id, pal) => {
    pal = pal || { slot: (x) => x, colorful: ['accent2', 'accent3', 'accent4', 'accent5', 'accent6'], text: () => 'lt1' };
    const clr = (vals, meth) => `<dgm:fillClrLst meth="${meth || 'repeat'}">${vals.join('')}</dgm:fillClrLst>`;
    const sc = (v, mods) => `<a:schemeClr val="${v}">${mods || ''}</a:schemeClr>`;
    let fills;
    const base = pal.slot(/^accent[1-6]$/.test(id) ? id : 'accent1');
    if (id === 'colorful') fills = clr(pal.colorful.map((v) => sc(v)), 'cycle');
    else if (id === 'range') fills = clr([sc(base), sc(base, '<a:alpha val="40000"/>')], 'span');
    else fills = clr([sc(base)]);
    const lbl = (n) => {
      let f = fills, line = `<dgm:linClrLst meth="repeat">${sc('lt1')}</dgm:linClrLst>`, tx = sc(pal.text(id === 'colorful' ? pal.colorful[0] : base));
      if (n === 'vennNode1') { f = fills.replace(/<\/a:schemeClr>/g, '<a:alpha val="50000"/></a:schemeClr>'); tx = sc('tx1'); }
      if (n === 'sibTrans2D1') { f = clr([sc(base, '<a:alpha val="45000"/>')]); line = `<dgm:linClrLst meth="repeat">${sc(base, '<a:alpha val="45000"/>')}</dgm:linClrLst>`; }
      if (n === 'parChTrans1D2') { f = clr([sc(base)]); line = `<dgm:linClrLst meth="repeat">${sc(base)}</dgm:linClrLst>`; }
      if (n === 'revTx') { f = clr([sc('lt1', '<a:alpha val="0"/>')]); line = `<dgm:linClrLst meth="repeat">${sc('dk1', '<a:alpha val="0"/>')}</dgm:linClrLst>`; tx = sc('tx1'); }
      return `<dgm:styleLbl name="${n}">${f}${line}<dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst meth="repeat">${tx}</dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl>`;
    };
    const name = (window.L.diagram && L.diagram.COLORS.find((c) => c[0] === id) || [0, 'Colored Fill - Accent 1'])[1];
    return HEAD + `<dgm:colorsDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${URN}colors/${X(id || 'accent1')}"><dgm:title val="${X(name)}"/><dgm:desc val=""/><dgm:catLst><dgm:cat type="${id === 'colorful' ? 'colorful' : 'accent1'}" pri="10000"/></dgm:catLst>${LABELS.map(lbl).join('')}</dgm:colorsDef>`;
  };

  /* ---------------------------------------------------------------- reading */
  /** a diagram written by Lectern (our layout id in the data model) back to its sa record; null for others */
  IO.read = (data) => {
    const doc = typeof data === 'string' ? new DOMParser().parseFromString(data, 'application/xml') : data;
    const pts = [...doc.getElementsByTagNameNS(NS_DGM, 'pt')];
    const docPt = pts.find((p) => p.getAttribute('type') === 'doc');
    const pr = docPt && docPt.getElementsByTagNameNS(NS_DGM, 'prSet')[0];
    const lo = pr && pr.getAttribute('loTypeId') || '';
    if (!lo.startsWith(URN + 'layout/')) return null;
    const layout = lo.slice((URN + 'layout/').length);
    if (!L.smartart.get(layout)) return null;
    const suffix = (attr, kind) => { const v = pr.getAttribute(attr) || ''; return v.startsWith(URN + kind + '/') ? v.slice((URN + kind + '/').length) : null; };
    const text = (p) => [...p.getElementsByTagNameNS(NS_A, 'p')].map((para) => [...para.getElementsByTagNameNS(NS_A, 't')].map((t) => t.textContent).join('')).join(' ');
    const byId = new Map(pts.filter((p) => !p.getAttribute('type') || p.getAttribute('type') === 'node').map((p) => [p.getAttribute('modelId'), { id: L.smartart.newId(), text: text(p), kids: [] }]));
    const kids = new Map();
    for (const c of doc.getElementsByTagNameNS(NS_DGM, 'cxn')) {
      if ((c.getAttribute('type') || 'parOf') !== 'parOf') continue;
      const src = c.getAttribute('srcId'), dst = c.getAttribute('destId');
      if (!byId.has(dst)) continue;
      if (!kids.has(src)) kids.set(src, []);
      kids.get(src).push([+c.getAttribute('srcOrd') || 0, dst]);
    }
    const build = (pid) => (kids.get(pid) || []).sort((a, b) => a[0] - b[0]).map(([, id]) => { const it = byId.get(id); it.kids = build(id); return it; });
    return { layout, colors: suffix('csTypeId', 'colors') || 'accent1', style: suffix('qsTypeId', 'style') || 'simple', items: build(docPt.getAttribute('modelId')) };
  };
})();
