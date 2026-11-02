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
  /* rPr: { attrs, inner } for a node's text (bold, italic, colour) */
  const tBody = (text, rPr) => { const a = (rPr && rPr.attrs) || '', inner = (rPr && rPr.inner) || ''; const pr = (tag) => inner ? `<a:${tag} lang="en-US"${a} dirty="0">${inner}</a:${tag}>` : `<a:${tag} lang="en-US"${a} dirty="0"/>`; return `<dgm:t><a:bodyPr/><a:lstStyle/><a:p>${text ? `<a:r>${pr('rPr')}<a:t>${X(text)}</a:t></a:r>` : pr('endParaRPr')}</a:p></dgm:t>`; };

  /** dgm:dataModel: the document point, one point per item, a parent/sibling transition pair per connection.
      fmt(item) → { spPr, rPr } writes the formatting the user gave the item's shape (pptx-write.js) */
  IO.dataXML = (sa, key, drawingRelId, fmt, ctx) => {
    const lay = L.smartart.get(sa.layout) || L.smartart.LAYOUTS[0];
    const doc = IO.guid(key + ':doc');
    /* loTypeId names the layout part's uniqueId exactly: PowerPoint finds the definition by it, and draws its
       default list when they differ */
    const loXML = IO.layoutXML(lay.id, ctx), loId = /uniqueId="([^"]*)"/.exec(loXML)[1];
    const pts = [`<dgm:pt modelId="${doc}" type="doc"><dgm:prSet loTypeId="${loId}" loCatId="${CAT[lay.cat] || 'list'}" qsTypeId="${URN}style/${sa.style || 'simple'}" qsCatId="simple" csTypeId="${URN}colors/${colorsId(sa.colors, L.diagram.styleLang(sa.style))}" csCatId="${colorsCat(sa.colors)}" phldr="1"/><dgm:spPr/>${tBody('')}</dgm:pt>`];
    const cxns = [];
    const walk = (list, parent) => list.forEach((it, i) => {
      const id = IO.guid(key + ':' + it.id), cx = IO.guid(key + ':' + it.id + ':cxn'), par = IO.guid(key + ':' + it.id + ':par'), sib = IO.guid(key + ':' + it.id + ':sib');
      const f = (fmt && it.fmt && it.fmt.node && fmt(it.fmt.node)) || {};
      pts.push(`<dgm:pt modelId="${id}"><dgm:prSet phldrT="[Text]"${it.text ? '' : ' phldr="1"'}/>${f.spPr ? `<dgm:spPr>${f.spPr}</dgm:spPr>` : '<dgm:spPr/>'}${tBody(it.text, f.rPr)}</dgm:pt>`);
      pts.push(`<dgm:pt modelId="${par}" type="parTrans" cxnId="${cx}"><dgm:prSet/><dgm:spPr/>${tBody('')}</dgm:pt>`);
      pts.push(`<dgm:pt modelId="${sib}" type="sibTrans" cxnId="${cx}"><dgm:prSet/><dgm:spPr/>${tBody('')}</dgm:pt>`);
      cxns.push(`<dgm:cxn modelId="${cx}" srcId="${parent}" destId="${id}" srcOrd="${i}" destOrd="0" parTransId="${par}" sibTransId="${sib}"/>`);
      walk(it.kids || [], id);
    });
    walk(sa.items || [], doc);
    /* presentation points: what each item becomes under the layout definition. PowerPoint writes them and
       LibreOffice needs them to lay a diagram out itself (without the drawing) */
    try { presPoints(sa, key, lay, doc, pts, cxns, loXML); } catch (e) { /* the diagram still opens; PowerPoint rebuilds them */ }
    const ext = drawingRelId ? `<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="${NS_DSP}" relId="${drawingRelId}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst>` : '';
    return HEAD + `<dgm:dataModel xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><dgm:ptLst>${pts.join('')}</dgm:ptLst><dgm:cxnLst>${cxns.join('')}</dgm:cxnLst><dgm:bg/><dgm:whole/>${ext}</dgm:dataModel>`;
  };

  /* walk our layout definition over the items, as a SmartArt engine does: every layoutNode met gives a pres
     point tied to its data point (presOf) and to the pres point above it (presParOf) */
  function presPoints(sa, key, lay, doc, pts, cxns, loXML) {
    const def = new DOMParser().parseFromString(loXML, 'application/xml');
    const el = (e) => [...e.children];
    const loc = (e) => e.localName;
    /* the data tree, with each node's transition points */
    const node = (it, i, list) => ({ id: IO.guid(key + ':' + it.id), par: IO.guid(key + ':' + it.id + ':par'), sib: IO.guid(key + ':' + it.id + ':sib'), last: i === list.length - 1, kids: [] });
    const build = (list) => list.map((it, i) => Object.assign(node(it, i, list), { kids: build(it.kids || []) }));
    const root = { id: doc, kids: build(sa.items || []) };
    const des = (n) => n.kids.flatMap((k) => [k, ...des(k)]);
    const out = [], links = [];
    let seq = 0;
    const srcOrd = new Map(), destOrd = new Map(), parOrd = new Map();
    const presOf = (src, dst) => { const so = srcOrd.get(src) || 0, dO = destOrd.get(dst) || 0; srcOrd.set(src, so + 1); destOrd.set(dst, dO + 1); links.push(['presOf', src, dst, so, dO]); };
    const byName = (name) => [...def.getElementsByTagNameNS(NS_DGM, 'forEach')].find((f) => f.getAttribute('name') === name);
    const test = (c, ctx) => {
      const func = c.getAttribute('func'), op = c.getAttribute('op') || 'equ', val = c.getAttribute('val');
      let v;
      if (func === 'var') v = c.getAttribute('arg') === 'dir' ? 'norm' : '';
      else if (func === 'cnt') {
        let set = ctx.n ? [ctx.n] : [];
        for (const ax of (c.getAttribute('axis') || 'ch').split(/\s+/)) set = ax === 'des' ? set.flatMap(des) : set.flatMap((x) => x.kids);
        v = String(set.length);
      }
      else return false;
      const a = isNaN(+v) ? v : +v, b = isNaN(+val) ? val : +val;
      return { equ: a === b, neq: a !== b, gt: a > b, gte: a >= b, lt: a < b, lte: a <= b }[op] || false;
    };
    const kids = (e, ctx, parent) => { for (const c of el(e)) visit(c, ctx, parent); };
    function visit(e, ctx, parent) {
      if (loc(e) === 'layoutNode') {
        const p = { id: IO.guid(key + ':pres:' + seq++), name: e.getAttribute('name'), style: e.getAttribute('styleLbl'), assoc: ctx.id, vars: '' };
        const vl = el(e).find((c) => loc(c) === 'varLst');
        if (vl) p.vars = new XMLSerializer().serializeToString(vl).replace(/^<dgm:varLst[^>]*>|<\/dgm:varLst>$/g, '').replace(/ xmlns(:\w+)?="[^"]*"/g, '');
        out.push(p);
        if (parent) { const o = parOrd.get(parent.id) || 0; parOrd.set(parent.id, o + 1); links.push(['presParOf', parent.id, p.id, o, 0]); }
        const po = el(e).find((c) => loc(c) === 'presOf');
        if (po) {
          const axis = po.getAttribute('axis');
          /* an empty presOf shows nothing, except on the root (as in Office's files) */
          if (axis === 'self' || (!axis && !parent)) presOf(ctx.id, p.id);
          else if (ctx.n && (axis === 'desOrSelf' || axis === 'des')) for (const d of axis === 'des' ? des(ctx.n) : [ctx.n, ...des(ctx.n)]) presOf(d.id, p.id);
        }
        kids(e, ctx, p);
      } else if (loc(e) === 'forEach') {
        const f = e.getAttribute('ref') ? byName(e.getAttribute('ref')) : e;
        if (!f) return;
        const axis = f.getAttribute('axis'), type = f.getAttribute('ptType'), cnt = +f.getAttribute('cnt') || Infinity, st = +f.getAttribute('st') || 1;
        let ctxs = [];
        if (axis === 'ch' && type === 'node' && ctx.n) ctxs = ctx.n.kids.map((k) => ({ id: k.id, n: k })).slice(st > 0 ? st - 1 : 0);
        else if (axis === 'des' && type === 'node' && ctx.n) ctxs = des(ctx.n).map((k) => ({ id: k.id, n: k })).slice(st > 0 ? st - 1 : 0);
        else if (axis === 'ch' && type === 'sibTrans' && ctx.n) ctxs = ctx.n.kids.map((k) => ({ id: k.sib })).slice(st > 0 ? st - 1 : 0);
        else if (axis === 'followSib' && type === 'sibTrans' && ctx.n && !ctx.n.last && ctx.n.sib) ctxs = [{ id: ctx.n.sib }];
        else if (axis === 'self' && type === 'parTrans' && ctx.n && ctx.n.par) ctxs = [{ id: ctx.n.par }];
        else if (axis === 'precedSib' && type === 'parTrans' && ctx.n && ctx.n.par && ctx.n !== root) ctxs = [{ id: ctx.n.par }];
        for (const c of ctxs.slice(0, cnt)) kids(f, c, parent);
      } else if (loc(e) === 'choose') {
        const pick = el(e).find((c) => loc(c) === 'if' && test(c, ctx)) || el(e).find((c) => loc(c) === 'else');
        if (pick) kids(pick, ctx, parent);
      }
    }
    const top = el(def.documentElement).find((c) => loc(c) === 'layoutNode');
    visit(top, { id: doc, n: root }, null);
    /* for the drawing: which pres point shows each item's text (pptx-write.js names the item's shape after it) */
    IO.lastPres = new Map();
    const textOf = new Map(links.filter((l) => l[0] === 'presOf' && l[3] === 0).map((l) => [l[2], l[1]]));
    for (const p of out) { const src = textOf.get(p.id); if (src && src !== doc && !IO.lastPres.has(src) && p.style) IO.lastPres.set(src, p.id); }
    /* each named layout node numbers its pres points (presStyleIdx of presStyleCnt), as PowerPoint does */
    const count = new Map();
    for (const p of out) if (p.style) count.set(p.style, (count.get(p.style) || 0) + 1);
    const idx = new Map();
    for (const p of out) {
      const cnt = p.style ? count.get(p.style) : 0, i = p.style ? idx.get(p.style) || 0 : null;
      if (p.style) idx.set(p.style, i + 1);
      pts.push(`<dgm:pt modelId="${p.id}" type="pres"><dgm:prSet presAssocID="${p.assoc}" presName="${X(p.name)}"${p.style ? ` presStyleLbl="${X(p.style)}" presStyleIdx="${i}"` : ''} presStyleCnt="${cnt}">${p.vars ? `<dgm:presLayoutVars>${p.vars}</dgm:presLayoutVars>` : ''}</dgm:prSet><dgm:spPr/></dgm:pt>`);
    }
    /* the layout part's uniqueId, as loTypeId: PowerPoint 2016 draws nothing for a diagram whose pres points name another */
    const presId = def.documentElement.getAttribute('uniqueId');
    links.forEach(([type, src, dst, so, dO], i) => cxns.push(`<dgm:cxn modelId="${IO.guid(key + ':pcx:' + i)}" type="${type}" srcId="${src}" destId="${dst}" srcOrd="${so}" destOrd="${dO}" presId="${presId}"/>`));
  }

  /* ---------------------------------------------------------------- layout definitions */
  const margins = (f) => ['lMarg', 'rMarg', 'tMarg', 'bMarg'].map((m) => `<dgm:constr type="${m}" refType="primFontSz" refFor="ch" refForName="node" fact="${f}"/>`).join('');
  const SAMPLE = '<dgm:sampData useDef="1"><dgm:dataModel><dgm:ptLst/><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:sampData><dgm:styleData useDef="1"><dgm:dataModel><dgm:ptLst/><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:styleData><dgm:clrData useDef="1"><dgm:dataModel><dgm:ptLst/><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:clrData>';
  const shape = (type, adj) => `<dgm:shape${type ? ` type="${type}"` : ''} xmlns:r="${NS_R}" r:blip=""><dgm:adjLst>${(adj || []).map(([i, v]) => `<dgm:adj idx="${i}" val="${v}"/>`).join('')}</dgm:adjLst></dgm:shape>`;
  /* axis 'self' where the item's sub-items have shapes of their own (trees, radial, the lists' bullet cards);
     'desOrSelf' where they are bullets in the item's own shape */
  const textNode = (name, styleLbl, geom, adj, extra, axis) => `<dgm:layoutNode name="${name}" styleLbl="${styleLbl}"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/>${shape(geom, adj)}<dgm:presOf axis="${axis || 'desOrSelf'}" ptType="node"/><dgm:constrLst>${['lMarg', 'rMarg', 'tMarg', 'bMarg'].map((m) => `<dgm:constr type="${m}" refType="primFontSz" fact="0.25"/>`).join('')}</dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst>${extra || ''}</dgm:layoutNode>`;
  const arrowNode = (styleLbl) => `<dgm:forEach name="arrows" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="arrow" styleLbl="${styleLbl}"><dgm:alg type="conn"><dgm:param type="begPts" val="auto"/><dgm:param type="endPts" val="auto"/></dgm:alg>${shape('conn')}<dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="h" refType="w" fact="0.6"/><dgm:constr type="connDist"/><dgm:constr type="begPad" refType="connDist" fact="0.25"/><dgm:constr type="endPad" refType="connDist" fact="0.22"/></dgm:constrLst><dgm:ruleLst/></dgm:layoutNode></dgm:forEach>`;
  const root = (alg, constr, body, vars) => `<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/>${vars || ''}</dgm:varLst>${alg}${shape('')}<dgm:presOf/><dgm:constrLst>${constr}</dgm:constrLst><dgm:ruleLst/>${body}</dgm:layoutNode>`;
  const nodeSize = (w, hf, sp) => `<dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="${hf}"/>${sp != null ? `<dgm:constr type="w" for="ch" forName="space" refType="w" refFor="ch" refForName="node" fact="${sp}"/>` : ''}<dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/>`;
  const LAYOUT = {
    blockList: () => root('<dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg>',
      nodeSize(1, 0.6) + '<dgm:constr type="sp" refType="w" refFor="ch" refForName="node" fact="0.1"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'rect')}</dgm:forEach>`),
    verticalBullet: () => root('<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/><dgm:param type="vertAlign" val="mid"/></dgm:alg>',
      '<dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="primFontSz" refFor="ch" refForName="node" fact="0.8"/><dgm:constr type="w" for="ch" forName="childText" refType="w"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/><dgm:constr type="primFontSz" for="ch" forName="childText" refType="primFontSz" refFor="ch" refForName="node" op="equ" fact="0.8"/><dgm:constr type="h" for="ch" forName="spacer" refType="primFontSz" refFor="ch" refForName="node" fact="0.15"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node">${textNode('node', 'node1', 'roundRect', [[1, 0.16667]], '', 'self')}<dgm:choose name="hasKids"><dgm:if name="kids" axis="ch" ptType="node" func="cnt" op="gte" val="1"><dgm:layoutNode name="childText" styleLbl="revTx"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"><dgm:param type="stBulletLvl" val="1"/></dgm:alg>${shape('rect')}<dgm:presOf axis="des" ptType="node"/><dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.2"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.2"/><dgm:constr type="lMarg" refType="w" fact="0.08"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode></dgm:if><dgm:else name="noKids"/></dgm:choose><dgm:forEach name="gap" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="spacer"><dgm:alg type="sp"/>${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:forEach>`),
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
      `<dgm:forEach name="centres" axis="ch" ptType="node" cnt="1">${textNode('centre', 'node0', 'ellipse', null, '', 'self')}<dgm:forEach name="kids" axis="ch" ptType="node"><dgm:forEach name="lines" axis="precedSib" ptType="parTrans" st="-1" cnt="1"><dgm:layoutNode name="line" styleLbl="parChTrans1D2"><dgm:alg type="conn"><dgm:param type="dim" val="1D"/><dgm:param type="endSty" val="noArr"/><dgm:param type="begPts" val="auto"/><dgm:param type="endPts" val="auto"/></dgm:alg>${shape('conn')}<dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="begPad"/><dgm:constr type="endPad"/></dgm:constrLst><dgm:ruleLst/></dgm:layoutNode></dgm:forEach>${textNode('node', 'node1', 'ellipse', null, '', 'self')}</dgm:forEach></dgm:forEach>`),
    /* the diagram arranges its top items side by side; each is a root over its sub-items (a branch) */
    /* the diagram arranges its top items side by side; each is a root over its sub-items (a branch): its text,
       then its children in a row, each a branch again (one definition, repeated through ref). Spacing is set once,
       at the top (PowerPoint rejects a definition whose branches carry constraints on their descendants). A
       child's connector is its parent transition: in PowerPoint's traversal the transition just before the child
       (axis precedSib, st -1, cnt 1), drawn from the parent's bottom to the child's top */
    orgChart: () => tree(TREE_OPTS.orgChart),
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
    hBullet: () => root('<dgm:alg type="lin"><dgm:param type="linDir" val="fromL"/><dgm:param type="nodeVertAlign" val="t"/></dgm:alg>',
      '<dgm:constr type="w" for="ch" forName="column" refType="w"/><dgm:constr type="h" for="ch" forName="column" refType="h"/><dgm:constr type="w" for="ch" forName="space" refType="w" refFor="ch" refForName="column" fact="0.06"/><dgm:constr type="primFontSz" for="des" forName="node" op="equ" val="65"/><dgm:constr type="primFontSz" for="des" forName="childText" refType="primFontSz" refFor="des" refForName="node" op="equ" fact="0.7"/>',
      `<dgm:forEach name="nodes" axis="ch" ptType="node"><dgm:layoutNode name="column"><dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/></dgm:alg>${shape('')}<dgm:presOf/><dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="h" fact="0.3"/><dgm:constr type="w" for="ch" forName="childText" refType="w"/><dgm:constr type="h" for="ch" forName="childText" refType="h" fact="0.7"/></dgm:constrLst><dgm:ruleLst/>${textNode('node', 'node1', 'rect', null, '', 'self')}<dgm:layoutNode name="childText" styleLbl="alignAcc1"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"><dgm:param type="stBulletLvl" val="1"/></dgm:alg>${shape('rect')}<dgm:presOf axis="des" ptType="node"/><dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.2"/><dgm:constr type="lMarg" refType="w" fact="0.08"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode></dgm:layoutNode><dgm:forEach name="gap" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="space"><dgm:alg type="sp"/>${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:forEach>`),
    stepUp: linear('roundRect', 'fromL', { sp: 0.1, adj: [[1, 0.1]] }),
    continuousArrow: linear('roundRect', 'fromL', { aspect: 0.5, sp: 0.12, adj: [[1, 0.16667]] }),
    vProcess: linear('roundRect', 'fromT', { aspect: 0.3, arrows: true, adj: [[1, 0.1]] }),
    timeline: linear('ellipse', 'fromL', { aspect: 1, sp: 1.5 }),
    continuousCycle: LAYOUT.cycle,
    hierarchy: () => tree(TREE_OPTS.hierarchy),
    hHierarchy: () => tree(TREE_OPTS.hHierarchy),
    linearVenn: linear('ellipse', 'fromL', { aspect: 1, sp: -0.25, styleLbl: 'vennNode1' }),
    funnel: LAYOUT.pyramid, pyramidList: LAYOUT.pyramid,
    invertedPyramid: () => LAYOUT.pyramid().replace('<dgm:param type="linDir" val="fromT"/>', '<dgm:param type="linDir" val="fromB"/>'),
    matrix: LAYOUT.blockList,
  });
  /* trees: top items (no connector) each over their children; every child is a branch with a connector from
     its parent (its parent transition, which PowerPoint visits as the child's preceding sibling) */
  const TREE_OPTS = {
    orgChart: { rout: 'bend', beg: 'bCtr', end: 'tCtr', dir: 'fromL', child: '' },
    hierarchy: { rout: 'stra', beg: 'bCtr', end: 'tCtr', dir: 'fromL', child: '', geom: 'roundRect', adj: [[1, 0.1]] },
    hHierarchy: { rout: 'bend', beg: 'midR', end: 'midL', dir: 'fromT', align: 'lCtrCh', child: '<dgm:param type="linDir" val="fromT"/><dgm:param type="chAlign" val="l"/>', geom: 'roundRect', adj: [[1, 0.1]] },
  };
  function tree(o) {
    const line = `<dgm:forEach name="line" axis="precedSib" ptType="parTrans" st="-1" cnt="1"><dgm:layoutNode name="conn" styleLbl="parChTrans1D2"><dgm:alg type="conn"><dgm:param type="dim" val="1D"/><dgm:param type="endSty" val="noArr"/><dgm:param type="connRout" val="${o.rout}"/><dgm:param type="begPts" val="${o.beg}"/><dgm:param type="endPts" val="${o.end}"/>${o.rout === 'bend' ? '<dgm:param type="bendPt" val="end"/>' : ''}</dgm:alg>${shape('conn')}<dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="begPad"/><dgm:constr type="endPad"/></dgm:constrLst><dgm:ruleLst/></dgm:layoutNode></dgm:forEach>`;
    const rootAlg = `<dgm:alg type="hierRoot"${o.align ? `><dgm:param type="hierAlign" val="${o.align}"/></dgm:alg>` : '/>'}`;
    const childAlg = o.child ? `<dgm:alg type="hierChild">${o.child}</dgm:alg>` : '<dgm:alg type="hierChild"/>';
    const branch = (b, t, k, inner) => `<dgm:layoutNode name="${b}">${rootAlg}${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>${textNode(t, 'node1', o.geom || 'rect', o.adj || null, '', 'self')}<dgm:layoutNode name="${k}">${childAlg}${shape('')}<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>${inner}</dgm:layoutNode></dgm:layoutNode>`;
    const same = (n) => `<dgm:constr type="w" for="des" forName="${n}" refType="w" refFor="des" refForName="node"/><dgm:constr type="h" for="des" forName="${n}" refType="h" refFor="des" refForName="node"/><dgm:constr type="primFontSz" for="des" forName="${n}" refType="primFontSz" refFor="des" refForName="node" op="equ"/>`;
    const gaps = (k, b) => `<dgm:constr type="sibSp" for="des" forName="${k}" refType="w" refFor="des" refForName="node" fact="0.25"/><dgm:constr type="secSibSp" for="des" forName="${k}" refType="w" refFor="des" refForName="node" fact="0.25"/><dgm:constr type="sp" for="des" forName="${b}" refType="h" refFor="des" refForName="node" fact="0.7"/>`;
    const make = o.parts ? (alg, constr, body) => ({ alg, constr, body }) : root;
    return make(`<dgm:alg type="hierChild"><dgm:param type="linDir" val="${o.dir}"/></dgm:alg>`,
      '<dgm:constr type="w" for="des" forName="node" refType="w" fact="0.24"/><dgm:constr type="h" for="des" forName="node" refType="w" refFor="des" refForName="node" fact="0.55"/><dgm:constr type="primFontSz" for="des" forName="node" op="equ" val="65"/>' + same('node2')
      + '<dgm:constr type="sibSp" refType="w" refFor="des" refForName="node" fact="0.25"/>' + gaps('kids', 'branch') + gaps('kids2', 'branch2'),
      `<dgm:forEach name="tops" axis="ch" ptType="node">${branch('branch', 'node', 'kids', `<dgm:forEach name="rep" axis="ch" ptType="node">${line}${branch('branch2', 'node2', 'kids2', '<dgm:forEach name="more" ref="rep"/>')}</dgm:forEach>`)}</dgm:forEach>`);
  }
  /* ---------------------------------------------------------------- composite definitions
     PowerPoint lays SmartArt out from the definition, never from the stored drawing. For every layout but the trees
     the definition is written from our own engine: for each item count (1 to MAXN) a branch places every shape where
     smartart.js puts it, as a fraction of the diagram, so PowerPoint draws what Lectern draws, also after Add Shape.
     Each shape is bound to its item (forEach st=i cnt=1), its arrow (the item's sibling transition) or its line (the
     satellite's parent transition); its text is a transparent layer over it, placed at our text box. */
  const MAXN = 10;
  /* a straight line as a preset drawn corner to corner of its box: 'line' (top left to bottom right) or, for lines
     that run down to the left, 'lineInv' (PowerPoint and LibreOffice ignore a turned line in a definition) */
  function lineBox(x0, y0, x1, y1) {
    const l = Math.min(x0, x1), t = Math.min(y0, y1), w = Math.abs(x1 - x0), h = Math.abs(y1 - y0);
    return { x: l, y: t, w: Math.max(w, 0.01), h: Math.max(h, 0.01), geom: w < 0.5 || h < 0.5 || (x1 - x0) * (y1 - y0) > 0 ? 'line' : 'lineInv' };
  }
  const TREES = new Set(['orgChart', 'hierarchy', 'hHierarchy']);
  /* PowerPoint rejects a whole definition whose branches give one layout node different algorithm types (none of
     Office's do): the fallback branches stay composite and hold one child that runs the row or the tree */
  const FILL = (n) => `<dgm:constr type="w" for="ch" forName="${n}" refType="w"/><dgm:constr type="h" for="ch" forName="${n}" refType="h"/>`;
  const r4 = (v) => Math.round(v * 10000) / 10000;
  function compositeXML(lay, ctx) {
    const W = ctx.w, H = ctx.h, SA = L.smartart, radial = lay.id === 'radial';
    const kidsOf = (i) => { const it = (ctx.items || [])[i - 1]; return it && it.kids ? it.kids.length : (lay.levels > 1 ? 1 : 0); };
    /* sample items for k: tops t1..tk (with the diagram's own sub-item counts); radial: one centre with k satellites */
    /* the diagram's own text where it has the item, so the branch for its count gets Lectern's text sizes */
    const its = ctx.items || [], textAt = (it) => (it && it.text) || 'Text';
    const sample = (k) => (radial
      ? [{ id: 't1', text: textAt(its[0]), kids: Array.from({ length: k }, (_, j) => ({ id: `s${j + 1}`, text: textAt(((its[0] || {}).kids || [])[j]), kids: [] })) }]
      : Array.from({ length: k }, (_, i) => ({ id: `t${i + 1}`, text: textAt(its[i]), kids: Array.from({ length: lay.levels > 1 ? kidsOf(i + 1) : 0 }, (_, j) => ({ id: `t${i + 1}k${j + 1}`, text: textAt(((its[i] || {}).kids || [])[j]), kids: [] })) })));
    const measure = ctx.measure || ((t, sz) => t.length * sz * 0.5);
    const fontOf = (n, s) => (s.fontSize ? `<dgm:constr type="primFontSz" for="ch" forName="${n}" op="equ" val="${r4(s.fontSize)}"/>` : '');
    const inside = lay.id === 'blockList' ? 'desOrSelf' : 'self';
    const pos = (n, x, y, w, h) => `<dgm:constr type="l" for="ch" forName="${n}" refType="w" fact="${r4(x / W)}"/><dgm:constr type="t" for="ch" forName="${n}" refType="h" fact="${r4(y / H)}"/><dgm:constr type="w" for="ch" forName="${n}" refType="w" fact="${r4(Math.max(w, 0.01) / W)}"/><dgm:constr type="h" for="ch" forName="${n}" refType="h" fact="${r4(Math.max(h, 0.01) / H)}"/>`;
    const adjOf = (s) => (s.adj || []).map((v, i) => `<dgm:adj idx="${i + 1}" val="${r4(v / 100000)}"/>`).join('');
    const sp = (n, geom, s, lbl, rot) => `<dgm:layoutNode name="${n}" styleLbl="${lbl}"><dgm:alg type="sp"/><dgm:shape type="${geom}"${rot ? ` rot="${r4(rot)}"` : ''} xmlns:r="${NS_R}" r:blip=""><dgm:adjLst>${s ? adjOf(s) : ''}</dgm:adjLst></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>`;
    const ALIGN = { l: ['l', 'r'], r: ['r', 'l'] };
    const tx = (n, lbl, axis, bullet, anchor, align) => `<dgm:layoutNode name="${n}" styleLbl="${lbl}"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"><dgm:param type="txAnchorVert" val="${anchor === 't' ? 't' : anchor === 'b' ? 'b' : 'mid'}"/>${ALIGN[align] ? `<dgm:param type="parTxLTRAlign" val="${ALIGN[align][0]}"/><dgm:param type="parTxRTLAlign" val="${ALIGN[align][1]}"/>` : ''}${bullet ? `<dgm:param type="stBulletLvl" val="${bullet}"/>` : ''}</dgm:alg><dgm:shape type="rect" xmlns:r="${NS_R}" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="${axis}" ptType="node"/><dgm:constrLst><dgm:constr type="lMarg" refType="primFontSz" fact="0.25"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.25"/><dgm:constr type="tMarg" refType="primFontSz" fact="0.2"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.2"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode>`;
    const bindItem = (i, inner) => (radial
      ? (i === 0 ? `<dgm:forEach name="@" axis="ch" ptType="node" st="1" cnt="1">${inner}</dgm:forEach>` : `<dgm:forEach name="@" axis="ch" ptType="node" st="1" cnt="1"><dgm:forEach name="@" axis="ch" ptType="node" st="${i}" cnt="1">${inner}</dgm:forEach></dgm:forEach>`)
      : `<dgm:forEach name="@" axis="ch" ptType="node" st="${i}" cnt="1">${inner}</dgm:forEach>`);
    const index = (id) => { const m = /^(t|s)(\d+)$/.exec(id || ''); return m ? (m[1] === 's' ? +m[2] : radial ? 0 : +m[2]) : null; };
    let branches = '', named = 0;
    /* one numbering for the whole definition (n000, n001, …): with names numbered per branch (k4a6: 4 items, arrow,
       sixth shape) PowerPoint ignored a few nodes' constraints and drew them across the whole diagram (round 7) */
    const nm = () => 'n' + String(named++).padStart(3, '0');
    for (let k = 1; k <= MAXN; k++) {
      const { shapes } = SA.layout(lay.id, sample(k), { x: 0, y: 0, w: W, h: H }, { measure });
      let cons = '', nodes = '';
      for (const s of shapes) {
        const i = index(s.itemId);
        if (s.role === 'arrow' && !s.big) {
          const n = nm('a');
          cons += pos(n, s.x, s.y, s.w, s.h);
          nodes += bindItem(i || 1, sp(n, s.geom, s, 'sibTrans2D1', s.rot));   // with the item it follows: shapes tied to nothing are not drawn everywhere
          continue;
        }
        if (s.role === 'conn' && s.path && s.path.length === 2) {
          const [[x0, y0], [x1, y1]] = s.path, lb = lineBox(x0, y0, x1, y1), n = nm('l');
          cons += pos(n, lb.x, lb.y, lb.w, lb.h);
          const line = sp(n, lb.geom, null, 'parChTrans1D2');
          nodes += i != null ? bindItem(i, `<dgm:forEach name="@" axis="precedSib" ptType="parTrans" st="-1" cnt="1">${line.replace('<dgm:presOf/>', '<dgm:presOf axis="self"/>')}</dgm:forEach>`) : line;
          continue;
        }
        if (s.role === 'conn') { const n = nm('l'); cons += pos(n, s.x, s.y + s.h / 2, s.w, 0.01); nodes += bindItem(i || 1, sp(n, 'line', null, 'parChTrans1D2')); continue; }
        if (s.role === 'ring') { const n = nm('r'); cons += pos(n, s.x, s.y, s.w, s.h); nodes += bindItem(radial ? 0 : 1, sp(n, 'ellipse', null, 'parChTrans1D2')); continue; }
        if (s.role === 'bg' || (s.role === 'arrow' && s.big)) { const n = nm('b'); cons += pos(n, s.x, s.y, s.w, s.h); nodes += bindItem(radial ? 0 : 1, sp(n, s.geom, s, s.role === 'bg' ? 'bgShp' : 'sibTrans2D1', s.rot)); continue; }
        if (s.role === 'marker') { const n = nm('m'); cons += pos(n, s.x, s.y, s.w, s.h); nodes += bindItem(i, sp(n, 'ellipse', null, 'node1')); continue; }
        if (i == null) continue;
        /* an item's shape and its text: the shape below (upside-down outlines as a turned preset), the text on top */
        let inner = '';
        const child = s.role === 'child';
        if (!s.noFill && !(child && !s.tint)) {
          const n = nm('s');
          let geom = s.geom, rot = s.rot || 0, src = s;
          if (geom === 'poly') {
            /* upside-down layers: a trapezoid (or the apex triangle) turned over; its inset from the outline */
            const ys = s.poly.map((q) => q[1]), apex = s.poly.length === 3;
            const bot = s.poly.filter((q) => q[1] === Math.max(...ys)), wBot = Math.max(...bot.map((q) => q[0])) - Math.min(...bot.map((q) => q[0]));
            geom = apex ? 'triangle' : 'trapezoid'; rot = 180;
            src = { adj: [apex ? 50000 : Math.round(((s.w - wBot) / 2) / Math.min(s.w, s.h) * 100000)] };
          }
          cons += pos(n, s.x, s.y, s.w, s.h);
          inner += sp(n, geom, src, child ? 'alignAcc1' : s.styleLbl === 'vennNode1' ? 'vennNode1' : 'node1', rot);
        }
        if (s.text) {
          const n = nm('t'), b = s.text.box;
          cons += pos(n, b.x, b.y, b.w, b.h) + fontOf(n, s);
          const lbl = child || s.noFill || s.styleLbl === 'vennNode1' ? 'revTx' : 'trAlignAcc1';
          /* sub-items are bulleted, the item itself is not */
          inner += tx(n, lbl, child ? 'des' : inside, child ? 1 : inside === 'desOrSelf' ? 2 : 0, s.text.anchor, s.text.align);
        }
        nodes += bindItem(i, inner);
      }
      const countAxis = radial ? 'axis="ch ch" ptType="node node"' : 'axis="ch" ptType="node"';
      branches += `<dgm:if name="@" func="cnt" ${countAxis} op="equ" val="${k}"><dgm:alg type="composite"><dgm:param type="ar" val="${r4(W / H)}"/></dgm:alg><dgm:constrLst>${cons}</dgm:constrLst>${nodes}</dgm:if>`;
    }
    /* more items than the branches cover: a simple row */
    const rest = `<dgm:else name="@"><dgm:alg type="composite"/><dgm:constrLst>${FILL('rest')}</dgm:constrLst><dgm:layoutNode name="rest"><dgm:alg type="lin"/>${shape('')}<dgm:presOf/><dgm:constrLst><dgm:constr type="w" for="ch" forName="more" refType="w"/><dgm:constr type="h" for="ch" forName="more" refType="w" refFor="ch" refForName="more" fact="0.6"/><dgm:constr type="sibSp" refType="w" refFor="ch" refForName="more" fact="0.15"/><dgm:constr type="primFontSz" for="ch" forName="more" op="equ" val="65"/></dgm:constrLst><dgm:ruleLst/><dgm:forEach name="@" axis="ch" ptType="node">${textNode('more', 'node1', 'rect', null, '', inside)}</dgm:forEach></dgm:layoutNode></dgm:else>`;
    let seq = 0;
    return `<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst><dgm:shape xmlns:r="${NS_R}" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:ruleLst/><dgm:choose name="byCount">${branches}${rest}</dgm:choose></dgm:layoutNode>`.replace(/name="@"/g, () => `name="f${seq++}"`);
  }
  /* a tree as it was saved, exactly (each item bound by its path of child positions), its lines as line
     shapes; a tree of any other size (an item added in PowerPoint) falls back to the hierarchy algorithms */
  function treeCompositeXML(lay, ctx) {
    const W = ctx.w, H = ctx.h, SA = L.smartart;
    const order = SA.outline(ctx.items || []).map((o) => o.item.id);
    const N = order.length;
    /* each item's path of child positions, from the top: nested forEach axis ch st=… */
    const path = new Map();
    const walk = (list, pre) => list.forEach((it, i) => { path.set(it.id, pre.concat(i + 1)); walk(it.kids || [], pre.concat(i + 1)); });
    walk(ctx.items || [], []);
    const { shapes } = SA.layout(lay.id, ctx.items || [], { x: 0, y: 0, w: W, h: H }, { measure: ctx.measure || ((t, sz) => t.length * sz * 0.5) });
    const pos = (n, x, y, w, h) => `<dgm:constr type="l" for="ch" forName="${n}" refType="w" fact="${r4(x / W)}"/><dgm:constr type="t" for="ch" forName="${n}" refType="h" fact="${r4(y / H)}"/><dgm:constr type="w" for="ch" forName="${n}" refType="w" fact="${r4(Math.max(w, 0.01) / W)}"/><dgm:constr type="h" for="ch" forName="${n}" refType="h" fact="${r4(Math.max(h, 0.01) / H)}"/>`;
    const bind = (j, inner) => path.get(order[j - 1]).reduceRight((acc, st) => `<dgm:forEach name="@" axis="ch" ptType="node" st="${st}" cnt="1">${acc}</dgm:forEach>`, inner);
    let cons = '', nodes = '', c = 0;
    for (const s of shapes) {
      const j = order.indexOf(s.itemId) + 1;
      if (!j) continue;
      if (s.role === 'conn') {
        for (let q = 1; q < s.path.length; q++) {
          const [x0, y0] = s.path[q - 1], [x1, y1] = s.path[q], n = `t${++c}`;
          if (Math.hypot(x1 - x0, y1 - y0) < 0.5) continue;
          const lb = lineBox(x0, y0, x1, y1);
          cons += pos(n, lb.x, lb.y, lb.w, lb.h);
          nodes += bind(j, `<dgm:layoutNode name="${n}" styleLbl="parChTrans1D2"><dgm:alg type="sp"/><dgm:shape type="${lb.geom}" xmlns:r="${NS_R}" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>`);
        }
        continue;
      }
      const n = `t${++c}`, adj = (s.adj || []).map((v, i) => `<dgm:adj idx="${i + 1}" val="${r4(v / 100000)}"/>`).join('');
      cons += pos(n, s.x, s.y, s.w, s.h) + (s.fontSize ? `<dgm:constr type="primFontSz" for="ch" forName="${n}" op="equ" val="${r4(s.fontSize)}"/>` : '');
      nodes += bind(j, `<dgm:layoutNode name="${n}" styleLbl="node1"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="${s.geom}" xmlns:r="${NS_R}" r:blip=""><dgm:adjLst>${adj}</dgm:adjLst></dgm:shape><dgm:presOf axis="self" ptType="node"/><dgm:constrLst><dgm:constr type="lMarg" refType="primFontSz" fact="0.25"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.25"/><dgm:constr type="tMarg" refType="primFontSz" fact="0.2"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.2"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode>`);
    }
    const fb = tree(Object.assign({ parts: true }, TREE_OPTS[lay.id]));
    let seq = 0;
    return `<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst><dgm:shape xmlns:r="${NS_R}" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:ruleLst/><dgm:choose name="asSaved"><dgm:if name="@" func="cnt" axis="des" ptType="node" op="equ" val="${N}"><dgm:alg type="composite"><dgm:param type="ar" val="${r4(W / H)}"/></dgm:alg><dgm:constrLst>${cons}</dgm:constrLst>${nodes}</dgm:if><dgm:else name="@"><dgm:alg type="composite"/><dgm:constrLst>${FILL('fallback')}</dgm:constrLst><dgm:layoutNode name="fallback">${fb.alg}${shape('')}<dgm:presOf/><dgm:constrLst>${fb.constr}</dgm:constrLst><dgm:ruleLst/>${fb.body}</dgm:layoutNode></dgm:else></dgm:choose></dgm:layoutNode>`.replace(/name="@"/g, () => `name="g${seq++}"`);
  }
  /** the layout definition; ctx { w, h, items } (the diagram's box and items) gives our composite definition */
  IO.layoutXML = (id, ctx) => {
    const lay = L.smartart.get(id) || L.smartart.LAYOUTS[0];
    const body = !ctx ? (LAYOUT[lay.id] || LAYOUT.blockList)() : TREES.has(lay.id) ? treeCompositeXML(lay, ctx) : compositeXML(lay, ctx);
    /* a composite definition depends on the diagram's shape: its own id, so two diagrams never share one */
    const uid = ctx ? `${URN}layout/${lay.id}#${fnv(body, 2166136261)}` : `${URN}layout/${lay.id}`;
    return HEAD + `<dgm:layoutDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${uid}"><dgm:title val="${X(lay.name)}"/><dgm:desc val="${X(lay.desc)}"/><dgm:catLst><dgm:cat type="${CAT[lay.cat] || 'list'}" pri="${10000 + L.smartart.LAYOUTS.indexOf(lay)}"/></dgm:catLst>${SAMPLE}${body}</dgm:layoutDef>`;
  };

  /* ---------------------------------------------------------------- quick styles */
  const LABELS = ['node0', 'node1', 'lnNode1', 'vennNode1', 'alignNode1', 'alignAcc1', 'trAlignAcc1', 'revTx', 'sibTrans2D1', 'parChTrans1D2', 'bgShp'];
  /** the style's [line, fill, effect] references (L.diagram.STYLES) on each label: shapes take all three; a Venn
      circle no shadow (it shows through), arrows and text boxes a plain fill, lines a 1.5 pt line */
  IO.styleXML = (id) => {
    const [ln, fl, fx] = L.diagram.styleRefs(id), lang = L.diagram.styleLang(id);
    const lbl = (n) => {
      const shape = ['node0', 'node1', 'lnNode1', 'vennNode1'].includes(n), text = n === 'revTx' || n === 'trAlignAcc1';
      const line = text || n === 'sibTrans2D1' ? 0 : n === 'parChTrans1D2' ? 2 : ln;
      const fill = text || n === 'parChTrans1D2' ? 0 : shape ? fl : 1;
      const effect = shape && n !== 'vennNode1' ? fx : 0;
      /* outlined and tinted shapes and the text over them: dark text (the colours give each item's own;
         LibreOffice takes this one) */
      const font = `<a:fontRef idx="minor"><a:schemeClr val="${['revTx', 'vennNode1', 'alignAcc1', 'bgShp'].includes(n) || ((shape || n === 'trAlignAcc1') && lang !== 'fill') ? 'tx1' : 'lt1'}"/></a:fontRef>`;
      return `<dgm:styleLbl name="${n}"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="${line}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="${fill}"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="${effect}"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>${font}</dgm:style></dgm:styleLbl>`;
    };
    const name = (L.diagram.STYLES.find((s) => s[0] === id) || L.diagram.STYLES[0])[1];
    return HEAD + `<dgm:styleDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${URN}style/${X(id || 'simple')}"><dgm:title val="${X(name)}"/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10000"/></dgm:catLst><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>${LABELS.map(lbl).join('')}</dgm:styleDef>`;
  };

  /* ---------------------------------------------------------------- colours */
  /** the colours id as saved (an older 'range' is accent 1's), and its category as Office files them */
  const colorsKey = (id) => (id === 'range' ? 'accent1-range' : id || 'accent1');
  /** the saved id of the colours with the style's design language */
  const colorsId = (id, lang) => colorsKey(id) + (lang && lang !== 'fill' ? '/' + lang : '');
  const colorsCat = (id) => { const [ref, kind] = L.diagram.colorsOf(id); return ref === 'colorful' || /^accent/.test(kind) || L.diagram.PALETTES[ref] ? 'colorful' : ref === 'dark2' ? 'mainScheme' : ref; };
  /** pal: the theme's usable colours (L.diagram.palette); lang: the style's design language (L.diagram.styleLang):
      outlined and tinted shapes have colours of their own, saved under their own id ('accent2-accent3/outline') */
  IO.colorsXML = (id, pal, lang) => {
    id = colorsKey(id);
    lang = lang || 'fill';
    /* a list of colours; a range ('span', the default) is written as Office writes it, without meth */
    const clr = (vals, meth, tag) => `<dgm:${tag || 'fillClrLst'}${meth === 'span' ? '' : ` meth="${meth || 'repeat'}"`}>${vals.join('')}</dgm:${tag || 'fillClrLst'}>`;
    /* a theme slot, or one of our palettes' own colours */
    const sc = (v, mods) => (v[0] === '#' ? `<a:srgbClr val="${v.slice(1).toUpperCase()}">${mods || ''}</a:srgbClr>` : `<a:schemeClr val="${v}">${mods || ''}</a:schemeClr>`);
    const mods = (list) => list.map(([n, v]) => `<a:${n} val="${v}"/>`).join('');
    /* the items' colours, how they spread over the items, and their text colours (L.diagram.colorSpec) */
    const { cols, meth, base, texts } = L.diagram.colorSpec(id, pal), [, kind] = L.diagram.colorsOf(id);
    /* extra: more colour changes, a string or (colour, changes) → string */
    const list = (extra, tag, only) => clr((only || cols).map(([v, m]) => sc(v, mods(m) + (typeof extra === 'function' ? extra(v, m) : extra || ''))), only ? 'repeat' : meth, tag);
    const fills = list();
    const textList = clr(texts.map(([v, m]) => sc(v, mods(m))), meth === 'cycle' ? 'cycle' : 'repeat', 'txFillClrLst');
    /* outlined and tinted: the colour on the edge and in the text (a range of shades writes in its deep end), the
       text deeper or lighter where it would be faint on the slide */
    const ink = (tag) => list((v, m) => mods(L.diagram.inkMods(pal, L.diagram.hexOf(pal, [v, m]))), tag, kind === 'range' ? [[base, []]] : null);
    const lbl = (n) => {
      let f = fills, line = `<dgm:linClrLst meth="repeat">${sc('lt1')}</dgm:linClrLst>`, tx = '', txl = textList;
      if (n === 'vennNode1') { f = list('<a:alpha val="50000"/>'); tx = sc('tx1'); txl = ''; }
      if (lang !== 'fill' && ['node0', 'node1', 'lnNode1', 'vennNode1'].includes(n)) {
        const venn = n === 'vennNode1';
        f = lang === 'outline' ? clr([sc('lt1', venn ? '<a:alpha val="50000"/>' : '')]) : list(mods(L.diagram.washMods(pal)) + (venn ? '<a:alpha val="70000"/>' : ''));
        line = list('', 'linClrLst');
        if (!venn) txl = ink('txFillClrLst');
      }
      if (n === 'sibTrans2D1') { f = clr([sc(base, '<a:alpha val="60000"/>')]); line = `<dgm:linClrLst meth="repeat">${sc(base, '<a:alpha val="60000"/>')}</dgm:linClrLst>`; }
      if (n === 'parChTrans1D2') { f = clr([sc(base)]); line = `<dgm:linClrLst meth="repeat">${sc(base)}</dgm:linClrLst>`; }
      /* the sub-item cards and the pyramid behind cards: a pale tint of the colour, dark text (as diagram.js paints them) */
      if (n === 'alignAcc1' || n === 'bgShp') { f = clr([sc(base, `<a:alpha val="${n === 'bgShp' ? 25000 : 18000}"/>`)]); line = `<dgm:linClrLst meth="repeat">${sc(base)}</dgm:linClrLst>`; tx = sc('tx1'); txl = ''; }
      /* the text over each shape: its colour, or each item's own on outlined and tinted shapes */
      if (n === 'trAlignAcc1') { f = clr([sc('lt1', '<a:alpha val="0"/>')]); line = `<dgm:linClrLst meth="repeat">${sc('lt1', '<a:alpha val="0"/>')}</dgm:linClrLst>`; if (lang !== 'fill') txl = ink('txFillClrLst'); }
      if (n === 'revTx') { f = clr([sc('lt1', '<a:alpha val="0"/>')]); line = `<dgm:linClrLst meth="repeat">${sc('dk1', '<a:alpha val="0"/>')}</dgm:linClrLst>`; tx = sc('tx1'); txl = ''; }
      return `<dgm:styleLbl name="${n}">${f}${line}<dgm:effectClrLst/><dgm:txLinClrLst/>${txl || `<dgm:txFillClrLst meth="repeat">${tx}</dgm:txFillClrLst>`}<dgm:txEffectClrLst/></dgm:styleLbl>`;
    };
    const name = L.diagram.colorsName(id) + (lang === 'fill' ? '' : lang === 'outline' ? ', Outlined' : ', Tinted');
    return HEAD + `<dgm:colorsDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${URN}colors/${X(colorsId(id, lang))}"><dgm:title val="${X(name)}"/><dgm:desc val=""/><dgm:catLst><dgm:cat type="${colorsCat(id)}" pri="10000"/></dgm:catLst>${LABELS.map(lbl).join('')}</dgm:colorsDef>`;
  };

  /* ---------------------------------------------------------------- reading */
  /* Office's layouts that are ours under the same name (by the id in the file; the definitions are not used) */
  const OFFICE = { default: 'blockList', orgChart1: 'orgChart', hierarchy1: 'hierarchy', hierarchy2: 'hHierarchy', vList2: 'verticalBullet', hList1: 'hBullet',
    process1: 'process', chevron1: 'chevron', process2: 'vProcess', hProcess11: 'timeline', cycle2: 'cycle', cycle3: 'continuousCycle', radial1: 'radial',
    venn1: 'venn', venn3: 'linearVenn', target1: 'target', funnel1: 'funnel', matrix1: 'matrix', pyramid1: 'pyramid', pyramid2: 'pyramidList', pyramid3: 'invertedPyramid' };
  const OFFICE_URN = 'urn:microsoft.com/office/officeart/2005/8/';
  /** a diagram's sa record from its data model: Lectern's own, or an Office layout that is ours by another
      id (sa.office: shown as PowerPoint drew it until edited here); null for any other layout */
  IO.read = (data) => {
    const doc = typeof data === 'string' ? new DOMParser().parseFromString(data, 'application/xml') : data;
    const pts = [...doc.getElementsByTagNameNS(NS_DGM, 'pt')];
    const docPt = pts.find((p) => p.getAttribute('type') === 'doc');
    const pr = docPt && docPt.getElementsByTagNameNS(NS_DGM, 'prSet')[0];
    const lo = pr && pr.getAttribute('loTypeId') || '';
    const office = lo.startsWith(OFFICE_URN + 'layout/');
    const layout = lo.startsWith(URN + 'layout/') ? lo.slice((URN + 'layout/').length).replace(/#.*$/, '') : office ? OFFICE[lo.slice((OFFICE_URN + 'layout/').length).replace(/#\d+$/, '')] : null;
    if (!layout || !L.smartart.get(layout)) return null;
    const suffix = (attr, kind) => {
      const v = pr.getAttribute(attr) || '';
      if (v.startsWith(URN + kind + '/')) { const k = v.slice((URN + kind + '/').length); return kind === 'colors' ? colorsKey(k.replace(/\/.*$/, '')) : k; }
      const o = v.startsWith(OFFICE_URN + kind + '/') ? v.slice((OFFICE_URN + kind + '/').length) : '';
      if (kind === 'colors') return /colorful/i.test(o) ? 'colorful' : /^accent([1-6])/.test(o) ? 'accent' + o[6] : null;
      if (kind === 'quickstyle') return /intense|3d|polished|cartoon/i.test(o) ? 'intense' : /moderate|simple[45]/i.test(o) ? 'moderate' : /simple[23]/i.test(o) ? 'subtle' : null;
      return null;
    };
    const text = (p) => [...p.getElementsByTagNameNS(NS_A, 'p')].map((para) => [...para.getElementsByTagNameNS(NS_A, 't')].map((t) => t.textContent).join('')).join(' ');
    const nodePts = pts.filter((p) => !p.getAttribute('type') || p.getAttribute('type') === 'node' || p.getAttribute('type') === 'asst');
    const byId = new Map(nodePts.map((p) => [p.getAttribute('modelId'), { id: L.smartart.newId(), text: text(p), kids: [] }]));
    /* the points, for the reader to take each item's own formatting from (pptx-read.js) */
    const points = nodePts.map((p) => [p, byId.get(p.getAttribute('modelId'))]);
    const kids = new Map();
    for (const c of doc.getElementsByTagNameNS(NS_DGM, 'cxn')) {
      if ((c.getAttribute('type') || 'parOf') !== 'parOf') continue;
      const src = c.getAttribute('srcId'), dst = c.getAttribute('destId');
      if (!byId.has(dst)) continue;
      if (!kids.has(src)) kids.set(src, []);
      kids.get(src).push([+c.getAttribute('srcOrd') || 0, dst]);
    }
    const build = (pid) => (kids.get(pid) || []).sort((a, b) => a[0] - b[0]).map(([, id]) => { const it = byId.get(id); it.kids = build(id); return it; });
    const sa = { layout, colors: suffix('csTypeId', 'colors') || 'accent1', style: suffix('qsTypeId', office ? 'quickstyle' : 'style') || 'simple', items: build(docPt.getAttribute('modelId')) };
    if (office) sa.office = true;
    Object.defineProperty(sa, 'points', { value: points, enumerable: false });
    return sa;
  };
})();
