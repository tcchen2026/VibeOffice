/* Preset type names derived from Apache POI sl/usermodel/ShapeType.java (Apache-2.0).
 * Derived from Apache OpenOffice filter/source/msfilter/msdffimp.cxx (Apache-2.0):
 * rotated anchor boxes and gradient kinds.
 * OfficeArt shape style/geometry: MS-ODRAW; paths also informed by POI HSLFAutoShape.
 * Host writers supply anchors, colours and text; unsupported geometry is explicit. */
(function (root) {
  'use strict';
  const N = s => s; // UI refusal text, translated where it is shown.
  const B = root.L.legacy, A = B.officeart, S = B.shapes = {}, el = B.el;
  S.presets = Object.freeze({"20":"line","5":"triangle","6":"rtTriangle","1":"rect","4":"diamond","7":"parallelogram","8":"trapezoid","56":"pentagon","9":"hexagon","10":"octagon","187":"star4","12":"star5","58":"star8","59":"star16","18":"star16","92":"star24","60":"star32","2":"roundRect","21":"plaque","3":"ellipse","15":"homePlate","55":"chevron","95":"blockArc","23":"donut","57":"noSmoking","13":"rightArrow","66":"leftArrow","68":"upArrow","67":"downArrow","93":"stripedRightArrow","94":"notchedRightArrow","90":"bentUpArrow","69":"leftRightArrow","70":"upDownArrow","89":"leftUpArrow","182":"leftRightUpArrow","76":"quadArrow","77":"leftArrowCallout","78":"rightArrowCallout","79":"upArrowCallout","80":"downArrowCallout","81":"leftRightArrowCallout","82":"upDownArrowCallout","83":"quadArrowCallout","91":"bentArrow","101":"uturnArrow","99":"circularArrow","102":"curvedRightArrow","103":"curvedLeftArrow","104":"curvedUpArrow","105":"curvedDownArrow","16":"cube","22":"can","73":"lightningBolt","74":"heart","183":"sun","184":"moon","96":"smileyFace","71":"irregularSeal1","72":"irregularSeal2","65":"foldedCorner","84":"bevel","75":"frame","19":"arc","85":"leftBracket","86":"rightBracket","87":"leftBrace","88":"rightBrace","185":"bracketPair","186":"bracePair","32":"straightConnector1","33":"bentConnector2","34":"bentConnector3","35":"bentConnector4","36":"bentConnector5","37":"curvedConnector2","38":"curvedConnector3","39":"curvedConnector4","40":"curvedConnector5","41":"callout1","42":"callout2","43":"callout3","44":"accentCallout1","45":"accentCallout2","46":"accentCallout3","47":"borderCallout1","48":"borderCallout2","49":"borderCallout3","50":"accentBorderCallout1","51":"accentBorderCallout2","52":"accentBorderCallout3","61":"wedgeRectCallout","62":"wedgeRoundRectCallout","63":"wedgeEllipseCallout","106":"cloudCallout","53":"ribbon","54":"ribbon2","107":"ellipseRibbon","108":"ellipseRibbon2","97":"verticalScroll","98":"horizontalScroll","64":"wave","188":"doubleWave","11":"plus","109":"flowChartProcess","110":"flowChartDecision","111":"flowChartInputOutput","112":"flowChartPredefinedProcess","113":"flowChartInternalStorage","114":"flowChartDocument","115":"flowChartMultidocument","116":"flowChartTerminator","117":"flowChartPreparation","118":"flowChartManualInput","119":"flowChartManualOperation","120":"flowChartConnector","121":"flowChartPunchedCard","122":"flowChartPunchedTape","123":"flowChartSummingJunction","124":"flowChartOr","125":"flowChartCollate","126":"flowChartSort","127":"flowChartExtract","128":"flowChartMerge","129":"flowChartOfflineStorage","130":"flowChartOnlineStorage","131":"flowChartMagneticTape","132":"flowChartMagneticDisk","133":"flowChartMagneticDrum","134":"flowChartDisplay","135":"flowChartDelay","176":"flowChartAlternateProcess","177":"flowChartOffpageConnector","189":"actionButtonBlank","190":"actionButtonHome","191":"actionButtonHelp","192":"actionButtonInformation","193":"actionButtonForwardNext","194":"actionButtonBackPrevious","195":"actionButtonEnd","196":"actionButtonBeginning","197":"actionButtonReturn","198":"actionButtonDocument","199":"actionButtonSound","200":"actionButtonMovie"});
  const get = (s, id, fallback = 0) => s.properties.get(id)?.value ?? fallback;
  const signed = (s, id, fallback = 0) => get(s, id, fallback) | 0;
  const bool = (s, id, bit, fallback) => { const n = get(s, id); return n & (1 << (bit + 16)) ? !!(n & (1 << bit)) : fallback; };
  const hex = v => v.toString(16).padStart(2, '0').toUpperCase();
  S.color = function (value, context = {}, shape, depth = 0) {
    shape ||= { properties: new Map() };
    if (depth > 8) B.fail('SHAPE_COLOR', N('Cyclic shape colour reference'));
    if (value & 0x08000000) { const colour = context.scheme?.[value & 255]; if (!colour) B.fail('SHAPE_COLOR', N('Missing document colour scheme')); return colour; }
    if (value & 0x01000000) { const colour = context.palette?.[value & 65535]; if (!colour) B.fail('SHAPE_COLOR', N('Missing document colour palette')); return colour; }
    if (value & 0x10000000) {
      const index = value & 255, operation = value >>> 8 & 255, parameter = value >>> 16 & 255;
      const sourceId = { 240: 385, 241: bool(shape, 511, 3, true) ? 448 : 385, 242: 448, 243: 513, 245: 387, 246: 450, 247: bool(shape, 447, 4, true) ? 385 : 448 }[index];
      let colour = sourceId == null ? context.system?.[index] : S.color(get(shape, sourceId, sourceId === 385 ? 0xffffff : 0), context, shape, depth + 1);
      if (!colour) B.fail('SHAPE_COLOR', N('Unsupported system colour reference'));
      return colour.match(/../g).map(c => { let n = parseInt(c, 16); switch (operation & 15) {
        case 1: n *= parameter / 255; break; case 2: n = 255 - (255 - n) * parameter / 255; break;
        case 3: n += parameter; break; case 4: n -= parameter; break; case 5: n = parameter - n; break; case 6: n = n < parameter ? 0 : 255; break;
      } n = Math.max(0, Math.min(255, Math.round(n))); if (operation & 32) n ^= 255; if (operation & 64) n ^= 128; return hex(n); }).join('');
    }
    return hex(value & 255) + hex(value >>> 8 & 255) + hex(value >>> 16 & 255);
  };
  S.path = shape => B.shapePath(shape);
  S.geometry = function (shape) {
    const presetName = shape.type === 202 || shape.type === 75 ? 'rect' : S.presets[shape.type];
    const needsNative = !presetName || Array.from({length:10},(_,i)=>shape.properties.has(327+i)).some(Boolean);
    if (shape.properties.has(325) || shape.properties.has(337) || needsNative) {
      const path = S.path(shape), names = { M: 'moveTo', L: 'lnTo', C: 'cubicBezTo', Q: 'quadBezTo', Z: 'close' };
      const xml = path.paths.map(p => el('a:path', { w: path.width, h: path.height, fill: p.fill ? 'norm' : 'none', stroke: p.stroke ? 1 : 0 }, p.commands.map(c => el('a:' + names[c.op], {}, c.pts.length ? c.pts.map(([x, y]) => el('a:pt', { x: Math.round(x), y: Math.round(y) })).join('') : undefined)).join(''))).join('');
      let connections=path.connections;
      if(!connections.length&&[1,2,75,202].includes(shape.type))connections=[{x:path.width/2,y:0,angle:270},{x:0,y:path.height/2,angle:180},{x:path.width/2,y:path.height,angle:90},{x:path.width,y:path.height/2,angle:0}];
      // Paths have their own w/h; connection sites and the text rectangle use
      // the shape coordinate system. Scale them with guides so resizing works.
      const guides=[];
      const scaled=(name,value,axis)=>{guides.push(el('a:gd',{name,fmla:'*/ '+axis+' '+Math.round(value)+' '+(axis==='w'?path.width:path.height)}));return name;};
      const cxn=el('a:cxnLst',{},connections.map((p,i)=>el('a:cxn',{ang:Math.round(((p.angle%360)+360)%360*60000)},el('a:pos',{x:scaled('cx'+i,p.x,'w'),y:scaled('cy'+i,p.y,'h')}))).join(''));
      const rect=path.textRect?Object.fromEntries(Object.entries(path.textRect).map(([k,v])=>[k,scaled('tx'+k,v,k==='l'||k==='r'?'w':'h')])):{l:'l',t:'t',r:'r',b:'b'};
      return el('a:custGeom', {}, '<a:avLst/>'+el('a:gdLst',{},guides.join(''))+'<a:ahLst/>' + cxn + el('a:rect',rect) + el('a:pathLst', {}, xml));
    }
    return el('a:prstGeom', { prst: presetName }, '<a:avLst/>');
  };
  S.anchorRect = function (shape, rect) {
    if (!rect || ![rect.x, rect.y, rect.cx, rect.cy].every(Number.isFinite)) B.fail('SHAPE_ANCHOR', N('The shape has no usable anchor'));
    // Binary Office swaps the anchor extents around a quarter turn. AOO tests
    // its counterclockwise angle; xfrm/VML want the original unrotated box.
    const angle = ((-signed(shape, 4) / 65536 % 360) + 360) % 360;
    return (angle > 45 && angle <= 135 || angle > 225 && angle <= 315) ?
      { ...rect, x: rect.x + (rect.cx - rect.cy) / 2, y: rect.y + (rect.cy - rect.cx) / 2, cx: rect.cy, cy: rect.cx } : rect;
  };
  S.transform = function (shape, rect, group = false, local = false) {
    rect = S.anchorRect(shape, rect);
    if (local) rect = { ...rect, x: 0, y: 0 };
    const attrs = { rot: shape.properties.has(4) ? Math.round(signed(shape, 4) / 65536 * 60000) : null, flipH: shape.flags & 64 ? 1 : null, flipV: shape.flags & 128 ? 1 : null };
    let body = el('a:off', { x: Math.round(rect.x), y: Math.round(rect.y) }) + el('a:ext', { cx: Math.max(0, Math.round(rect.cx)), cy: Math.max(0, Math.round(rect.cy)) });
    if (group) { const c = shape.coord || { left: 0, top: 0, right: rect.cx, bottom: rect.cy }; body += el('a:chOff', { x: c.left, y: c.top }) + el('a:chExt', { cx: Math.max(1, c.right - c.left), cy: Math.max(1, c.bottom - c.top) }); }
    return el('a:xfrm', attrs, body);
  };
  S.style = function (shape, context = {}) {
    const colour = (value, opacity = 65536) => el('a:srgbClr', { val: S.color(value, context, shape) }, opacity === 65536 ? undefined : el('a:alpha', { val: Math.max(0, Math.min(100000, Math.round(opacity * 100000 / 65536))) }));
    const solid = (value, opacity) => el('a:solidFill', {}, colour(value, opacity));
    let fill;
    if (!bool(shape, 447, 4, ![20, 32, 33, 34, 35, 36, 37, 38, 39, 40, 202].includes(shape.type))) fill = '<a:noFill/>';
    else {
      const type = get(shape, 384), front = get(shape, 385, 0xffffff), back = get(shape, 387, 0xffffff);
      if (type === 0) fill = solid(front, get(shape, 386, 65536));
      else if ([1, 2, 3].includes(type)) {
        const id = context.imageRel?.(get(shape, 390)); if (!id) B.fail('SHAPE_FILL', N('The shape fill picture is missing'));
        fill = el('a:blipFill', {}, el('a:blip', { 'r:embed': id }) + (type === 3 ? '<a:stretch><a:fillRect/></a:stretch>' : '<a:tile/>'));
      } else if ([4, 5, 6, 7].includes(type)) {
        let stops = A.array(shape.properties.get(407)).map(s => ({ value: s.u32(0), pos: Math.max(0, Math.min(100000, Math.round(s.u32(4) * 100000 / 65536))) }));
        if (!stops.length) stops = [{ value: back, pos: 0, opacity: get(shape,388,65536) }, { value: front, pos: 100000, opacity: get(shape,386,65536) }]; stops.sort((a, b) => a.pos - b.pos);
        // POI HSLFFill and Office's embedded metro shape XML agree: focus 100
        // reverses the two endpoints; an intermediate focus reflects the ramp.
        let focusValue=Math.max(-100,Math.min(100,signed(shape,396)));
        if(Math.abs(focusValue)===100){const colors=stops.map(s=>({value:s.value,opacity:s.opacity})).reverse();stops=stops.map((s,i)=>({...s,...colors[i]}));}
        else if(focusValue){if(focusValue<0)focusValue+=100;const f=focusValue/100;stops=[...stops.slice(1).reverse().map(s=>({...s,pos:Math.round((100000-s.pos)*f)})),...stops.map(s=>({...s,pos:Math.round(100000*f+s.pos*(1-f))}))];}
        const edge = id => Math.max(0, Math.min(100000, Math.round(signed(shape, id) * 100000 / 65536)));
        const focus = type === 6 ? {l:50000,t:50000,r:50000,b:50000} : {l:edge(397),t:edge(398),r:100000-edge(399),b:100000-edge(400)};
        const dir = type === 5 || type === 6 ? el('a:path', { path: type === 6 ? 'shape' : 'rect' }, el('a:fillToRect', focus)) : el('a:lin', { ang: ((Math.round(5400000-signed(shape, 395)*60000/65536)%21600000)+21600000)%21600000, scaled: type === 7 ? 1 : 0 });
        fill = el('a:gradFill', { rotWithShape: bool(shape,447,5,false)?1:0 }, el('a:gsLst', {}, stops.map(s => el('a:gs', { pos: s.pos }, colour(s.value,s.opacity))).join('')) + dir);
      } else B.fail('SHAPE_FILL', N('This shape uses an unsupported fill'), type);
    }
    let line = el(context.lineTag||'a:ln',{},el('a:noFill'));
    if (bool(shape, 511, 3, shape.type !== 202 && shape.type !== 75)) {
      const arrows = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'], sizes = ['sm', 'med', 'lg'];
      let body = solid(get(shape, 448), get(shape, 449, 65536));
      const dash = ['solid', 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot'][get(shape, 462)];
      if (dash) body += el('a:prstDash', { val: dash });
      body += el(['a:bevel', 'a:miter', 'a:round'][get(shape, 470, 2)] || 'a:round', get(shape, 470, 2) === 1 ? { lim: Math.max(0, Math.round(get(shape, 460, 8 * 65536) * 100000 / 65536)) } : {});
      for (const [tag, head, width, length] of [['a:headEnd', 464, 466, 467], ['a:tailEnd', 465, 468, 469]]) body += el(tag, { type: arrows[get(shape, head)] || 'none', w: sizes[get(shape, width, 1)] || 'med', len: sizes[get(shape, length, 1)] || 'med' });
      line = el(context.lineTag||'a:ln', { w: Math.max(0, signed(shape, 459, 9525)), cap: ['rnd', 'sq', 'flat'][get(shape, 471, 2)] || 'flat', cmpd: ['sng', 'dbl', 'thickThin', 'thinThick', 'tri'][get(shape, 461)] || 'sng' }, body);
    }
    let effects = '';
    if (bool(shape, 575, 1, false)) {
      const x = signed(shape, 517, 25400), y = signed(shape, 518, 25400);
      if (get(shape, 512) > 1) B.fail('SHAPE_SHADOW', N('Unsupported perspective shape shadow'));
      effects = el('a:effectLst', {}, el('a:outerShdw', { dist: Math.round(Math.hypot(x, y)), dir: (Math.round(Math.atan2(y, x) * 180 / Math.PI * 60000) + 21600000) % 21600000, rotWithShape: 0 }, colour(get(shape, 513, 0x808080), get(shape, 516, 65536))));
    }
    return context.components ? {fill,line,effects} : fill + line + effects;
  };
  S.bodyPr = function (shape) {
    const anchor = get(shape, 135), flow = get(shape, 136), attrs = {
      lIns: signed(shape, 129, 91440), tIns: signed(shape, 130, 45720), rIns: signed(shape, 131, 91440), bIns: signed(shape, 132, 45720),
      wrap: get(shape, 133) === 2 ? 'none' : 'square', anchor: ['t', 'ctr', 'b', 't', 'ctr', 'b'][anchor] || 't', anchorCtr: anchor >= 3 && anchor <= 5 ? 1 : null,
      // MS-ODRAW MSOTXFL: 0/4 are horizontal; 1/3/5 turn clockwise.
      vert: ['horz', 'vert', 'vert270', 'vert', 'horz', 'vert'][flow] || 'horz',
      rot: shape.properties.has(137) ? (((get(shape,137)%4)+4)%4===3?-5400000:((get(shape,137)%4)+4)%4*5400000) : null
    };
    return el('a:bodyPr', attrs, bool(shape, 191, 1, false) ? '<a:spAutoFit/>' : '<a:noAutofit/>');
  };
  S.dml = (shape, rect, context) => S.transform(shape, rect) + S.geometry(shape) + S.style(shape, context);
  S.vml = function (shape, rect, context = {}) {
    rect = S.anchorRect(shape, rect);
    const id = context.id || '_x0000_s' + shape.id;
    let tag = { 1: 'v:rect', 2: 'v:roundrect', 3: 'v:oval', 20: 'v:line', 202: 'v:rect', 75: 'v:rect' }[shape.type], path;
    if (shape.properties.has(325) || !tag) { path = S.path(shape); tag = 'v:shape'; }
    if (!tag) B.fail('SHAPE_VML', N('This legacy shape needs a VML preset definition'));
    const style = 'position:absolute;margin-left:' + rect.x / 12700 + 'pt;margin-top:' + rect.y / 12700 + 'pt;width:' + rect.cx / 12700 + 'pt;height:' + rect.cy / 12700 + 'pt;rotation:' + signed(shape, 4) / 65536 + ((shape.flags & 192) ? ';flip:' + (shape.flags & 64 ? 'x' : '') + (shape.flags & 128 ? 'y' : '') : '');
    const attrs = { id, 'o:spid': id, style: context.style ?? style, filled: bool(shape, 447, 4, ![20, 202].includes(shape.type)) ? 't' : 'f', stroked: bool(shape, 511, 3, ![75, 202].includes(shape.type)) ? 't' : 'f', fillcolor: '#' + S.color(get(shape, 385, 0xffffff), context, shape), strokecolor: '#' + S.color(get(shape, 448), context, shape), strokeweight: signed(shape, 459, 9525) / 12700 + 'pt', ...context.attributes };
    if (path) { attrs.coordorigin = '0,0'; attrs.coordsize = path.width + ',' + path.height; attrs.path = path.paths.map(p => p.commands.map(c => ({ M: 'm', L: 'l', C: 'c', Q: 'qb', Z: 'x' })[c.op] + c.pts.flat().map(Math.round).join(',')).join('') + (p.fill ? '' : 'nf') + (p.stroke ? '' : 'ns') + 'e').join(''); }
    if (tag === 'v:line') { attrs.from = '0,0'; attrs.to = rect.cx / 12700 + 'pt,' + rect.cy / 12700 + 'pt'; }
    if (tag === 'v:roundrect') attrs.arcsize = signed(shape, 327, 5400) / 21600;
    const fillType = get(shape, 384), fillAttrs = { opacity: get(shape, 386, 65536) / 65536 };
    if ([1, 2, 3].includes(fillType)) { const rid = context.imageRel?.(get(shape, 390)); if (!rid) B.fail('SHAPE_FILL', N('The shape fill picture is missing')); Object.assign(fillAttrs, { type: fillType === 3 ? 'frame' : 'tile', 'r:id': rid }); }
    else if ([4, 5, 6, 7].includes(fillType)) {
      Object.assign(fillAttrs, { type: fillType === 5 || fillType === 6 ? 'gradientRadial' : 'gradient', color2: '#' + S.color(get(shape, 387, 0xffffff), context, shape), angle: signed(shape, 395) / 65536, focus: signed(shape, 396) + '%',rotate:bool(shape,447,5,false)?'t':'f' });
      if(fillType===4||fillType===7){const f=Math.max(-100,Math.min(100,signed(shape,396))),at=Math.abs(f)===100?100:f<0?100+f:f;fillAttrs.focus=(100-at)+'%';}
      if (fillType === 5) Object.assign(fillAttrs, {focusposition:signed(shape,397)/65536+','+signed(shape,398)/65536, focussize:(signed(shape,399)-signed(shape,397))/65536+','+(signed(shape,400)-signed(shape,398))/65536});
      if (fillType === 6) Object.assign(fillAttrs, {focusposition:'0.5,0.5',focussize:'0,0'});
      const stops = A.array(shape.properties.get(407)); if (stops.length) fillAttrs.colors = stops.map(s => s.u32(4) / 65536 + ' #' + S.color(s.u32(0), context, shape)).join(';');
    } else if (fillType) B.fail('SHAPE_FILL', N('This shape uses an unsupported fill'), fillType);
    let body = el('v:fill', fillAttrs);
    const arrows = ['none', 'block', 'classic', 'diamond', 'oval', 'open'], sizes = ['narrow', 'medium', 'wide'], lengths = ['short', 'medium', 'long'];
    body += el('v:stroke', { opacity: get(shape, 449, 65536) / 65536, startarrow: arrows[get(shape, 464)] || 'none', endarrow: arrows[get(shape, 465)] || 'none', startarrowwidth: sizes[get(shape, 466, 1)], startarrowlength: lengths[get(shape, 467, 1)], endarrowwidth: sizes[get(shape, 468, 1)], endarrowlength: lengths[get(shape, 469, 1)], dashstyle: ['solid', 'shortdash', 'shortdot', 'shortdashdot', 'shortdashdotdot', 'dot', 'dash', 'longdash', 'dashdot', 'longdashdot', 'longdashdotdot'][get(shape, 462)] });
    if (bool(shape, 575, 1, false)) body += el('v:shadow', { on: 't', type: 'single', color: '#' + S.color(get(shape, 513, 0x808080), context, shape), opacity: get(shape, 516, 65536) / 65536, offset: signed(shape, 517, 25400) / 12700 + 'pt,' + signed(shape, 518, 25400) / 12700 + 'pt' });
    if (context.imageRel) { const rid = context.imageRel(get(shape, 260)); if (rid) body += el('v:imagedata', { 'r:id': rid, 'o:title': '', ...context.imageAttributes }); }
    const text = typeof context.text === 'function' ? context.text() : context.text;
    if (text) body += el('v:textbox', { inset: [129, 130, 131, 132].map((id, i) => signed(shape, id, i % 2 ? 45720 : 91440) / 12700 + 'pt').join(','), ...context.textboxAttributes }, text);
    return el(tag, attrs, body + (context.tail || ''));
  };
})(typeof window !== 'undefined' ? window : globalThis);
