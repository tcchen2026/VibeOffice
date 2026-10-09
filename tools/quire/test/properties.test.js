const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../../public/common/core.js');
require('../../../public/common/xml.js');
require('../../../public/common/opc.js');
require('../../../public/common/opc-order.js');
require('../../../public/quire/js/dmodel.js');
require('../../../public/quire/js/ops.js');
require('../../../public/quire/js/preserve.js');
const { D, O, preserve: P, opc: K } = L;
const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const doc = () => { const d = D.newDoc(); D.doc = d; D.resetHistory(); return d; };
const context = d => ({ writer: new K.Writer(null, { doc: d }), doc: d, part: 'word/document.xml' });
const property = (xml, model = {}, kind = 'rPr', owner = kind === 'rPr' ? 'r' : 'p') => P.properties(model, K.parse(`<w:${owner} xmlns:w="${K.NS.w}" xmlns:w14="${w14}" xmlns:mc="${K.NS.mc}" mc:Ignorable="w14"><w:${kind}>${xml}</w:${kind}></w:${owner}>`).children[0], kind, { part: 'word/document.xml' });

test('copied source formatting repairs property order and invalid color spelling without losing extensions', () => {
  const d = doc(), r = property('<w:sz w:val="24"/><w:color w:val="#123456" w:themeColor="accent1"/><w14:glow w14:rad="190500"/>', { sz: 12, color: '123456' });
  const xml = P.propertyXML('rPr', r, '<w:color w:val="123456"/><w:sz w:val="24"/>', context(d));
  assert.match(xml, /w:val="123456"/); assert.doesNotMatch(xml, /#123456/);
  assert.match(xml, /w:themeColor="accent1"/); assert.match(xml, /w14:glow/);
  assert.ok(xml.indexOf('<w:color') < xml.indexOf('<w:sz'));
  assert.equal(d.losses.length, 1);
});

test('theme-only colors receive their required fallback when retained across text splits', () => {
  const d = doc(), r = property('<w:u w:val="single"/><w:color w:themeColor="hyperlink"/>', { u: 'single' });
  const xml = P.propertyXML('rPr', r, '<w:u w:val="single"/>', context(d));
  assert.match(xml, /w:themeColor="hyperlink"/); assert.match(xml, /w:val="auto"/);
  assert.ok(xml.indexOf('<w:color') < xml.indexOf('<w:u'));
});

test('redundant invalid properties collapse without removing unknown children', () => {
  const d = doc(), p = property('<w:pBdr/><w:framePr w:w="1000"/><w:pBdr/><w14:unknown/><w14:unknown/>', {}, 'pPr');
  const xml = P.propertyXML('pPr', p, '', context(d));
  assert.equal((xml.match(/<w:pBdr\b/g) || []).length, 1);
  assert.equal((xml.match(/<w14:unknown\b/g) || []).length, 2);
  assert.ok(xml.indexOf('<w:framePr') < xml.indexOf('<w:pBdr'));
  P.propertyXML('rPr', property('<w:b/><w:b/>', { b: true }), '<w:b/>', context(d));
  // A document with hundreds of malformed runs gets one notice per source part.
  assert.equal(d.losses.length, 1);
});

test('automatic repairs leave styles and ambiguous source properties unchanged', () => {
  for (const [kind, owner, raw] of [
    ['rPr', 'style', '<w:highlight w:val="cyan"/><w:rFonts w:cs="Times New Roman"/>'],
    ['rPr', 'r', '<w:b/><w:b/><w:sz w:val="24"/><w:sz w:val="26"/>'],
    ['pPr', 'p', '<w:spacing w:after="200"/><w:suppressAutoHyphens/><w:hyphenationLines w:val="0"/>'],
  ]) {
    const d = doc(), model = property(raw, {}, kind, owner);
    const xml = P.propertyXML(kind, model, '', context(d));
    assert.ok(xml.includes(raw)); assert.equal(d.losses.length, 0);
  }
});

test('formatting changes replace their own property and undo restores unsupported effects', () => {
  const d = doc(), r = property('<w:lang w:val="en-US" w:bidi="ar-SA"/><w14:glow w14:rad="190500"/><w14:shadow w14:blurRad="76200"/><w14:reflection w14:blurRad="12700"/>', { lang: 'en-US' });
  const p = D.para([D.text('hello', r)]); d.main.blocks = [p];
  const emit = () => P.propertyXML('rPr', p.runs[0].rPr, '<w:b/><w:lang w:val="en-US"/><w:shadow w:val="0"/>', context(d));
  const original = P.propertyXML('rPr', r, '<w:lang w:val="en-US"/>', context(d));
  D.tx('Change shadow and bold', () => O.setRunProps(D.pos(p, 0), D.pos(p, 5), { shadow: false, b: true }));
  const changed = emit();
  assert.match(changed, /w14:glow/); assert.match(changed, /w14:reflection/);
  assert.doesNotMatch(changed, /w14:shadow/); assert.match(changed, /w:shadow w:val="0"/);
  assert.match(changed, /w:bidi="ar-SA"/);
  assert.ok(changed.indexOf('<w:b ') < changed.indexOf('<w14:glow'));
  D.undo(); assert.equal(P.propertyXML('rPr', p.runs[0].rPr, '<w:lang w:val="en-US"/>', context(d)), original);
  assert.equal(d.losses?.length || 0, 0);
});

test('property alternatives keep both branches and change only the edited property', () => {
  const d = doc(), r = property('<mc:AlternateContent><mc:Choice Requires="w14"><w14:shadow w14:blurRad="76200"/><w14:glow w14:rad="190500"/></mc:Choice><mc:Fallback><w:shadow/></mc:Fallback></mc:AlternateContent>', { shadow: true });
  const before = P.propertyXML('rPr', r, '<w:shadow/>', context(d));
  assert.match(before, /mc:Choice/); assert.match(before, /mc:Fallback/); assert.match(before, /w14:shadow/);
  r.shadow = false;
  const after = P.propertyXML('rPr', r, '<w:b/><w:shadow w:val="0"/>', context(d));
  assert.match(after, /w14:glow/); assert.doesNotMatch(after, /w14:shadow/);
  assert.equal((after.match(/<w:shadow /g) || []).length, 1);
  assert.ok(after.indexOf('<w:b ') < after.indexOf('<mc:AlternateContent'));
  assert.match(after, /mc:Fallback/); assert.equal(d.losses?.length || 0, 0);
});

test('section preservation keeps absent defaults, unread settings and individual margin attributes', () => {
  const d = doc(), s = property('<w:pgSz w:w="11901" w:h="9000"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0" w14:other="kept"/><w:rtlGutter/>', { ...D.defaultSect(), pgW: 595.05, pgH: 450, orient: 'landscape' }, 'sectPr', 'body');
  const generated = '<w:pgSz w:w="11901" w:h="9000" w:orient="landscape"/><w:pgMar w:top="1800" w:left="1800"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/>';
  const before = P.propertyXML('sectPr', s, generated, context(d));
  assert.doesNotMatch(before, /w:orient=|w:cols|w:docGrid/);
  assert.match(before, /<w:rtlGutter/);
  s.mt = 90;
  const after = P.propertyXML('sectPr', s, generated, context(d));
  assert.match(after, /w:top="1800"/); assert.match(after, /w:left="1440"/);
  assert.match(after, /w14:other="kept"/); assert.match(after, /<w:rtlGutter/);
  assert.doesNotMatch(after, /w:orient=|w:cols|w:docGrid/);
  s.cols.n = 2;
  const columns = P.propertyXML('sectPr', s, generated.replace('<w:cols w:space="720"/>', '<w:cols w:num="2" w:space="720"/>'), context(d));
  assert.match(columns, /w:num="2"/); assert.ok(columns.indexOf('<w:cols') < columns.indexOf('<w:rtlGutter'));
  assert.equal(d.losses?.length || 0, 0);
});

test('section numbering edits keep chapter, grid and revision details through JSON history', () => {
  const d = doc(), s = property('<w:pgNumType w:start="1" w:chapStyle="1" w:chapSep="hyphen"/><w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="4096"/><w:sectPrChange w:id="7" w:author="Author"><w:sectPr><w:rtlGutter/></w:sectPr></w:sectPrChange>', { ...D.defaultSect(), pgNum: { start: 1, fmt: 'decimal', chapStyle: '1', chapSep: 'hyphen' }, docGrid: { type: 'linesAndChars', linePitch: 312, charSpace: 4096 } }, 'sectPr', 'body');
  d.sect = s;
  const generated = '<w:pgNumType w:start="5"/><w:docGrid w:type="linesAndChars" w:linePitch="360"/>';
  const before = P.propertyXML('sectPr', s, generated, context(d));
  D.tx('Page numbering', () => { D.touchKey(d, 'sect'); d.sect.pgNum.start = 5; });
  const after = P.propertyXML('sectPr', d.sect, generated, context(d));
  assert.match(after, /w:start="5"/); assert.match(after, /w:chapStyle="1"/); assert.match(after, /w:chapSep="hyphen"/);
  assert.match(after, /w:linePitch="312"/); assert.match(after, /w:charSpace="4096"/); assert.match(after, /w:sectPrChange/);
  D.undo(); assert.equal(P.propertyXML('sectPr', d.sect, generated, context(d)), before);
  D.redo(); assert.equal(P.propertyXML('sectPr', JSON.parse(JSON.stringify(d.sect)), generated, context(d)), after);
  assert.equal(d.losses?.length || 0, 0);
});

test('new section attributes do not rebind a foreign prefix on retained children', () => {
  const d = doc(), s = P.properties(D.defaultSect(), K.parse(`<sectPr xmlns="${K.NS.w}" xmlns:q="${K.NS.w}" xmlns:w="urn:foreign"><pgNumType q:fmt="decimal"><w:unknown/></pgNumType></sectPr>`), 'sectPr', { part: 'word/document.xml' });
  s.pgNum.start = 3;
  const xml = P.propertyXML('sectPr', s, '<w:pgNumType w:start="3"/>', context(d));
  const element = K.parse(xml).children[0];
  const number = Array.from(element.attributes).find(a => a.localName === 'start');
  assert.equal(number.value, '3'); assert.equal(element.lookupNamespaceURI(number.prefix), K.NS.w);
  assert.equal(element.children[0].namespaceURI, 'urn:foreign');
  assert.match(xml, /q:fmt="decimal"/);
});

test('numbering edits retain legacy details, level attributes, alternatives and absent defaults', () => {
  const d = doc(), source = `<w:lvl xmlns:w="${K.NS.w}" xmlns:w14="${w14}" xmlns:mc="${K.NS.mc}" mc:Ignorable="w14" w:ilvl="2" w:tplc="12345678" w:tentative="1"><mc:AlternateContent><mc:Choice Requires="w14"><w:numFmt w:val="custom" w:format="001, 002, 003, ..."/></mc:Choice><mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback></mc:AlternateContent><w:legacy w:legacy="1" w:legacySpace="120" w:legacyIndent="-80"/><w:lvlJc w:val="left"/></w:lvl>`;
  const level = P.properties({ start: 1, fmt: 'custom', custFmt: '001, 002, 003, ...', text: '%3.', jc: 'left', suff: 'tab', legacy: true }, K.parse(source), 'lvl', { part: 'word/numbering.xml' });
  const ctx = { ...context(d), part: 'word/numbering.xml' };
  const generated = '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%3."/><w:legacy w:legacy="1"/><w:lvlJc w:val="right"/>';
  level.jc = 'right';
  const xml = P.propertyXML('lvl', level, generated, ctx, { ilvl: 2 });
  assert.match(xml, /w:legacySpace="120"/); assert.match(xml, /w:legacyIndent="-80"/); assert.match(xml, /w:tplc="12345678"/);
  assert.match(xml, /w:val="custom"/); assert.match(xml, /w:format="001, 002, 003, ..."/); assert.match(xml, /w:val="right"/);
  assert.doesNotMatch(xml, /<w:start|<w:lvlText/);
  level.fmt = 'decimal'; delete level.custFmt;
  const changed = P.propertyXML('lvl', level, generated, ctx, { ilvl: 3 });
  assert.doesNotMatch(changed, /w:val="custom"|w:format=/); assert.match(changed, /mc:Choice/); assert.match(changed, /mc:Fallback/);
  assert.equal((changed.match(/w:val="decimal"/g) || []).length, 2);
  assert.match(changed, /w:ilvl="3"/); assert.match(changed, /w:tentative="1"/);
  assert.equal(d.losses?.length || 0, 0);
});

test('numbering metadata follows its owner while deleted levels stay deleted and new IDs remain linked', () => {
  const d = doc(), ctx = { ...context(d), part: 'word/numbering.xml' };
  const a = P.properties({ id: '42', multi: 'singleLevel', levels: [D.numLevel(0, 'decimal', '%1.'), D.numLevel(1, 'decimal', '%2.')] }, K.parse(`<w:abstractNum xmlns:w="${K.NS.w}" w:abstractNumId="42"><w:nsid w:val="ABCD0123"/><w:multiLevelType w:val="singleLevel"/><w:tmpl w:val="01020304"/><w:lvl w:ilvl="0"/></w:abstractNum>`), 'abstractNum', { part: ctx.part, source: 'test' });
  P.numberingLevels(a, [0]);
  assert.equal(P.writeNumberingLevel(a, 0), true); assert.equal(P.writeNumberingLevel(a, 1), false);
  a.levels[1].start = 3; assert.equal(P.writeNumberingLevel(a, 1), true);
  const xml = P.propertyXML('abstractNum', a, '<w:multiLevelType w:val="singleLevel"/>', ctx, { abstractNumId: 42 });
  assert.match(xml, /w:abstractNumId="42"/); assert.match(xml, /w:nsid w:val="ABCD0123"/); assert.match(xml, /w:tmpl w:val="01020304"/); assert.doesNotMatch(xml, /<w:lvl /);
  d.numbering = { abs: { 42: a }, nums: {} };
  const numId = D.addNum(d, a), copied = d.numbering.abs[d.numbering.nums[numId].abs];
  P.prepareNumbering(d.numbering, ctx.writer);
  assert.notEqual(copied.x.fragment.copy, a.x.fragment.copy);
  assert.equal(ctx.writer.ids.resolve('test', 'numbering', 'abstractNumId', '42', { copy: copied.x.fragment.copy }), String(copied.id));
  assert.equal(ctx.writer.ids.resolve('test', 'numbering', 'abstractNumId', '42'), '42');
});

test('copied list formatting replaces an invalid highlight with equivalent shading', () => {
  const d = doc(), ctx = { ...context(d), part: 'word/numbering.xml' };
  L.R = { HIGHLIGHT: { yellow: '#FFFF00' } };
  const original = property('<w:highlight w:val="yellow"/><w:color w:val="123456"/>', { hl: 'yellow', color: '123456' }, 'rPr', 'lvl');
  const copied = K.duplicate(original);
  const xml = P.propertyXML('rPr', copied, '<w:color w:val="123456"/><w:highlight w:val="yellow"/>', ctx);
  assert.doesNotMatch(xml, /<w:highlight/); assert.match(xml, /w:fill="FFFF00"/);
  assert.ok(xml.indexOf('<w:color') < xml.indexOf('<w:shd'));
  assert.equal(d.losses.length, 1);
});

test('permission copies remap both endpoints, partial copies omit them, and incomplete saves report loss', () => {
  const d = doc(), start = P.rangeMarker(K.parse(`<w:permStart xmlns:w="${K.NS.w}" w:id="42" w:edGrp="everyone" w:colFirst="0" w:colLast="2"/>`), { part: 'word/document.xml' });
  const end = P.rangeMarker(K.parse(`<w:permEnd xmlns:w="${K.NS.w}" w:id="42"/>`), { part: 'word/document.xml' });
  const p = D.para([start, D.text('hello'), end]); d.main.blocks = [p];
  const copy = D.cloneBlocks(O.copyRange(D.pos(p, 0), D.pos(p, 5)))[0];
  d.main.blocks.push(copy);
  const ctx = { ...context(d), ranges: P.prepareRanges(d) };
  const sourceXML = [start, end].map(it => P.markerXML(it, ctx)).join('');
  const copyXML = copy.runs.filter(it => it.t === 'perm').map(it => P.markerXML(it, ctx)).join('');
  const ids = xml => [...xml.matchAll(/w:id="(\d+)"/g)].map(m => m[1]);
  assert.equal(ids(sourceXML)[0], ids(sourceXML)[1]); assert.equal(ids(copyXML)[0], ids(copyXML)[1]);
  assert.notEqual(ids(sourceXML)[0], ids(copyXML)[0]); assert.match(copyXML, /w:colLast="2"/);
  assert.equal(D.cloneBlocks(O.copyRange(D.pos(p, 0), D.pos(p, 2)))[0].runs.some(it => it.t === 'perm'), false);
  p.runs.pop(); ctx.ranges = P.prepareRanges(d);
  assert.equal(P.markerXML(start, ctx), ''); assert.equal(d.losses.length, 1);
});
