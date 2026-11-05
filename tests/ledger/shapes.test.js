/* Shapes Ledger writes itself (drawn, or edited): a line keeps the direction it was drawn in (flipH/flipV), an arrow
   its arrowhead (tailEnd), a preset its adjustments and a freeform the geometry it came with. They used to be written
   from top left to bottom right, without arrowheads or adjustments, and a freeform as prst="custom". */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const L = require('./load.js');
const K = L.opc;

const anchor = { type: 'two', from: { r: 1, c: 1, rOff: 0, cOff: 0 }, to: { r: 5, c: 4, rOff: 0, cOff: 0 }, editAs: 'twoCell' };
async function saved(drawings) {
  const wb = new L.model.Workbook(), sh = wb.addSheet('Sheet1');
  sh.drawings = drawings.map((d, i) => Object.assign({ kind: 'shape', id: i + 2, anchor, fill: '#FFFFFF', line: '#000000', lineW: 0.75, text: '', dirty: true }, d));
  const pkg = await K.open(await L.xlsxWrite.bytes(wb));
  return pkg.text('xl/drawings/drawing1.xml').match(/<xdr:sp\b[\s\S]*?<\/xdr:sp>/g);
}

test('a line drawn up and leftwards, an arrow, a preset with adjustments and a freeform are written as they are', async () => {
  const custGeom = '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo></a:path></a:pathLst></a:custGeom>';
  const [line, arrow, round, free] = await saved([
    { name: 'Line 2', geom: 'line', fill: null, flipH: true, flipV: true },
    { name: 'Line 3', geom: 'line', fill: null, tail: { type: 'triangle', w: 'med', len: 'med' } },
    { name: 'Rounded Rectangle 4', geom: 'roundRect', adj: { adj: 25000 } },
    { name: 'Freeform 5', geom: 'custom', xml: `<xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:spPr>${custGeom}</xdr:spPr></xdr:sp>` },
  ]);
  assert.match(line, /<a:xfrm flipH="1" flipV="1">/);
  assert.match(arrow, /<\/a:solidFill><a:tailEnd type="triangle"\/><\/a:ln>/, 'the arrowhead after the line\'s fill');
  assert.match(round, /<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"\/><\/a:avLst><\/a:prstGeom>/);
  assert.ok(free.includes(custGeom), 'the freeform\'s own geometry');
  assert.doesNotMatch(free, /prst="custom"/);
});

test('a shape from the file, written by Ledger after an edit, keeps what the edit didn\'t touch', async () => {
  const xml = '<xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Text Box 2"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>' +
    '<xdr:spPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="254000" cy="2540000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst></xdr:spPr>' +
    '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>' +
    '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" upright="1"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="hu-HU" sz="1200" b="1"/><a:t>Label</a:t></a:r></a:p></xdr:txBody></xdr:sp>';
  /* as Format AutoShape leaves it: dirty, nothing set; a one-cell anchor 200 × 20 pt (turned: the box is 20 × 200) */
  const [sp] = await saved([{ name: 'Text Box 2', geom: 'rect', txBox: true, xml, tag: 'sp', rot: 90, text: 'Label', vAlign: 'ctr', align: 'ctr',
    paras: [{ algn: 'ctr', runs: [{ t: 'Label', lang: 'hu-HU', sz: 12, b: true }] }], anchor: { type: 'one', from: { r: 1, c: 1, rOff: 0, cOff: 0 }, w: 200, h: 20 } }]);
  assert.match(sp, /<a:xfrm rot="5400000"><a:off x="0" y="0"\/><a:ext cx="254000" cy="2540000"\/><\/a:xfrm>/, 'the turned box\'s own size');
  assert.match(sp, /<\/a:prstGeom><a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"\/>/, 'its theme fill and shaded line');
  assert.match(sp, /<\/a:ln><a:effectLst><a:outerShdw/, 'its shadow');
  assert.match(sp, /<\/xdr:spPr><xdr:style><a:lnRef idx="2">/, 'its style');
  assert.match(sp, /<a:bodyPr vertOverflow="clip" horzOverflow="clip" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" upright="1"\/>/, 'its text body');
  assert.match(sp, /<a:r><a:rPr lang="hu-HU" sz="1200" b="1"\/><a:t>Label<\/a:t><\/a:r>/, 'its run, the rest from the style');
});
