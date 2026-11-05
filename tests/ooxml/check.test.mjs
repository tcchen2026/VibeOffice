import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import '../../public/common/zip.js';

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vo-package-tests-'));
test.after(() => fs.rmSync(dir, { recursive: true, force: true }));
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const rels = content => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
const rel = (id, target, type = 'image') => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
const slide = content => `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}">${content}</p:sld>`;
const picture = id => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture"/></p:nvPicPr></p:pic>`;
async function pkg(name, parts) {
  const types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/></Types>';
  const blob = await L.zip.write(Object.entries({ '[Content_Types].xml': types, ...parts }).map(([name, data]) => ({ name, data })));
  const file = path.join(dir, name + '.zip');
  fs.writeFileSync(file, new Uint8Array(await blob.arrayBuffer()));
  return file;
}
function run(command, ...args) {
  const r = spawnSync('python3', [fileURLToPath(new URL('../../tools/ooxml/package.py', import.meta.url)), command, ...args], { encoding: 'utf8' });
  assert.equal(r.error, undefined);
  assert.ok(r.stdout, r.stderr);
  return JSON.parse(r.stdout);
}
test('checks every relationship attribute and internal target, while retaining external URLs', async () => {
  const file = await pkg('rels', { 'a.xml': `<x xmlns:r="${R}" r:embed="missing" r:dm="ok"/>`, '_rels/a.xml.rels': rels(rel('ok', 'missing.xml') + `<Relationship Id="url" Type="${R}/hyperlink" Target="https://example.invalid/" TargetMode="External"/>`) });
  const codes = run('check', file).issues.map(i => i.code);
  assert.ok(codes.includes('missing-target'));
  assert.ok(codes.includes('unresolved-rid'));
  assert.equal(codes.length, 2);
});
test('valid shared targets and cyclic relationships do not look like missing parts', async () => {
  const file = await pkg('cycle', { 'a.xml': '<x/>', 'b.xml': '<y/>', '_rels/a.xml.rels': rels(rel('one', 'b.xml') + rel('two', 'b.xml')), '_rels/b.xml.rels': rels(rel('back', 'a.xml')) });
  assert.equal(run('check', file).status, 'ok');
});
test('OLE spid resolves in its related VML preview rather than slide shape IDs', async () => {
  const parts = id => ({ 'slide.xml': slide(`<p:oleObj spid="${id}"/>`),
    '_rels/slide.xml.rels': rels(rel('vml', 'vml.xml', 'vmlDrawing')),
    'vml.xml': '<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="_x0000_s1030"/></xml>' });
  assert.equal(run('check', await pkg('ole-vml', parts('_x0000_s1030'))).status, 'ok');
  assert.ok(run('check', await pkg('ole-missing-vml', parts('_x0000_s1031'))).issues.some(i => i.code === 'unresolved-vml-preview'));
});
test('an ActiveX control\'s spid names its VML shape by o:spid; Office\'s conditional VML markers parse', async () => {
  const parts = id => ({ 'slide.xml': slide(`<p:cSld><p:spTree/><p:controls><p:control spid="${id}" name="CommandButton1" r:id="ax"/></p:controls></p:cSld>`),
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>',
    '_rels/slide.xml.rels': rels(rel('vml', 'vml.vml', 'vmlDrawing') + rel('ax', 'ax.xml', 'control')),
    'ax.xml': '<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"/>',
    'vml.vml': '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><![if gte mso 9]><v:shape id="CommandButton1" o:spid="_x0000_s1030"/><![endif]></xml>' });
  assert.deepEqual(run('check', await pkg('control-vml', parts('1030'))).issues, []);
  assert.ok(run('check', await pkg('control-missing-vml', parts('1031'))).issues.some(i => i.code === 'unresolved-vml-preview'));
});
test('PowerPoint graphic frames cannot carry DrawingML rotation or flips', async () => {
  const frame = attrs => slide(`<p:graphicFrame><p:xfrm${attrs}><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm></p:graphicFrame>`);
  assert.equal(run('check', await pkg('plain-frame', { 'slide.xml': frame('') })).status, 'ok');
  const bad = run('check', await pkg('rotated-frame', { 'slide.xml': frame(' rot="60000" flipH="0" flipV="1"') }));
  assert(bad.issues.some(i => i.code === 'graphic-frame-orientation'));
});
test('PowerPoint table style lists name their default style even when it is built in', async () => {
  const styles = attr => ({ 'ppt/tableStyles.xml': `<a:tblStyleLst xmlns:a="${A}"${attr}/>` });
  assert.equal(run('check', await pkg('table-style-default', styles(' def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"'))).status, 'ok');
  assert.ok(run('check', await pkg('table-style-no-default', styles(''))).issues.some(i => i.code === 'table-style-default'));
});
test('synthetic unread graphics need an unsupported compatibility choice with a fallback', async () => {
  const graphic = '<p:graphicFrame><a:graphic><a:graphicData uri="urn:vibeoffice:test:unread"><probe:object xmlns:probe="urn:vibeoffice:test:unread"/></a:graphicData></a:graphic></p:graphicFrame>';
  const check = async (name, xml) => run('check', await pkg(name, { 'slide.xml': slide(xml) }));
  assert((await check('unguarded-probe', graphic)).issues.some(i => i.code === 'unguarded-test-graphic'));
  const choice = (requires, fallback) => `<mc:AlternateContent xmlns:mc="${MC}" xmlns:probe="urn:vibeoffice:test:unread" xmlns:p="${P}"><mc:Choice Requires="${requires}">${graphic}</mc:Choice>${fallback}</mc:AlternateContent>`;
  assert.equal((await check('guarded-probe', choice('probe', '<mc:Fallback/>'))).status, 'ok');
  assert((await check('probe-without-fallback', choice('probe', ''))).issues.some(i => i.code === 'unguarded-test-graphic'));
  assert((await check('probe-in-supported-choice', choice('p', '<mc:Fallback/>'))).issues.some(i => i.code === 'unguarded-test-graphic'));
  assert.equal((await check('unknown-real-graphic', graphic.replaceAll('urn:vibeoffice:test:unread', 'urn:vendor:future-graphic'))).status, 'ok');
});
test('VML drawing parts name each shape once and each embedded object has its own preview', async () => {
  const parts = (ids, spids) => ({ 'slide.xml': slide(spids.map(id => `<p:oleObj spid="${id}"/>`).join('')),
    '_rels/slide.xml.rels': rels(rel('vml', 'vml.vml', 'vmlDrawing')),
    'vml.vml': `<xml xmlns:v="urn:schemas-microsoft-com:vml">${ids.map(id => `<v:shape id="${id}"/>`).join('')}</xml>`,
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>' });
  assert.equal(run('check', await pkg('vml-unique', parts(['_x0000_s1025', '_x0000_s1026'], ['_x0000_s1025', '_x0000_s1026']))).status, 'ok');
  const codes = run('check', await pkg('vml-duplicate', parts(['_x0000_s1026', '_x0000_s1025', '_x0000_s1026'], ['_x0000_s1026', '_x0000_s1025', '_x0000_s1026']))).issues.map(i => i.code);
  assert.ok(codes.includes('duplicate-vml-id'));
  assert.ok(codes.includes('shared-vml-preview'));
});
test('each spreadsheet VML drawing has its own shape id block', async () => {
  const vml = (data, id) => `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${data}"/></o:shapelayout><v:shape id="_x0000_s${id}"/></xml>`;
  const types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>';
  assert.equal(run('check', await pkg('vml-blocks', { '[Content_Types].xml': types, 'xl/drawings/a.vml': vml(1, 1025), 'xl/drawings/b.vml': vml(2, 2049) })).status, 'ok');
  assert.ok(run('check', await pkg('vml-id-across-sheets', { '[Content_Types].xml': types, 'xl/drawings/a.vml': vml('1,2', 2049), 'xl/drawings/b.vml': vml(3, 2049) })).issues.some(i => i.code === 'duplicate-vml-id'));
  assert.ok(run('check', await pkg('vml-block-clash', { '[Content_Types].xml': types, 'xl/drawings/a.vml': vml(1, 1025), 'xl/drawings/b.vml': vml(1, 1025) })).issues.some(i => i.code === 'duplicate-vml-block'));
});
test('ActiveX names are VBA identifiers and match their VML preview, unlike form-control captions', async () => {
  const parts = (name, id, type = 'control') => ({
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><controls><control name="${name}" shapeId="1025" r:id="ctrl"/></controls></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('ctrl', '../activeX/activeX1.xml', type) + rel('vml', '../drawings/vml.xml', 'vmlDrawing')),
    'xl/activeX/activeX1.xml': '<control/>',
    'xl/drawings/vml.xml': `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="${id}" o:spid="_x0000_s1025"/></xml>` });
  assert.equal(run('check', await pkg('activex-good', parts('cmdOK1', 'cmdOK1'))).status, 'ok');
  const codes = run('check', await pkg('activex-bad', parts('cmdOK 1', '1'))).issues.map(i => i.code);
  assert(codes.includes('activex-control-name')); assert(codes.includes('activex-vml-name'));
  assert.equal(run('check', await pkg('form-control-name', parts('Button 1', '_x0000_s1025', 'ctrlProp'))).status, 'ok');
});
test('SmartArt drawing caches may repeat shape id 0, as PowerPoint writes them', async () => {
  const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
  const sp = `<dsp:sp modelId="{1}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr></dsp:sp>`;
  assert.equal(run('check', await pkg('dsp-zero', { 'drawing.xml': `<dsp:drawing xmlns:dsp="${DSP}"><dsp:spTree>${sp}${sp}</dsp:spTree></dsp:drawing>` })).status, 'ok');
  assert.ok(run('check', await pkg('slide-zero', { 'slide.xml': slide(picture(0) + picture(0)) })).issues.some(i => i.code === 'duplicate-shape-id'));
});
test('a SmartArt data part names its layout part by uniqueId, or PowerPoint draws its default list', async () => {
  const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
  const frame = `<p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" r:dm="d" r:lo="l" r:qs="d" r:cs="d"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>`;
  const parts = (lo) => ({
    'slide.xml': slide(frame), '_rels/slide.xml.rels': rels(rel('d', 'data.xml', 'diagramData') + rel('l', 'layout.xml', 'diagramLayout')),
    'data.xml': `<dgm:dataModel xmlns:dgm="${DGM}"><dgm:ptLst><dgm:pt modelId="1" type="doc"><dgm:prSet loTypeId="${lo}"/></dgm:pt></dgm:ptLst></dgm:dataModel>`,
    'layout.xml': `<dgm:layoutDef xmlns:dgm="${DGM}" uniqueId="urn:example/layout/process#1f2e3d4c"/>`,
  });
  assert.equal(run('check', await pkg('dgm-same', parts('urn:example/layout/process#1f2e3d4c'))).status, 'ok');
  assert.ok(run('check', await pkg('dgm-differ', parts('urn:example/layout/process'))).issues.some(i => i.code === 'diagram-layout-id'));
});
test('text moved away with Track Changes stays w:t inside w:moveFrom, or Word refuses the file', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const doc = (text) => ({ 'word/document.xml': `<w:document xmlns:w="${W}"><w:body><w:p><w:moveFrom w:id="1" w:author="A"><w:r>${text}</w:r></w:moveFrom><w:del w:id="2" w:author="A"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p></w:body></w:document>` });
  assert.equal(run('check', await pkg('move-text', doc('<w:t>moved</w:t>'))).status, 'ok');
  assert.ok(run('check', await pkg('move-deltext', doc('<w:delText>moved</w:delText>'))).issues.some(i => i.code === 'moved-deleted-text'));
});
test('a styles part holds at most 4,092 styles, or Word 2016 refuses the file', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const styles = (n) => ({ 'word/styles.xml': `<w:styles xmlns:w="${W}">${Array.from({ length: n }, (_, i) => `<w:style w:type="paragraph" w:styleId="S${i}"><w:name w:val="S ${i}"/></w:style>`).join('')}</w:styles>` });
  assert.equal(run('check', await pkg('styles-4092', styles(4092))).status, 'ok');
  assert.ok(run('check', await pkg('styles-4093', styles(4093))).issues.some(i => i.code === 'too-many-styles'));
});
test('a sheet comment drawn by DrawingML names a shape of the sheet\'s drawing, or Excel 2016 refuses the file', async () => {
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const parts = (id) => ({
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${S}"/>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('c', '../comments1.xml', 'comments') + rel('d', '../drawings/drawing1.xml', 'drawing')),
    'xl/comments1.xml': `<comments xmlns="${S}" xmlns:mc="${MC}" xmlns:v="urn:schemas-microsoft-com:vml"><commentList><comment ref="A1" shapeId="1025"><text><t>Note</t></text><mc:AlternateContent><mc:Choice Requires="v"/><mc:Fallback><commentPr/></mc:Fallback></mc:AlternateContent></comment></commentList></comments>`,
    'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:mc="${MC}" xmlns:v="urn:schemas-microsoft-com:vml"><mc:AlternateContent><mc:Choice Requires="v"/><mc:Fallback><xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Comment"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor></mc:Fallback></mc:AlternateContent></xdr:wsDr>`,
  });
  assert.deepEqual(run('check', await pkg('comment-shape-kept', parts(1025))).issues.map(i => i.code + ' ' + i.detail), []);
  assert.ok(run('check', await pkg('comment-shape-gone', parts(1026))).issues.some(i => i.code === 'comment-shape'));
});
test('comment extensions use the w15 namespace the other parts bind, or Word 2016 refuses the file', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', DRAFT = 'http://schemas.microsoft.com/office/word/2010/11/wordml';
  const parts = (ex) => ({ 'word/document.xml': `<w:document xmlns:w="${W}" xmlns:w15="${DRAFT}"><w:body/></w:document>`, 'word/commentsExtended.xml': `<w15:commentsEx xmlns:w15="${ex}"/>` });
  assert.equal(run('check', await pkg('w15-draft', parts(DRAFT))).status, 'ok');
  assert.ok(run('check', await pkg('w15-mixed', parts('http://schemas.microsoft.com/office/word/2012/wordml'))).issues.some(i => i.code === 'w15-namespaces'));
});
test('a SmartArt data part\'s presentation links name its layout part, or PowerPoint 2016 draws nothing', async () => {
  const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram', id = 'urn:example/layout/process#1f2e3d4c';
  const frame = `<p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" r:dm="d" r:lo="l" r:qs="d" r:cs="d"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>`;
  const parts = (presId) => ({
    'slide.xml': slide(frame), '_rels/slide.xml.rels': rels(rel('d', 'data.xml', 'diagramData') + rel('l', 'layout.xml', 'diagramLayout')),
    'data.xml': `<dgm:dataModel xmlns:dgm="${DGM}"><dgm:ptLst><dgm:pt modelId="1" type="doc"><dgm:prSet loTypeId="${id}"/></dgm:pt><dgm:pt modelId="2" type="pres"/></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="3" type="presOf" srcId="1" destId="2" presId="${presId}"/></dgm:cxnLst></dgm:dataModel>`,
    'layout.xml': `<dgm:layoutDef xmlns:dgm="${DGM}" uniqueId="${id}"/>`,
  });
  assert.equal(run('check', await pkg('dgm-pres-same', parts(id))).status, 'ok');
  assert.ok(run('check', await pkg('dgm-pres-differ', parts('urn:example/layout/process'))).issues.some(i => i.code === 'diagram-pres-id'));
});
test('every branch of a SmartArt choice keeps the same algorithm type, or PowerPoint draws its default list', async () => {
  const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
  const def = (other) => `<dgm:layoutDef xmlns:dgm="${DGM}" uniqueId="urn:example/layout/x"><dgm:layoutNode name="diagram"><dgm:choose name="c"><dgm:if name="i" func="cnt" axis="ch" ptType="node" op="equ" val="3"><dgm:alg type="composite"/></dgm:if><dgm:else name="e"><dgm:alg type="${other}"/></dgm:else></dgm:choose></dgm:layoutNode></dgm:layoutDef>`;
  assert.equal(run('check', await pkg('dgm-same-alg', { 'layout.xml': def('composite') })).status, 'ok');
  assert.ok(run('check', await pkg('dgm-mixed-alg', { 'layout.xml': def('lin') })).issues.some(i => i.code === 'diagram-mixed-algorithm'));
});
test('diagram, chart and OLE part animations target a graphic frame', async () => {
  const target = '<p:timing><p:tnLst><p:par><p:cTn id="1"><p:childTnLst><p:set><p:cBhvr><p:cTn id="2"/><p:tgtEl><p:spTgt spid="4"><p:graphicEl><a:dgm id="{00000000-0000-0000-0000-000000000001}"/></p:graphicEl></p:spTgt></p:tgtEl></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldGraphic spid="4" grpId="0"/></p:bldLst></p:timing>';
  const frame = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr></p:graphicFrame>';
  const group = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Diagram"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr></p:grpSp>';
  assert.equal(run('check', await pkg('part-anim-frame', { 'slide.xml': slide(`<p:cSld><p:spTree>${frame}</p:spTree></p:cSld>${target}`) })).status, 'ok');
  const issues = run('check', await pkg('part-anim-group', { 'slide.xml': slide(`<p:cSld><p:spTree>${group}</p:spTree></p:cSld>${target}`) })).issues;
  assert.equal(issues.filter(i => i.code === 'animation-part-target').length, 2);
});
test('Word repair regressions: duplicate comment references and promoted pre-release threads', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const duplicate = await pkg('comment-reference', { 'word/document.xml': `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:commentReference w:id="0"/><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>` });
  assert(run('check', duplicate).issues.some(i => i.code === 'duplicate-comment-reference-id'));
  const two = await pkg('two-comment-definitions', {
    '_rels/.rels': rels(rel('main', 'word/document.xml', 'officeDocument')),
    'word/_rels/document.xml.rels': rels(rel('comments', 'comments.xml', 'comments')),
    'word/document.xml': `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:commentReference w:id="0"/><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>`,
    'word/comments.xml': `<w:comments xmlns:w="${W}"><w:comment w:id="0"/><w:comment w:id="0"/></w:comments>`,
  });
  assert(!run('check', two).issues.some(i => i.code === 'duplicate-comment-reference-id'));
  const old = await pkg('old-comments', { 'word/commentsExtended.xml': '<commentsEx xmlns="http://schemas.microsoft.com/office/word/2010/11/wordml"/>' });
  const modern = await pkg('modern-comments', { 'word/commentsExtended.xml': '<commentsEx xmlns="http://schemas.microsoft.com/office/word/2012/wordml"/>' });
  assert(run('compare', old, modern).newIssues.some(i => i.code === 'promoted-comment-extension'));
  assert.equal(run('compare', old, old).newIssues.length, 0);
});
test('IDs in mutually exclusive alternatives may repeat, active duplicate IDs may not', async () => {
  const ac = `<mc:AlternateContent xmlns:mc="${MC}" xmlns:new="urn:new"><mc:Choice Requires="new">${picture(2)}</mc:Choice><mc:Fallback>${picture(2)}</mc:Fallback></mc:AlternateContent>`;
  assert.equal(run('check', await pkg('ac', { 'slide.xml': slide(ac) })).status, 'ok');
  const invalid = run('check', await pkg('duplicate', { 'slide.xml': slide(ac + picture(2)) }));
  assert.ok(invalid.issues.some(i => i.code === 'duplicate-shape-id'));
  assert.equal(run('check', await pkg('scoped', { 'one.xml': slide(picture(2)), 'two.xml': slide(picture(2)) })).status, 'ok');
});
test('a Strict format conversion matches existing diagnostics but still detects new duplicates', async () => {
  const original = await pkg('strict-duplicates', { 'slide.xml': slide(picture(2) + picture(2)).replaceAll(P, 'http://purl.oclc.org/ooxml/presentationml/main') });
  const same = await pkg('transitional-duplicates', { 'slide.xml': slide(picture(2) + picture(2)) });
  assert.equal(run('check', same).issues.length, 1);
  assert.equal(run('compare', original, same).newIssues.length, 0);
  const extra = await pkg('transitional-more-duplicates', { 'slide.xml': slide(picture(2) + picture(2) + picture(2)) });
  assert.equal(run('compare', original, extra).newIssues.length, 2);
});
test('a slide layout may not reuse its master\'s ID', async () => {
  const pres = id => `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst></p:presentation>`;
  const master = id => `<p:sldMaster xmlns:p="${P}" xmlns:r="${R}"><p:sldLayoutIdLst><p:sldLayoutId id="${id}" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`;
  const parts = id => ({ 'presentation.xml': pres(), '_rels/presentation.xml.rels': rels(rel('rId1', 'master.xml', 'slideMaster')), 'master.xml': master(id), '_rels/master.xml.rels': rels(rel('rId1', 'layout.xml', 'slideLayout')), 'layout.xml': `<p:sldLayout xmlns:p="${P}"/>` });
  assert.ok(run('check', await pkg('master-layout-dup', parts('2147483648'))).issues.some(i => i.code === 'duplicate-master-layout-id'));
  assert.ok(!run('check', await pkg('master-layout-ok', parts('2147483649'))).issues.some(i => i.code.startsWith('duplicate')));
});

test('a preset shape lists all of its adjustment values or none', async () => {
  const sp = av => `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Callout"/></p:nvSpPr><p:spPr><a:prstGeom prst="wedgeRoundRectCallout"><a:avLst>${av}</a:avLst></a:prstGeom></p:spPr></p:sp>`;
  const gd = (n, v) => `<a:gd name="${n}" fmla="val ${v}"/>`;
  assert.ok(run('check', await pkg('adj-partial', { 'slide.xml': slide(sp(gd('adj1', -62000) + gd('adj2', 48000))) })).issues.some(i => i.code === 'partial-preset-adjust'));
  assert.equal(run('check', await pkg('adj-full', { 'slide.xml': slide(sp(gd('adj1', -62000) + gd('adj2', 48000) + gd('adj3', 16667))) })).status, 'ok');
  assert.equal(run('check', await pkg('adj-none', { 'slide.xml': slide(sp('')) })).status, 'ok');
});

test('a cell\'s text is at most 32,767 characters, as Excel opens it: runs joined, phonetic text not counted', async () => {
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const sst = (...items) => ({ 'xl/sharedStrings.xml': `<sst xmlns="${S}">${items.map((i) => `<si>${i}</si>`).join('')}</sst>` });
  const t = (n, ch = 'a') => `<t>${ch.repeat(n)}</t>`;
  const codes = async (name, parts) => run('check', await pkg(name, parts)).issues.map((i) => i.code);
  assert.ok((await codes('cell-text-over', sst(t(32768)))).includes('cell-text-too-long'));
  assert.ok((await codes('cell-text-runs', sst(`<r>${t(20000)}</r><r>${t(12768)}</r>`))).includes('cell-text-too-long'));
  assert.ok((await codes('cell-text-cached', { 'xl/worksheets/sheet1.xml': `<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1" t="str"><f>REPT("a",32768)</f><v>${'a'.repeat(32768)}</v></c></row></sheetData></worksheet>` })).includes('cell-text-too-long'));
  assert.ok(!(await codes('cell-text-limit', sst(t(32767), t(32766) + '<rPh sb="0" eb="1"><t>ア</t></rPh>', t(32763) + '<t>_x000D_</t>'))).includes('cell-text-too-long'));
});

test('what Excel refuses in a converted workbook: an empty external cache set, a built-in short of operands, GET.CELL in a plain workbook\'s name', async () => {
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const codes = async (name, parts) => run('check', await pkg(name, parts)).issues.map((i) => i.code);
  const link = (cache) => ({ 'xl/externalLinks/externalLink1.xml': `<externalLink xmlns="${S}"><externalBook><sheetNames><sheetName val="Data"/></sheetNames>${cache}</externalBook></externalLink>` });
  assert.ok((await codes('ext-empty', link('<sheetDataSet/>'))).includes('external-empty-sheet-data-set'));
  assert.ok(!(await codes('ext-cache', link('<sheetDataSet><sheetData sheetId="0"/></sheetDataSet>'))).includes('external-empty-sheet-data-set'));
  assert.ok(!(await codes('ext-none', link(''))).includes('external-empty-sheet-data-set'));
  const sheet = (f) => ({ 'xl/worksheets/sheet1.xml': `<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1"><f>${f}</f><v>0</v></c></row></sheetData></worksheet>` });
  for (const f of ['sumif(D9:D42)', '_xlfn.SUMIF(D9:D42)', 'SUM(1,_xll.SUMIF(D9:D42))'])
    assert.ok((await codes('short-' + f.length, sheet(f))).includes('function-missing-operands'), f);
  for (const f of ['SUMIF(D9:D42,"a,b")', 'ABS("1")', 'SUMIF(D9:D42,)', 'CUSTOMFN(D9:D42)', 'SUM((A1,B1))', 'SUM(T[[#This Row],[a,b]])'])
    assert.ok(!(await codes('enough-' + f.length, sheet(f))).includes('function-missing-operands'), f);
  const book = `<workbook xmlns="${S}"><definedNames><definedName name="Zellfarbe">GET.CELL(63,'Ressourcen'!$2:$2)</definedName></definedNames></workbook>`;
  const types = (type) => '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.${type}.main+xml"/></Types>`;
  assert.ok((await codes('xlm-plain', { '[Content_Types].xml': types('sheet').replace('ms-excel.sheet', 'openxmlformats-officedocument.spreadsheetml.sheet'), 'xl/workbook.xml': book })).includes('xlm-name-in-plain-workbook'));
  assert.ok(!(await codes('xlm-macro', { '[Content_Types].xml': types('sheet.macroEnabled'), 'xl/workbook.xml': book })).includes('xlm-name-in-plain-workbook'));
});

test('what PowerPoint refuses in a converted presentation: unsigned font bytes, huge transform extents, unknown numbering schemes, an overlong editing time', async () => {
  const codes = async (name, parts) => run('check', await pkg(name, parts)).issues.map((i) => i.code);
  const run1 = (rPr) => ({ 'ppt/slides/slide1.xml': slide(`<p:sp><p:txBody><a:p><a:pPr>${rPr.pPr || ''}</a:pPr><a:r><a:rPr>${rPr.r || ''}</a:rPr><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>`) });
  for (const [attr, bad] of [['charset', 128], ['pitchFamily', 192]]) {
    assert.ok((await codes('font-' + attr, run1({ r: `<a:latin typeface="MS Gothic" ${attr}="${bad}"/>` }))).includes('drawingml-font-byte'), attr);
    for (const ok of [127, -128, -64]) assert.ok(!(await codes('font-ok-' + attr + ok, run1({ r: `<a:ea typeface="MS Gothic" ${attr}="${ok}"/>` }))).includes('drawingml-font-byte'), attr + ok);
  }
  const xfrm = (cx) => ({ 'ppt/slides/slide1.xml': slide(`<p:grpSp><p:grpSpPr><a:xfrm><a:off x="-5" y="0"/><a:ext cx="${cx}" cy="10"/><a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="10"/></a:xfrm></p:grpSpPr></p:grpSp>`) });
  assert.ok((await codes('extent-over', xfrm(2147483648))).includes('drawingml-transform-extent'));
  assert.ok(!(await codes('extent-max', xfrm(2147483647))).includes('drawingml-transform-extent'), 'the largest, and a negative offset, open');
  assert.ok((await codes('numbering-poi', run1({ pPr: '<a:buAutoNum type="arabicParenRight"/>' }))).includes('drawingml-numbering-scheme'));
  assert.ok(!(await codes('numbering-ok', run1({ pPr: '<a:buAutoNum type="arabicParenR"/>' }))).includes('drawingml-numbering-scheme'));
  const app = (t) => ({ 'docProps/app.xml': `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><TotalTime>${t}</TotalTime></Properties>` });
  assert.ok((await codes('total-time-over', app(26185585450))).includes('extended-property-total-time'));
  assert.ok(!(await codes('total-time-ok', app(2147483647))).includes('extended-property-total-time'));
});

test('WordArt fill belongs to textFill/textOutline and insertion encloses deletion', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const doc = content => `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body><w:p>${content}</w:p></w:body></w:document>`;
  const bad = await pkg('word-conventions-bad', { 'word/document.xml': doc('<w:del><w:ins><w:r><w:rPr><w14:solidFill/></w:rPr><w:delText>edited</w:delText></w:r></w:ins></w:del>') });
  assert.deepEqual(run('check', bad).issues.map(i => i.code).sort(), ['inserted-inside-deleted', 'wordart-fill-outside-property']);
  const good = await pkg('word-conventions-good', { 'word/document.xml': doc('<w:ins><w:del><w:r><w:rPr><w14:textFill><w14:solidFill/></w14:textFill><w14:textOutline><w14:noFill/></w14:textOutline></w:rPr><w:delText>edited</w:delText></w:r></w:del></w:ins>') });
  assert.equal(run('check', good).status, 'ok');
});

test('timing and connector targets must identify existing shapes and timing nodes', async () => {
  const file = await pkg('timing', { 'slide.xml': slide(picture(2) + '<p:timing><p:cTn id="1"/><p:spTgt spid="3"/><p:tn val="4"/><a:stCxn id="5"/></p:timing>') });
  const codes = run('check', file).issues.map(i => i.code);
  assert.equal(codes.filter(c => c === 'unresolved-shape-id').length, 2);
  assert.ok(codes.includes('unresolved-timing-id'));
});
test('VML spid defines the VML shape rather than referring to a DrawingML shape', async () => {
  const file = await pkg('vml', { 'drawing.xml': '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="Watermark" o:spid="_x0000_s2050"/>' });
  assert.equal(run('check', file).status, 'ok');
});
test('content types and pivot cache identities are checked', async () => {
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const file = await pkg('pivot', { 'xl/workbook.xml': `<workbook xmlns="${ns}"><pivotCaches><pivotCache cacheId="1"/></pivotCaches></workbook>`, 'xl/pivot.xml': `<pivotTableDefinition xmlns="${ns}" cacheId="2"/>`, 'raw.bin': new Uint8Array([0, 1]) });
  assert.deepEqual(run('check', file).issues.map(i => i.code).sort(), ['missing-content-type', 'unresolved-cache-id']);
});
test('opaque bytes cannot change unnoticed; QName-aware merged XML tolerates prefix changes', async () => {
  const a = await pkg('cmp-a', { 'a.xml': `<x xmlns:mc="${MC}" xmlns:v="urn:vendor" mc:Ignorable="v"><v:z v:id="1"/></x>` });
  const b = await pkg('cmp-b', { 'a.xml': `<x xmlns:mc="${MC}" xmlns:t="urn:vendor" mc:Ignorable="t"><t:z t:id="1"/></x>` });
  assert.equal(run('compare', a, b).status, 'failed');
  const policy = path.join(dir, 'merged.json'); fs.writeFileSync(policy, JSON.stringify({ parts: { 'a.xml': { mode: 'merged' } } }));
  assert.equal(run('compare', a, b, '--policy', policy).status, 'ok');
  const changed = await pkg('cmp-c', { 'a.xml': `<x xmlns:mc="${MC}" xmlns:t="urn:different" mc:Ignorable="t"><t:z t:id="1"/></x>` });
  assert.equal(run('compare', a, changed, '--policy', policy).status, 'failed');
});
test('reference comparison detects a valid rId redirected to the wrong existing part', async () => {
  const contents = { 'a.xml': `<x xmlns:r="${R}" r:id="r1"/>`, 'one.xml': '<x/>', 'two.xml': '<x/>' };
  const a = await pkg('target-a', { ...contents, '_rels/a.xml.rels': rels(rel('r1', 'one.xml')) });
  const b = await pkg('target-b', { ...contents, '_rels/a.xml.rels': rels(rel('r1', 'two.xml')) });
  const policy = path.join(dir, 'targets.json'); fs.writeFileSync(policy, JSON.stringify({ parts: { 'a.xml': { mode: 'merged' }, '_rels/a.xml.rels': { mode: 'merged' } } }));
  assert.equal(run('check', b).status, 'ok');
  assert.ok(run('compare', a, b, '--policy', policy).differences.some(d => d.part === 'a.xml'));
  const swapA = await pkg('swap-a', { ...contents, '_rels/a.xml.rels': rels(rel('r1', 'one.xml') + rel('r2', 'two.xml')) });
  const swapB = await pkg('swap-b', { ...contents, '_rels/a.xml.rels': rels(rel('r1', 'two.xml') + rel('r2', 'one.xml')) });
  assert.ok(run('compare', swapA, swapB, '--policy', policy).differences.some(d => d.part === 'a.xml' && d.what === 'relationship-retargeted'));
});
test('regenerated text compares case, punctuation and order, while allowing a run split', async () => {
  const make = text => `<a:txBody xmlns:a="${A}"><a:p>${text}</a:p></a:txBody>`;
  const a = await pkg('text-a', { 'a.xml': make('<a:r><a:t>Hello, world.</a:t></a:r>') });
  const b = await pkg('text-b', { 'a.xml': make('<a:r><a:t>Hello, </a:t></a:r><a:r><a:t>world.</a:t></a:r>') });
  const c = await pkg('text-c', { 'a.xml': make('<a:r><a:t>world hello</a:t></a:r>') });
  const policy = path.join(dir, 'text.json'); fs.writeFileSync(policy, JSON.stringify({ parts: { 'a.xml': { mode: 'regenerated' } } }));
  assert.equal(run('compare', a, b, '--policy', policy).status, 'ok');
  assert.equal(run('compare', a, c, '--policy', policy).status, 'failed');
});
