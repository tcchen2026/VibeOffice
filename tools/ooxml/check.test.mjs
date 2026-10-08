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
  const r = spawnSync('python3', [fileURLToPath(new URL('./package.py', import.meta.url)), command, ...args], { encoding: 'utf8' });
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
