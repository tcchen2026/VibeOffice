// A block control spans two paragraphs; only its first boundary is copied.
// node --test tests/quire/duplicate-controls.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const root = fileURLToPath(new URL('../../', import.meta.url));
const output = path.resolve(root, process.env.DUPLICATE_CONTROLS_OUT || 'tmp/tests/duplicate-controls');
let page, stop;
before(async () => {
  fs.mkdirSync(output, { recursive: true });
  stop = await ensureServer();
  page = await openPage('quire', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

async function duplicate(selection) {
  const { D, preserve: P, opc: K } = L, N = K.NS;
  const part = 'word/document.xml';
  const content = '<w:sdt><w:sdtPr><w:id w:val="101"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'
    + '<w:p><w:sdt><w:sdtPr><w:id w:val="102"/><w:alias w:val="Title"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Contents</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    + '<w:p><w:r><w:t>Entry</w:t></w:r></w:p></w:sdtContent></w:sdt>';
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="${part}"/></Relationships>` },
    { name: part, data: `<w:document xmlns:w="${N.w}"><w:body>${content}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>` },
  ]);
  const encode = async blob => {
    let text = '';
    const bytes = new Uint8Array(await blob.arrayBuffer());
    for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(text);
  };
  const { doc, warnings } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  D.doc = doc; D.resetHistory();
  const scan = value => [...P.controls(value, false).records.values()].map(r => ({ key: r.control.key, complete: r.complete, type: r.control.type, text: r.text }));
  const originalModel = scan(doc), first = doc.main.blocks[0], last = doc.main.blocks[1];
  const selected = selection === 'whole' ? [first, last] : selection === 'end' ? [last] : [first];
  const result = { warnings, original: await encode(original), before: originalModel,
    sourceBoundary: { start: first.pPr.sdts?.map(c => c.key), end: last.pPr.sdte },
    selectedText: selected.map(p => D.ptext(p)) };
  let copy;
  try {
    D.tx('Duplicate paragraphs', () => {
      D.touchList(doc.main);
      copy = D.cloneBlocks(selected);
      doc.main.blocks.push(...copy);
    });
  } catch (error) { return { ...result, error: error.stack }; }
  result.copied = scan({ main: { blocks: copy }, hf: {}, fn: {}, en: {}, comments: {} });
  result.copiedText = copy.map(p => D.ptext(p));
  result.after = scan(doc);
  const saved = await L.docx.write(doc);
  result.saved = await encode(saved);
  const pkg = await K.open(new Uint8Array(await saved.arrayBuffer()));
  result.savedIds = [...pkg.xml(pkg.main).getElementsByTagNameNS(N.w, 'sdtPr')].map(pr =>
    [...pr.children].find(c => c.localName === 'id')?.getAttributeNS(N.w, 'val'));
  const reread = await L.docx.read(new Uint8Array(await saved.arrayBuffer()));
  result.reopened = scan(reread.doc);
  result.reopenWarnings = reread.warnings;
  result.losses = doc.losses || [];
  return result;
}

for (const [selection, copiedCount] of [['start', 1], ['whole', 2], ['end', 0]]) {
  test(`duplicate ${selection} of a block content control: strip incomplete boundaries, renew complete controls`, async () => {
    const result = await page.evaluate(`(${duplicate.toString()})(${JSON.stringify(selection)})`);
    fs.writeFileSync(path.join(output, selection + '-original.docx'), Buffer.from(result.original, 'base64'));
    assert.deepEqual(page.errors, [], 'no browser errors or CSP refusals');
    assert.deepEqual(page.requests, [], 'no external requests');
    assert.deepEqual(page.windows, [], 'no additional windows');
    assert.deepEqual(page.navigations, [], 'no unexpected navigation');
    assert.deepEqual(result.warnings, []);
    assert.equal(result.before.length, 2);
    assert.ok(result.before.every(c => c.complete), 'the source controls are complete');
    assert.deepEqual(result.sourceBoundary.start, result.sourceBoundary.end, 'the block control spans both paragraphs');
    assert.equal(result.error, undefined, 'duplicating the paragraph must not throw: ' + result.error);
    assert.deepEqual(result.copiedText, result.selectedText, 'copy retains all selected text');
    assert.equal(result.copied.length, copiedCount);
    assert.ok(result.copied.every(c => c.complete), 'no incomplete copied control survives');
    assert.ok(result.copied.every(c => !result.before.some(b => b.key === c.key)), 'copies get fresh model keys');
    assert.equal(result.after.length, 2 + copiedCount);
    assert.ok(result.after.every(c => c.complete));
    assert.deepEqual(result.after.slice(0, 2), result.before, 'the source control is unchanged');
    assert.equal(result.savedIds.length, 2 + copiedCount);
    assert.ok(result.savedIds.every(id => id !== undefined && id !== null));
    assert.equal(new Set(result.savedIds).size, result.savedIds.length, 'saved w:id values are distinct');
    assert.ok(result.savedIds.includes('101') && result.savedIds.includes('102'), 'source Word ids survive');
    assert.equal(result.reopened.length, result.after.length);
    assert.ok(result.reopened.every(c => c.complete), 'saved boundaries reopen complete');
    assert.deepEqual(result.reopenWarnings, []);
    assert.deepEqual(result.losses, [], 'a deliberate partial copy leaves its source intact');
    const savedPath = path.join(output, selection + '-saved.docx');
    fs.writeFileSync(savedPath, Buffer.from(result.saved, 'base64'));
    const check = spawnSync('python3', ['-I', path.join(root, 'tools/ooxml/package.py'), 'check', savedPath], { encoding: 'utf8' });
    assert.equal(check.status, 0, 'independent OOXML package validation: ' + check.stdout + check.stderr);
    assert.deepEqual(JSON.parse(check.stdout).issues, []);
  });
}
