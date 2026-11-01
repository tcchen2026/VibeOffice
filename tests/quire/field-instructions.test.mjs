// Synthetic field instructions through Quire's browser reader and writer.
// node --test tests/quire/field-instructions.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

let page, stop;
before(async () => {
  stop = await ensureServer(process.env.BASE);
  page = await openPage('quire', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

async function roundTrip(kind) {
  const K = L.opc, N = K.NS, part = 'word/document.xml';
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const singleRun = kind === 'single-run';
  const field = type => `<w:fldChar w:fldCharType="${type}"/>`;
  const instruction = text => `<w:instrText xml:space="preserve">${K.esc(text)}</w:instrText>`;
  const drawing = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Conditional logo"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="${pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  // Word puts the outer delimiters, each instruction, the nested PAGE field's
  // delimiters/result, and the picture in separate sibling runs.
  const contents = [field('begin'), instruction(' IF '), field('begin'), instruction(' PAGE '),
    field('separate'), '<w:t>1</w:t>', field('end'),
    ...(singleRun ? [instruction(' = 1 "yes" "no" ')] : [instruction(' = 1 "'), drawing, instruction('" "" ')]),
    field('separate'), '<w:t>cached IF result</w:t>', field('end')];
  const runs = singleRun ? `<w:r>${contents.join('')}</w:r>` : contents.map(c => `<w:r>${c}</w:r>`).join('');
  const relationship = (id, type, target) => `<Relationship Id="${id}" Type="${N.rel}/${type}" Target="${target}"/>`;
  const relationships = content => `<Relationships xmlns="${N.pkg}">${content}</Relationships>`;
  const files = [
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: '_rels/.rels', data: relationships(relationship('rId1', 'officeDocument', part)) },
    { name: part, data: `<w:document xmlns:w="${N.w}" xmlns:r="${N.rel}" xmlns:wp="${wp}" xmlns:a="${N.a}" xmlns:pic="${pic}"><w:body><w:p>${runs}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>` },
  ];
  if (!singleRun) {
    const canvas = document.createElement('canvas');
    canvas.width = canvas.height = 1;
    canvas.getContext('2d').fillRect(0, 0, 1, 1);
    const png = Uint8Array.from(atob(canvas.toDataURL('image/png').split(',')[1]), c => c.charCodeAt(0));
    files.push({ name: 'word/media/logo.png', data: png },
      { name: K.relsPath(part), data: relationships(relationship('rIdImage', 'image', 'media/logo.png')) });
  }
  const original = await L.zip.write(files);
  const { doc, warnings } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  let outer;
  L.D.walk(doc.main, p => {
    if (p.t === 'p') outer ||= p.runs.find(it => it.t === 'fb' && L.D.fieldType(it.instr) === 'IF');
  });
  if (!outer) throw new Error('The synthetic IF field was not read into the model');
  if (kind === 'edited') outer.instr = ' IF 2 = 2 "changed" "unchanged" ';
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));

  const inspect = pkg => {
    const root = pkg.xml(pkg.main);
    const elements = (ns, name) => Array.from(root.getElementsByTagNameNS(ns, name));
    return {
      fieldChars: elements(N.w, 'fldChar').map(e => e.getAttributeNS(N.w, 'fldCharType')),
      instructions: elements(N.w, 'instrText').map(e => e.textContent),
      texts: elements(N.w, 't').map(e => e.textContent),
      drawings: elements(N.w, 'drawing').length,
      images: elements(N.a, 'blip').map(e => {
        const id = e.getAttributeNS(N.rel, 'embed'), rel = pkg.rels(pkg.main).find(r => r.id === id);
        return { id, type: rel?.type, part: rel?.part, external: rel?.external,
          contentType: rel && pkg.type(rel.part), bytes: rel && pkg.has(rel.part) ? Array.from(pkg.bytes(rel.part)) : null };
      }),
    };
  };
  return { before: inspect(doc.pkg), after: inspect(saved), warnings, fieldId: outer.fid,
    code: outer.instr, losses: doc.losses || [], notices: K.noticeLosses(doc) };
}

async function run(kind) {
  const result = await page.evaluate(`(${roundTrip.toString()})(${JSON.stringify(kind)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  assert.deepEqual(page.requests, [], 'no external requests');
  assert.deepEqual(page.windows, [], 'no additional windows');
  assert.deepEqual(page.navigations, [], 'no navigation away from the test server');
  assert.deepEqual(result.warnings, [], 'the synthetic package reads without repairs');
  assert.deepEqual(result.before.fieldChars, ['begin', 'begin', 'separate', 'end', 'separate', 'end']);
  return result;
}

function fieldLoss(result) {
  const losses = result.losses.filter(e => e.id.startsWith('field-code:'));
  assert.equal(losses.length, 1, 'losing nested field instructions must record field-code');
  const loss = losses[0];
  assert.equal(loss.id, 'field-code:word/document.xml:' + result.fieldId);
  assert.equal(loss.action, 'drop');
  assert.equal(loss.where, 'word/document.xml');
  assert.equal(loss.notify, true);
  assert.equal(result.notices.filter(e => e.id === loss.id).length, 1, 'the field-code loss reaches the user');
  return loss;
}

test('unchanged IF instructions keep the nested PAGE field and inline picture with its image part', async () => {
  const result = await run('unchanged');
  assert.deepEqual(result.after.fieldChars, result.before.fieldChars, 'save must keep both field begins and their delimiters');
  assert.deepEqual(result.after.instructions, result.before.instructions);
  assert.deepEqual(result.after.texts, result.before.texts);
  assert.equal(result.before.drawings, 1);
  assert.equal(result.after.drawings, 1, 'save must keep the instruction picture');
  assert.equal(result.after.images.length, 1);
  const image = result.after.images[0];
  assert.equal(image.type, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
  assert.equal(image.part, 'word/media/logo.png');
  assert.equal(image.external, false);
  assert.equal(image.contentType, 'image/png');
  assert.ok(image.bytes?.length, 'the saved picture relationship resolves to image bytes');
  assert.deepEqual(image.bytes, result.before.images[0].bytes);
  assert.deepEqual(result.losses.filter(e => /^content:.*:fields$|^object:|^unreferenced:|^field-code:/.test(e.id)), []);
  assert.deepEqual(result.notices, []);
});

test('editing the IF field code writes plain instructions and reports the lost nested field and picture', async () => {
  const result = await run('edited');
  assert.deepEqual(result.after.fieldChars, ['begin', 'separate', 'end']);
  assert.deepEqual(result.after.instructions, [result.code]);
  assert.equal(result.before.drawings, 1);
  assert.equal(result.after.drawings, 0);
  assert.deepEqual(result.after.images, []);
  assert.match(fieldLoss(result).what, /field code was edited/);
});

test('nested fields in one run report that their instructions cannot be kept', async () => {
  const result = await run('single-run');
  assert.deepEqual(result.after.fieldChars, ['begin', 'separate', 'end']);
  assert.deepEqual(result.after.instructions, [result.code]);
  assert.match(fieldLoss(result).what, /can't be kept/);
});
