// Actual object clipboard transfer, history and suite draft recovery.
import fs from 'node:fs';
import path from 'node:path';
import { openPage, ensureServer } from '../../ooxml/cdp.mjs';
const [source, output] = process.argv.slice(2);
if (!source || !output) throw Error('Usage: objects-ui.mjs INPUT.xlsx OUTPUT');
fs.mkdirSync(output, { recursive: true });
const data = fs.readFileSync(source).toString('base64'), stop = await ensureServer();
const page = await openPage('ledger', { persistence: false });
try {
  const result = await page.evaluate(`(${run.toString()})(${JSON.stringify(data)})`);
  for (const kind of ['draft', 'recovered']) {
    fs.writeFileSync(path.join(output, kind + '.xlsx'), Buffer.from(result[kind], 'base64')); delete result[kind];
  }
  result.consoleErrors = page.errors; result.requests = page.requests;
  if (page.errors.length) result.status = 'failed';
  fs.writeFileSync(path.join(output, 'browser.json'), JSON.stringify(result, null, 2) + '\n');
  const shot = await page.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(output, 'paste-draft.png'), Buffer.from(shot.data, 'base64'));
  console.log(JSON.stringify(result));
  if (result.status !== 'ok') process.exitCode = 1;
} finally { await page.close(); stop(); }

async function run(data) {
  const h = __corpusHooks; VO.opened = () => {};
  await h.open(new File([Uint8Array.from(atob(data), c => c.charCodeAt(0))], 'objects.xlsx'));
  const source = h.current().wb, object = source.sheets[0].drawings.find(d => d.objectKeep);
  if (!object) throw Error('Input needs a preserved object on its first sheet');
  L.drawing.copy(object);
  const blank = new L.model.Workbook(); blank.addSheet('Pasted objects');
  await h.open(new File([await L.xlsxWrite.write(blank)], 'pasted.xlsx'));
  let doc = h.current(), wb = doc.wb;
  L.grid.selectRange({ r1: 2, c1: 1, r2: 2, c2: 1 }); L.drawing.pasteObject();
  if (wb.sheets[0].drawings.length !== 1) throw Error('Object paste failed');
  wb.undo.undo(); if (wb.sheets[0].drawings.length) throw Error('Object paste undo failed');
  wb.undo.redo();
  const draft = await h.snapshot(doc), info = h.docInfo(doc);
  await h.open(new File([draft], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: structuredClone(info.lossState) });
  doc = h.current(); wb = doc.wb;
  const copy = wb.sheets[0].drawings.find(d => d.objectKeep);
  if (!copy?.objectKeep.vml || !copy.objectKeep.groups.length || copy.anchor.from.r !== 2) throw Error('Draft lost the object or its position');
  const saved = await L.xlsxWrite.write(wb), pkg = await L.opc.open(await saved.arrayBuffer());
  const originalPayloads = source.pkg.names.filter(n => /\/(ctrlProps|activeX|embeddings)\//.test(n) && !n.endsWith('.rels'));
  const retained = originalPayloads.filter(n => pkg.names.some(p => {
    const a = source.pkg.bytes(n), b = pkg.bytes(p); return a.length === b.length && a.every((v, i) => b[i] === v);
  })).length;
  if (!retained) throw Error('Pasted payload bytes were not retained');
  const encode = async blob => {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(text);
  };
  return { status: 'ok', objects: wb.sheets[0].drawings.filter(d => d.objectKeep).length, retainedPayloads: retained,
    draft: await encode(draft), recovered: await encode(saved) };
}
