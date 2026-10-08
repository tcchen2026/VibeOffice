// Exercise the actual drawing clipboard, undo and suite drafts for a slicer view.
import fs from 'node:fs';
import path from 'node:path';
import { openPage, ensureServer } from '../../ooxml/cdp.mjs';
const [source, output] = process.argv.slice(2);
if (!source || !output) throw Error('Usage: slicers-ui.mjs INPUT.xlsx OUTPUT');
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
  await h.open(new File([Uint8Array.from(atob(data), c => c.charCodeAt(0))], 'slicers.xlsx'));
  let doc = h.current(), wb = doc.wb;
  const count = () => wb.sheets.flatMap(s => s.drawings).filter(d => d.slicerKeep).length;
  const d = wb.sheets[0].drawings.find(d => d.slicerKeep), before = count();
  if (!d) throw Error('Input needs a slicer on its first sheet');
  L.drawing.copy(d); L.grid.selectRange({ r1: 17, c1: 5, r2: 17, c2: 5 }); L.drawing.pasteObject();
  if (count() !== before + 1) throw Error('Paste failed');
  wb.undo.undo(); if (count() !== before) throw Error('Paste undo failed'); wb.undo.redo();
  const draft = await h.snapshot(doc), info = h.docInfo(doc);
  await h.open(new File([draft], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: structuredClone(info.lossState) });
  doc = h.current(); wb = doc.wb;
  if (count() !== before + 1) throw Error('Draft lost the copied view');
  const names = wb.sheets.flatMap(s => s.drawings.flatMap(d => d.slicerKeep?.names || []));
  if (new Set(names).size !== names.length) throw Error('Draft has colliding view names');
  const saved = await L.xlsxWrite.write(wb);
  const encode = async blob => {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(text);
  };
  return { status: 'ok', views: count(), names, draft: await encode(draft), recovered: await encode(saved) };
}
