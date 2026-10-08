// Real suite draft hooks after extended-format edits, history and a sheet copy.
import fs from 'node:fs';
import path from 'node:path';
import { openPage, ensureServer } from '../../ooxml/cdp.mjs';
const [input, output] = process.argv.slice(2);
if (!input || !output) throw Error('Usage: extensions-ui.mjs INPUT.xlsx OUTPUT');
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('ledger', { persistence: false });
try {
  const result = await page.evaluate(`(${run.toString()})(${JSON.stringify(fs.readFileSync(input).toString('base64'))})`);
  for (const name of ['draft', 'recovered']) { fs.writeFileSync(path.join(output, name + '.xlsx'), Buffer.from(result[name], 'base64')); delete result[name]; }
  result.consoleErrors = page.errors; result.requests = page.requests;
  if (page.errors.length) result.status = 'failed';
  fs.writeFileSync(path.join(output, 'browser.json'), JSON.stringify(result, null, 2) + '\n');
  const shot = await page.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(output, 'bars-draft.png'), Buffer.from(shot.data, 'base64'));
  console.log(JSON.stringify(result)); if (result.status !== 'ok') process.exitCode = 1;
} finally { await page.close(); stop(); }
async function run(data) {
  const h = __corpusHooks; VO.opened = () => {};
  await h.open(new File([Uint8Array.from(atob(data), c => c.charCodeAt(0))], 'extensions.xlsx'));
  let doc = h.current(), wb = doc.wb;
  const sh = wb.sheets.find(sh => sh.cf.some(c => c.rules.some(r => r.bar && r.extKeep)));
  if (!sh) throw Error('Input needs an extended data bar');
  const cf = structuredClone(sh.cf), bar = cf.flatMap(c => c.rules).find(r => r.bar && r.extKeep);
  bar.bar.gradient = false;
  L.ops.tx(wb, 'Bar appearance', () => L.ops.setCF(sh, cf));
  L.ops.insertLines(sh, 'r', 0, 2); wb.undo.undo(); wb.undo.redo();
  L.ops.copySheet(wb, sh, wb.sheets.length); wb.undo.undo(); wb.undo.redo();
  const draft = await h.snapshot(doc), info = h.docInfo(doc);
  await h.open(new File([draft], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: structuredClone(info.lossState) });
  doc = h.current(); wb = doc.wb;
  const bars = wb.sheets.flatMap(sh => sh.cf.flatMap(c => c.rules.filter(r => r.bar && r.extKeep)));
  if (bars.length !== 2 || bars.some(r => r.bar.gradient) || new Set(bars.map(r => r.x14id)).size !== 2) throw Error('Draft lost extended settings or duplicated rule identities');
  const encode = async blob => {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let s = '';
    for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(s);
  };
  return { status: 'ok', bars: bars.length, draft: await encode(draft), recovered: await encode(await L.xlsxWrite.write(wb)) };
}
