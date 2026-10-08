// Query field edits and sheet copying through actual suite draft/recovery hooks.
import fs from 'node:fs';
import path from 'node:path';
import { openPage, ensureServer } from '../../ooxml/cdp.mjs';
const [input, output] = process.argv.slice(2);
if (!input || !output) throw Error('Usage: tables-ui.mjs INPUT.xlsx OUTPUT');
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('ledger', { persistence: false });
try {
  const result = await page.evaluate(`(${run.toString()})(${JSON.stringify(fs.readFileSync(input).toString('base64'))})`);
  for (const name of ['draft', 'recovered']) { fs.writeFileSync(path.join(output, name + '.xlsx'), Buffer.from(result[name], 'base64')); delete result[name]; }
  result.consoleErrors = page.errors; result.requests = page.requests;
  if (page.errors.length) result.status = 'failed';
  fs.writeFileSync(path.join(output, 'browser.json'), JSON.stringify(result, null, 2) + '\n');
  const shot = await page.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(output, 'query-draft.png'), Buffer.from(shot.data, 'base64'));
  console.log(JSON.stringify(result)); if (result.status !== 'ok') process.exitCode = 1;
} finally { await page.close(); stop(); }
async function run(data) {
  const h = __corpusHooks; VO.opened = () => {};
  await h.open(new File([Uint8Array.from(atob(data), c => c.charCodeAt(0))], 'queries.xlsx'));
  let doc = h.current(), wb = doc.wb, table = wb.tables.find(t => t.queries?.length);
  if (!table) throw Error('Input needs a query table');
  const sh = table.sheet, count = table.columns.length;
  L.ops.insertLines(sh, 'c', table.ref.c1 + 1, 1); wb.undo.undo(); wb.undo.redo();
  L.ops.copySheet(wb, sh, wb.sheets.length); wb.undo.undo(); wb.undo.redo();
  const draft = await h.snapshot(doc), info = h.docInfo(doc);
  await h.open(new File([draft], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: structuredClone(info.lossState) });
  doc = h.current(); wb = doc.wb;
  const tables = wb.tables.filter(t => t.queries?.length);
  if (tables.length !== 2 || tables.some(t => t.columns.length !== count + 1) || new Set(tables.map(t => t.queries[0].part)).size !== 2) throw Error('Draft lost query field ownership');
  for (const t of tables) {
    const query = wb.pkg.xml(t.queries[0].part), fields = query.getElementsByTagName('*').filter(e => e.localName === 'queryTableField');
    for (const col of t.columns) if (!fields.some(f => +f.getAttribute('id') === col.queryField && +f.getAttribute('tableColumnId') === col.id)) throw Error('Draft has unmatched query field IDs');
  }
  const encode = async blob => {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let s = '';
    for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(s);
  };
  return { status: 'ok', tables: tables.length, fields: tables.reduce((n, t) => n + t.columns.length, 0), draft: await encode(draft), recovered: await encode(await L.xlsxWrite.write(wb)) };
}
