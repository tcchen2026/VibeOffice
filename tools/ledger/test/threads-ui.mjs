// Actual clipboard, undo and suite snapshot/recovery hooks. Input may be the
// mentions-no-legacy.xlsx fixture emitted by threads.test.js with THREAD_RESULTS.
import fs from 'node:fs';
import path from 'node:path';
import { openPage, ensureServer } from '../../ooxml/cdp.mjs';
const [source, output] = process.argv.slice(2);
if (!source || !output) throw Error('Usage: threads-ui.mjs INPUT.xlsx OUTPUT');
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
  await h.open(new File([Uint8Array.from(atob(data), c => c.charCodeAt(0))], 'threads.xlsx'));
  let doc = h.current(), wb = doc.wb;
  const first = [...wb.sheets[0].comments.values()].find(c => c.thread);
  if (!first) throw Error('Input needs a thread on its first sheet');
  const before = [...wb.sheets[0].comments.values()].filter(c => c.thread).length;
  L.grid.selectRange({ r1: first.r, c1: first.c, r2: first.r, c2: first.c });
  L.clip.copy({ clipboardData: { setData() {} } }, false);
  L.grid.selectRange({ r1: 20, c1: 10, r2: 20, c2: 10 });
  L.clip.pasteInternal({ what: 'comments' }); L.clip.clearMarquee();
  let comments = [...wb.sheets[0].comments.values()].filter(c => c.thread);
  if (comments.length !== before + 1 || new Set(comments.map(c => c.thread[0].id)).size !== comments.length) throw Error('Paste IDs were not remapped');
  wb.undo.undo(); if ([...wb.sheets[0].comments.values()].filter(c => c.thread).length !== before) throw Error('Paste undo failed');
  wb.undo.redo();
  const draft = await h.snapshot(doc), info = h.docInfo(doc);
  await h.open(new File([draft], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: structuredClone(info.lossState) });
  doc = h.current(); wb = doc.wb; comments = [...wb.sheets[0].comments.values()].filter(c => c.thread);
  const copy = comments.find(c => c.r === 20 && c.c === 10);
  if (comments.length !== before + 1 || copy?.thread.length !== first.thread.length || copy.thread[0].mentions.length !== first.thread[0].mentions.length || copy.thread.some(t => t.parent && !copy.thread.some(p => p.id === t.parent))) throw Error('Draft lost threads, parent links or mentions');
  const saved = await L.xlsxWrite.write(wb);
  const encode = async blob => {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(text);
  };
  return { status: 'ok', threads: comments.length, comments: comments.reduce((n, c) => n + c.thread.length, 0), draft: await encode(draft), recovered: await encode(saved) };
}
