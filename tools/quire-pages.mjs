// Quire's page count for every Word-saved .docx of a corpus, next to the count Word stored (docProps/app.xml <Pages>,
// files whose Application is Microsoft Office Word): run it before and after a layout change and compare.
//
//   node tools/quire-pages.mjs [SOURCE] OUT.jsonl [--jobs 4] [--list files.txt] [--limit N]
//   node tools/quire-pages.mjs --compare BEFORE.jsonl AFTER.jsonl [--printed tests/quire/word2016-pages.json] [--all]
//
// SOURCE defaults to /work/corpora/word. Each line of OUT: {file, word, quire, compat, ms} (quire is null and error set
// when the file didn't open). BASE (default http://127.0.0.1:8760) is the copy of the site to run, so an older one
// served by tools/serve.py PORT FOLDER gives the "before"; CDP picks the Chromium (tools/ooxml/cdp.mjs). Word's stored
// count is a coarse reference (errors can cancel out, some producers never update it): --compare lists every file
// that moved away from it, to be looked at, rather than a score to chase. --printed takes the pages Word 2016 printed
// for the files listed there instead (some stored counts are stale).
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { ensureServer, openPage } from './ooxml/cdp.mjs';

const args = process.argv.slice(2);
const opt = (name, def) => { const i = args.indexOf(name); return i < 0 ? def : args.splice(i, 2)[1]; };

if (args[0] === '--compare') {
  /* --printed FILE: pages Word printed for some files (tests/quire/word2016-pages.json), in place of the stored count */
  const printedFile = opt('--printed', null), printed = printedFile ? JSON.parse(fs.readFileSync(printedFile, 'utf8')).pages : {};
  const read = (f) => new Map(fs.readFileSync(f, 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l)).map((r) => [r.file, printed[r.file] ? Object.assign(r, { word: printed[r.file] }) : r]));
  const a = read(args[1]), b = read(args[2]);
  const both = [...a.keys()].filter((f) => b.has(f) && a.get(f).quire != null && b.get(f).quire != null);
  const exact = (m) => both.filter((f) => m.get(f).quire === m.get(f).word).length;
  const moved = both.filter((f) => a.get(f).quire !== b.get(f).quire);
  const away = moved.filter((f) => Math.abs(b.get(f).quire - b.get(f).word) > Math.abs(a.get(f).quire - a.get(f).word));
  const toward = moved.filter((f) => !away.includes(f));
  console.log(`${both.length} files counted both times. Word's count: ${exact(a)} before, ${exact(b)} after. ${moved.length} changed: ${toward.length} toward Word's count or level, ${away.length} away from it.`);
  for (const f of away) console.log(`  away: ${f}  Word ${b.get(f).word}, Quire ${a.get(f).quire} -> ${b.get(f).quire}`);
  if (args.includes('--all')) for (const f of toward) console.log(`  toward: ${f}  Word ${b.get(f).word}, Quire ${a.get(f).quire} -> ${b.get(f).quire}`);
  process.exit(0);
}

const jobs = +opt('--jobs', 4), limit = +opt('--limit', 0), list = opt('--list', null);
const [source, out] = args.length > 1 ? args : ['/work/corpora/word', args[0]];
if (!out) { console.error('usage: node tools/quire-pages.mjs [SOURCE] OUT.jsonl [--jobs 4] [--list files.txt] [--limit N]'); process.exit(2); }
let names = list ? fs.readFileSync(list, 'utf8').split('\n').map((s) => s.trim()).filter(Boolean) : fs.readdirSync(source).filter((f) => /\.(docx|docm)$/i.test(f)).sort();
const wordSaved = [];
for (const f of names) {
  let app = '';
  try { app = execFileSync('unzip', ['-p', path.join(source, f), 'docProps/app.xml'], { stdio: ['ignore', 'pipe', 'ignore'] }).toString(); } catch { continue; }
  const pages = /<Pages>(\d+)<\/Pages>/.exec(app);
  if (/Microsoft Office Word/.test(app) && pages) wordSaved.push({ file: f, word: +pages[1] });
}
names = limit ? wordSaved.slice(0, limit) : wordSaved;
console.error(`${names.length} Word-saved files with a stored page count`);

const stop = await ensureServer();
fs.writeFileSync(out, '');
let next = 0, done = 0;
async function worker() {
  let page = null, used = 0;
  while (next < names.length) {
    const it = names[next++];
    if (!page || used++ >= 25) { if (page) await page.close().catch(() => {}); page = await openPage('quire', { persistence: false, automated: true }); used = 1; }
    const t0 = Date.now();
    let rec;
    try {
      const b64 = fs.readFileSync(path.join(source, it.file)).toString('base64');
      const r = await Promise.race([page.evaluate(`(async () => {
        VO.opened = () => {}; L.layout.view = 'print'; L.layout.setZoom(1);
        const bytes = Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0));
        const { doc } = await L.docx.read(bytes.buffer);
        L.app.loadDoc(doc, 'doc', { saved: true });
        await document.fonts.ready; L.layout.render();
        let last = -1, same = 0;
        for (let k = 0; k < 120 && same < 3; k++) { await new Promise((r) => setTimeout(r, 250)); const n = L.layout.pages.length; same = n === last ? same + 1 : 0; last = n; }
        return { pages: L.layout.pages.length, compat: (L.D.doc.settings || {}).compat || null };
      })()`), new Promise((_, no) => setTimeout(() => no(new Error('timeout')), 90000))]);
      rec = { file: it.file, word: it.word, quire: r.pages, compat: r.compat, ms: Date.now() - t0 };
    } catch (e) {
      rec = { file: it.file, word: it.word, quire: null, error: String(e.message || e).slice(0, 200), ms: Date.now() - t0 };
      await page.close().catch(() => {}); page = null;
    }
    fs.appendFileSync(out, JSON.stringify(rec) + '\n');
    if (++done % 100 === 0) console.error(`${done}/${names.length}`);
  }
  if (page) await page.close().catch(() => {});
}
await Promise.all(Array.from({ length: jobs }, worker));
stop();
const recs = fs.readFileSync(out, 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l));
const ok = recs.filter((r) => r.quire != null);
console.log(`${ok.length} counted (${recs.length - ok.length} failed): ${ok.filter((r) => r.quire === r.word).length} show Word's stored page count`);
process.exit(0);
