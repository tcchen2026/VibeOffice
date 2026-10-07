// Markdown round trips in Quire, in the Chromium that tools/shot.mjs drives (CDP 127.0.0.1:9222).
//
//   node tools/quire/test/mdroundtrip.mjs DIR          every .md file in DIR (not checked in: download READMEs)
//   node tools/quire/test/mdroundtrip.mjs --spec FILE  the examples of the CommonMark spec.json
//
// Per file: open → save unchanged (must be byte for byte the same), edit one paragraph → save (how many
// lines changed, against the edited block's own lines), and every block rewritten → save, compared with
// the original as rendered HTML (raw HTML in the file is expected to differ: it becomes formatting).
// Per spec example: only the rewritten comparison.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execFileSync } from 'node:child_process';

const here = path.dirname(new URL(import.meta.url).pathname);
const shot = path.join(here, '../../shot.mjs');
const args = process.argv.slice(2);
const spec = args[0] === '--spec';
const src = spec ? args[1] : args[0];
if (!src) { console.error('usage: mdroundtrip.mjs DIR | --spec spec.json'); process.exit(2); }

const files = spec
  ? JSON.parse(fs.readFileSync(src, 'utf8')).map((e) => [`${e.example} ${e.section}`, e.markdown])
  : fs.readdirSync(src).filter((f) => /\.(md|markdown)$/i.test(f)).sort().map((f) => [f, fs.readFileSync(path.join(src, f), 'utf8')]);

/* runs in the page; results stay in window.__res and come back in pieces (the tool prints at most ~50 kB) */
const run = `
const norm = (h) => h.replace(/\\s+/g, ' ').replace(/> </g, '><').replace(/ ?(<\\/?(p|li|ul|ol|blockquote|pre|h\\d|table|thead|tbody|tr|th|td|hr)[^>]*>) ?/g, '$1').trim();
const span = (a, b) => { const x = a.split('\\n'), y = b.split('\\n'); let i = 0; while (i < x.length && i < y.length && x[i] === y[i]) i++; let j = 0; while (j < x.length - i && j < y.length - i && x[x.length - 1 - j] === y[y.length - 1 - j]) j++; return Math.max(x.length - i - j, y.length - i - j); };
window.__res = [];
for (const [name, text] of window.__files) {
  const r = { name };
  try {
    const d = await L.mdio.read(text);
    if (!${spec}) {
      r.same = L.mdio.write(d).text === text;
      const paras = d.main.blocks.filter((b) => b.t === 'p' && b.md && !b.pPr.num && !b.pPr.style && b.runs.some((x) => x.t === 'text' && x.text.length > 20));
      if (paras.length) {
        const p = paras[Math.floor(paras.length / 2)];
        const run = p.runs.find((x) => x.t === 'text' && x.text.length > 20);
        run.text += ' EDITED';
        r.edit = span(text, L.mdio.write(d).text);
        r.block = d.md.groups[p.md.g].body.length;
        run.text = run.text.slice(0, -7);
      }
    }
    const m = d.md; d.md = null;
    const all = L.mdio.write(d).text; d.md = m;
    const body = m && m.front ? text.split('\\n').slice(m.front.length).join('\\n') : text; /* front matter is not rendered */
    r.html = norm(L.md.toHTML(L.md.parse(body))) === norm(L.md.toHTML(L.md.parse(all)));
    r.rawHtml = /<[a-z!/]/i.test(text);
  } catch (e) { r.err = String(e && e.message || e); }
  window.__res.push(r);
}
return window.__res.length;`;
const steps = [{ wait: 1500, js: `window.__files = ${JSON.stringify(files)}; return 1` }, { js: run }];
for (let i = 0; i < files.length; i += 200) steps.push({ js: `return JSON.stringify(window.__res.slice(${i}, ${i + 200}))` });
const tmp = path.join(os.tmpdir(), `mdroundtrip-${process.pid}.json`);
fs.writeFileSync(tmp, JSON.stringify(steps));
const out = execFileSync('node', [shot, 'quire/', tmp], { encoding: 'utf8', maxBuffer: 1 << 28 });
fs.unlinkSync(tmp);
const res = out.split('\n').filter((l) => l.startsWith('js -> "[')).flatMap((l) => JSON.parse(JSON.parse(l.slice(6))));

const errs = res.filter((r) => r.err);
for (const r of errs) console.log(`ERROR ${r.name}: ${r.err}`);
if (spec) {
  const ok = res.filter((r) => r.html).length;
  const noHtml = res.filter((r) => !r.rawHtml);
  console.log(`rewritten and still the same HTML: ${ok} of ${res.length} (without raw HTML: ${noHtml.filter((r) => r.html).length} of ${noHtml.length})`);
  console.log('differ:', res.filter((r) => !r.html).map((r) => r.name.split(' ')[0]).join(' '));
} else {
  for (const r of res) if (!r.err) console.log(`${r.same ? 'same' : 'DIFF'}  edit ${String(r.edit ?? '-').padStart(3)} / ${String(r.block ?? '-').padEnd(3)} rewritten ${r.html ? 'same HTML' : r.rawHtml ? 'differs (raw HTML)' : 'DIFFERS'}  ${r.name}`);
  const n = res.length - errs.length;
  console.log(`\n${res.filter((r) => r.same).length} of ${n} saved unchanged byte for byte; edits confined to the edited block in ${res.filter((r) => r.edit != null && r.edit <= r.block + 1).length} of ${res.filter((r) => r.edit != null).length}; rewritten with the same HTML: ${res.filter((r) => r.html).length} (${res.filter((r) => !r.rawHtml && r.html).length} of ${res.filter((r) => !r.rawHtml).length} without raw HTML)`);
}
