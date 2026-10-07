/* Checks Quire's Markdown parser (public/quire/js/markdown.js) against the spec examples.
 *
 *   node tools/quire/test/mdspec.js DIR [--verbose] [--section NAME]
 *
 * DIR holds the downloaded references (not checked in):
 *   spec.json      https://spec.commonmark.org/0.31.2/spec.json        CommonMark, 652 examples
 *   gfm-spec.txt   https://raw.githubusercontent.com/github/cmark-gfm/master/test/spec.txt
 *                  (only its extension examples: tables, task lists, strikethrough, autolinks, disallowed raw HTML)
 *   entities.json  https://html.spec.whatwg.org/entities.json           the HTML5 entity names
 */
'use strict';
const fs = require('fs');
const path = require('path');
global.L = {};
require(path.join(__dirname, '../../../public/quire/js/markdown.js'));
const M = global.L.md;

const dir = process.argv[2];
if (!dir) { console.error('usage: mdspec.js DIR [--verbose] [--section NAME]'); process.exit(2); }
const verbose = process.argv.includes('--verbose');
const si = process.argv.indexOf('--section');
const only = si > 0 ? process.argv[si + 1] : null;

const ents = JSON.parse(fs.readFileSync(path.join(dir, 'entities.json'), 'utf8'));
M.entities = {};
for (const k in ents) if (k.endsWith(';')) M.entities[k.slice(1, -1)] = ents[k].characters;

/* the reference runner compares normalized HTML; exact comparison is stricter, so only line ends are unified */
const norm = (s) => s.replace(/\r\n/g, '\n');

function run(name, examples, opts) {
  let pass = 0;
  const failed = {};
  for (const ex of examples) {
    if (only && ex.section !== only) continue;
    let got;
    try { got = M.toHTML(M.parse(ex.markdown, opts), opts); } catch (e) { got = 'EXCEPTION ' + e.stack; }
    if (norm(got) === norm(ex.html)) pass++;
    else {
      (failed[ex.section] = failed[ex.section] || []).push(ex.example);
      if (verbose) console.log(`--- ${name} example ${ex.example} (${ex.section})\n${JSON.stringify(ex.markdown)}\nwant ${JSON.stringify(ex.html)}\ngot  ${JSON.stringify(got)}`);
    }
  }
  const total = examples.filter((e) => !only || e.section === only).length;
  console.log(`${name}: ${pass} of ${total} examples pass`);
  for (const s in failed) console.log(`  ${s}: ${failed[s].join(', ')}`);
  return { pass, total };
}

const cm = JSON.parse(fs.readFileSync(path.join(dir, 'spec.json'), 'utf8'));
run('CommonMark 0.31.2', cm, { gfm: false, math: false, tagfilter: false });

/* GFM extension examples: ```` example table ... ```` blocks */
const txt = fs.readFileSync(path.join(dir, 'gfm-spec.txt'), 'utf8').replace(/→/g, '\t');
const gfm = [];
const re = /^`{32} example (\w+)\n([\s\S]*?)^\.\n([\s\S]*?)^`{32}$/gm;
let m, n = 0;
while ((m = re.exec(txt))) { n++; gfm.push({ example: n, section: m[1], markdown: m[2], html: m[3] }); }
const ext = gfm.filter((e) => e.section !== 'disabled');
run('GFM extensions', ext, { gfm: true, math: false, tagfilter: true });
