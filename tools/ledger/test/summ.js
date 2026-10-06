/* Summarise a recalc .jsonl: node summ.js file.jsonl [maxSamples] */
const fs = require('fs');
const lines = fs.readFileSync(process.argv[2], 'utf8').trim().split('\n').map((l) => JSON.parse(l));
const max = +(process.argv[3] || 60);
let files = 0, readErr = [], crash = [], calcErr = [], tot = { formulas: 0, compared: 0, ok: 0, kept: 0, vol: 0, nocache: 0 };
const fn = {};
const samples = [];
for (const o of lines) {
  files++;
  if (o.readError) { readErr.push(o.file + ': ' + o.readError); continue; }
  if (o.crash || o.fatal) { crash.push(o.file + ': ' + (o.crash || o.fatal)); continue; }
  if (o.calcError) { calcErr.push(o.file + ': ' + o.calcError); continue; }
  for (const k in tot) tot[k] += o[k] || 0;
  for (const k in o.fn) { const s = fn[k] || (fn[k] = [0, 0]); s[0] += o.fn[k][0]; s[1] += o.fn[k][1]; }
  for (const m of o.mism) samples.push(Object.assign({ file: o.file }, m));
}
console.log('files', files, JSON.stringify(tot), 'match', (tot.ok / tot.compared * 100).toFixed(3) + '%');
console.log('readErrors', readErr.length); readErr.slice(0, 30).forEach((x) => console.log('  ' + x));
console.log('crashes', crash.length); crash.slice(0, 30).forEach((x) => console.log('  ' + x));
console.log('calcErrors', calcErr.length); calcErr.slice(0, 10).forEach((x) => console.log('  ' + x));
const bad = Object.entries(fn).filter(([, s]) => s[1]).sort((a, b) => b[1][1] - a[1][1]);
console.log('functions with mismatches:', bad.slice(0, 50).map(([k, s]) => k + ' ' + s[1] + '/' + (s[0] + s[1])).join(', '));
console.log('mismatch samples', samples.length);
const byFile = {};
for (const s of samples) { (byFile[s.file] = byFile[s.file] || []).push(s); }
let n = 0;
for (const f in byFile) { if (n >= max) break; console.log('# ' + f); for (const s of byFile[f].slice(0, 4)) { console.log(`   ${s.s}!${s.a} =${s.f}  exp ${s.exp} got ${s.got}${s.m ? ' (member)' : ''}`); n++; } }
