/* Summarise a corpus.js results file: node summarize.js results.jsonl [refJsonl] */
const fs = require('fs');
const rs = fs.readFileSync(process.argv[2], 'utf8').trim().split('\n').map((l) => JSON.parse(l));
const ref = process.argv[3] ? new Map(fs.readFileSync(process.argv[3], 'utf8').trim().split('\n').map((l) => JSON.parse(l)).map((r) => [r.file, r])) : null;
const by = (pred) => rs.filter(pred);
const tally = (list, key) => { const m = new Map(); for (const r of list) { const k = key(r); m.set(k, (m.get(k) || 0) + 1); } return Array.from(m.entries()).sort((a, b) => b[1] - a[1]); };
const opened = by((r) => r.slides != null);
console.log('files', rs.length, 'opened', opened.length, '(with a password:', by((r) => r.password).length + ')', 'readErrors', by((r) => r.readError).length, 'fatal', by((r) => r.fatal).length);
console.log('read errors:', tally(by((r) => r.readError), (r) => r.readError.slice(0, 70)));
console.log('fatal:', by((r) => r.fatal).map((r) => r.file + ' ' + r.fatal.slice(0, 160)));
console.log('loadErrors', by((r) => r.loadError).length, 'renderErrors(files)', by((r) => r.renderErrors).length, 'pngErrors', by((r) => r.pngErrors).length, 'pageErrors', by((r) => r.pageErrors).length);
console.log('writeErrors', by((r) => r.writeError).length, 'rereadErrors', by((r) => r.rereadError).length, 'rtDiff files', by((r) => r.rtDiffs).length, 'with warnings', by((r) => r.warnCount).length);
const show = (label, list, f) => { if (!list.length) return; console.log('\n## ' + label); for (const r of list.slice(0, +(process.env.N || 15))) console.log('  ' + r.file + '  ' + f(r)); };
show('load errors', by((r) => r.loadError), (r) => r.loadError.slice(0, 200));
show('render errors', by((r) => r.renderErrors), (r) => r.renderErrors[0].slice(0, 220));
show('png errors', by((r) => r.pngErrors), (r) => r.pngErrors[0]);
show('page errors', by((r) => r.pageErrors), (r) => r.pageErrors[0]);
show('write errors', by((r) => r.writeError), (r) => r.writeError.slice(0, 220));
show('reread errors', by((r) => r.rereadError), (r) => r.rereadError);
console.log('\nwarning kinds:', tally(rs.flatMap((r) => (r.warns || []).map((w) => ({ w }))), (x) => x.w.replace(/\d+/g, '#').slice(0, 90)).slice(0, 25));
console.log('\nround-trip diff kinds:', tally(rs.flatMap((r) => (r.rtDiffs || []).map((d) => ({ d }))), (x) => x.d.replace(/^slide \d+ /, '').replace(/".*$/, '').replace(/\d+ vs \d+/, 'n vs m').slice(0, 40)));
show('round-trip diffs', by((r) => r.rtDiffs), (r) => r.rtDiffs.slice(0, 2).join(' ; ').slice(0, 260));
if (ref) {
  const mism = opened.filter((r) => { const x = ref.get(r.file); return x && x.slides != null && x.slides !== r.slides; });
  console.log('\nslide count vs python-pptx: compared', opened.filter((r) => ref.get(r.file) && ref.get(r.file).slides != null).length, 'mismatch', mism.length);
  show('slide count mismatches', mism, (r) => r.slides + ' vs python-pptx ' + ref.get(r.file).slides + (r.repaired ? ' (repaired)' : ''));
  const notOpenedButPyOk = rs.filter((r) => r.readError && ref.get(r.file) && ref.get(r.file).slides != null);
  show('python-pptx opens but Lectern does not', notOpenedButPyOk, (r) => r.readError);
}
const ms = opened.map((r) => r.readMs + (r.renderMs || 0)).sort((a, b) => a - b);
console.log('\nopen+render ms: median', ms[Math.floor(ms.length / 2)], 'p90', ms[Math.floor(ms.length * 0.9)], 'max', ms[ms.length - 1]);
