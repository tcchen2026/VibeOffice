/* XML Spreadsheet 2003 round trip: read an .xlsx, write SpreadsheetML 2003, read that back, compare values,
 * formulas, merges and a few style properties. Usage: node xmlss-rt.js file.xlsx [outdir] */
const L = require('./load.js');
const fs = require('fs');
const M = L.model, F = L.formula;
const same = (a, b) => {
  if (a === b) return true;
  if ((a == null || a === '') && (b == null || b === '')) return true;
  if (M.isErr(a) || M.isErr(b)) return M.isErr(a) && M.isErr(b) && a.e === b.e;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= Math.abs(a) * 1e-14 + 1e-300;
  if (typeof a === 'string' && typeof b === 'string') return a.replace(/\r\n?/g, '\n') === b.replace(/\r\n?/g, '\n');
  return false;
};
(async () => {
  const file = process.argv[2], outdir = process.argv[3];
  const res = { file: file.split('/').pop(), counts: {}, diffs: [] };
  const add = (k, m) => { res.counts[k] = (res.counts[k] || 0) + 1; if (res.diffs.length < 8) res.diffs.push(k + ': ' + m); };
  let a;
  try { a = await L.xlsxRead.read(fs.readFileSync(file)); } catch (e) { res.skip = String(e.message || e).slice(0, 80); console.log(JSON.stringify(res)); return; }
  let text;
  try { text = L.xmlss.text(a); } catch (e) { res.writeError = String(e.stack || e).slice(0, 300); console.log(JSON.stringify(res)); return; }
  if (outdir) fs.writeFileSync(outdir + '/' + res.file.replace(/\.xls[xm]$/i, '.xml'), text);
  let b;
  try { b = await L.xmlss.read(new TextEncoder().encode(text)); } catch (e) { res.readError = String(e.stack || e).slice(0, 300); console.log(JSON.stringify(res)); return; }
  let cells = 0;
  if (a.sheets.length !== b.sheets.length) add('sheets', a.sheets.length + ' vs ' + b.sheets.length);
  a.sheets.forEach((s1, i) => {
    const s2 = b.sheets[i];
    if (!s2) return;
    if (s1.name !== s2.name) add('name', s1.name + ' vs ' + s2.name);
    s1.rows.forEach((row, r) => {
      if (!row || r >= 65536) return; /* SpreadsheetML 2003 stops at Excel 2003's grid */
      row.cells.forEach((c1, c) => {
        if (!c1 || c >= 256) return;
        if (c1.v == null && c1.f == null) return;
        cells++;
        const c2 = s2.get(r, c);
        const at = s1.name + '!' + F.cellName(r, c);
        if (!c2) { add('missing', at); return; }
        if (!same(c1.v, c2.v)) add('value', at + ' ' + JSON.stringify(c1.v).slice(0, 50) + ' vs ' + JSON.stringify(c2.v).slice(0, 50));
        if ((c1.f || null) !== (c2.f || null) && !c1.am) add('formula', at + ' ' + c1.f + ' vs ' + c2.f);
        const st1 = a.styles.get(c1.s || 0), st2 = b.styles.get(c2.s || 0);
        if ((st1.nf || 'General') !== (st2.nf || 'General')) add('numFmt', at + ' ' + st1.nf + ' vs ' + st2.nf);
        if (!!(st1.font && st1.font.b) !== !!(st2.font && st2.font.b)) add('bold', at);
      });
    });
    const mk = (s) => JSON.stringify(s.merges.filter((m) => m.r2 < 65536 && m.c2 < 256).map((m) => [m.r1, m.c1, m.r2, m.c2]).sort());
    if (mk(s1) !== mk(s2)) add('merges', s1.name);
  });
  res.cells = cells;
  console.log(JSON.stringify(res));
})();
