/* xlsx round trip: read → write → read, compare the two models. Usage: node roundtrip.js file.xlsx [outdir] */
const L = require('./load.js');
const fs = require('fs');
const M = L.model, F = L.formula;
function same(a, b) {
  if (a === b) return true;
  if (a == null && b == null) return true;
  if (M.isErr(a) || M.isErr(b)) return M.isErr(a) && M.isErr(b) && a.e === b.e;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= Math.abs(a) * 1e-15;
  return false;
}
const J = (x) => JSON.stringify(x);
const canon = (x) => (x && typeof x === 'object' ? (Array.isArray(x) ? x.map(canon) : Object.keys(x).sort().reduce((o, k) => { if (x[k] !== undefined) o[k] = canon(x[k]); return o; }, {})) : x);
const CJ = (x) => JSON.stringify(canon(x));
/* a font without name / size means the workbook's default font; the writer spells those out */
const stJ = (wb, s) => {
  const st = Object.assign({}, wb.styles.get(s || 0));
  delete st.xs;
  const d = wb.defaultFont || {};
  if (st.font && (st.font.name == null || st.font.sz == null)) st.font = Object.assign({}, st.font, { name: st.font.name == null ? d.name || 'Calibri' : st.font.name, sz: st.font.sz == null ? d.sz || 11 : st.font.sz });
  return CJ(st);
};
(async () => {
  const file = process.argv[2], outdir = process.argv[3];
  const res = { file: file.split('/').pop(), diffs: [], counts: {} };
  const add = (k, msg) => { res.counts[k] = (res.counts[k] || 0) + 1; if (res.diffs.length < 25) res.diffs.push(k + ': ' + msg); };
  let a;
  try { a = await L.xlsxRead.read(fs.readFileSync(file)); } catch (e) { res.skip = String(e.message || e); console.log(J(res)); return; }
  let bytes;
  const t0 = Date.now();
  try { bytes = await L.xlsxWrite.bytes(a, { type: a.macroEnabled ? 'xlsm' : 'xlsx' }); } catch (e) { res.writeError = String(e.stack || e).slice(0, 400); console.log(J(res)); return; }
  res.writeMs = Date.now() - t0;
  res.size = [fs.statSync(file).size, bytes.length];
  if (outdir) fs.writeFileSync(outdir + '/' + res.file.replace(/\.(xlsx|xlsm)$/i, '') + (a.macroEnabled ? '.xlsm' : '.xlsx'), bytes);
  let b;
  try { b = await L.xlsxRead.read(bytes); } catch (e) { res.rereadError = String(e.stack || e).slice(0, 400); console.log(J(res)); return; }
  if (a.sheets.length !== b.sheets.length) add('sheets', a.sheets.length + ' vs ' + b.sheets.length);
  let cells = 0;
  a.sheets.forEach((s1, i) => {
    const s2 = b.sheets[i];
    if (!s2) return;
    if (s1.name !== s2.name) add('sheetName', s1.name + ' vs ' + s2.name);
    if (s1.state !== s2.state) add('state', s1.name);
    if (J(s1.tabColor) !== J(s2.tabColor)) add('tabColor', s1.name);
    s1.rows.forEach((row, r) => {
      if (!row) return;
      row.cells.forEach((c1, c) => {
        if (!c1) return;
        cells++;
        const c2 = s2.get(r, c);
        const at = s1.name + '!' + F.cellName(r, c);
        if (!c2) { if (c1.v != null || c1.f != null || c1.s) add('missing', at); return; }
        if (!same(c1.v, c2.v) && !(c1.v === '' && c2.v == null)) {
          const inTable = a.tables.some((t) => t.sheet === s1 && t.header !== false && t.ref.r1 === r && c >= t.ref.c1 && c <= t.ref.c2);
          if (!inTable) add('value', at + ' ' + J(c1.v && c1.v.e || c1.v) + ' vs ' + J(c2.v && c2.v.e || c2.v));
        }
        if ((c1.f || null) !== (c2.f || null) && !c1.am) add('formula', at + ' ' + c1.f + ' vs ' + c2.f);
        if (!!c1.af !== !!c2.af || (c1.af && J(c1.af) !== J(c2.af))) add('array', at);
        if (stJ(a, c1.s) !== stJ(b, c2.s)) add('style', at + ' ' + stJ(a, c1.s).slice(0, 160) + ' vs ' + stJ(b, c2.s).slice(0, 160));
        if (J(c1.rt || null) !== J(c2.rt || null)) add('rich', at);
      });
      const r2 = s2.rows[r];
      if (Math.abs(M.rowPt(s1, r) - M.rowPt(s2, r)) > 0.01) add('rowHeight', s1.name + ' row ' + (r + 1) + ' ' + M.rowPt(s1, r) + ' vs ' + M.rowPt(s2, r));
      if (!!row.hidden !== !!(r2 && r2.hidden)) add('rowHidden', s1.name + ' ' + (r + 1));
      if ((row.level || 0) !== ((r2 && r2.level) || 0)) add('rowLevel', s1.name + ' ' + (r + 1));
      if ((row.s || 0) && stJ(a, row.s) !== stJ(b, r2 && r2.s)) add('rowStyle', s1.name + ' ' + (r + 1));
    });
    for (let c = 0; c < Math.min(300, Math.max(s1.cols.length, s2.cols.length) + 2); c++) { if (M.colPx(s1, c) !== M.colPx(s2, c)) { add('colWidth', s1.name + ' col ' + F.colName(c) + ' ' + M.colPx(s1, c) + ' vs ' + M.colPx(s2, c)); break; } }
    if (J(s1.merges) !== J(s2.merges)) add('merges', s1.name);
    if (J(s1.view.freeze) !== J(s2.view.freeze)) add('freeze', s1.name + ' ' + J(s1.view.freeze) + ' vs ' + J(s2.view.freeze));
    for (const k of ['grid', 'headings', 'zeros', 'rtl', 'zoom', 'formulas']) if (s1.view[k] !== s2.view[k]) add('view.' + k, s1.name);
    if (s1.cf.reduce((n, x) => n + x.rules.length, 0) !== s2.cf.reduce((n, x) => n + x.rules.length, 0)) add('cf', s1.name + ' ' + s1.cf.length + ' vs ' + s2.cf.length);
    if (s1.dv.length !== s2.dv.length) add('dv', s1.name + ' ' + s1.dv.length + ' vs ' + s2.dv.length);
    else s1.dv.forEach((d, k) => { const d2 = s2.dv[k]; if (d.type !== d2.type || (d.f1 || '') !== (d2.f1 || '') || J(d.ranges) !== J(d2.ranges)) add('dvDetail', s1.name + ' ' + J([d.type, d.f1]) + ' vs ' + J([d2.type, d2.f1])); });
    if (s1.links.length !== s2.links.length) add('links', s1.name);
    else s1.links.forEach((l, k) => { const l2 = s2.links[k]; if ((l.target || '') !== (l2.target || '') || (l.location || '') !== (l2.location || '')) add('linkDetail', J(l) + ' vs ' + J(l2)); });
    if (s1.comments.size !== s2.comments.size) add('comments', s1.name + ' ' + s1.comments.size + ' vs ' + s2.comments.size);
    else for (const [k, cm] of s1.comments) { const c2 = s2.comments.get(k); if (!c2 || c2.text !== cm.text) add('commentText', s1.name + ' ' + J(cm.text).slice(0, 80) + ' vs ' + J(c2 && c2.text).slice(0, 80)); }
    const dk = (s) => s.drawings.map((d) => d.kind).sort().join(',');
    if (dk(s1) !== dk(s2)) add('drawings', s1.name + ' ' + dk(s1) + ' vs ' + dk(s2));
    if (J(s1.autoFilter && s1.autoFilter.ref) !== J(s2.autoFilter && s2.autoFilter.ref)) add('autoFilter', s1.name);
    const pk = (s) => J([s.print.orientation, s.print.margins, s.print.paper, s.print.scale, s.print.fit, s.print.header, s.print.footer, s.print.rowBreaks, s.print.colBreaks]);
    if (pk(s1) !== pk(s2)) add('print', s1.name + ' ' + pk(s1).slice(0, 200) + ' vs ' + pk(s2).slice(0, 200));
    if (!!s1.protection !== !!s2.protection) add('protection', s1.name);
    const ol = (o) => CJ({ below: o.below !== false, right: o.right !== false, r: o.levelRow || 0, c: o.levelCol || 0, a: !!o.applyStyles });
    if (ol(s1.outline) !== ol(s2.outline)) add('outline', s1.name + ' ' + J(s1.outline) + ' vs ' + J(s2.outline));
  });
  const nm = (w) => w.names.filter((n) => !/_FilterDatabase/i.test(n.name)).map((n) => n.name + '|' + n.scope + '|' + n.ref).sort().join('\n');
  if (nm(a) !== nm(b)) add('names', nm(a).slice(0, 200) + ' vs ' + nm(b).slice(0, 200));
  const tb = (w) => w.tables.map((t) => t.name + '|' + J(t.ref) + '|' + t.totals + '|' + (t.style && t.style.name)).join(',');
  if (tb(a) !== tb(b)) add('tables', tb(a) + ' vs ' + tb(b));
  if (a.date1904 !== b.date1904) add('date1904', '');
  if (a.active !== b.active) add('active', a.active + ' vs ' + b.active);
  if (J(a.theme.colors) !== J(b.theme.colors)) add('theme', '');
  if (!!a.extra.vba !== !!b.extra.vba) add('vba', '');
  res.cells = cells;
  console.log(J(res));
})();
