/* Compare cell values of our written files with LibreOffice's re-saved copies. Usage: node lo-compare.js ourDir loDir */
const L = require('./load.js');
const fs = require('fs');
const path = require('path');
const M = L.model;
const [ours, lo] = process.argv.slice(2);
const same = (a, b) => {
  if (a === b) return true;
  /* LibreOffice stores line breaks as \n */
  if (typeof a === 'string' && typeof b === 'string' && a.replace(/\r\n?/g, '\n') === b.replace(/\r\n?/g, '\n')) return true;
  if ((a == null || a === '') && (b == null || b === '')) return true;
  if (M.isErr(a) || M.isErr(b)) return M.isErr(a) && M.isErr(b) && a.e === b.e;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= Math.max(Math.abs(a), Math.abs(b)) * 1e-9 + 1e-12;
  if (typeof a === 'boolean' && typeof b === 'number') return +a === b;
  return false;
};
(async () => {
  let files = 0, cells = 0, bad = 0, skipped = 0, sheetsBad = 0, formulas = 0, fbad = 0;
  const notes = [];
  for (const f of fs.readdirSync(lo)) {
    const a = path.join(ours, f), b = path.join(lo, f);
    if (!fs.existsSync(a)) continue;
    let wa, wb;
    try { wa = await L.xlsxRead.read(fs.readFileSync(a)); wb = await L.xlsxRead.read(fs.readFileSync(b)); } catch (e) { notes.push(f + ': read error ' + e.message); continue; }
    files++;
    if (wa.sheets.length !== wb.sheets.length) { sheetsBad++; notes.push(f + ': sheets ' + wa.sheets.length + ' vs ' + wb.sheets.length); continue; }
    let fileBad = 0;
    wa.sheets.forEach((s1, i) => {
      const s2 = wb.sheets[i];
      s1.rows.forEach((row, r) => { if (!row || r > 5000) return; row.cells.forEach((c1, c) => {
        if (!c1 || c > 200) return;
        if (c1.v == null && c1.f == null) return;
        cells++;
        if (c1.f != null) formulas++;
        const v2 = s2.val(r, c);
        /* not comparable: volatile functions, external links, values never cached by the writer of the source */
        if (c1.f != null && (/\b(RAND|RANDBETWEEN|RANDARRAY|NOW|TODAY|INFO|CELL)\(|\[\d+\]/i.test(c1.f) || c1.v == null)) { skipped++; return; }
        if (!same(c1.v, v2)) { bad++; fileBad++; if (c1.f != null) fbad++; if (fileBad <= 2 || process.env.ALL) notes.push(f + ' ' + s1.name + '!' + L.formula.cellName(r, c) + ' ' + JSON.stringify(c1.v && c1.v.e || c1.v).slice(0, 60) + ' vs ' + JSON.stringify(v2 && v2.e || v2).slice(0, 60) + (c1.f ? ' f=' + c1.f.slice(0, 60) : '')); }
      }); });
    });
  }
  console.log(JSON.stringify({ files, cells, skipped, bad, formulas, fbad, sheetsBad }));
  console.log(notes.slice(0, process.env.ALL ? 1e6 : 60).join('\n'));
})();
