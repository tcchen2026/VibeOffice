/* list TEXT() mismatches: node textfmt.js file.xlsx */
const L = require('./load.js');
const fs = require('fs');
const M = L.model, C = L.calc, F = L.formula;
(async () => {
  const wb = await L.xlsxRead.read(fs.readFileSync(process.argv[2]));
  const exp = [];
  for (const sh of wb.sheets) sh.rows.forEach((row, r) => row && row.cells.forEach((cell, c) => { if (cell && cell.f && /TEXT\(/i.test(cell.f) && !cell.dirty) exp.push({ sh, r, c, cell, v: cell.v }); }));
  C.recalcAll(wb);
  {
    for (const e of exp) {
      const sh = e.sh;
      if (e.cell.v === e.v) continue;
      const m = /TEXT\(\s*([A-Z]+\d+)\s*,\s*([A-Z]+\d+)\s*\)/i.exec(e.cell.f);
      let fmt = '', val = '';
      if (m) { const a = F.parseCell(m[1]), b = F.parseCell(m[2]); val = sh.val(a.r, a.c); fmt = sh.val(b.r, b.c); }
      console.log(F.cellName(e.r, e.c), JSON.stringify(fmt), JSON.stringify(val), 'exp', JSON.stringify(M.isErr(e.v) ? e.v.e : e.v), 'got', JSON.stringify(M.isErr(e.cell.v) ? e.cell.v.e : e.cell.v), m ? '' : e.cell.f);
    }
  }
})();
