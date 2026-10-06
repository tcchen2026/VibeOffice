/* node dbg.js file.xlsx "Sheet!A1" [more cells...] — show cells and recalculated values */
const L = require('./load.js');
const fs = require('fs');
(async () => {
  const file = process.argv[2];
  const wb = await L.xlsxRead.read(fs.readFileSync(file));
  wb.fileName = file.split('/').pop();
  const before = {};
  const cells = process.argv.slice(3).map((a) => { const m = /^(?:'?(.*?)'?!)?([A-Z]+\d+)$/.exec(a); const sh = m[1] ? wb.sheetByName(m[1]) : wb.sheets[0]; const p = L.formula.parseCell(m[2]); return { a, sh, p }; });
  for (const c of cells) { const cell = c.sh.get(c.p.r, c.p.c); before[c.a] = JSON.stringify(cell && { v: cell.v, f: cell.f, af: cell.af, am: cell.am, s: cell.s }); }
  L.calc.recalcAll(wb);
  for (const c of cells) { const cell = c.sh.get(c.p.r, c.p.c); console.log(c.a, before[c.a], '=>', JSON.stringify(cell && cell.v)); }
  if (process.env.NAMES) console.log(JSON.stringify(wb.names.filter((n) => new RegExp(process.env.NAMES, 'i').test(n.name))));
  if (process.env.EVAL) { const [sa, f] = process.env.EVAL.split('|'); const m = /^(?:'?(.*?)'?!)?([A-Z]+\d+)$/.exec(sa); const sh = m[1] ? wb.sheetByName(m[1]) : wb.sheets[0]; const p = L.formula.parseCell(m[2]); console.log('EVAL', f, '=>', JSON.stringify(L.calc.evalScalar(wb, sh, p.r, p.c, f))); }
})();
