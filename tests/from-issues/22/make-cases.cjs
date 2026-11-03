// Regenerate the four small merge controls used with the win10 Excel PDF API.
// node tests/from-issues/22/make-cases.cjs [OUTPUT.xlsx]
const fs = require('node:fs'), path = require('node:path');
const L = require('../../ledger/load.js');

async function main() {
  const wb = new L.model.Workbook(), sh = wb.addSheet('Merged borders');
  const line = (rgb = 'FF000000', style = 'thin') => ({ color: { rgb }, style });
  const set = (r, c, border) => sh.put(r, c, { v: null, s: wb.styles.add({ border }) });
  sh.view.grid = false;
  const names = ['Anchor only', 'Separate perimeter segments', 'No anchor cell', 'Diagonal with perimeter'];
  for (const [i, title] of names.entries()) {
    const r = 2 + i * 5;
    sh.put(r - 1, 1, { v: title });
    sh.merges.push({ r1: r, c1: 1, r2: r + 2, c2: 3 });
    if (i === 0) set(r, 1, { t: line(), l: line(), b: line(), r: line() });
    if (i === 1) {
      for (let c = 1; c <= 3; c++) {
        const sides = { ...(c === 1 ? { l: line() } : {}), ...(c === 3 ? { r: line() } : {}) };
        set(r, c, { t: line(['FFFF0000', 'FF008000', 'FF0000FF'][c - 1], 'medium'), ...sides });
        set(r + 2, c, { b: line('FF0000FF'), ...sides });
      }
      set(r + 1, 1, { l: line() }); set(r + 1, 3, { r: line() });
    }
    if (i === 2) {
      for (let c = 1; c <= 3; c++) set(r + 2, c, { b: line() });
      for (let rr = r; rr <= r + 2; rr++) set(rr, 3, { ...(rr === r + 2 ? { b: line() } : {}), r: line() });
    }
    if (i === 3) {
      set(r, 1, { t: line(), l: line(), d: line(), dd: true });
      set(r, 2, { t: line() }); set(r, 3, { t: line(), r: line() });
      for (let c = 1; c <= 3; c++) set(r + 2, c, { b: line() });
    }
  }
  fs.writeFileSync(process.argv[2] || path.join(__dirname, 'merge-cases.xlsx'), await L.xlsxWrite.bytes(wb));
}
main().catch(e => { console.error(e); process.exitCode = 1; });
