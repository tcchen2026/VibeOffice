/* Recalculate every formula of one workbook and compare with the values Excel saved. */
const L = require('./load.js');
const fs = require('fs');
const M = L.model, C = L.calc, F = L.formula;
const VOL = /\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|INFO|CELL|OFFSET|INDIRECT)\s*\(/i;
const HARDVOL = /\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|INFO)\s*\(/i;
function same(a, b) {
  if (a === b) return true;
  if (M.isErr(a) || M.isErr(b)) return M.isErr(a) && M.isErr(b) && a.e === b.e;
  if (typeof a === 'number' && typeof b === 'number') {
    const d = Math.abs(a - b), m = Math.max(Math.abs(a), Math.abs(b));
    return d <= m * 1e-9 || d < 1e-10;
  }
  if (typeof a === 'string' && typeof b === 'string') return a === b;
  if ((a === '' && b == null) || (b === '' && a == null)) return true;
  return false;
}
const show = (v) => (M.isErr(v) ? v.e : v === null ? 'null' : typeof v === 'string' ? JSON.stringify(v.length > 60 ? v.slice(0, 60) + '…' : v) : String(v));
(async () => {
  const file = process.argv[2];
  const out = { file: file.split('/').pop(), formulas: 0, compared: 0, ok: 0, kept: 0, vol: 0, nocache: 0, mism: [], fn: {} };
  const t0 = Date.now();
  let wb;
  const bytes = fs.readFileSync(file);
  try { wb = await L.xlsxRead.read(bytes); } catch (e) {
    if (e && e.code === 'password') {
      for (const pw of ['password', 'Password', 'test', '1234', 'pass', 'secret']) { try { wb = await L.xlsxRead.read(bytes, { password: pw }); out.password = pw; break; } catch (e2) { /* next */ } }
    }
    if (!wb) { out.readError = String(e && e.message || e); console.log(JSON.stringify(out)); return; }
  }
  if (wb.calcPr.fullCalcOnLoad) out.fullCalc = true;
  out.readMs = Date.now() - t0;
  out.app = wb.props.application || '';
  wb.fileName = file.split('/').pop().replace(/^[a-z]+__[0-9a-f]+__/, '');
  const exp = [];
  for (const sh of wb.sheets) {
    sh.rows.forEach((row, r) => {
      if (!row) return;
      row.cells.forEach((cell, c) => {
        if (!cell) return;
        if (cell.f != null) {
          out.formulas++;
          if (cell.dirty) { out.nocache++; return; }
          exp.push({ sh, r, c, cell, v: cell.v, f: cell.f });
          if (cell.af) {
            const a = cell.af;
            for (let rr = a.r1; rr <= a.r2 && rr <= a.r1 + 200; rr++) for (let cc = a.c1; cc <= a.c2 && cc <= a.c1 + 50; cc++) {
              if (rr === r && cc === c) continue;
              const m = sh.get(rr, cc);
              if (m) exp.push({ sh, r: rr, c: cc, cell: m, v: m.v, f: cell.f, member: true, master: cell });
            }
          }
        }
      });
    });
  }
  const t1 = Date.now();
  try { C.recalcAll(wb); } catch (e) { out.calcError = String(e && e.stack || e).slice(0, 500); console.log(JSON.stringify(out)); return; }
  out.calcMs = Date.now() - t1;
  for (const e of exp) {
    const fcell = e.master || e.cell;
    if (fcell.kept) { out.kept++; continue; }
    if (HARDVOL.test(e.f)) { out.vol++; continue; }
    if (e.v == null) continue;
    if (out.fullCalc) { out.unverified = (out.unverified || 0) + 1; continue; }
    out.compared++;
    const fns = new Set((e.f.toUpperCase().match(/[A-Z][A-Z0-9.]*(?=\()/g) || []).map((x) => x.replace(/^_XLFN\.|^_XLWS\./, '')));
    const ok = same(e.cell.v, e.v);
    for (const fn of fns) { const s = out.fn[fn] || (out.fn[fn] = [0, 0]); s[ok ? 0 : 1]++; }
    if (ok) out.ok++;
    else if (out.mism.length < 40) out.mism.push({ s: e.sh.name, a: F.cellName(e.r, e.c), f: e.f.length > 200 ? e.f.slice(0, 200) + '…' : e.f, exp: show(e.v), got: show(e.cell.v), vol: VOL.test(e.f) || undefined, m: e.member || undefined });
  }
  console.log(JSON.stringify(out));
})().catch((e) => { console.log(JSON.stringify({ file: process.argv[2].split('/').pop(), fatal: String(e && e.stack || e).slice(0, 500) })); });
