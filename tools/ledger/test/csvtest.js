/* CSV import / export test for one file.
 * 1. decode + sniff + parse with Ledger's parser, report the delimiter, encoding and a digest of the parsed rows
 *    (test/csvref.py parses the same file with Python's csv module and compares the digests);
 * 2. build the workbook as File > Open does, write it back as CSV, re-import that, and compare every cell value.
 * Usage: node csvtest.js file.csv */
const fs = require('fs');
const crypto = require('crypto');
const path = require('path');
const L = require('./load');
globalThis.Blob = globalThis.Blob || require('buffer').Blob;
const CSV = L.csv, M = L.model;

const file = process.argv[2];
const res = { file: path.basename(file) };
try {
  const bytes = new Uint8Array(fs.readFileSync(file));
  res.size = bytes.length;
  const t0 = Date.now();
  const { text, encoding } = CSV.decode(bytes);
  const delim = CSV.sniff(text);
  const rows = CSV.parse(text, { delims: delim });
  res.encoding = encoding;
  res.delim = delim;
  res.rows = rows.length;
  /* blank lines: [] and [''] are the same thing */
  const norm = rows.map((r) => (r.length === 1 && r[0] === '' ? [] : r));
  res.digest = crypto.createHash('md5').update(JSON.stringify(norm)).digest('hex');
  const wb = CSV.toWorkbook(rows, { name: res.file });
  res.importMs = Date.now() - t0;
  const sh = wb.sheets[0];
  let cells = 0, nums = 0, dates = 0, texts = 0, formulas = 0;
  sh.rows.forEach((row) => row && row.cells.forEach((c) => {
    if (!c) return;
    cells++;
    if (c.f != null) formulas++;
    else if (typeof c.v === 'number') { nums++; if (c.s && /[dmy]/.test(wb.styles.get(c.s).nf || '')) dates++; } else if (typeof c.v === 'string') texts++;
  }));
  Object.assign(res, { cells, nums, dates, texts, formulas });
  /* write back and re-import */
  const t1 = Date.now();
  const out = CSV.text(sh, ',');
  res.exportMs = Date.now() - t1;
  const wb2 = CSV.toWorkbook(CSV.parse(out, { delims: ',' }), { name: res.file });
  const sh2 = wb2.sheets[0];
  const diffs = [];
  const same = (a, b) => {
    if (a == null || a === '') return b == null || b === '';
    if (typeof a === 'number' && typeof b === 'number') return a === b || Math.abs(a - b) <= Math.abs(a) * 1e-14;
    if (M.isErr(a)) return M.isErr(b) && a.e === b.e;
    return a === b;
  };
  const R = Math.max(sh.rows.length, sh2.rows.length);
  for (let r = 0; r < R; r++) {
    const a = sh.rows[r], b = sh2.rows[r];
    const C = Math.max(a ? a.cells.length : 0, b ? b.cells.length : 0);
    for (let c = 0; c < C; c++) {
      const x = a && a.cells[c], y = b && b.cells[c];
      const fx = x && x.f, fy = y && y.f;
      if (fx != null || fy != null) continue; /* formulas export as their values: not comparable */
      const vx = x ? x.v : null, vy = y ? y.v : null;
      if (same(vx, vy)) continue;
      /* Excel writes what the cell shows: a value typed as 3.73929E-09 gets the 0.00E+00 format and is saved as
         3.74E-09, a time with seconds gets m/d/yyyy h:mm. Those losses are Excel's own; count them apart. */
      const nf = x && x.s ? wb.styles.get(x.s).nf || 'General' : 'General';
      if (!/^General$/i.test(nf) && typeof vx === 'number') { res.displayRounded = (res.displayRounded || 0) + 1; continue; }
      if (diffs.length < 5) diffs.push({ at: r + ',' + c, a: vx, b: vy });
      res.diffCount = (res.diffCount || 0) + 1;
    }
  }
  res.diffs = diffs;
} catch (e) {
  res.error = String(e && e.stack || e).slice(0, 400);
}
console.log(JSON.stringify(res));
