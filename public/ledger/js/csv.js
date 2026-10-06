/* Ledger — text files: CSV / TXT import (with the Text Import Wizard), CSV / TXT / HTML export. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, NF = L.numfmt, F = L.formula;
  const CSV = (L.csv = {});

  /* ------------------------------------------------------------ decoding */
  /** bytes → {text, encoding}; honours BOMs, tries strict UTF-8, falls back to Windows-1252 */
  /* Windows-1252's 0x80–0x9F block (some runtimes' TextDecoder treat the label as plain Latin-1) */
  const CP1252 = '€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008DŽ\u008F\u0090‘’“”•–—˜™š›œ\u009DžŸ';
  function decode1252(b) {
    let out = '';
    for (let i = 0; i < b.length; i += 8192) {
      const part = b.subarray(i, i + 8192);
      const codes = new Array(part.length);
      for (let k = 0; k < part.length; k++) { const x = part[k]; codes[k] = x >= 0x80 && x < 0xa0 ? CP1252.charCodeAt(x - 0x80) : x; }
      out += String.fromCharCode.apply(null, codes);
    }
    return out;
  }
  CSV.decode = function (bytes, enc) {
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const dec = (e, fatal) => (/^(windows-1252|cp1252|latin1|iso-8859-1|ascii)$/i.test(e) ? decode1252(b) : new TextDecoder(e, { fatal: !!fatal }).decode(b));
    if (enc && enc !== 'auto') {
      let t = dec(enc);
      if (t.charCodeAt(0) === 0xfeff) t = t.slice(1);
      return { text: t, encoding: enc };
    }
    if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return { text: dec('utf-8').replace(/^﻿/, ''), encoding: 'utf-8' };
    if (b[0] === 0xff && b[1] === 0xfe) return { text: dec('utf-16le').replace(/^﻿/, ''), encoding: 'utf-16le' };
    if (b[0] === 0xfe && b[1] === 0xff) return { text: dec('utf-16be').replace(/^﻿/, ''), encoding: 'utf-16be' };
    /* UTF-16 without BOM: many zero bytes in alternating positions */
    if (b.length >= 4) {
      let z0 = 0, z1 = 0;
      const n = Math.min(b.length, 2000);
      for (let i = 0; i + 1 < n; i += 2) { if (b[i] === 0) z0++; if (b[i + 1] === 0) z1++; }
      if (z1 > n / 6 && z0 < n / 50) return { text: dec('utf-16le'), encoding: 'utf-16le' };
      if (z0 > n / 6 && z1 < n / 50) return { text: dec('utf-16be'), encoding: 'utf-16be' };
    }
    try { return { text: dec('utf-8', true), encoding: 'utf-8' }; } catch (e) { return { text: dec('windows-1252'), encoding: 'windows-1252' }; }
  };

  /* ------------------------------------------------------------ parsing */
  /**
   * Delimited text → array of rows (arrays of strings).
   * o: {delims: string of delimiter chars, quote: '"' | "'" | '', merge: treat consecutive delimiters as one, maxRows}
   */
  CSV.parse = function (text, o) {
    o = o || {};
    const delims = o.delims != null ? o.delims : ',';
    const q = o.quote === undefined ? '"' : o.quote;
    const rows = [];
    let row = [], field = '', i = 0, quoted = false, wasQuoted = false;
    const n = text.length;
    const isDelim = (ch) => delims.indexOf(ch) >= 0;
    const endField = () => { row.push(field); field = ''; wasQuoted = false; };
    const endRow = () => { endField(); rows.push(row); row = []; };
    const max = o.maxRows || Infinity;
    while (i < n) {
      const ch = text[i];
      if (quoted) {
        if (ch === q) {
          if (text[i + 1] === q) { field += q; i += 2; continue; }
          quoted = false; i++; continue;
        }
        field += ch; i++; continue;
      }
      if (q && ch === q && (field === '' || (o.lenientQuotes && /^\s*$/.test(field))) && !wasQuoted) { quoted = true; wasQuoted = true; field = ''; i++; continue; }
      if (isDelim(ch)) {
        endField();
        i++;
        if (o.merge) while (i < n && isDelim(text[i])) i++;
        continue;
      }
      if (ch === '\r' || ch === '\n') {
        endRow();
        if (ch === '\r' && text[i + 1] === '\n') i++;
        i++;
        if (rows.length >= max) return rows;
        continue;
      }
      field += ch; i++;
    }
    if (field !== '' || row.length || quoted) endRow();
    return rows;
  };
  /** fixed width: breaks = sorted character positions */
  CSV.parseFixed = function (text, breaks, maxRows) {
    const lines = text.split(/\r\n|\r|\n/);
    if (lines.length && lines[lines.length - 1] === '') lines.pop();
    const out = [];
    for (const ln of lines.slice(0, maxRows || Infinity)) {
      const row = [];
      let p = 0;
      for (const b of breaks.concat([Infinity])) { row.push(ln.slice(p, b === Infinity ? undefined : b).trim()); p = b; }
      out.push(row);
    }
    return out;
  };
  /** the most likely delimiter of a delimited text */
  CSV.sniff = function (text) {
    const sample = text.slice(0, 20000);
    const lines = CSV.parse(sample, { delims: '\u0000', quote: '' }).map((r) => r[0]).filter((l) => l && l.trim()).slice(0, 30);
    if (!lines.length) return ',';
    let best = ',', bestScore = -1;
    for (const d of [',', '\t', ';', '|']) {
      const counts = lines.map((l) => CSV.parse(l, { delims: d }).reduce((m, r) => Math.max(m, r.length), 0));
      const first = counts[0];
      if (first < 2) continue;
      const consistent = counts.filter((c) => c === first).length / counts.length;
      const score = consistent * 10 + Math.min(first, 50) / 50 + (d === ',' ? 0.05 : d === '\t' ? 0.04 : 0);
      if (score > bestScore) { bestScore = score; best = d; }
    }
    return best;
  };
  /** fixed-width guess: columns of spaces common to all sample lines */
  CSV.guessBreaks = function (text) {
    const lines = text.split(/\r\n|\r|\n/).filter((l) => l.trim()).slice(0, 50);
    if (!lines.length) return [];
    const w = Math.max(...lines.map((l) => l.length));
    const blank = new Array(w).fill(true);
    for (const l of lines) for (let i = 0; i < w; i++) if (l[i] && l[i] !== ' ') blank[i] = false;
    const out = [];
    for (let i = 1; i < w; i++) if (!blank[i] && blank[i - 1]) out.push(i);
    return out;
  };

  /* ------------------------------------------------------------ value conversion (Excel's General rules) */
  const MONTH3 = NF.MONTHS.map((m) => m.slice(0, 3).toLowerCase());
  /**
   * text → {v, nf?, f?} as Excel's General column conversion does.
   * fmt: 'general' | 'text' | 'MDY' | 'DMY' | 'YMD' | 'MYD' | 'DYM' | 'YDM' | 'skip'
   */
  CSV.convert = function (s, fmt, o) {
    o = o || {};
    if (s == null) return { v: null };
    if (fmt === 'text') return { v: s === '' ? null : s, nf: s === '' ? undefined : '@' };
    const t = s.trim();
    if (t === '') return { v: null };
    if (fmt && /^[MDY]{3}$/.test(fmt)) { const d = parseDate(t, fmt, o.date1904); if (d) return d; }
    if (t[0] === '=' && t.length > 1 && o.formulas !== false) {
      try { F.parse(t.slice(1)); return { f: F.toStore ? F.toStore(t.slice(1)) : t.slice(1) }; } catch (e) { return { v: s }; }
    }
    const p = NF.parseInput(t, { date1904: o.date1904 });
    if (p && typeof p.v === 'number') return { v: p.v, nf: p.fmt || undefined };
    if (p && typeof p.v === 'boolean') return { v: p.v };
    if (p && p.v && typeof p.v === 'object' && p.v.e) return { v: M.err(p.v.e) };
    /* a trailing minus (1234-) is negative, as the wizard's "Trailing minus" option does */
    if (/^[\d,.]+-$/.test(t) && o.trailingMinus !== false) { const n = NF.parseNumber(t.slice(0, -1)); if (n && typeof n.v === 'number') return { v: -n.v }; }
    return { v: s };
  };
  function parseDate(t, order, d1904) {
    const parts = t.split(/[\/\-. ]+/).filter(Boolean);
    if (parts.length < 3) return null;
    const o = {};
    for (let i = 0; i < 3; i++) {
      const k = order[i], x = parts[i];
      if (k === 'M' && /^[a-z]{3,}/i.test(x)) { const m = MONTH3.indexOf(x.slice(0, 3).toLowerCase()); if (m < 0) return null; o.M = m + 1; continue; }
      if (!/^\d+$/.test(x)) return null;
      o[k] = +x;
    }
    if (o.Y < 100) o.Y += o.Y < 30 ? 2000 : 1900;
    if (!(o.M >= 1 && o.M <= 12 && o.D >= 1 && o.D <= 31)) return null;
    let v = NF.dateToSerial(o.Y, o.M, o.D, d1904);
    const tm = /(\d{1,2}):(\d\d)(?::(\d\d))?\s*(am|pm)?$/i.exec(t);
    if (tm) { let hh = +tm[1]; if (tm[4]) { if (/pm/i.test(tm[4]) && hh < 12) hh += 12; if (/am/i.test(tm[4]) && hh === 12) hh = 0; } v += (hh * 3600 + +tm[2] * 60 + (+tm[3] || 0)) / 86400; return { v, nf: 'm/d/yyyy h:mm' }; }
    return { v, nf: 'm/d/yyyy' };
  }

  /** rows of strings → a one-sheet workbook */
  CSV.toWorkbook = function (rows, o) {
    o = o || {};
    const wb = L.app && L.app.blankWorkbook ? L.app.blankWorkbook() : new M.Workbook();
    if (!wb.sheets.length) wb.addSheet('Sheet1');
    wb.sheets.splice(1);
    const sh = wb.sheets[0];
    sh.name = (o.name || 'Sheet1').replace(/\.[^.]+$/, '').replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet1';
    CSV.fill(sh, rows, o, 0, 0);
    return wb;
  };
  /** write parsed rows into a sheet at (r0, c0); colFormats[i] per column */
  CSV.fill = function (sh, rows, o, r0, c0) {
    o = o || {};
    const wb = sh.wb;
    const fmts = o.colFormats || [];
    const nfStyle = new Map();
    const styleFor = (base, nf) => { const k = base + '|' + nf; let s = nfStyle.get(k); if (s == null) { s = wb.styles.derive(base, { nf }); nfStyle.set(k, s); } return s; };
    const start = o.startRow || 0;
    let r = r0 || 0;
    for (let i = start; i < rows.length; i++, r++) {
      if (r >= M.MAXR) break;
      const row = rows[i];
      let c = c0 || 0;
      for (let j = 0; j < row.length; j++) {
        const fmt = fmts[j] || 'general';
        if (fmt === 'skip') continue;
        if (c >= M.MAXC) break;
        const x = CSV.convert(row[j], fmt, o);
        if (x.v != null || x.f != null) {
          const cell = sh.cell(r, c);
          cell.v = x.v == null ? null : x.v;
          if (x.f != null) { cell.f = x.f; cell.dirty = true; }
          if (x.nf) cell.s = styleFor(cell.s || 0, x.nf);
        }
        c++;
      }
    }
    return r;
  };
  /** import bytes without UI (used for .csv and by tests) */
  CSV.importBytes = function (bytes, name, o) {
    o = o || {};
    const { text, encoding } = CSV.decode(bytes, o.encoding);
    let rows;
    if (o.fixed) rows = CSV.parseFixed(text, o.breaks || CSV.guessBreaks(text));
    else {
      const ext = (String(name || '').split('.').pop() || '').toLowerCase();
      const delims = o.delims != null ? o.delims : ext === 'tsv' || ext === 'txt' || ext === 'tab' ? (text.indexOf('\t') >= 0 ? '\t' : CSV.sniff(text)) : CSV.sniff(text);
      rows = CSV.parse(text, { delims, quote: o.quote === undefined ? '"' : o.quote, merge: o.merge });
    }
    const wb = CSV.toWorkbook(rows, Object.assign({ name }, o));
    wb.textImport = { encoding };
    return wb;
  };

  /* ------------------------------------------------------------ Text Import Wizard */
  CSV.importDialog = function (bytes, name, opts) {
    opts = opts || {};
    const ext = (String(name).split('.').pop() || '').toLowerCase();
    if (ext === 'csv' && !opts.wizard) return Promise.resolve({ wb: CSV.importBytes(bytes, name) });
    const ui = L.ui, h = L.h;
    let enc = 'auto';
    let decoded = CSV.decode(bytes);
    const st = { fixed: false, start: 1, delims: { tab: true, semi: false, comma: false, space: false, other: false, otherCh: '' }, merge: false, quote: '"', breaks: null, formats: [], step: 1 };
    const sn = CSV.sniff(decoded.text);
    st.delims.tab = sn === '\t'; st.delims.comma = sn === ','; st.delims.semi = sn === ';'; if (sn === '|') { st.delims.other = true; st.delims.otherCh = '|'; }
    const lines = () => decoded.text.split(/\r\n|\r|\n/);
    st.fixed = !(/[\t,;|]/.test(decoded.text.slice(0, 5000))) && /  /.test(decoded.text.slice(0, 5000));
    const delimStr = () => (st.delims.tab ? '\t' : '') + (st.delims.semi ? ';' : '') + (st.delims.comma ? ',' : '') + (st.delims.space ? ' ' : '') + (st.delims.other && st.delims.otherCh ? st.delims.otherCh[0] : '');
    const parsedRows = (max) => (st.fixed ? CSV.parseFixed(decoded.text, st.breaks || (st.breaks = CSV.guessBreaks(decoded.text)), max) : CSV.parse(decoded.text, { delims: delimStr(), quote: st.quote, merge: st.merge, maxRows: max }));
    const body = h('div', { class: 'tiw' });
    const title = h('div', { class: 'tiw-step' });
    const content = h('div', { class: 'tiw-content' });
    const preview = h('div', { class: 'tiw-preview', tabindex: '0' });
    body.append(title, content, h('div', { class: 'tiw-cap', text: 'Preview of selected data:' }), preview);
    let selCol = 0;
    function renderPreview() {
      preview.textContent = '';
      if (st.step === 1) {
        const pre = h('pre', { class: 'tiw-pre' });
        lines().slice(st.start - 1, st.start + 14).forEach((l, i) => pre.appendChild(document.createTextNode(String(st.start + i).padStart(3) + ' ' + l.replace(/\t/g, '→') + '\n')));
        preview.appendChild(pre);
        return;
      }
      if (st.step === 2 && st.fixed) {
        /* the ruler: click to add a break, click a break to remove it */
        const ls = lines().slice(st.start - 1, st.start + 14);
        const cw = 7.2;
        const w = Math.max(40, ...ls.map((l) => l.length));
        const box = h('div', { class: 'tiw-fixed', style: `width:${w * cw + 10}px` });
        const ruler = h('div', { class: 'tiw-ruler' });
        for (let i = 0; i <= w; i += 10) ruler.appendChild(h('span', { style: `left:${i * cw}px`, text: String(i) }));
        box.appendChild(ruler);
        const pre = h('pre', { class: 'tiw-pre', text: ls.join('\n') });
        box.appendChild(pre);
        for (const b of st.breaks || []) box.appendChild(h('div', { class: 'tiw-break', style: `left:${b * cw + 4}px` }));
        box.addEventListener('click', (e) => {
          const r = box.getBoundingClientRect();
          const pos = Math.round((e.clientX - r.left - 4) / cw);
          if (pos <= 0) return;
          const i = st.breaks.indexOf(pos);
          const near = st.breaks.findIndex((b) => Math.abs(b - pos) <= 0);
          if (i >= 0 || near >= 0) st.breaks.splice(i >= 0 ? i : near, 1); else { st.breaks.push(pos); st.breaks.sort((a, b) => a - b); }
          renderPreview();
        });
        preview.appendChild(box);
        return;
      }
      const rows = parsedRows(st.start - 1 + 15).slice(st.start - 1);
      const ncol = Math.max(1, ...rows.map((r) => r.length));
      const tbl = h('table', { class: 'tiw-table' });
      if (st.step === 3) {
        const hr = h('tr');
        for (let c = 0; c < ncol; c++) {
          const f = st.formats[c] || 'general';
          const th = h('th', { class: c === selCol ? 'on' : '', text: f === 'general' ? 'General' : f === 'text' ? 'Text' : f === 'skip' ? 'Skip Column' : f });
          th.addEventListener('click', () => { selCol = c; sync3(); renderPreview(); });
          hr.appendChild(th);
        }
        tbl.appendChild(hr);
      }
      for (const r of rows) {
        const tr = h('tr');
        for (let c = 0; c < ncol; c++) {
          const td = h('td', { text: r[c] == null ? '' : r[c] });
          if (st.step === 3) { if (c === selCol) td.classList.add('on'); if (st.formats[c] === 'skip') td.classList.add('skip'); td.addEventListener('click', () => { selCol = c; sync3(); renderPreview(); }); }
          tr.appendChild(td);
        }
        tbl.appendChild(tr);
      }
      preview.appendChild(tbl);
    }
    let sync3 = () => {};
    function step1() {
      title.textContent = 'Text Import Wizard - Step 1 of 3';
      content.textContent = '';
      content.appendChild(h('p', { text: 'The Text Wizard has determined that your data is ' + (st.fixed ? 'Fixed Width.' : 'Delimited.') }));
      content.appendChild(h('p', { text: 'If this is correct, choose Next, or choose the data type that best describes your data.' }));
      const g = ui.group('Original data type',
        h('div', { class: 'tiw-dt' }, ui.radio('tiw-type', '&Delimited', !st.fixed, () => { st.fixed = false; renderPreview(); }), h('span', { class: 'tiw-hint', text: '- Characters such as commas or tabs separate each field.' })),
        h('div', { class: 'tiw-dt' }, ui.radio('tiw-type', 'Fixed &width', st.fixed, () => { st.fixed = true; renderPreview(); }), h('span', { class: 'tiw-hint', text: '- Fields are aligned in columns with spaces between each field.' })));
      content.appendChild(g);
      const start = ui.spin({ value: st.start, min: 1, max: Math.max(1, lines().length), step: 1, dec: 0, onChange: (v) => { st.start = Math.round(v); renderPreview(); } });
      const encSel = ui.select([['auto', 'Detect automatically (' + decoded.encoding + ')'], ['utf-8', 'Unicode (UTF-8)'], ['windows-1252', 'Windows (ANSI)'], ['utf-16le', 'Unicode (UTF-16 LE)'], ['utf-16be', 'Unicode (UTF-16 BE)'], ['iso-8859-1', 'Western European (ISO)'], ['shift_jis', 'Japanese (Shift-JIS)'], ['gb18030', 'Chinese Simplified (GB18030)'], ['big5', 'Chinese Traditional (Big5)'], ['euc-kr', 'Korean (EUC-KR)'], ['windows-1251', 'Cyrillic (Windows)'], ['ibm866', 'MS-DOS (OEM 866)']], enc, (v) => { enc = v; decoded = CSV.decode(bytes, v === 'auto' ? null : v); renderPreview(); });
      content.appendChild(h('div', { class: 'row' }, ui.field('Start import at &row:', start), ui.field('File &origin:', encSel)));
      renderPreview();
    }
    function step2() {
      title.textContent = 'Text Import Wizard - Step 2 of 3';
      content.textContent = '';
      if (st.fixed) {
        if (!st.breaks) st.breaks = CSV.guessBreaks(decoded.text);
        content.appendChild(h('p', { text: 'This screen lets you set field widths (column breaks). Lines with arrows signify a column break. To CREATE a break line, click at the desired position. To DELETE a break line, click on it.' }));
        renderPreview();
        return;
      }
      content.appendChild(h('p', { text: 'This screen lets you set the delimiters your data contains. You can see how your text is affected in the preview below.' }));
      const d = st.delims;
      const other = h('input', { type: 'text', maxlength: '1', value: d.otherCh, style: 'width:24px' });
      other.addEventListener('input', () => { d.otherCh = other.value; d.other = !!other.value; oc.input.checked = d.other; renderPreview(); });
      const oc = ui.check('&Other:', d.other, (v) => { d.other = v; renderPreview(); });
      const g = ui.group('Delimiters',
        h('div', { class: 'tiw-delims' }, ui.check('&Tab', d.tab, (v) => { d.tab = v; renderPreview(); }), ui.check('Se&micolon', d.semi, (v) => { d.semi = v; renderPreview(); }), ui.check('&Comma', d.comma, (v) => { d.comma = v; renderPreview(); }), ui.check('&Space', d.space, (v) => { d.space = v; renderPreview(); }), h('span', null, oc, other)));
      content.appendChild(h('div', { class: 'row' }, g, h('div', { class: 'col' }, ui.check('Treat consecutive delimiters as &one', st.merge, (v) => { st.merge = v; renderPreview(); }),
        ui.field('Text &qualifier:', ui.select([['"', '"'], ["'", "'"], ['', '{none}']], st.quote, (v) => { st.quote = v; renderPreview(); })))));
      renderPreview();
    }
    function step3() {
      title.textContent = 'Text Import Wizard - Step 3 of 3';
      content.textContent = '';
      content.appendChild(h('p', { text: "This screen lets you select each column and set the Data Format. 'General' converts numeric values to numbers, date values to dates, and all remaining values to text." }));
      const radios = [];
      const dateSel = ui.select(['MDY', 'DMY', 'YMD', 'MYD', 'DYM', 'YDM'], 'MDY', (v) => { if (/^[MDY]{3}$/.test(st.formats[selCol] || '')) { st.formats[selCol] = v; renderPreview(); } });
      const set = (f) => { st.formats[selCol] = f === 'date' ? dateSel.value : f; renderPreview(); };
      radios.push(ui.radio('tiw-fmt', '&General', true, () => set('general')));
      radios.push(ui.radio('tiw-fmt', '&Text', false, () => set('text')));
      const dr = ui.radio('tiw-fmt', '&Date:', false, () => set('date'));
      radios.push(dr);
      radios.push(ui.radio('tiw-fmt', 'Do not import column (s&kip)', false, () => set('skip')));
      sync3 = () => { const f = st.formats[selCol] || 'general'; radios[0].input.checked = f === 'general'; radios[1].input.checked = f === 'text'; radios[2].input.checked = /^[MDY]{3}$/.test(f); radios[3].input.checked = f === 'skip'; if (/^[MDY]{3}$/.test(f)) dateSel.value = f; };
      content.appendChild(ui.group('Column data format', radios[0], radios[1], h('div', { class: 'row' }, dr, dateSel), radios[3]));
      sync3();
      renderPreview();
    }
    const steps = [step1, step2, step3];
    return new Promise((resolve) => {
      let result = null;
      const d = ui.dialog({
        title: 'Text Import Wizard - Step 1 of 3', body, width: 620,
        buttons: [
          { label: 'Cancel' },
          { label: '< &Back', onClick: () => { if (st.step > 1) { st.step--; steps[st.step - 1](); upd(); } return false; } },
          { label: '&Next >', onClick: () => { if (st.step < 3) { st.step++; steps[st.step - 1](); upd(); } return false; } },
          { label: '&Finish', primary: true, onClick: () => {
            const rows = parsedRows();
            const wb = CSV.toWorkbook(rows, { name, startRow: st.start - 1, colFormats: st.formats });
            wb.textImport = { encoding: decoded.encoding };
            result = { wb };
          } },
        ],
      });
      const upd = () => { d.el.querySelector('.dlg-ttl').textContent = title.textContent; d.buttons[1].disabled = st.step === 1; d.buttons[2].disabled = st.step === 3; };
      step1(); upd();
      d.done.then(() => resolve(result));
    });
  };

  /* ------------------------------------------------------------ export */
  /** text a cell shows, without column-width effects (what CSV export writes) */
  CSV.cellText = function (sh, r, c) {
    const cell = sh.get(r, c);
    if (!cell) return '';
    let v = cell.v;
    if (cell.dirty && L.calc && L.calc.cellValue) v = L.calc.cellValue(sh, r, c);
    if (v == null) return '';
    if (M.isErr(v)) return v.e;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    const st = sh.wb.styles.get(L.layout && L.layout.styleIdOf ? L.layout.styleIdOf(sh, r, c, cell) : cell.s || 0);
    const nf = st.nf || 'General';
    if (typeof v === 'number') {
      if (/^General$/i.test(nf)) return generalCSV(v);
      try { return NF.format(nf, v, { date1904: sh.wb.date1904 }).text.replace(/ | /g, ' ').replace(/^\s+|\s+$/g, (m) => (/_|\*/.test(nf) ? '' : m)); } catch (e) { return generalCSV(v); }
    }
    if (nf && /@/.test(nf) && !/^@$/.test(nf)) { try { return NF.format(nf, v).text; } catch (e) { /* ignore */ } }
    return String(v);
  };
  function generalCSV(v) {
    if (v === 0) return '0';
    if (!isFinite(v)) return '#NUM!';
    const a = Math.abs(v);
    if (a >= 1e15 || a < 1e-9) { const e = Number(v.toPrecision(15)).toExponential(); const [m, x] = e.split('e'); return m + 'E' + (x[0] === '-' ? '-' : '+') + x.replace(/^[+-]/, '').padStart(2, '0'); }
    return String(Number(v.toPrecision(15)));
  }
  CSV.generalCSV = generalCSV;
  /** one sheet → delimited text */
  CSV.text = function (sh, delim, o) {
    o = o || {};
    delim = delim || ',';
    const lines = [];
    const end = { r: sh.maxR, c: sh.maxC };
    /* trailing blank rows / columns are not written */
    let lastR = -1, lastC = -1;
    sh.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cell, c) => { if (cell && (cell.v != null || cell.f != null) && cell.v !== '') { if (r > lastR) lastR = r; if (c > lastC) lastC = c; } }); });
    void end;
    const needQ = new RegExp('[' + delim.replace(/[\]\\^-]/g, '\\$&') + '"\\r\\n]');
    for (let r = 0; r <= lastR; r++) {
      const row = sh.rows[r];
      const out = [];
      let lastNon = -1;
      for (let c = 0; c <= lastC; c++) {
        let t = row && row.cells[c] ? CSV.cellText(sh, r, c) : '';
        if (t !== '') lastNon = c;
        if (needQ.test(t) || (o.quoteSpaces && /^\s|\s$/.test(t))) t = '"' + t.replace(/"/g, '""') + '"';
        out.push(t);
      }
      void lastNon;
      lines.push(out.join(delim));
    }
    return lines.join('\r\n') + (lines.length ? '\r\n' : '');
  };
  CSV.write = function (wb, sh, delim, o) {
    o = o || {};
    const text = CSV.text(sh, delim, o);
    const bom = o.bom === false ? '' : '﻿';
    return new Blob([bom + text], { type: delim === '\t' ? 'text/plain;charset=utf-8' : 'text/csv;charset=utf-8' });
  };

  /* ------------------------------------------------------------ HTML ("Save as Web Page") */
  CSV.htmlSheet = function (wb, sh) {
    const LY = L.layout;
    const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const lastR = Math.min(sh.maxR, 65535), lastC = Math.min(sh.maxC, 255);
    const covered = new Set();
    let html = `<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;table-layout:fixed">`;
    html += '<colgroup>';
    for (let c = 0; c <= lastC; c++) html += `<col style="width:${M.colPx(sh, c)}px">`;
    html += '</colgroup>';
    const cssColor = (c, d) => M.colorHex(wb, c, d);
    const bcss = (b) => (b && b.style ? `${b.style === 'thick' ? 3 : b.style === 'medium' || /^medium/.test(b.style) ? 2 : b.style === 'double' ? 3 : 1}px ${b.style === 'double' ? 'double' : /dash|dot/i.test(b.style) ? (/dot/i.test(b.style) ? 'dotted' : 'dashed') : 'solid'} ${cssColor(b.color, '#000000')}` : '');
    for (let r = 0; r <= lastR; r++) {
      const hpx = M.ptToPx(M.rowPt(sh, r));
      if (!hpx) continue;
      html += `<tr style="height:${hpx}px">`;
      for (let c = 0; c <= lastC; c++) {
        if (covered.has(r * 16384 + c)) continue;
        if (!M.colPx(sh, c)) continue;
        const m = sh.mergeAt(r, c);
        let span = '';
        if (m && m.r1 === r && m.c1 === c) { span = `${m.r2 > m.r1 ? ` rowspan="${m.r2 - m.r1 + 1}"` : ''}${m.c2 > m.c1 ? ` colspan="${m.c2 - m.c1 + 1}"` : ''}`; for (let rr = m.r1; rr <= m.r2; rr++) for (let cc = m.c1; cc <= m.c2; cc++) if (rr !== r || cc !== c) covered.add(rr * 16384 + cc); }
        const cell = sh.get(r, c);
        const st = LY ? LY.styleOf(sh, r, c, cell) : wb.styles.get(cell ? cell.s : 0);
        const css = [];
        const f = st.font || {};
        if (f.name) css.push(`font-family:'${f.name}'`);
        if (f.sz) css.push(`font-size:${f.sz}pt`);
        if (f.b) css.push('font-weight:bold');
        if (f.i) css.push('font-style:italic');
        if (f.u || f.strike) css.push(`text-decoration:${[f.u ? 'underline' : '', f.strike ? 'line-through' : ''].join(' ').trim()}`);
        if (f.color && !f.color.auto) css.push(`color:${cssColor(f.color, '#000000')}`);
        if (st.fill && st.fill.pattern && st.fill.pattern !== 'none') css.push(`background:${cssColor(st.fill.pattern === 'solid' ? st.fill.fg : st.fill.bg || st.fill.fg, '#FFFFFF')}`);
        const b = st.border || {};
        if (b.t) css.push('border-top:' + bcss(b.t));
        if (b.b) css.push('border-bottom:' + bcss(b.b));
        if (b.l) css.push('border-left:' + bcss(b.l));
        if (b.r) css.push('border-right:' + bcss(b.r));
        const a = st.align || {};
        const v = cell ? cell.v : null;
        const hal = a.h && a.h !== 'general' ? (a.h === 'centerContinuous' ? 'center' : a.h === 'distributed' || a.h === 'fill' ? 'left' : a.h) : typeof v === 'number' ? 'right' : typeof v === 'boolean' || M.isErr(v) ? 'center' : 'left';
        css.push('text-align:' + hal);
        css.push('vertical-align:' + (a.v === 'top' ? 'top' : a.v === 'center' ? 'middle' : 'bottom'));
        css.push(a.wrap ? 'white-space:pre-wrap' : 'white-space:nowrap');
        css.push('overflow:hidden');
        if (a.indent) css.push(`padding-left:${a.indent * 9}px`);
        const text = cell ? CSV.cellText(sh, r, c) : '';
        const link = sh.links.find((l) => M.rangeContains(l.ref, r, c));
        let inner = esc(text);
        if (link && link.target && /^(https?:|mailto:)/i.test(link.target)) inner = `<a href="${esc(link.target)}">${inner}</a>`;
        html += `<td${span} style="${css.join(';')}">${inner || '&nbsp;'}</td>`;
      }
      html += '</tr>';
    }
    return html + '</table>';
  };
  CSV.writeHTML = function (wb) {
    const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const visible = wb.sheets.filter((s) => s.state === 'visible' && s.kind !== 'chartsheet');
    let html = `<!DOCTYPE html>\n<html><head><meta charset="utf-8"><meta name="Generator" content="Ledger 2003"><title>${esc((wb.props && wb.props.title) || (wb.fileName || 'Book1').replace(/\.[^.]+$/, ''))}</title>` +
      '<style>body{font-family:Arial,sans-serif;font-size:10pt;margin:8px}h2{font:bold 11pt Tahoma,sans-serif;margin:16px 0 6px}td{padding:0 2px;border:1px solid #DADADA;font-size:10pt}nav a{margin-right:10px}</style></head><body>';
    if (visible.length > 1) html += '<nav>' + visible.map((s, i) => `<a href="#sheet${i + 1}">${esc(s.name)}</a>`).join('') + '</nav>';
    visible.forEach((s, i) => { html += `<h2 id="sheet${i + 1}">${esc(s.name)}</h2>` + CSV.htmlSheet(wb, s); });
    html += '</body></html>';
    return new Blob([html], { type: 'text/html;charset=utf-8' });
  };
})(typeof window !== 'undefined' ? window : globalThis);
