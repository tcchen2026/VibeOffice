/* Ledger — XML Spreadsheet 2003 (SpreadsheetML 2003, *.xml): read and write.
 * Formulas in that format use R1C1 references; they are converted to and from A1 here. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, F = L.formula, NF = L.numfmt;
  const XS = (L.xmlss = {});
  const SS = 'urn:schemas-microsoft-com:office:spreadsheet';

  /* ------------------------------------------------------------ R1C1 ⇄ A1 */
  /** R1C1 formula text (as stored, with leading =) for the cell at (r, c) → A1 text without = */
  XS.fromR1C1 = function (text, r, c) {
    let out = '', i = 0;
    const s = String(text).replace(/^=/, '');
    const part = (kind, m, base) => { if (m == null || m === '') return { v: base, abs: false }; if (m[0] === '[') return { v: base + parseInt(m.slice(1, -1), 10), abs: false }; return { v: parseInt(m, 10) - 1, abs: true }; };
    const cellA1 = (rr, cc) => (cc.abs ? '$' : '') + F.colName(Math.max(0, cc.v)) + (rr.abs ? '$' : '') + (Math.max(0, rr.v) + 1);
    while (i < s.length) {
      const ch = s[i];
      if (ch === '"') { const j = s.indexOf('"', i + 1); let k = j; while (k >= 0 && s[k + 1] === '"') k = s.indexOf('"', k + 2); const e = k < 0 ? s.length : k + 1; out += s.slice(i, e); i = e; continue; }
      if (ch === "'") { const j = s.indexOf("'", i + 1); out += s.slice(i, j < 0 ? s.length : j + 1); i = j < 0 ? s.length : j + 1; continue; }
      const prev = i > 0 ? s[i - 1] : '';
      if (/[RrCc]/.test(ch) && !/[A-Za-z0-9_.\]]/.test(prev)) {
        const m = /^(?:[Rr](\[-?\d+\]|\d+)?)?(?:[Cc](\[-?\d+\]|\d+)?)?/.exec(s.slice(i));
        const tok = m[0];
        const after = s[i + tok.length] || '';
        const hasR = /^[Rr]/.test(tok), hasC = /[Cc]/.test(tok.replace(/^[Rr](\[-?\d+\]|\d+)?/, ''));
        if (tok && !/[A-Za-z0-9_(.]/.test(after) && (hasR || hasC)) {
          const rr = hasR ? part('r', m[1], r) : null, cc = hasC ? part('c', m[2], c) : null;
          if (rr && cc) out += cellA1(rr, cc);
          else if (rr) out += (rr.abs ? '$' : '') + (rr.v + 1) + ':' + (rr.abs ? '$' : '') + (rr.v + 1);
          else out += (cc.abs ? '$' : '') + F.colName(cc.v) + ':' + (cc.abs ? '$' : '') + F.colName(cc.v);
          i += tok.length;
          continue;
        }
      }
      out += ch; i++;
    }
    /* "1:1:3:3" from row ranges R1:R3 → 1:3 */
    out = out.replace(/(\$?\d+):\$?\d+:(\$?\d+):\$?\d+/g, '$1:$2').replace(/(\$?[A-Z]+):\$?[A-Z]+:(\$?[A-Z]+):\$?[A-Z]+/g, '$1:$2');
    return out;
  };
  /** A1 formula (store form, no =) at (r, c) → R1C1 text with = */
  /* token by token, so the author's spacing survives; wb (optional) names external workbooks */
  XS.toR1C1 = function (f, r, c, wb) {
    /* [0]!Name is a workbook-level name in this workbook */
    const src = String(f).replace(/\[0\]!/g, '');
    let toks;
    try { toks = F.tokenize(src); } catch (e) { return '=' + src; }
    const rel = (v, base, abs) => (abs ? String(v + 1) : v === base ? '' : '[' + (v - base) + ']');
    const rc = (rr, ra, cc, ca) => 'R' + rel(rr, r, ra) + 'C' + rel(cc, c, ca);
    const bookName = (b) => {
      const el = wb && wb.externalLinks && wb.externalLinks[+b - 1];
      const t = el && el.relTarget ? String(el.relTarget) : '';
      try { return t ? decodeURIComponent(t.split(/[\\/]/).pop()) : 'Book' + b; } catch (e) { return t.split(/[\\/]/).pop(); }
    };
    let out = '';
    for (let k = 0; k < toks.length; k++) {
      const t = toks[k];
      const raw = src.slice(t.pos, k + 1 < toks.length ? toks[k + 1].pos : src.length);
      if (t.t !== 'ref' || t.kind === 'err' || t.kind === 'name') { out += raw; continue; }
      const sheet = t.sheet ? t.sheet + (t.sheet2 ? ':' + t.sheet2 : '') : '';
      const quoted = (x) => F.quoteSheet(x) !== x;
      let pre = '';
      if (t.book != null) pre = "'[" + bookName(t.book) + ']' + sheet.replace(/'/g, "''") + "'!";
      else if (t.sheet2) pre = (quoted(t.sheet) || quoted(t.sheet2) ? "'" + sheet.replace(/'/g, "''") + "'" : sheet) + '!';
      else if (sheet) pre = F.quoteSheet(sheet) + '!';
      let x;
      if (t.kind === 'cell') x = rc(t.r, t.ra, t.c, t.ca);
      else if (t.kind === 'cols') x = 'C' + rel(t.c1, c, t.ca1) + ':C' + rel(t.c2, c, t.ca2);
      else if (t.kind === 'rows') x = 'R' + rel(t.r1, r, t.ra1) + ':R' + rel(t.r2, r, t.ra2);
      else if (t.kind === 'area') x = rc(t.r1, t.ra1, t.c1, t.ca1) + ':' + rc(t.r2, t.ra2, t.c2, t.ca2);
      else { out += raw; continue; }
      out += pre + x;
    }
    return '=' + out;
  };

  /* ------------------------------------------------------------ reading */
  const NAMED_FMT = { 'General': 'General', 'General Number': 'General', 'General Date': 'm/d/yyyy h:mm', 'Long Date': 'dddd, mmmm dd, yyyy', 'Medium Date': 'dd-mmm-yy', 'Short Date': 'm/d/yyyy', 'Long Time': 'h:mm:ss AM/PM', 'Medium Time': 'h:mm AM/PM', 'Short Time': 'h:mm', 'Currency': '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', 'Euro Currency': '[$€-2] #,##0.00', 'Fixed': '0.00', 'Standard': '#,##0.00', 'Percent': '0.00%', 'Scientific': '0.00E+00', 'Yes/No': '"Yes";"Yes";"No"', 'True/False': '"True";"True";"False"', 'On/Off': '"On";"On";"Off"', '@': '@' };
  const LINE = { Continuous: { 1: 'thin', 2: 'medium', 3: 'thick', 0: 'hair' }, Dash: { 1: 'dashed', 2: 'mediumDashed' }, Dot: { 1: 'dotted' }, DashDot: { 1: 'dashDot', 2: 'mediumDashDot' }, DashDotDot: { 1: 'dashDotDot', 2: 'mediumDashDotDot' }, SlantDashDot: { 2: 'slantDashDot' }, Double: { 3: 'double' } };
  const PATTERNS = { Solid: 'solid', Gray75: 'darkGray', Gray50: 'mediumGray', Gray25: 'lightGray', Gray125: 'gray125', Gray0625: 'gray0625', HorzStripe: 'darkHorizontal', VertStripe: 'darkVertical', ReverseDiagStripe: 'darkDown', DiagStripe: 'darkUp', DiagCross: 'darkGrid', ThickDiagCross: 'darkTrellis', ThinHorzStripe: 'lightHorizontal', ThinVertStripe: 'lightVertical', ThinReverseDiagStripe: 'lightDown', ThinDiagStripe: 'lightUp', ThinHorzCross: 'lightGrid', ThinDiagCross: 'lightTrellis' };
  const at = (el, n) => { if (!el || !el.attrs) return null; if (el.attrs['ss:' + n] != null) return el.attrs['ss:' + n]; if (el.attrs[n] != null) return el.attrs[n]; for (const k in el.attrs) if (k.endsWith(':' + n)) return el.attrs[k]; return null; };
  const kids = (el, n) => (el ? el.children.filter((c) => !n || c.localName === n) : []);
  const kid = (el, n) => (el ? el.children.find((c) => c.localName === n) || null : null);
  const hexColor = (v) => (v && /^#[0-9a-f]{6}$/i.test(v) ? M.rgb(v) : undefined);
  function readStyle(el, parent) {
    const st = parent ? JSON.parse(JSON.stringify(parent)) : M.defaultStyle({ name: 'Arial', sz: 10 });
    for (const c of el.children) {
      switch (c.localName) {
        case 'Font': {
          const f = st.font = Object.assign({}, st.font);
          if (at(c, 'FontName')) f.name = at(c, 'FontName');
          if (at(c, 'Size')) f.sz = +at(c, 'Size');
          if (at(c, 'Bold') === '1') f.b = true;
          if (at(c, 'Italic') === '1') f.i = true;
          if (at(c, 'StrikeThrough') === '1') f.strike = true;
          const u = at(c, 'Underline'); if (u && u !== 'None') f.u = { Single: 'single', Double: 'double', SingleAccounting: 'singleAccounting', DoubleAccounting: 'doubleAccounting' }[u] || 'single';
          const v = at(c, 'VerticalAlign'); if (v === 'Superscript') f.vert = 'superscript'; else if (v === 'Subscript') f.vert = 'subscript';
          const col = hexColor(at(c, 'Color')); if (col) f.color = col;
          break;
        }
        case 'Interior': {
          const pat = at(c, 'Pattern');
          if (!pat || pat === 'None') { if (at(c, 'Color')) st.fill = { pattern: 'solid', fg: hexColor(at(c, 'Color')) }; break; }
          const p = PATTERNS[pat] || 'solid';
          st.fill = p === 'solid' ? { pattern: 'solid', fg: hexColor(at(c, 'Color')) || M.rgb('#FFFFFF') } : { pattern: p, fg: hexColor(at(c, 'PatternColor')) || M.rgb('#000000'), bg: hexColor(at(c, 'Color')) || M.rgb('#FFFFFF') };
          break;
        }
        case 'Borders': {
          const b = {};
          for (const bd of kids(c, 'Border')) {
            const pos = { Top: 't', Bottom: 'b', Left: 'l', Right: 'r', DiagonalLeft: 'dd', DiagonalRight: 'du' }[at(bd, 'Position')];
            if (!pos) continue;
            const ls = at(bd, 'LineStyle') || 'Continuous';
            if (ls === 'None') continue;
            const wt = +(at(bd, 'Weight') || 1);
            const style = (LINE[ls] || LINE.Continuous)[wt] || (LINE[ls] || LINE.Continuous)[1] || 'thin';
            const line = { style };
            const col = hexColor(at(bd, 'Color')); if (col) line.color = col;
            if (pos === 'dd' || pos === 'du') { b.d = line; b[pos] = true; } else b[pos] = line;
          }
          st.border = Object.keys(b).length ? b : null;
          break;
        }
        case 'Alignment': {
          const a = Object.assign({}, st.align || {});
          const hh = at(c, 'Horizontal'); if (hh && hh !== 'Automatic') a.h = { Left: 'left', Center: 'center', Right: 'right', Fill: 'fill', Justify: 'justify', CenterAcrossSelection: 'centerContinuous', Distributed: 'distributed' }[hh];
          const vv = at(c, 'Vertical'); if (vv && vv !== 'Automatic' && vv !== 'Bottom') a.v = { Top: 'top', Center: 'center', Justify: 'justify', Distributed: 'distributed' }[vv];
          if (at(c, 'WrapText') === '1') a.wrap = true;
          if (at(c, 'ShrinkToFit') === '1') a.shrink = true;
          if (at(c, 'Indent')) a.indent = +at(c, 'Indent');
          if (at(c, 'Rotate')) { const r = +at(c, 'Rotate'); a.rot = r < 0 ? 90 - r : r; }
          if (at(c, 'VerticalText') === '1') a.rot = 255;
          st.align = Object.keys(a).length ? a : null;
          break;
        }
        case 'NumberFormat': { const f = at(c, 'Format'); if (f) st.nf = NAMED_FMT[f] || f.replace(/&quot;/g, '"'); break; }
        case 'Protection': { const p = {}; if (at(c, 'Protected') === '0') p.locked = false; if (at(c, 'HideFormula') === '1') p.hidden = true; st.prot = Object.keys(p).length ? p : null; break; }
        default: break;
      }
    }
    return st;
  }
  XS.read = async function (bytes) {
    const text = L.csv ? L.csv.decode(bytes).text : new TextDecoder().decode(bytes);
    if (!/urn:schemas-microsoft-com:office:spreadsheet/.test(text)) throw Object.assign(new Error('This file is not an XML Spreadsheet 2003 workbook.'), { code: 'format' });
    const doc = L.xml.parse(text.replace(/^[\s\S]*?(<Workbook)/, '$1'));
    const wb = new M.Workbook({ font: { name: 'Arial', sz: 10 } });
    /* styles */
    const styles = new Map();
    const stEls = kids(kid(doc, 'Styles'), 'Style');
    const byId = new Map(stEls.map((e) => [at(e, 'ID'), e]));
    const resolve = (id, depth) => {
      if (styles.has(id)) return styles.get(id);
      const el = byId.get(id);
      if (!el || depth > 10) return 0;
      const parent = at(el, 'Parent') ? wb.styles.get(resolve(at(el, 'Parent'), depth + 1)) : byId.has('Default') && id !== 'Default' ? wb.styles.get(resolve('Default', depth + 1)) : null;
      const idx = wb.styles.add(readStyle(el, parent));
      styles.set(id, idx);
      return idx;
    };
    if (byId.has('Default')) { const d = readStyle(byId.get('Default'), null); wb.styles = new M.StyleTable(d); styles.set('Default', 0); }
    for (const id of byId.keys()) resolve(id, 0);
    const docProps = kid(doc, 'DocumentProperties');
    if (docProps) { const g = (n) => { const e = kid(docProps, n); return e ? e.textContent : ''; }; Object.assign(wb.props, { title: g('Title'), subject: g('Subject'), creator: g('Author'), keywords: g('Keywords'), description: g('Description'), lastModifiedBy: g('LastAuthor'), company: g('Company'), manager: g('Manager'), category: g('Category'), created: g('Created') || null }); }
    const ewb = kid(doc, 'ExcelWorkbook');
    if (ewb && kid(ewb, 'Date1904')) wb.date1904 = true;
    const names = [];
    const nameEls = (el) => kids(kid(el, 'Names'), 'NamedRange');
    for (const ws of kids(doc, 'Worksheet')) {
      const sh = wb.addSheet(at(ws, 'Name') || wb.nextSheetName());
      const table = kid(ws, 'Table');
      if (table) {
        if (at(table, 'DefaultColumnWidth')) sh.defColW = M.pxToWidth(Math.round(+at(table, 'DefaultColumnWidth') / 0.75), M.mdw(wb));
        if (at(table, 'DefaultRowHeight')) sh.defRowH = +at(table, 'DefaultRowHeight');
        let col = -1;
        for (const ce of kids(table, 'Column')) {
          col = at(ce, 'Index') ? +at(ce, 'Index') - 1 : col + 1;
          const span = +(at(ce, 'Span') || 0);
          const o = {};
          if (at(ce, 'Width')) { o.w = M.pxToWidth(Math.round(+at(ce, 'Width') / 0.75), M.mdw(wb)); o.custom = at(ce, 'AutoFitWidth') !== '1'; }
          if (at(ce, 'Hidden') === '1') o.hidden = true;
          if (at(ce, 'StyleID')) o.s = styles.get(at(ce, 'StyleID')) || undefined;
          for (let k = 0; k <= span; k++) sh.cols[col + k] = Object.assign({}, o);
          col += span;
        }
        let r = -1;
        for (const re of kids(table, 'Row')) {
          r = at(re, 'Index') ? +at(re, 'Index') - 1 : r + 1;
          const span = +(at(re, 'Span') || 0);
          const row = sh.rowObj(r);
          if (at(re, 'Height')) { row.ht = +at(re, 'Height'); if (at(re, 'AutoFitHeight') === '0') row.customHeight = true; }
          if (at(re, 'Hidden') === '1') row.hidden = true;
          if (at(re, 'StyleID')) row.s = styles.get(at(re, 'StyleID')) || undefined;
          let c = -1;
          for (const ce of kids(re, 'Cell')) {
            c = at(ce, 'Index') ? +at(ce, 'Index') - 1 : c + 1;
            const cell = sh.cell(r, c);
            if (at(ce, 'StyleID')) { const s = styles.get(at(ce, 'StyleID')); if (s) cell.s = s; }
            const d = kid(ce, 'Data');
            if (d) {
              const t = at(d, 'Type');
              const raw = d.textContent;
              if (t === 'Number') cell.v = +raw;
              else if (t === 'Boolean') cell.v = raw === '1' || /^true$/i.test(raw);
              else if (t === 'Error') cell.v = M.err(raw);
              else if (t === 'DateTime') { const m = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d(?:\.\d+)?))?)?/.exec(raw); cell.v = m ? NF.dateToSerial(+m[1], +m[2], +m[3], wb.date1904) + (m[4] ? (+m[4] * 3600 + +m[5] * 60 + (+m[6] || 0)) / 86400 : 0) : raw; }
              else cell.v = raw;
            }
            const f = at(ce, 'Formula');
            if (f) { try { const a1 = XS.fromR1C1(f, r, c); F.parse(a1); cell.f = F.toStore(a1); if (cell.v == null) cell.dirty = true; } catch (e) { /* keep the value */ } }
            const ma = +(at(ce, 'MergeAcross') || 0), md = +(at(ce, 'MergeDown') || 0);
            if (ma || md) sh.merges.push({ r1: r, c1: c, r2: r + md, c2: c + ma });
            const href = at(ce, 'HRef');
            if (href) sh.links.push(Object.assign({ ref: { r1: r, c1: c, r2: r, c2: c } }, /^#/.test(href) ? { location: href.slice(1) } : { target: href }));
            const cm = kid(ce, 'Comment');
            if (cm) sh.comments.set(M.key(r, c), { r, c, author: at(cm, 'Author') || '', text: (kid(cm, 'Data') || cm).textContent.trim(), visible: at(cm, 'ShowAlways') === '1' });
            if (cell.v == null && cell.f == null && !cell.s) { delete row.cells[c]; }
            c += ma;
          }
          for (let k = 1; k <= span; k++) Object.assign(sh.rowObj(r + k), { ht: row.ht, customHeight: row.customHeight, hidden: row.hidden, s: row.s });
          r += span;
        }
      }
      for (const n of nameEls(ws)) names.push({ el: n, scope: wb.sheets.length - 1 });
      const opts = kid(ws, 'WorksheetOptions');
      if (opts) {
        if (kid(opts, 'FreezePanes') || kid(opts, 'FrozenNoSplit')) { const ys = +((kid(opts, 'SplitHorizontal') || {}).textContent || 0), xs = +((kid(opts, 'SplitVertical') || {}).textContent || 0); if (xs || ys) sh.view.freeze = { r: ys, c: xs, top: { r: +((kid(opts, 'TopRowBottomPane') || {}).textContent || ys), c: +((kid(opts, 'LeftColumnRightPane') || {}).textContent || xs) } }; }
        if (kid(opts, 'DoNotDisplayGridlines')) sh.view.grid = false;
        if (kid(opts, 'DoNotDisplayHeadings')) sh.view.headings = false;
        if (kid(opts, 'DoNotDisplayZeros')) sh.view.zeros = false;
        if (kid(opts, 'Zoom')) sh.view.zoom = +kid(opts, 'Zoom').textContent || 100;
        if (kid(opts, 'Selected')) wb.active = wb.sheets.length - 1;
        if (kid(opts, 'Visible') && /Hidden/.test(kid(opts, 'Visible').textContent)) sh.state = /VeryHidden/.test(kid(opts, 'Visible').textContent) ? 'veryHidden' : 'hidden';
        const ps = kid(opts, 'PageSetup');
        if (ps) { const lay = kid(ps, 'Layout'); if (lay && at(lay, 'Orientation') === 'Landscape') sh.print.orientation = 'landscape'; const pm = kid(ps, 'PageMargins'); if (pm) sh.print.margins = { l: +(at(pm, 'Left') || 0.75), r: +(at(pm, 'Right') || 0.75), t: +(at(pm, 'Top') || 1), b: +(at(pm, 'Bottom') || 1), header: +(at(kid(ps, 'Header'), 'Margin') || 0.5), footer: +(at(kid(ps, 'Footer'), 'Margin') || 0.5) }; const hd = kid(ps, 'Header'), ft = kid(ps, 'Footer'); if (hd && at(hd, 'Data')) sh.print.header = at(hd, 'Data'); if (ft && at(ft, 'Data')) sh.print.footer = at(ft, 'Data'); }
        const pr = kid(opts, 'Print');
        if (pr) { if (kid(pr, 'Gridlines')) sh.print.gridLines = true; if (kid(pr, 'RowColHeadings')) sh.print.headings = true; if (kid(pr, 'Scale')) sh.print.scale = +kid(pr, 'Scale').textContent || 100; if (kid(pr, 'PaperSizeIndex')) sh.print.paper = +kid(pr, 'PaperSizeIndex').textContent || 1; }
        if (kid(opts, 'ProtectContents') && kid(opts, 'ProtectContents').textContent === 'True') sh.protection = { attrs: { sheet: '1' } };
      }
      const af = kid(ws, 'AutoFilter');
      if (af && at(af, 'Range')) { const a1 = XS.fromR1C1(at(af, 'Range'), 0, 0); const rg = F.parseRange(a1.replace(/\$/g, '')); if (rg) sh.autoFilter = { ref: rg, cols: [] }; }
      for (const dv of kids(ws, 'DataValidation')) {
        const rng = kid(dv, 'Range');
        if (!rng) continue;
        const ranges = rng.textContent.split(',').map((x) => F.parseRange(XS.fromR1C1(x.trim(), 0, 0).replace(/\$/g, ''))).filter(Boolean);
        const type = (kid(dv, 'Type') || {}).textContent || 'Any';
        const o = { ranges, type: { List: 'list', Whole: 'whole', Decimal: 'decimal', Date: 'date', Time: 'time', TextLength: 'textLength', Custom: 'custom' }[type] || 'none', allowBlank: true, showDrop: true, showInput: true, showError: true };
        const val = kid(dv, 'Value'), mn = kid(dv, 'Min'), mx = kid(dv, 'Max');
        const conv = (t) => (t == null ? undefined : XS.fromR1C1(t, ranges[0] ? ranges[0].r1 : 0, ranges[0] ? ranges[0].c1 : 0));
        if (val) o.f1 = conv(val.textContent); if (mn) o.f1 = conv(mn.textContent); if (mx) o.f2 = conv(mx.textContent);
        const q = kid(dv, 'Qualifier'); if (q) o.op = { Equal: 'equal', NotEqual: 'notEqual', Greater: 'greaterThan', Less: 'lessThan', GreaterOrEqual: 'greaterThanOrEqual', LessOrEqual: 'lessThanOrEqual', NotBetween: 'notBetween' }[q.textContent] || 'between';
        if (kid(dv, 'InputMessage')) o.prompt = kid(dv, 'InputMessage').textContent;
        if (kid(dv, 'ErrorMessage')) o.error = kid(dv, 'ErrorMessage').textContent;
        sh.dv.push(o);
      }
    }
    for (const n of nameEls(doc)) names.push({ el: n, scope: null });
    for (const { el, scope } of names) {
      const name = at(el, 'Name');
      const ref = at(el, 'RefersTo');
      if (!name || !ref) continue;
      let a1 = XS.fromR1C1(ref, 0, 0);
      if (/^Print_Area$/i.test(name)) wb.names.push({ name: '_xlnm.Print_Area', ref: a1, scope });
      else if (/^Print_Titles$/i.test(name)) wb.names.push({ name: '_xlnm.Print_Titles', ref: a1, scope });
      else if (!/^_FilterDatabase$/i.test(name)) wb.names.push({ name, ref: a1, scope, hidden: at(el, 'Hidden') === '1' || undefined });
    }
    if (!wb.sheets.length) wb.addSheet('Sheet1');
    if (wb.sheets[wb.active].state !== 'visible') wb.active = Math.max(0, wb.sheets.findIndex((s) => s.state === 'visible'));
    return wb;
  };

  /* ------------------------------------------------------------ writing */
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/\r?\n/g, '&#10;');
  const escT = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const hex = (wb, c, d) => { const x = M.colorHex(wb, c, d); return x ? x.toUpperCase() : null; };
  const RLINE = { thin: ['Continuous', 1], medium: ['Continuous', 2], thick: ['Continuous', 3], hair: ['Continuous', 0], dashed: ['Dash', 1], mediumDashed: ['Dash', 2], dotted: ['Dot', 1], dashDot: ['DashDot', 1], mediumDashDot: ['DashDot', 2], dashDotDot: ['DashDotDot', 1], mediumDashDotDot: ['DashDotDot', 2], slantDashDot: ['SlantDashDot', 2], double: ['Double', 3] };
  const RPAT = Object.fromEntries(Object.entries(PATTERNS).map(([k, v]) => [v, k]));
  function styleXml(wb, st, id) {
    let x = `<Style ss:ID="${id}"${id === 'Default' ? ' ss:Name="Normal"' : ''}>`;
    const a = st.align || {};
    const H = { left: 'Left', center: 'Center', right: 'Right', fill: 'Fill', justify: 'Justify', centerContinuous: 'CenterAcrossSelection', distributed: 'Distributed' };
    const V = { top: 'Top', center: 'Center', justify: 'Justify', distributed: 'Distributed' };
    x += `<Alignment${a.h ? ` ss:Horizontal="${H[a.h] || 'Left'}"` : ''} ss:Vertical="${V[a.v] || 'Bottom'}"${a.wrap ? ' ss:WrapText="1"' : ''}${a.shrink ? ' ss:ShrinkToFit="1"' : ''}${a.indent ? ` ss:Indent="${a.indent}"` : ''}${a.rot === 255 ? ' ss:VerticalText="1"' : a.rot ? ` ss:Rotate="${a.rot > 90 ? 90 - a.rot : a.rot}"` : ''}/>`;
    const b = st.border;
    if (b) {
      x += '<Borders>';
      for (const [k, pos] of [['b', 'Bottom'], ['l', 'Left'], ['r', 'Right'], ['t', 'Top']]) if (b[k]) { const [ls, wt] = RLINE[b[k].style] || RLINE.thin; x += `<Border ss:Position="${pos}" ss:LineStyle="${ls}" ss:Weight="${wt}"${b[k].color ? ` ss:Color="${hex(wb, b[k].color, '#000000')}"` : ''}/>`; }
      if (b.d) { const [ls, wt] = RLINE[b.d.style] || RLINE.thin; if (b.dd) x += `<Border ss:Position="DiagonalLeft" ss:LineStyle="${ls}" ss:Weight="${wt}"/>`; if (b.du) x += `<Border ss:Position="DiagonalRight" ss:LineStyle="${ls}" ss:Weight="${wt}"/>`; }
      x += '</Borders>';
    } else x += '<Borders/>';
    const f = st.font || {};
    x += `<Font ss:FontName="${esc(L.layout ? L.layout.fontName(wb, f) : f.name || 'Arial')}" x:Family="Swiss" ss:Size="${f.sz || 10}"${f.color && !f.color.auto ? ` ss:Color="${hex(wb, f.color, '#000000')}"` : ''}${f.b ? ' ss:Bold="1"' : ''}${f.i ? ' ss:Italic="1"' : ''}${f.strike ? ' ss:StrikeThrough="1"' : ''}${f.u ? ` ss:Underline="${{ single: 'Single', double: 'Double', singleAccounting: 'SingleAccounting', doubleAccounting: 'DoubleAccounting' }[f.u] || 'Single'}"` : ''}${f.vert ? ` ss:VerticalAlign="${f.vert === 'superscript' ? 'Superscript' : 'Subscript'}"` : ''}/>`;
    const fl = st.fill;
    if (fl && fl.pattern && fl.pattern !== 'none') x += fl.pattern === 'solid' ? `<Interior ss:Color="${hex(wb, fl.fg, '#FFFFFF')}" ss:Pattern="Solid"/>` : `<Interior ss:Color="${hex(wb, fl.bg, '#FFFFFF')}" ss:Pattern="${RPAT[fl.pattern] || 'Solid'}" ss:PatternColor="${hex(wb, fl.fg, '#000000')}"/>`;
    else x += '<Interior/>';
    x += st.nf && st.nf !== 'General' ? `<NumberFormat ss:Format="${esc(st.nf)}"/>` : '<NumberFormat/>';
    x += st.prot && (st.prot.locked === false || st.prot.hidden) ? `<Protection${st.prot.locked === false ? ' ss:Protected="0"' : ''}${st.prot.hidden ? ' x:HideFormula="1"' : ''}/>` : '<Protection/>';
    return x + '</Style>';
  }
  const isoOf = (v, d1904) => { const p = NF.serialToParts(v, d1904, 3); const pad = (n, k) => String(n).padStart(k || 2, '0'); return `${pad(p.y, 4)}-${pad(p.m)}-${pad(p.d)}T${pad(p.H)}:${pad(p.M)}:${pad(p.S)}.${pad(Math.round((p.sub || 0) * 1000), 3)}`; };
  XS.text = function (wb) {
    const used = new Set([0]);
    for (const s of wb.sheets) s.rows.forEach((row) => { if (row) { if (row.s) used.add(row.s); row.cells.forEach((c) => { if (c && c.s) used.add(c.s); }); } });
    const sid = (i) => (i ? 's' + (20 + i) : 'Default');
    const out = ['<?xml version="1.0"?>\r\n<?mso-application progid="Excel.Sheet"?>\r\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">'];
    const p = wb.props || {};
    out.push('<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">' + [['Title', p.title], ['Subject', p.subject], ['Author', p.creator], ['Keywords', p.keywords], ['Description', p.description], ['LastAuthor', p.lastModifiedBy || p.creator], ['Category', p.category], ['Manager', p.manager], ['Company', p.company]].filter(([, v]) => v).map(([k, v]) => `<${k}>${escT(v)}</${k}>`).join('') + `<Created>${new Date(p.created || Date.now()).toISOString().replace(/\.\d+Z$/, 'Z')}</Created><Version>11.5606</Version></DocumentProperties>`);
    out.push(`<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel"><ActiveSheet>${wb.active}</ActiveSheet>${wb.date1904 ? '<Date1904/>' : ''}<ProtectStructure>${wb.protection && wb.protection.structure ? 'True' : 'False'}</ProtectStructure><ProtectWindows>False</ProtectWindows></ExcelWorkbook>`);
    out.push('<Styles>' + Array.from(used).sort((a, b) => a - b).map((i) => styleXml(wb, wb.styles.get(i), sid(i))).join('') + '</Styles>');
    const gnames = wb.names.filter((n) => n.scope == null && !/^_xlnm\./i.test(n.name));
    if (gnames.length) out.push('<Names>' + gnames.map((n) => `<NamedRange ss:Name="${esc(n.name)}" ss:RefersTo="${esc(XS.toR1C1(String(n.ref).replace(/^=/, ''), 0, 0, wb))}"${n.hidden ? ' ss:Hidden="1"' : ''}/>`).join('') + '</Names>');
    wb.sheets.forEach((sh, si) => {
      if (sh.kind === 'chartsheet') return;
      out.push(`<Worksheet ss:Name="${esc(sh.name)}"${sh.protection ? ' ss:Protected="1"' : ''}>`);
      const lnames = wb.names.filter((n) => n.scope === si);
      if (lnames.length) out.push('<Names>' + lnames.map((n) => `<NamedRange ss:Name="${esc(n.name.replace(/^_xlnm\./i, ''))}" ss:RefersTo="${esc(XS.toR1C1(String(n.ref).replace(/^=/, ''), 0, 0, wb))}"/>`).join('') + '</Names>');
      const maxR = Math.max(0, sh.maxR), maxC = Math.max(0, sh.maxC);
      out.push(`<Table ss:ExpandedColumnCount="${maxC + 1}" ss:ExpandedRowCount="${maxR + 1}" x:FullColumns="1" x:FullRows="1" ss:DefaultColumnWidth="${(M.defaultColPx(sh) * 0.75).toFixed(2)}" ss:DefaultRowHeight="${M.defaultRowPt(sh)}">`);
      let lastC = -1;
      sh.cols.forEach((o, c) => {
        if (!o || c > Math.max(maxC, 255)) return;
        out.push(`<Column${c !== lastC + 1 ? ` ss:Index="${c + 1}"` : ''}${o.s && used.has(o.s) ? ` ss:StyleID="${sid(o.s)}"` : ''}${o.hidden ? ' ss:Hidden="1"' : ''} ss:AutoFitWidth="0" ss:Width="${(M.colPx(sh, c) * 0.75 || M.defaultColPx(sh) * 0.75).toFixed(2)}"/>`);
        lastC = c;
      });
      const merges = new Map(sh.merges.map((m) => [M.key(m.r1, m.c1), m]));
      const covered = new Set();
      for (const m of sh.merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(M.key(r, c));
      let lastR = -1;
      sh.rows.forEach((row, r) => {
        if (!row) return;
        const cells = [];
        let lc = -1;
        row.cells.forEach((cell, c) => {
          if (!cell || covered.has(M.key(r, c))) return;
          const m = merges.get(M.key(r, c));
          const link = sh.links.find((l) => l.ref.r1 === r && l.ref.c1 === c);
          const cm = sh.comments.get(M.key(r, c));
          if (cell.v == null && cell.f == null && !cell.s && !m && !cm) return;
          let x = `<Cell${c !== lc + 1 ? ` ss:Index="${c + 1}"` : ''}${m ? `${m.c2 > m.c1 ? ` ss:MergeAcross="${m.c2 - m.c1}"` : ''}${m.r2 > m.r1 ? ` ss:MergeDown="${m.r2 - m.r1}"` : ''}` : ''}${cell.s && used.has(cell.s) ? ` ss:StyleID="${sid(cell.s)}"` : ''}${cell.f != null && !cell.am ? ` ss:Formula="${esc(XS.toR1C1(cell.f, r, c, wb))}"${cell.af ? ` ss:ArrayRange="${esc(XS.toR1C1(F.rangeName(cell.af), r, c).slice(1))}"` : ''}` : ''}${link ? ` ss:HRef="${esc(link.target || '#' + link.location)}"` : ''}>`;
          const v = cell.v;
          if (v != null) {
            const st = wb.styles.get(cell.s || 0);
            if (typeof v === 'number') x += NF.isDate(st.nf || 'General') && v >= 0 && v < 2958466 ? `<Data ss:Type="DateTime">${isoOf(v, wb.date1904)}</Data>` : `<Data ss:Type="Number">${isFinite(v) ? v : 0}</Data>`;
            else if (typeof v === 'boolean') x += `<Data ss:Type="Boolean">${v ? 1 : 0}</Data>`;
            else if (M.isErr(v)) x += `<Data ss:Type="Error">${escT(v.e)}</Data>`;
            else x += `<Data ss:Type="String">${escT(v)}</Data>`;
          }
          if (cm) x += `<Comment ss:Author="${esc(cm.author || '')}"${cm.visible ? ' ss:ShowAlways="1"' : ''}><ss:Data xmlns="http://www.w3.org/TR/REC-html40">${escT(cm.text || '')}</ss:Data></Comment>`;
          cells.push(x + '</Cell>');
          lc = m ? m.c2 : c;
        });
        const rest = `${row.ht != null ? ` ss:AutoFitHeight="${row.customHeight ? 0 : 1}" ss:Height="${row.ht}"` : ''}${row.hidden ? ' ss:Hidden="1"' : ''}${row.s && used.has(row.s) ? ` ss:StyleID="${sid(row.s)}"` : ''}`;
        if (!cells.length && !rest) return;
        /* short gaps are written as empty rows (some readers ignore Row ss:Index), long ones by index */
        const gap = r - lastR - 1;
        let idx = '';
        if (gap > 0 && gap <= 100) for (let k = 0; k < gap; k++) out.push('<Row/>');
        else if (gap > 0) idx = ` ss:Index="${r + 1}"`;
        out.push(`<Row${idx}${rest}>${cells.join('')}</Row>`);
        lastR = r;
      });
      out.push('</Table>');
      const v = sh.view, pr = sh.print;
      let o = '<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><PageSetup>' + `<Layout x:Orientation="${pr.orientation === 'landscape' ? 'Landscape' : 'Portrait'}"/>` + (pr.header ? `<Header x:Margin="${pr.margins.header}" x:Data="${esc(pr.header)}"/>` : `<Header x:Margin="${pr.margins.header}"/>`) + (pr.footer ? `<Footer x:Margin="${pr.margins.footer}" x:Data="${esc(pr.footer)}"/>` : `<Footer x:Margin="${pr.margins.footer}"/>`) + `<PageMargins x:Bottom="${pr.margins.b}" x:Left="${pr.margins.l}" x:Right="${pr.margins.r}" x:Top="${pr.margins.t}"/></PageSetup>`;
      o += `<Print><ValidPrinterInfo/>${pr.paper && pr.paper !== 1 ? `<PaperSizeIndex>${pr.paper}</PaperSizeIndex>` : ''}${pr.scale && pr.scale !== 100 ? `<Scale>${pr.scale}</Scale>` : ''}${pr.gridLines ? '<Gridlines/>' : ''}${pr.headings ? '<RowColHeadings/>' : ''}</Print>`;
      if (si === wb.active) o += '<Selected/>';
      if (v.zoom && v.zoom !== 100) o += `<Zoom>${v.zoom}</Zoom>`;
      if (v.grid === false) o += '<DoNotDisplayGridlines/>';
      if (v.headings === false) o += '<DoNotDisplayHeadings/>';
      if (v.zeros === false) o += '<DoNotDisplayZeros/>';
      if (sh.state !== 'visible') o += `<Visible>${sh.state === 'veryHidden' ? 'SheetVeryHidden' : 'SheetHidden'}</Visible>`;
      if (v.freeze && (v.freeze.r || v.freeze.c)) o += `<FreezePanes/><FrozenNoSplit/>${v.freeze.r ? `<SplitHorizontal>${v.freeze.r}</SplitHorizontal><TopRowBottomPane>${(v.freeze.top || {}).r || v.freeze.r}</TopRowBottomPane>` : ''}${v.freeze.c ? `<SplitVertical>${v.freeze.c}</SplitVertical><LeftColumnRightPane>${(v.freeze.top || {}).c || v.freeze.c}</LeftColumnRightPane>` : ''}<ActivePane>${v.freeze.r && v.freeze.c ? 0 : v.freeze.r ? 2 : 1}</ActivePane>`;
      if (sh.protection) o += '<ProtectObjects>True</ProtectObjects><ProtectScenarios>True</ProtectScenarios>';
      o += '</WorksheetOptions>';
      out.push(o);
      if (sh.autoFilter && sh.autoFilter.ref) out.push(`<AutoFilter x:Range="${esc(XS.toR1C1(F.rangeName(sh.autoFilter.ref), 0, 0).slice(1))}" xmlns="urn:schemas-microsoft-com:office:excel"></AutoFilter>`);
      for (const d of sh.dv) {
        if (!d.ranges.length || d.type === 'none') continue;
        const a = d.ranges[0];
        const T = { list: 'List', whole: 'Whole', decimal: 'Decimal', date: 'Date', time: 'Time', textLength: 'TextLength', custom: 'Custom' }[d.type];
        const Q = { equal: 'Equal', notEqual: 'NotEqual', greaterThan: 'Greater', lessThan: 'Less', greaterThanOrEqual: 'GreaterOrEqual', lessThanOrEqual: 'LessOrEqual', notBetween: 'NotBetween' }[d.op];
        const two = !d.op || d.op === 'between' || d.op === 'notBetween';
        const fx = (f) => escT(XS.toR1C1(f, a.r1, a.c1).replace(/^=/, /^"/.test(f) || !/[A-Za-z(]/.test(f) ? '' : '='));
        out.push(`<DataValidation xmlns="urn:schemas-microsoft-com:office:excel"><Range>${escT(d.ranges.map((g) => XS.toR1C1(F.rangeName(g), 0, 0).slice(1)).join(','))}</Range><Type>${T}</Type>${Q ? `<Qualifier>${Q}</Qualifier>` : ''}${d.f1 != null ? (two && d.type !== 'list' && d.type !== 'custom' ? `<Min>${fx(d.f1)}</Min>` : `<Value>${fx(d.f1)}</Value>`) : ''}${d.f2 != null ? `<Max>${fx(d.f2)}</Max>` : ''}${d.prompt ? `<InputMessage>${escT(d.prompt)}</InputMessage>` : ''}${d.error ? `<ErrorMessage>${escT(d.error)}</ErrorMessage>` : ''}</DataValidation>`);
      }
      out.push('</Worksheet>');
    });
    out.push('</Workbook>');
    return out.join('\r\n');
  };
  XS.write = (wb) => new Blob([XS.text(wb)], { type: 'application/xml' });
  void SS;
})(typeof window !== 'undefined' ? window : globalThis);
