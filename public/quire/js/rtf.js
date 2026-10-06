/* Quire — Rich Text Format (RTF 1.x) reader and writer.
 * Reader: fonts, colours, character and paragraph formatting, styles (headings), lists, tables,
 * hyperlinks and other field results, pictures (PNG/JPEG/EMF/WMF), page/line/column breaks, section
 * page setup and document information.  Writer: the same subset, so RTF round-trips through Quire.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D;
  const RTF = (L.rtf = {});
  const CP1252 = '€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008DŽ\u008F\u0090‘’“”•–—˜™š›œ\u009DžŸ';
  const decodeByte = (b) => (b >= 0x80 && b < 0xa0 ? CP1252[b - 0x80] : String.fromCharCode(b));
  const HL = ['', 'black', 'blue', 'cyan', 'green', 'magenta', 'red', 'yellow', '', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray'];
  const HLRGB = { black: '000000', blue: '0000FF', cyan: '00FFFF', green: '00FF00', magenta: 'FF00FF', red: 'FF0000', yellow: 'FFFF00', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0' };
  const tw = (n) => n / 20;
  const UL = { ul: 'single', uld: 'dotted', uldash: 'dash', uldashd: 'dotDash', uldashdd: 'dotDotDash', uldb: 'double', ulhwave: 'wavyHeavy', ulldash: 'dashLong', ulth: 'thick', ulthd: 'dottedHeavy', ulthdash: 'dashedHeavy', ulwave: 'wave', ulw: 'words', uldbwave: 'wavyDouble' };

  /* ================= reader ================= */
  RTF.read = function (src) {
    if (src instanceof ArrayBuffer || ArrayBuffer.isView(src)) { const u = new Uint8Array(src.buffer || src); let s = ''; for (let i = 0; i < u.length; i += 8192) s += String.fromCharCode.apply(null, u.subarray(i, i + 8192)); src = s; }
    const doc = D.newDoc();
    doc.main.blocks = [];
    const fonts = {}, colors = [null], styles = {};
    const info = {};
    let i = 0;
    const n = src.length;
    /* state */
    const baseC = () => ({ rPr: {}, uc: 1, dest: null, skip: false });
    let st = { rPr: {}, pPr: {}, uc: 1, dest: 'body', skip: false, intbl: false, fontNum: null };
    const stack = [];
    let curP = null;
    const out = doc.main.blocks;
    /* table state */
    let rowDef = null, curRow = null, curCellBlocks = [], tbl = null;
    let lastBorderTarget = null;
    /* field state */
    let fldInst = '', fldDepth = -1, fldLink = null;
    /* picture */
    let pict = null;
    /* list state */
    let listBullet = null, listNum = null;
    const listIds = {};
    let fontTblCur = { num: null, name: '' }, colorCur = { r: 0, g: 0, b: 0, any: false };
    let styleCur = null;
    let infoKey = null;
    let pendingTextForPict = false;
    const sect = doc.sect;

    const flushPara = (force) => {
      if (!curP) { if (!force) return; curP = D.para([], {}); }
      const pPr = L.clone(st.pPr);
      delete pPr._ls; delete pPr._ilvl; delete pPr._pn;
      curP.pPr = Object.assign(curP.pPr, pPr);
      if (st.pPr._ls != null || st.pPr._pn) {
        const isNum = st.pPr._pn === 'num' || (st.pPr._ls != null && listIds[st.pPr._ls] === 'num');
        if (isNum) { if (!listNum) listNum = D.addNum(doc, D.makeNumberAbs('decimal', '%1.')); curP.pPr.num = { id: listNum, lvl: st.pPr._ilvl || 0 }; }
        else { if (!listBullet) listBullet = D.addNum(doc, D.makeBulletAbs('•', 'Symbol', '')); curP.pPr.num = { id: listBullet, lvl: st.pPr._ilvl || 0 }; }
      }
      if (st.intbl || (st.pPr && st.pPr._intbl)) curCellBlocks.push(curP);
      else { flushTable(); out.push(curP); }
      curP = null;
    };
    const flushTable = () => {
      if (!tbl) return;
      out.push(tbl);
      tbl = null;
    };
    const ensureP = () => { if (!curP) curP = D.para([], {}); return curP; };
    const addText = (t) => {
      if (!t) return;
      if (st.skip) return;
      if (st.dest === 'fonttbl') { fontTblCur.name += t; return; }
      if (st.dest === 'stylesheet') { if (styleCur) styleCur.name += t; return; }
      if (st.dest === 'info' && infoKey) { info[infoKey] = (info[infoKey] || '') + t; return; }
      if (st.dest === 'fldinst') { fldInst += t; return; }
      if (st.dest === 'pict') { if (pict) pict.hex += t; return; }
      if (st.dest !== 'body' && st.dest !== 'fldrslt') return;
      const p = ensureP();
      const rPr = L.clone(st.rPr);
      if (fldLink && st.dest === 'fldrslt') { rPr.link = { url: fldLink }; rPr.style = rPr.style || 'Hyperlink'; }
      const last = p.runs[p.runs.length - 1];
      if (last && last.t === 'text' && D.sameRPr(last.rPr, rPr)) last.text += t;
      else p.runs.push(D.text(t, rPr));
    };
    const addItem = (it) => { if (st.skip || (st.dest !== 'body' && st.dest !== 'fldrslt')) return; it.rPr = L.clone(st.rPr); ensureP().runs.push(it); };
    const charProp = (w, p, has) => {
      const r = st.rPr;
      const on = !has || p !== 0;
      switch (w) {
        case 'plain': st.rPr = {}; return true;
        case 'b': r.b = on || undefined; if (!on) delete r.b; return true;
        case 'i': if (on) r.i = true; else delete r.i; return true;
        case 'strike': if (on) r.strike = true; else delete r.strike; return true;
        case 'striked': if (on) r.dstrike = true; else delete r.dstrike; return true;
        case 'caps': if (on) r.caps = true; else delete r.caps; return true;
        case 'scaps': if (on) r.smallCaps = true; else delete r.smallCaps; return true;
        case 'v': if (on) r.hidden = true; else delete r.hidden; return true;
        case 'outl': if (on) r.outline = true; else delete r.outline; return true;
        case 'shad': if (on) r.shadow = true; else delete r.shadow; return true;
        case 'embo': if (on) r.emboss = true; else delete r.emboss; return true;
        case 'impr': if (on) r.imprint = true; else delete r.imprint; return true;
        case 'ulnone': delete r.u; return true;
        case 'super': r.vert = 'superscript'; return true;
        case 'sub': r.vert = 'subscript'; return true;
        case 'nosupersub': delete r.vert; return true;
        case 'fs': r.sz = p / 2; return true;
        case 'f': if (fonts[p]) r.font = fonts[p]; st.fontNum = p; return true;
        case 'cf': if (colors[p]) r.color = colors[p]; else delete r.color; return true;
        case 'highlight': if (HL[p]) r.hl = HL[p]; else delete r.hl; return true;
        case 'cb': case 'chcbpat': if (colors[p]) r.shd = { val: 'clear', fill: colors[p] }; else delete r.shd; return true;
        case 'expndtw': r.spacing = tw(p); if (!p) delete r.spacing; return true;
        case 'up': r.pos = p / 2; return true;
        case 'dn': r.pos = -p / 2; return true;
        case 'charscalex': r.w = p; return true;
        case 'cs': { const s = styles['c' + p]; if (s) r.style = s.id; return true; }
        case 'lang': return true;
      }
      if (UL[w]) { if (has && p === 0) delete r.u; else r.u = UL[w]; return true; }
      return false;
    };
    const paraProp = (w, p, has) => {
      const pp = st.pPr;
      switch (w) {
        case 'pard': st.pPr = {}; st.intbl = false; return true;
        case 'ql': pp.jc = 'left'; return true;
        case 'qc': pp.jc = 'center'; return true;
        case 'qr': pp.jc = 'right'; return true;
        case 'qj': pp.jc = 'both'; return true;
        case 'qd': pp.jc = 'distribute'; return true;
        case 'li': case 'lin': pp.ind = Object.assign({}, pp.ind, { l: tw(p) }); return true;
        case 'ri': case 'rin': pp.ind = Object.assign({}, pp.ind, { r: tw(p) }); return true;
        case 'fi': pp.ind = Object.assign({}, pp.ind, { fl: tw(p) }); return true;
        case 'sb': pp.sp = Object.assign({}, pp.sp, { b: tw(p) }); return true;
        case 'sa': pp.sp = Object.assign({}, pp.sp, { a: tw(p) }); return true;
        case 'sl': if (p) pp.sp = Object.assign({}, pp.sp, p < 0 ? { line: tw(-p), rule: 'exact' } : { line: tw(p), rule: 'atLeast' }); return true;
        case 'slmult': if (p && pp.sp && pp.sp.line) { pp.sp.line = pp.sp.line / 12; pp.sp.rule = 'auto'; } return true;
        case 'keepn': pp.keepNext = true; return true;
        case 'keep': pp.keepLines = true; return true;
        case 'pagebb': pp.pageBreakBefore = true; return true;
        case 'widctlpar': pp.widow = true; return true;
        case 'nowidctlpar': pp.widow = false; return true;
        case 'intbl': st.intbl = true; pp._intbl = true; return true;
        case 's': { const s = styles['p' + p]; if (s) pp.style = s.id; return true; }
        case 'outlinelevel': pp.outline = p; return true;
        case 'ls': pp._ls = p; return true;
        case 'ilvl': pp._ilvl = p; return true;
        case 'tx': pp.tabs = (pp.tabs || []).concat([{ pos: tw(p), al: pp._tqal || 'left', leader: pp._tldr || undefined }]); delete pp._tqal; delete pp._tldr; return true;
        case 'tqr': pp._tqal = 'right'; return true;
        case 'tqc': pp._tqal = 'center'; return true;
        case 'tqdec': pp._tqal = 'decimal'; return true;
        case 'tldot': pp._tldr = 'dot'; return true;
        case 'tlhyph': pp._tldr = 'hyphen'; return true;
        case 'tlul': pp._tldr = 'underscore'; return true;
        case 'cbpat': if (colors[p]) pp.shd = { val: 'clear', fill: colors[p] }; return true;
        case 'brdrt': lastBorderTarget = ['p', 'top']; return true;
        case 'brdrb': lastBorderTarget = ['p', 'bottom']; return true;
        case 'brdrl': lastBorderTarget = ['p', 'left']; return true;
        case 'brdrr': lastBorderTarget = ['p', 'right']; return true;
      }
      return false;
    };
    const borderProp = (w, p) => {
      if (!lastBorderTarget) return false;
      const [kind, side] = lastBorderTarget;
      const obj = kind === 'p' ? (st.pPr.borders = st.pPr.borders || {}) : kind === 'c' ? ((rowDef.pend.borders = rowDef.pend.borders || {})) : null;
      if (!obj) return false;
      const b = (obj[side] = obj[side] || { val: 'single', sz: 0.5, color: 'auto' });
      switch (w) {
        case 'brdrs': b.val = 'single'; return true;
        case 'brdrdb': b.val = 'double'; return true;
        case 'brdrdot': b.val = 'dotted'; return true;
        case 'brdrdash': b.val = 'dashed'; return true;
        case 'brdrth': b.val = 'thick'; return true;
        case 'brdrnone': b.val = 'nil'; return true;
        case 'brdrw': b.sz = tw(p); return true;
        case 'brdrcf': b.color = colors[p] || 'auto'; return true;
        case 'brsp': b.space = tw(p); return true;
      }
      return false;
    };
    const tableProp = (w, p) => {
      switch (w) {
        case 'trowd': rowDef = { cells: [], pend: {}, trPr: {}, left: 0, gap: 0 }; return true;
        case 'trgaph': if (rowDef) rowDef.gap = tw(p); return true;
        case 'trleft': if (rowDef) rowDef.left = tw(p); return true;
        case 'trrh': if (rowDef && p) { rowDef.trPr.h = tw(Math.abs(p)); rowDef.trPr.hRule = p < 0 ? 'exact' : 'atLeast'; } return true;
        case 'trhdr': if (rowDef) rowDef.trPr.header = true; return true;
        case 'trkeep': if (rowDef) rowDef.trPr.cantSplit = true; return true;
        case 'trqc': if (rowDef) rowDef.jc = 'center'; return true;
        case 'trqr': if (rowDef) rowDef.jc = 'right'; return true;
        case 'clbrdrt': lastBorderTarget = ['c', 'top']; return true;
        case 'clbrdrb': lastBorderTarget = ['c', 'bottom']; return true;
        case 'clbrdrl': lastBorderTarget = ['c', 'left']; return true;
        case 'clbrdrr': lastBorderTarget = ['c', 'right']; return true;
        case 'clcbpat': if (rowDef && colors[p]) rowDef.pend.shd = { val: 'clear', fill: colors[p] }; return true;
        case 'clvertalc': if (rowDef) rowDef.pend.vAlign = 'center'; return true;
        case 'clvertalb': if (rowDef) rowDef.pend.vAlign = 'bottom'; return true;
        case 'clvmgf': if (rowDef) rowDef.pend.vMerge = 'restart'; return true;
        case 'clvmrg': if (rowDef) rowDef.pend.vMerge = 'continue'; return true;
        case 'clmgf': if (rowDef) rowDef.pend.hMergeFirst = true; return true;
        case 'clmrg': if (rowDef) rowDef.pend.hMerge = true; return true;
        case 'cellx': if (rowDef) { rowDef.cells.push(Object.assign({ x: tw(p) }, rowDef.pend)); rowDef.pend = {}; lastBorderTarget = null; } return true;
        case 'cell': {
          if (curP || !curCellBlocks.length) flushPara(true);
          if (!curRow) curRow = [];
          curRow.push(curCellBlocks.length ? curCellBlocks : [D.para()]);
          curCellBlocks = [];
          return true;
        }
        case 'row': {
          if (curP) flushPara();
          endRow();
          return true;
        }
        case 'nestcell': return tableProp('cell', p);
        case 'nestrow': return tableProp('row', p);
      }
      return false;
    };
    const endRow = () => {
      const cellsIn = curRow || [];
      curRow = null;
      const def = rowDef || { cells: [], trPr: {}, left: 0 };
      if (!tbl) {
        tbl = D.table([], [], { style: 'TableGrid', w: { type: 'auto', v: 0 }, look: {}, borders: {}, jc: def.jc });
        if (!doc.styles.TableGrid && D.builtinStyles().TableGrid) doc.styles.TableGrid = D.builtinStyles().TableGrid;
      }
      /* grid from cell boundaries */
      let x0 = def.left || 0;
      const cells = [];
      cellsIn.forEach((blocks, k) => {
        const cd = def.cells[k] || { x: x0 + 108 };
        const w = Math.max(9, cd.x - x0);
        x0 = cd.x;
        if (cd.hMerge && cells.length) { const prev = cells[cells.length - 1]; prev.tcPr.span = (prev.tcPr.span || 1) + 1; prev.tcPr.w = (prev.tcPr.w || 0) + w; return; }
        const tcPr = { w };
        if (cd.shd) tcPr.shd = cd.shd;
        if (cd.vAlign) tcPr.vAlign = cd.vAlign;
        if (cd.vMerge) tcPr.vMerge = cd.vMerge;
        if (cd.borders) { tcPr.borders = cd.borders; }
        cells.push(D.cell(blocks, tcPr));
      });
      if (!tbl.grid.length) { let gx = def.left || 0; for (const cd of def.cells) { tbl.grid.push(Math.max(9, cd.x - gx)); gx = cd.x; } }
      const hasBorders = def.cells.some((c) => c.borders && Object.values(c.borders).some((b) => b.val !== 'nil'));
      if (!hasBorders && tbl.rows.length === 0) tbl.tblPr.style = 'TableNormal';
      if (hasBorders) tbl.tblPr.style = 'TableGrid';
      tbl.rows.push(D.row(cells, Object.assign({}, def.trPr)));
    };
    const special = (w, p, has) => {
      switch (w) {
        case 'par': flushPara(true); return true;
        case 'line': addItem(D.item('br', { type: 'line' })); return true;
        case 'page': addItem(D.item('br', { type: 'page' })); return true;
        case 'column': addItem(D.item('br', { type: 'column' })); return true;
        case 'tab': addItem(D.item('tab')); return true;
        case 'emdash': addText('—'); return true;
        case 'endash': addText('–'); return true;
        case 'emspace': addText('\u2003'); return true;
        case 'enspace': addText('\u2002'); return true;
        case 'bullet': addText('•'); return true;
        case 'lquote': addText('‘'); return true;
        case 'rquote': addText('’'); return true;
        case 'ldblquote': addText('“'); return true;
        case 'rdblquote': addText('”'); return true;
        case 'sect': { flushPara(); const last = out[out.length - 1]; if (last && last.t === 'p') last.sect = L.clone(sect); return true; }
        case 'paperw': sect.pgW = tw(p); return true;
        case 'paperh': sect.pgH = tw(p); return true;
        case 'margl': case 'marglsxn': sect.ml = tw(p); return true;
        case 'margr': case 'margrsxn': sect.mr = tw(p); return true;
        case 'margt': case 'margtsxn': sect.mt = tw(p); return true;
        case 'margb': case 'margbsxn': sect.mb = tw(p); return true;
        case 'pgwsxn': sect.pgW = tw(p); return true;
        case 'pghsxn': sect.pgH = tw(p); return true;
        case 'landscape': case 'lndscpsxn': sect.orient = 'landscape'; return true;
        case 'cols': sect.cols = Object.assign({}, sect.cols, { num: p }); return true;
        case 'uc': st.uc = p; return true;
        case 'u': { const code = p < 0 ? p + 65536 : p; addText(String.fromCharCode(code)); skipFallback(st.uc); return true; }
        case 'deftab': doc.settings.defTab = tw(p); return true;
      }
      return false;
    };
    const skipFallback = (cnt) => {
      while (cnt > 0 && i < n) {
        const c = src[i];
        if (c === '\\') {
          if (src[i + 1] === "'") { i += 4; cnt--; continue; }
          break;
        }
        if (c === '{' || c === '}') break;
        if (c === '\r' || c === '\n') { i++; continue; }
        i++; cnt--;
      }
    };
    const DESTS = { fonttbl: 'fonttbl', colortbl: 'colortbl', stylesheet: 'stylesheet', info: 'info', fldinst: 'fldinst', fldrslt: 'fldrslt', pict: 'pict', field: null };
    const SKIPDEST = new Set(['header', 'headerl', 'headerr', 'headerf', 'footer', 'footerl', 'footerr', 'footerf', 'footnote', 'annotation', 'pntext', 'pntxta', 'pntxtb', 'listtable', 'listoverridetable', 'revtbl', 'rsidtbl', 'generator', 'xmlnstbl', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'mmathPr', 'object', 'filetbl', 'nonshppict', 'bkmkstart', 'bkmkend', 'shpinst', 'pgdsctbl', 'listtext', 'operator', 'company', 'manager', 'category', 'comment', 'doccomm', 'hlinkbase', 'userprops', 'xe', 'tc']);

    while (i < n) {
      const c = src[i];
      if (c === '{') {
        stack.push({ rPr: L.clone(st.rPr), pPr: L.clone(st.pPr), uc: st.uc, dest: st.dest, skip: st.skip, intbl: st.intbl, fontNum: st.fontNum, starred: false });
        i++;
        /* the first control word in the group decides its destination */
        let j = i;
        let star = false;
        if (src[j] === '\\' && src[j + 1] === '*') { star = true; j += 2; while (src[j] === ' ' || src[j] === '\r' || src[j] === '\n') j++; }
        if (src[j] === '\\') {
          const m = /^\\([a-zA-Z]+)(-?\d+)? ?/.exec(src.slice(j, j + 40));
          if (m) {
            const w = m[1];
            if (w in DESTS && DESTS[w]) {
              st.dest = DESTS[w];
              if (w === 'fldinst') { fldInst = ''; }
              if (w === 'fldrslt') { fldLink = parseHyperlink(fldInst); }
              if (w === 'pict') { pict = { hex: '', type: null, w: 0, h: 0, gw: 0, gh: 0, sx: 100, sy: 100 }; }
              if (w === 'fonttbl') fontTblCur = { num: null, name: '' };
              if (w === 'stylesheet') styleCur = null;
              i = j + m[0].length;
              continue;
            } else if (w === 'field') { fldDepth = stack.length; }
            else if (w === 'shppict') { /* picture container: keep reading, its \pict is what we want */ i = j + m[0].length; pendingTextForPict = true; continue; }
            else if (w === 'pn') { st.dest = 'pn'; st.pPr._pn = 'bullet'; i = j + m[0].length; continue; }
            else if (st.dest === 'info' && /^(title|subject|author|keywords|operator|doccomm|company|manager|category)$/.test(w)) { infoKey = w; i = j + m[0].length; continue; }
            else if (SKIPDEST.has(w) || star) { st.skip = true; st.dest = 'skip'; i = j + m[0].length; continue; }
            else if (st.dest === 'fonttbl' && w === 'f') { fontTblCur = { num: +(m[2] || 0), name: '' }; }
            else if (st.dest === 'stylesheet' && (w === 's' || w === 'cs' || w === 'ds' || w === 'ts')) { styleCur = { key: (w === 's' ? 'p' : w === 'cs' ? 'c' : 'x') + (m[2] || 0), name: '', type: w === 's' ? 'paragraph' : 'character' }; }
          }
        } else if (star) { st.skip = true; st.dest = 'skip'; i = j; continue; }
        continue;
      }
      if (c === '}') {
        /* group end */
        const leaving = st;
        if (leaving.dest === 'fonttbl' && fontTblCur.num != null && fontTblCur.name) { fonts[fontTblCur.num] = fontTblCur.name.replace(/;\s*$/, '').trim(); fontTblCur = { num: null, name: '' }; }
        if (leaving.dest === 'stylesheet' && styleCur && styleCur.name) { finishStyle(styleCur); styleCur = null; }
        if (leaving.dest === 'pict' && pict) { finishPict(); }
        if (leaving.dest === 'info') infoKey = null;
        const prev = stack.pop() || baseC();
        if (fldDepth >= 0 && stack.length < fldDepth) { fldDepth = -1; fldLink = null; fldInst = ''; }
        if (fldDepth >= 0 && leaving.dest === 'fldrslt') fldLink = null;
        st = Object.assign({}, prev);
        if (leaving.dest === 'pn') { st.pPr._pn = leaving.pPr._pn; }
        if (leaving.dest === 'stylesheet' && prev.dest === 'stylesheet') {}
        i++;
        continue;
      }
      if (c === '\\') {
        const nx = src[i + 1];
        if (nx === "'") { const b = parseInt(src.substr(i + 2, 2), 16); i += 4; if (st.dest === 'pict') continue; addText(decodeByte(b)); continue; }
        if (nx === '\\' || nx === '{' || nx === '}') { addText(nx); i += 2; continue; }
        if (nx === '~') { addText('\u00A0'); i += 2; continue; }
        if (nx === '-') { addText('\u00AD'); i += 2; continue; }
        if (nx === '_') { addText('\u2011'); i += 2; continue; }
        if (nx === '\n' || nx === '\r') { flushPara(true); i += 2; continue; }
        if (nx === '*') { i += 2; continue; }
        const m = /^([a-zA-Z]+)(-?\d+)? ?/.exec(src.slice(i + 1, i + 48));
        if (!m) { i += 2; continue; }
        i += 1 + m[0].length;
        const w = m[1], has = m[2] != null, p = has ? +m[2] : 0;
        if (st.dest === 'fonttbl') continue;
        if (st.dest === 'colortbl') { if (w === 'red') { colorCur.r = p; colorCur.any = true; } else if (w === 'green') { colorCur.g = p; colorCur.any = true; } else if (w === 'blue') { colorCur.b = p; colorCur.any = true; } continue; }
        if (st.dest === 'stylesheet') { if (styleCur) { if (w === 'sbasedon') styleCur.basedOn = p; else if (w === 'outlinelevel') styleCur.outline = p; else { const saveR = st.rPr, saveP = st.pPr; st.rPr = styleCur.rPr = styleCur.rPr || {}; st.pPr = styleCur.pPr = styleCur.pPr || {}; charProp(w, p, has) || paraProp(w, p, has); st.rPr = saveR; st.pPr = saveP; } } continue; }
        if (st.dest === 'pict' && pict) {
          if (w === 'pngblip') pict.type = 'image/png'; else if (w === 'jpegblip') pict.type = 'image/jpeg'; else if (w === 'emfblip') pict.type = 'image/x-emf'; else if (w === 'wmetafile') pict.type = 'image/x-wmf'; else if (w === 'dibitmap' || w === 'wbitmap') pict.type = 'image/bmp';
          else if (w === 'picw') pict.w = p; else if (w === 'pich') pict.h = p; else if (w === 'picwgoal') pict.gw = tw(p); else if (w === 'pichgoal') pict.gh = tw(p); else if (w === 'picscalex') pict.sx = p; else if (w === 'picscaley') pict.sy = p;
          else if (w === 'bin') { const bytes = src.substr(i, p); i += p; for (let k = 0; k < bytes.length; k++) pict.hex += bytes.charCodeAt(k).toString(16).padStart(2, '0'); }
          continue;
        }
        if (st.dest === 'pn') { if (w === 'pnlvlbody' || w === 'pndec' || w === 'pnucrm' || w === 'pnlcrm' || w === 'pnucltr' || w === 'pnlcltr') st.pPr._pn = 'num'; else if (w === 'pnlvlblt') st.pPr._pn = 'bullet'; continue; }
        if (st.skip) { if (w === 'bin') i += p; continue; }
        if (w === 'rtf' || w === 'ansi' || w === 'ansicpg' || w === 'deff' || w === 'mac' || w === 'pc' || w === 'pca') continue;
        if (special(w, p, has)) continue;
        if (charProp(w, p, has)) continue;
        if (paraProp(w, p, has)) continue;
        if (tableProp(w, p)) continue;
        if (borderProp(w, p)) continue;
        if (w === 'sectd') { continue; }
        continue;
      }
      if (c === '\r' || c === '\n') { i++; continue; }
      /* plain text run */
      let j = i;
      while (j < n && src[j] !== '\\' && src[j] !== '{' && src[j] !== '}' && src[j] !== '\r' && src[j] !== '\n') j++;
      const t = src.slice(i, j);
      if (st.dest === 'colortbl') { for (const ch of t) if (ch === ';') { colors.push(colorCur.any ? [colorCur.r, colorCur.g, colorCur.b].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase() : null); colorCur = { r: 0, g: 0, b: 0, any: false }; } }
      else if (st.dest === 'fonttbl') { fontTblCur.name += t; if (t.includes(';') && fontTblCur.num != null) { fonts[fontTblCur.num] = fontTblCur.name.replace(/;.*$/, '').trim(); fontTblCur.name = ''; } }
      else if (st.dest === 'stylesheet') { if (styleCur) { styleCur.name += t; if (t.includes(';')) { finishStyle(styleCur); styleCur = null; } } }
      else addText(t);
      i = j;
    }
    if (curP && curP.runs.length) flushPara();
    flushTable();
    function finishStyle(s) {
      const name = s.name.replace(/;.*$/, '').trim();
      if (!name) return;
      const bi = D.builtinStyles();
      const lname = name.toLowerCase();
      let id = Object.keys(bi).find((k) => (bi[k].name || '').toLowerCase() === lname) || null;
      if (id) { if (!doc.styles[id]) doc.styles[id] = L.clone(bi[id]); }
      else {
        id = name.replace(/[^A-Za-z0-9]/g, '') || 'Style' + s.key;
        if (!doc.styles[id]) doc.styles[id] = { id, name, type: s.type, basedOn: s.type === 'paragraph' ? 'Normal' : undefined, pPr: s.pPr ? cleanP(s.pPr) : {}, rPr: s.rPr || {}, custom: true, q: true };
      }
      styles[s.key] = { id };
    }
    function cleanP(pp) { const o = L.clone(pp); delete o._ls; delete o._ilvl; delete o._pn; delete o._intbl; delete o.style; return o; }
    function finishPict() {
      const hex = pict.hex.replace(/[^0-9a-fA-F]/g, '');
      const pc = pict;
      pict = null;
      if (!hex || !pc.type) return;
      const bytes = new Uint8Array(hex.length >> 1);
      for (let k = 0; k < bytes.length; k++) bytes[k] = parseInt(hex.substr(k * 2, 2), 16);
      const blob = new Blob([bytes], { type: pc.type });
      const media = L.media.add(blob, 'image.' + L.mimeToExt(pc.type));
      const w = (pc.gw || tw(pc.w * 15) || 72) * (pc.sx || 100) / 100, hh = (pc.gh || tw(pc.h * 15) || 72) * (pc.sy || 100) / 100;
      const it = { t: 'img', media, w: L.round(w, 2), h: L.round(hh, 2), natW: pc.gw || w, natH: pc.gh || hh, rPr: L.clone(st.rPr) };
      if (st.dest === 'body' || st.dest === 'fldrslt' || stack.some((x) => x.dest === 'body')) ensureP().runs.push(it);
    }
    function parseHyperlink(inst) {
      const m = /^\s*HYPERLINK\s+(?:\\l\s+)?"([^"]*)"/i.exec(inst || '');
      return m ? m[1] : null;
    }
    for (const b of out) if (b.t === 'p') for (const r of b.runs) if (r.rPr) { delete r.rPr._x; }
    if (!out.length) out.push(D.para());
    if (out[out.length - 1].t !== 'p') out.push(D.para());
    if (!doc.styles.Hyperlink && D.builtinStyles().Hyperlink) doc.styles.Hyperlink = D.builtinStyles().Hyperlink;
    if (info.title) doc.props.title = info.title.trim();
    if (info.subject) doc.props.subject = info.subject.trim();
    if (info.author) doc.props.creator = info.author.trim();
    if (info.keywords) doc.props.keywords = info.keywords.trim();
    D.stylesChanged && D.stylesChanged();
    return doc;
  };

  /* ================= writer ================= */
  const esc = (t) => {
    let s = '';
    for (const ch of t) {
      const c = ch.codePointAt(0);
      if (ch === '\\' || ch === '{' || ch === '}') s += '\\' + ch;
      else if (ch === '\t') s += '\\tab ';
      else if (ch === '\u00A0') s += '\\~';
      else if (ch === '\u00AD') s += '\\-';
      else if (ch === '\u2011') s += '\\_';
      else if (c < 0x80) s += ch;
      else if (c > 0xffff) { const hi = Math.floor((c - 0x10000) / 0x400) + 0xd800, lo = ((c - 0x10000) % 0x400) + 0xdc00; s += `\\u${hi - 65536}?\\u${lo - 65536}?`; }
      else s += `\\u${c > 32767 ? c - 65536 : c}?`;
    }
    return s;
  };
  const T = (pt) => Math.round((pt || 0) * 20);
  RTF.write = async function (doc) {
    const fonts = ['Times New Roman'], colors = [null];
    const fontIx = (f) => { if (!f) return 0; let k = fonts.indexOf(f); if (k < 0) { fonts.push(f); k = fonts.length - 1; } return k; };
    const colIx = (c) => { if (!c || c === 'auto') return 0; c = String(c).replace('#', '').toUpperCase(); let k = colors.indexOf(c); if (k < 0) { colors.push(c); k = colors.length - 1; } return k; };
    const defFont = (doc.defaults && doc.defaults.rPr && doc.defaults.rPr.font) || 'Times New Roman';
    fonts[0] = defFont;
    const styleIds = Object.keys(doc.styles).filter((k) => doc.styles[k].type === 'paragraph' || doc.styles[k].type === 'character');
    const sIdx = {};
    styleIds.forEach((k, ix) => { sIdx[k] = ix; });
    const imgData = {};
    for (const st of D.stories(doc)) D.walk(st, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'img' && it.media && L.media.has(it.media)) imgData[it.media] = null; });
    for (const id of Object.keys(imgData)) {
      const m = L.media.get(id);
      let blob = m.blob, type = m.type;
      if (!/png|jpe?g|emf|wmf/.test(type)) { blob = m.view; type = blob.type; }
      if (!/png|jpe?g|emf|wmf/.test(type)) continue;
      const u = new Uint8Array(await L.readAsArrayBuffer(blob));
      let hex = '';
      for (let k = 0; k < u.length; k++) { hex += u[k].toString(16).padStart(2, '0'); if (k % 64 === 63) hex += '\n'; }
      imgData[id] = { hex, type };
    }
    const rpr = (r, styleR) => {
      r = r || {};
      let s = '';
      if (r.style && sIdx[r.style] != null && doc.styles[r.style].type === 'character') s += `\\cs${sIdx[r.style]}`;
      const sr = r.style && doc.styles[r.style] ? D.charStyleProps ? Object.assign({}, doc.styles[r.style].rPr || {}) : {} : {};
      const R = Object.assign({}, styleR || {}, sr, r);
      if (R.font) s += `\\f${fontIx(R.font)}`;
      if (R.sz) s += `\\fs${Math.round(R.sz * 2)}`;
      if (R.b) s += '\\b';
      if (R.i) s += '\\i';
      if (R.u && R.u !== 'none') s += { single: '\\ul', double: '\\uldb', dotted: '\\uld', dash: '\\uldash', thick: '\\ulth', wave: '\\ulwave', words: '\\ulw', dotDash: '\\uldashd', dotDotDash: '\\uldashdd', dashLong: '\\ulldash' }[R.u] || '\\ul';
      if (R.strike) s += '\\strike';
      if (R.dstrike) s += '\\striked1';
      if (R.caps) s += '\\caps';
      if (R.smallCaps) s += '\\scaps';
      if (R.hidden) s += '\\v';
      if (R.outline) s += '\\outl';
      if (R.shadow) s += '\\shad';
      if (R.emboss) s += '\\embo';
      if (R.imprint) s += '\\impr';
      if (R.vert === 'superscript') s += '\\super';
      if (R.vert === 'subscript') s += '\\sub';
      if (R.color && R.color !== 'auto') s += `\\cf${colIx(R.color)}`;
      if (R.hl && R.hl !== 'none') { const k = HL.indexOf(R.hl); s += `\\highlight${k > 0 ? k : 7}`; }
      if (R.shd && R.shd.fill && R.shd.fill !== 'auto') s += `\\chcbpat${colIx(R.shd.fill)}`;
      if (R.spacing) s += `\\expndtw${T(R.spacing)}`;
      if (R.pos) s += R.pos > 0 ? `\\up${Math.round(R.pos * 2)}` : `\\dn${Math.round(-R.pos * 2)}`;
      return s;
    };
    const brd = (b) => { if (!b || !b.val || b.val === 'nil' || b.val === 'none') return ''; return `${{ double: '\\brdrdb', dotted: '\\brdrdot', dashed: '\\brdrdash', thick: '\\brdrth' }[b.val] || '\\brdrs'}\\brdrw${Math.max(1, T(b.sz || 0.5))}${b.space ? `\\brsp${T(b.space)}` : ''}${b.color && b.color !== 'auto' ? `\\brdrcf${colIx(b.color)}` : ''}`; };
    const ppr = (p, inTable) => {
      const pp = p.pPr || {};
      let s = '\\pard\\plain';
      if (pp.style && sIdx[pp.style] != null) s += `\\s${sIdx[pp.style]}`;
      const eff = D.pProps(doc, p);
      const jc = eff.jc;
      if (jc === 'center') s += '\\qc'; else if (jc === 'right') s += '\\qr'; else if (jc === 'both' || jc === 'distribute') s += '\\qj';
      const ind = eff.ind || {};
      if (ind.l) s += `\\li${T(ind.l)}`;
      if (ind.r) s += `\\ri${T(ind.r)}`;
      if (ind.fl) s += `\\fi${T(ind.fl)}`;
      const sp = eff.sp || {};
      if (sp.b) s += `\\sb${T(sp.b)}`;
      if (sp.a) s += `\\sa${T(sp.a)}`;
      if (sp.line) s += sp.rule === 'auto' || !sp.rule ? `\\sl${Math.round(sp.line * 240)}\\slmult1` : sp.rule === 'exact' ? `\\sl-${T(sp.line)}\\slmult0` : `\\sl${T(sp.line)}\\slmult0`;
      if (eff.keepNext) s += '\\keepn';
      if (eff.keepLines) s += '\\keep';
      if (eff.pageBreakBefore) s += '\\pagebb';
      if (eff.outline != null && eff.outline < 9) s += `\\outlinelevel${eff.outline}`;
      for (const t of eff.tabs || []) { if (t.al === 'clear') continue; s += ({ right: '\\tqr', center: '\\tqc', decimal: '\\tqdec' }[t.al] || '') + ({ dot: '\\tldot', hyphen: '\\tlhyph', underscore: '\\tlul' }[t.leader] || '') + `\\tx${T(t.pos)}`; }
      const b = eff.borders || {};
      for (const [side, kw] of [['top', 'brdrt'], ['bottom', 'brdrb'], ['left', 'brdrl'], ['right', 'brdrr']]) { const x = brd(b[side]); if (x) s += `\\${kw}${x}`; }
      if (eff.shd && eff.shd.fill && eff.shd.fill !== 'auto') s += `\\cbpat${colIx(eff.shd.fill)}`;
      if (inTable) s += '\\intbl';
      return s;
    };
    const paraBody = (p) => {
      let s = '';
      const pp = D.pProps(doc, p);
      if (pp.num && pp.num.id && pp.num.id !== '0') {
        const lv = D.numLevelDef(doc, pp.num.id, pp.num.lvl || 0);
        let label = lv && lv.fmt !== 'bullet' ? '' : '•';
        if (listText && listText.get(p.id)) label = listText.get(p.id).text || label;
        s += `{\\listtext\\pard\\plain ${esc(label)}\\tab}`;
      }
      const styleR = p.pPr && p.pPr.style && doc.styles[p.pPr.style] ? doc.styles[p.pPr.style].rPr : null;
      let fld = null;
      for (const it of p.runs) {
        const r = it.rPr || {};
        if (r.del) continue;
        if (it.t === 'fb') { fld = { instr: it.instr || '', res: '' }; s += `{\\field{\\*\\fldinst {${esc(it.instr || '')}}}{\\fldrslt {`; continue; }
        if (it.t === 'fs') continue;
        if (it.t === 'fe') { if (fld) { s += '}}}'; fld = null; } continue; }
        if (fld && it.t !== 'text') continue;
        const rp = rpr(r, styleR);
        const pre = rp ? `{${rp} ` : '{';
        if (it.t === 'text') {
          if (r.link && r.link.url && !fld) s += `{\\field{\\*\\fldinst {HYPERLINK "${esc(r.link.url)}"}}{\\fldrslt ${pre}${esc(it.text)}}}}`;
          else s += pre + esc(it.text) + '}';
        } else if (it.t === 'tab' || it.t === 'ptab') s += pre + '\\tab}';
        else if (it.t === 'br') s += pre + (it.type === 'page' ? '\\page' : it.type === 'column' ? '\\column' : '\\line') + '}';
        else if (it.t === 'sym') s += pre + esc(it.char || '') + '}';
        else if (it.t === 'ruby') s += pre + esc(it.base || '') + '}';
        else if (it.t === 'img' && imgData[it.media]) {
          const im = imgData[it.media];
          const kind = /png/.test(im.type) ? '\\pngblip' : /jpe?g/.test(im.type) ? '\\jpegblip' : /emf/.test(im.type) ? '\\emfblip' : '\\wmetafile8';
          s += `{\\*\\shppict{\\pict${kind}\\picwgoal${T(it.w)}\\pichgoal${T(it.h)}\n${im.hex}}}`;
        } else if (it.t === 'fn' || it.t === 'en') {
          const story = (it.t === 'fn' ? doc.fn : doc.en)[it.id];
          if (it.self) { s += `{${rp}\\chftn}`; continue; }
          const txt = story ? story.blocks.filter((b) => b.t === 'p').map((b) => D.plainText(b)).join(' ').trim() : '';
          s += `{\\super\\chftn}{\\footnote\\pard\\plain{\\super\\chftn} ${esc(txt)}}`;
        }
      }
      return s;
    };
    let body = '';
    let listText = null;
    const blocks = (list, inTable) => {
      for (const b of list) {
        if (b.t === 'p') { body += `${ppr(b, inTable)} ${paraBody(b)}${inTable ? '' : '\\par'}\n`; if (b.sect && !inTable) body += sectWords(b.sect) + '\\sect\\sectd\n'; }
        else if (b.t === 'tbl') table(b);
        else if (b.t === 'sdt' && b.blocks) blocks(b.blocks, inTable);
      }
    };
    const table = (t) => {
      const grid = t.grid && t.grid.length ? t.grid : [468];
      const vm = {};
      for (const row of t.rows) {
        let def = `\\trowd\\trgaph${T(5.4)}\\trleft${T(-5.4)}`;
        const tp = row.trPr || {};
        if (tp.h) def += `\\trrh${tp.hRule === 'exact' ? '-' : ''}${T(tp.h)}`;
        if (tp.header) def += '\\trhdr';
        if (tp.cantSplit) def += '\\trkeep';
        if (t.tblPr && t.tblPr.jc === 'center') def += '\\trqc'; else if (t.tblPr && t.tblPr.jc === 'right') def += '\\trqr';
        let x = 0, gi = tp.gridBefore || 0;
        for (let k = 0; k < gi; k++) x += grid[k] || 0;
        const cellBodies = [];
        for (const cell of row.cells) {
          const cp = cell.tcPr || {};
          const span = cp.span || 1;
          for (let k = 0; k < span; k++) x += grid[gi + k] || (cp.w || 72) / span;
          gi += span;
          const bd = Object.assign({}, D.tableStyleProps ? {} : {}, (t.tblPr && t.tblPr.borders) || {}, cp.borders || {});
          const styled = t.tblPr && (t.tblPr.style === 'TableGrid' || /Grid|Table/.test(t.tblPr.style || '')) && !(t.tblPr.borders && Object.keys(t.tblPr.borders).length);
          for (const [side, kw, alt] of [['top', 'clbrdrt', 'insideH'], ['left', 'clbrdrl', 'insideV'], ['bottom', 'clbrdrb', 'insideH'], ['right', 'clbrdrr', 'insideV']]) {
            const b = bd[side] || bd[alt] || (styled ? { val: 'single', sz: 0.5 } : null);
            const xx = brd(b);
            if (xx) def += `\\${kw}${xx}`;
          }
          if (cp.shd && cp.shd.fill && cp.shd.fill !== 'auto') def += `\\clcbpat${colIx(cp.shd.fill)}`;
          if (cp.vAlign === 'center') def += '\\clvertalc'; else if (cp.vAlign === 'bottom') def += '\\clvertalb';
          if (cp.vMerge === 'restart') def += '\\clvmgf'; else if (cp.vMerge === 'continue' || cp.vMerge === true) def += '\\clvmrg';
          def += `\\cellx${T(x)}`;
          cellBodies.push(cell);
        }
        body += def + '\n';
        for (const cell of cellBodies) {
          const bl = cell.blocks.filter((b) => b.t === 'p');
          bl.forEach((p, k) => { body += `${ppr(p, true)} ${paraBody(p)}${k < bl.length - 1 ? '\\par' : ''}`; });
          if (!bl.length) body += '\\pard\\intbl ';
          body += '\\cell\n';
        }
        body += '\\row\n';
      }
      body += '\\pard\n';
    };
    const sectWords = (s) => `\\pgwsxn${T(s.pgW)}\\pghsxn${T(s.pgH)}\\marglsxn${T(s.ml)}\\margrsxn${T(s.mr)}\\margtsxn${T(s.mt)}\\margbsxn${T(s.mb)}${s.orient === 'landscape' ? '\\lndscpsxn' : ''}${s.cols && s.cols.num > 1 ? `\\cols${s.cols.num}` : ''}`;
    /* list labels: compute numbering text so lists read correctly in other programs */
    try { listText = D.computeLists(doc); } catch (e) { listText = null; }
    blocks(doc.main.blocks, false);
    /* styles */
    let ss = '';
    styleIds.forEach((k) => {
      const s = doc.styles[k];
      const nm = s.name || k;
      if (s.type === 'paragraph') ss += `{\\s${sIdx[k]}${s.pPr && s.pPr.outline != null ? `\\outlinelevel${s.pPr.outline}` : ''}${rpr(s.rPr || {})} ${esc(nm)};}`;
      else ss += `{\\*\\cs${sIdx[k]} \\additive${rpr(s.rPr || {})} ${esc(nm)};}`;
    });
    const s0 = doc.sect;
    const p = doc.props || {};
    const infoS = `{\\info${p.title ? `{\\title ${esc(p.title)}}` : ''}${p.subject ? `{\\subject ${esc(p.subject)}}` : ''}${p.creator ? `{\\author ${esc(p.creator)}}` : ''}${p.keywords ? `{\\keywords ${esc(p.keywords)}}` : ''}}`;
    const defSz = (doc.defaults && doc.defaults.rPr && doc.defaults.rPr.sz) || 12;
    const head = `{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1\\deflang1033` +
      `{\\fonttbl${fonts.map((f, k) => `{\\f${k}\\fnil\\fcharset0 ${esc(f)};}`).join('')}}` +
      `{\\colortbl;${colors.slice(1).map((c) => `\\red${parseInt(c.slice(0, 2), 16)}\\green${parseInt(c.slice(2, 4), 16)}\\blue${parseInt(c.slice(4, 6), 16)};`).join('')}}` +
      `{\\stylesheet${ss}}` + infoS +
      `\\paperw${T(s0.pgW)}\\paperh${T(s0.pgH)}\\margl${T(s0.ml)}\\margr${T(s0.mr)}\\margt${T(s0.mt)}\\margb${T(s0.mb)}${s0.orient === 'landscape' ? '\\landscape' : ''}\\deftab${T((doc.settings && doc.settings.defTab) || 36)}\\widowctrl\\fs${Math.round(defSz * 2)}\n` +
      `\\sectd${sectWords(s0)}\n`;
    /* fonts and colours are collected while writing the body, so build the header last */
    return head + body + '}';
  };
})();
