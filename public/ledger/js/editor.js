/* Ledger — cell editing: in-cell editor, formula bar, point mode, reference highlighting, AutoComplete. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, LY = L.layout, RD = L.render, O = L.ops, F = L.formula, NF = L.numfmt;
  const { h } = L;
  const E = (L.editor = {});
  const G = () => L.grid;

  E.active = false;
  E.mode = 'ready'; /* ready | enter | edit | point */
  let fbar, fbwrap, cell, surface = null, orig = '', point = null, acList = null, tipEl = null;

  E.init = function () {
    fbar = L.$('#fbinput');
    fbwrap = L.$('#fbar');
    cell = G().input;
    fbar.addEventListener('focus', () => {
      if (!E.active) { if (G().guardProtect(null, true)) { fbar.blur(); G().focus(); return; } E.begin('edit', null, { surface: 'bar' }); }
      surface = fbar;
    });
    fbar.addEventListener('keydown', (e) => { if (L.ui.menuKey(e) || L.ui.dialogKey(e)) { e.preventDefault(); return; } if (E.active) E.key(e); });
    fbar.addEventListener('input', () => { if (E.active) E.onInput(); });
    fbar.addEventListener('pointerdown', () => { if (E.active && surface !== fbar) { surface = fbar; } });
    fbar.addEventListener('dblclick', () => { fbar.classList.toggle('multi'); });
    L.$('#fb-cancel').addEventListener('pointerdown', (e) => { e.preventDefault(); E.cancel(); });
    L.$('#fb-enter').addEventListener('pointerdown', (e) => { e.preventDefault(); E.commit({ stay: true }); });
    L.$('#fb-cancel').innerHTML = '<svg width="13" height="13" viewBox="0 0 13 13"><path d="M2 2l9 9M11 2l-9 9" stroke="#c0281a" stroke-width="2.2"/></svg>';
    L.$('#fb-enter').innerHTML = '<svg width="13" height="13" viewBox="0 0 13 13"><path d="M1.5 7l3.5 3.5L11.5 2.5" fill="none" stroke="#1b8a2a" stroke-width="2.2"/></svg>';
    L.$('#fb-fx').innerHTML = '<svg width="18" height="14" viewBox="0 0 18 14"><text x="1" y="11" font-family="Times New Roman,Tinos,serif" font-style="italic" font-size="13" fill="#1a3f8a">f</text><text x="6" y="11" font-family="Times New Roman,Tinos,serif" font-style="italic" font-size="11" fill="#1a3f8a">x</text></svg>';
    L.$('#fb-fx').addEventListener('pointerdown', (e) => { e.preventDefault(); L.ui.exec('insertFunction'); });
    cell.addEventListener('pointerdown', (e) => { if (E.active) e.stopPropagation(); });
    cell.addEventListener('blur', () => {
      /* clicking elsewhere in the page (not the grid or formula bar) keeps editing alive; menus commit first */
    });
    L.bus.on('selection', () => { if (!E.active) E.show(); });
    L.bus.on('changed', () => { if (!E.active) E.show(); });
  };

  /* ------------------------------------------------------------ display in the formula bar */
  /** text shown for a cell when not editing */
  E.cellText = function (sh, r, c) {
    const cl = sh.get(r, c);
    if (!cl) return '';
    if (cl.am) { const m = sh.get(cl.am.r, cl.am.c); if (m && m.f != null) return (m.dyn ? '=' : '{=') + F.display(m.f) + (m.dyn ? '' : '}'); }
    if (cl.dt) return '{=TABLE(' + [cl.dt.r1 || '', cl.dt.dt2D ? cl.dt.r2 || '' : ''].join(',') + ')}';
    if (cl.f != null) return cl.af && !cl.dyn ? '{=' + F.display(cl.f) + '}' : '=' + F.display(cl.f);
    const v = cl.v;
    if (v == null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (M.isErr(v)) return v.e;
    if (typeof v === 'number') {
      const st = LY.styleOf(sh, r, c, cl);
      const nf = st.nf || 'General';
      /* Excel shows dates and times in a standard editable form */
      if (NF.isDate(nf)) {
        const p = NF.serialToParts(v, sh.wb.date1904, 0);
        const hasT = v % 1 !== 0, hasD = Math.floor(v) !== 0 || !NF.isTimeOnly(nf);
        const t = p.H % 12 === 0 ? 12 : p.H % 12;
        const time = `${t}:${String(p.M).padStart(2, '0')}:${String(p.S).padStart(2, '0')} ${p.H >= 12 ? 'PM' : 'AM'}`;
        if (!hasD || (v < 1 && NF.isTimeOnly(nf))) return time;
        return `${p.m}/${p.d}/${p.y}` + (hasT ? ' ' + time : '');
      }
      if (/%/.test(nf) && !/"[^"]*%/.test(nf)) return barNumber(v * 100) + '%';
      return barNumber(v);
    }
    const st = LY.styleOf(sh, r, c, cl);
    return (st.prot && st.prot.hidden && sh.protection ? '' : '') + String(v);
  };
  /** a number as the formula bar shows it: up to 15 significant digits, scientific for very large / small values */
  function barNumber(v) {
    const x = NF.r15(v);
    if (x === 0) return '0';
    const a = Math.abs(x);
    if (a >= 1e15 || a < 1e-9) {
      const [m, e] = x.toExponential().split('e');
      const ex = parseInt(e, 10);
      return m.replace(/\.?0+$/, '') + 'E' + (ex < 0 ? '-' : '+') + String(Math.abs(ex)).padStart(2, '0');
    }
    let s = String(x);
    if (/e/.test(s)) s = x.toFixed(20).replace(/\.?0+$/, '');
    return s;
  }
  E.barNumber = barNumber;
  E.show = function () {
    if (!fbar) return;
    const sh = G().sheet(), s = G().sel();
    const st = LY.styleOf(sh, s.r, s.c, sh.get(s.r, s.c));
    const hidden = sh.protection && st.prot && st.prot.hidden;
    const cl = sh.get(s.r, s.c);
    fbar.value = hidden && cl && cl.f != null ? '' : E.cellText(sh, s.r, s.c);
    fbar.classList.remove('multi');
    fbwrap.classList.remove('editing');
  };

  /* ------------------------------------------------------------ begin / end */
  E.canPoint = () => E.active && E.mode === 'point';
  /**
   * Start editing the active cell. mode: 'enter' (typing replaces) | 'edit' (F2 / double-click)
   */
  E.begin = function (mode, text, opts) {
    opts = opts || {};
    const g = G(), sh = g.sheet(), s = g.sel();
    if (E.active) return;
    /* refuse when the cell is part of an array formula and the selection doesn't cover it */
    E.sheet = sh; E.r = s.r; E.c = s.c;
    orig = E.cellText(sh, s.r, s.c);
    E.wasArray = /^\{=/.test(orig) && !(sh.get(s.r, s.c) || {}).dyn;
    if (E.wasArray) orig = orig.slice(1, -1);
    E.active = true;
    E.mode = mode === 'edit' ? 'edit' : 'enter';
    point = null;
    surface = opts.surface === 'bar' ? fbar : cell;
    const val = text != null ? text : orig;
    cell.value = val;
    fbar.value = val;
    cell.style.opacity = '1';
    cell.style.pointerEvents = 'auto';
    cell.classList.add('on');
    fbwrap.classList.add('editing');
    E.styleBox();
    E.place();
    if (surface === cell) {
      cell.focus({ preventScroll: true });
      if (opts.caretFromPoint && cell.value) { const pos = caretAt(opts.caretFromPoint); cell.setSelectionRange(pos, pos); }
      else cell.setSelectionRange(cell.value.length, cell.value.length);
    }
    E.mode = E.mode === 'enter' ? E.pointableAfterInput() : E.mode;
    E.highlight();
    G().paint();
    L.bus.emit('mode');
    E.tip();
  };
  /** after typing in Enter mode, are we pointing (formula expecting an operand)? */
  E.pointableAfterInput = function () {
    const t = surface.value;
    if (!/^[=+-]/.test(t)) return 'enter';
    const pos = surface.selectionStart;
    if (pos !== surface.selectionEnd) return E.mode === 'edit' ? 'edit' : 'enter';
    const before = t.slice(0, pos);
    if (/(^[=+\-]|[=+\-*/^&(,;<>:%\s])$/.test(before) && !/"[^"]*$/.test(before.replace(/"[^"]*"/g, ''))) return 'point';
    return E.mode === 'edit' ? 'edit' : 'enter';
  };
  /** typed-text styling of the in-cell box */
  E.styleBox = function () {
    const sh = E.sheet;
    const st = LY.styleOf(sh, E.r, E.c, sh.get(E.r, E.c));
    const z = G().zoomOf(sh);
    const f = Object.assign({}, st.font || {}, { name: LY.fontName(sh.wb, st.font || {}) });
    cell.style.font = LY.cssFont(f, z);
    cell.style.color = M.colorHex(sh.wb, f.color, '#000000');
    cell.style.fontWeight = f.b ? 'bold' : 'normal';
    cell.style.fontStyle = f.i ? 'italic' : 'normal';
    cell.style.textDecoration = (f.u ? 'underline ' : '') + (f.strike ? 'line-through' : '');
    const fill = st.fill && st.fill.pattern === 'solid' ? M.colorHex(sh.wb, st.fill.fg || st.fill.bg, '#FFFFFF') : '#FFFFFF';
    cell.style.background = fill;
    const wrap = st.align && st.align.wrap;
    cell.classList.toggle('wrap', !!wrap);
    cell.style.textAlign = st.align && st.align.h === 'right' ? 'right' : st.align && st.align.h === 'center' ? 'center' : 'left';
  };
  /** position / size the in-cell box (grows to the right like Excel) */
  E.place = function () {
    if (!E.active) return;
    const g = G();
    if (g.sheet() !== E.sheet) { cell.style.opacity = '0'; return; }
    cell.style.opacity = '1';
    const q = g.cellRect(E.r, E.c);
    const fr = q.fr;
    const sh = E.sheet;
    const st = LY.styleOf(sh, E.r, E.c, sh.get(E.r, E.c));
    const wrap = st.align && st.align.wrap;
    const font = cell.style.font;
    const lines = cell.value.split('\n');
    let need = 0;
    for (const ln of lines) need = Math.max(need, LY.measure(font, ln + 'W'));
    const maxW = g.view().w - q.x - 2;
    let w = Math.max(q.w + 1, Math.min(maxW, need + 6));
    if (wrap) w = q.w + 1;
    const lh = parseFloat(font) * 1.2 || 18;
    let hh = Math.max(q.h + 1, Math.ceil(lines.length * lh + 4));
    if (wrap) { const ls = RD.wrapLines(font, cell.value, Math.max(10, q.w - 6)).length; hh = Math.max(q.h + 1, Math.ceil(ls * lh + 4)); }
    Object.assign(cell.style, { left: (q.x - 1) + 'px', top: (q.y - 1) + 'px', width: w + 'px', height: hh + 'px', lineHeight: (q.h > lh * 1.4 && lines.length === 1 && !wrap ? q.h - 2 : lh) + 'px' });
    if (lines.length === 1 && !wrap && q.h > lh * 1.4) cell.style.paddingTop = '0';
    void fr;
  };
  function caretAt(p) {
    /* approximate caret position from a click x within the cell */
    const q = G().cellRect(E.r, E.c);
    const font = cell.style.font;
    const t = cell.value;
    let x = q.x + 2;
    for (let i = 0; i < t.length; i++) { const w = LY.measure(font, t[i]); if (x + w / 2 > p.x) return i; x += w; }
    return t.length;
  }

  E.onInput = function () {
    const other = surface === cell ? fbar : cell;
    other.value = surface.value;
    if (E.mode !== 'edit') E.mode = E.pointableAfterInput();
    point = null;
    E.place();
    E.highlight();
    E.autoComplete();
    E.tip();
    L.bus.emit('mode');
  };
  E.text = () => (surface ? surface.value : '');
  function setText(t, caret) {
    cell.value = t; fbar.value = t;
    const pos = caret == null ? t.length : caret;
    try { surface.setSelectionRange(pos, pos); } catch (e) { /* ignore */ }
    E.place(); E.highlight(); E.tip();
  }
  E.setText = setText;
  E.insertText = function (s) {
    const t = surface.value, a = surface.selectionStart, b = surface.selectionEnd;
    setText(t.slice(0, a) + s + t.slice(b), a + s.length);
    if (E.mode !== 'edit') E.mode = E.pointableAfterInput();
  };

  /** finish editing; returns false when the entry was refused (editor stays open) */
  E.commit = function (opts) {
    opts = opts || {};
    if (!E.active) return true;
    const sh = E.sheet, r = E.r, c = E.c;
    let text = surface.value;
    const g = G();
    hideAux();
    const unchanged = text === orig && !opts.array && !opts.fill && !E.wasArray;
    if (unchanged || (text === orig && E.wasArray && !opts.array)) { end(); return true; }
    if (L.autocorrect && L.autocorrect.applyCell && !/^[=+\-@']/.test(text)) { const t2 = L.autocorrect.applyCell(text); if (t2 !== text) { text = t2; setText(t2); } }
    /* formula autocorrect proposals and errors */
    if (/^[=]/.test(text) || (/^[+-]/.test(text) && O.parseInput(sh.wb, text).kind === 'formula')) {
      const body = /^=/.test(text) ? text.slice(1) : text;
      const chk = O.checkFormula(body);
      if (chk.error) {
        if (chk.fixed) {
          L.ui.msg('Ledger found an error in the formula you entered. Do you want to accept the correction proposed below?\n\n=' + chk.fixed + '\n\n• To accept the correction, click Yes.\n• To close this message and correct the formula yourself, click No.', { icon: 'question', buttons: ['&Yes', '&No'], width: 430 }).then((i) => {
            if (i === 0) { setText('=' + chk.fixed); E.commit(opts); }
            else surface.focus();
          });
          return false;
        }
        L.ui.msg('The formula you typed contains an error.\n\n• For information about fixing common formula problems, click Help.\n• To get assistance in entering a function, click OK, and then click Function on the Insert menu.\n• If you are not trying to enter a formula, avoid using an equal sign (=) or minus sign (-), or precede it with a single quotation mark (\').', { icon: 'warn', width: 440 }).then(() => {
          surface.focus();
          if (chk.pos != null) { const p = Math.min(surface.value.length, chk.pos + 1); try { surface.setSelectionRange(p, p); } catch (e) { /* ignore */ } }
        });
        return false;
      }
    }
    /* data validation */
    if (!opts.skipValidation && !/^=/.test(text)) {
      const p = O.parseInput(sh.wb, text);
      const bad = p.kind === 'value' ? L.cf.dvCheck(sh, r, c, p.v) : null;
      if (bad && bad.showError !== false) {
        const style = bad.errorStyle || 'stop';
        const msg = bad.error || 'The value you entered is not valid.\nA user has restricted values that can be entered into this cell.';
        const title = bad.errorTitle || L.APP;
        if (style === 'stop') { L.ui.msg(msg, { icon: 'error', title, buttons: ['&Retry', 'Cancel'] }).then((i) => { if (i === 0) { surface.focus(); surface.select(); } else E.cancel(); }); return false; }
        L.ui.msg(msg + (style === 'warning' ? '\n\nContinue?' : ''), { icon: style === 'warning' ? 'warn' : 'info', title, buttons: style === 'warning' ? ['&Yes', '&No', 'Cancel'] : ['OK', 'Cancel'] }).then((i) => {
          if (i === 0) E.commit(Object.assign({}, opts, { skipValidation: true }));
          else if (style === 'warning' && i === 1) { surface.focus(); surface.select(); }
          else E.cancel();
        });
        return false;
      }
    }
    try {
      const s = g.sel();
      const arr = O.arrayAt(sh, r, c);
      if (arr && !opts.array && !(arr.r1 === arr.r2 && arr.c1 === arr.c2)) { L.ui.msg('You cannot change part of an array.', { icon: 'warn' }); return false; }
      O.tx(sh.wb, 'Typing', () => {
        if (opts.array || (E.wasArray && arr)) {
          const rg = opts.array ? s.ranges[s.active] : arr;
          O.enter(sh, rg.r1, rg.c1, text, { array: rg });
        } else if (opts.fill) {
          /* Ctrl+Enter: the entry goes to every selected cell */
          for (const rg of s.ranges) {
            const rr = O.clip(sh, rg);
            for (let rr2 = rr.r1; rr2 <= rr.r2; rr2++) for (let cc = rr.c1; cc <= rr.c2; cc++) {
              const m = LY.mergeAt(sh, rr2, cc);
              if (m && (m.r1 !== rr2 || m.c1 !== cc)) continue;
              const t = /^[=+-]/.test(text) && O.parseInput(sh.wb, text).kind === 'formula' ? '=' + F.display(F.translate(O.checkFormula(text.replace(/^=/, '')).f, rr2 - r, cc - c)) : text;
              O.enter(sh, rr2, cc, t);
            }
          }
        } else O.enter(sh, r, c, text);
        O.autoRow(sh, r);
        autoWiden(sh, r, c);
      });
    } catch (e) {
      if (e.code === 'formula') { L.ui.msg('The formula you typed contains an error.', { icon: 'warn' }); return false; }
      L.app.error(e);
      return false;
    }
    end();
    L.app.afterEntry(r, c, text);
    return true;
  };
  /** a typed date / number that does not fit widens a column that still has the default width */
  function autoWiden(sh, r, c) {
    const cl = sh.get(r, c);
    if (!cl || typeof cl.v !== 'number' || cl.f != null) return;
    const co = sh.cols[c];
    if (co && co.custom) return;
    const st = LY.styleOf(sh, r, c, cl);
    if (!st.nf || st.nf === 'General') return;
    const g = LY.geo(sh);
    const f = Object.assign({}, st.font, { name: LY.fontName(sh.wb, st.font || {}) });
    const font = LY.cssFont(f, 1);
    const d = LY.display(sh, cl, st, null, font, { noHash: true });
    if (!d) return;
    const w = (d.segs ? LY.segWidth(d.segs, font) : LY.measure(font, d.text)) + 7;
    if (w > g.cols.size(c)) {
      const width = Math.round(M.pxToWidth(Math.ceil(w), M.mdw(sh.wb)) * 256) / 256;
      const before = co ? Object.assign({}, co) : null;
      const o = sh.colObj(c); o.w = width; o.bestFit = true;
      const after = Object.assign({}, o);
      LY.invalidate(sh);
      sh.wb.undo.op(() => { sh.cols[c] = before ? Object.assign({}, before) : undefined; LY.invalidate(sh); }, () => { sh.cols[c] = Object.assign({}, after); LY.invalidate(sh); });
    }
  }
  E.cancel = function () {
    if (!E.active) return;
    hideAux();
    end();
  };
  function end() {
    E.active = false;
    E.mode = 'ready';
    point = null;
    cell.value = '';
    cell.style.opacity = '0';
    cell.style.pointerEvents = 'none';
    cell.classList.remove('on');
    cell.style.width = '20px'; cell.style.height = '18px';
    fbwrap.classList.remove('editing');
    fbar.classList.remove('multi');
    G().state.hl = null;
    E.show();
    G().paint();
    G().focus();
    L.bus.emit('mode');
  }
  function hideAux() { if (acList) { acList.remove(); acList = null; } if (tipEl) { tipEl.remove(); tipEl = null; } }

  /* ------------------------------------------------------------ keys */
  E.key = function (e) {
    const k = e.key, ctrl = e.ctrlKey || e.metaKey, shift = e.shiftKey, alt = e.altKey;
    if (e.isComposing) return;
    if (acList && (k === 'ArrowDown' || k === 'ArrowUp' || (k === 'Enter' && acList._sel != null) || k === 'Escape')) { acKey(e); return; }
    const g = G();
    const moveAfter = (dr, dc) => { const ok = E.commit(); if (ok) { if (dr || dc) { const s = g.sel(); const rg = s.ranges[s.active]; if (s.ranges.length > 1 || rg.r1 !== rg.r2 || rg.c1 !== rg.c2) g.moveInSelection(dr, dc); else g.move(dr, dc); } } };
    switch (k) {
      case 'Enter':
        e.preventDefault();
        if (alt) { E.insertText('\n'); E.mode = 'edit'; return; }
        if (ctrl && shift) { E.commit({ array: true }); return; }
        if (ctrl) { E.commit({ fill: true }); return; }
        { const d = L.app.opts.moveAfterEnter === false ? null : L.app.opts.enterDir || 'down'; const m = d ? { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1] }[d] : [0, 0]; moveAfter(shift ? -m[0] : m[0], shift ? -m[1] : m[1]); }
        return;
      case 'Tab': e.preventDefault(); moveAfter(0, shift ? -1 : 1); return;
      case 'Escape': e.preventDefault(); E.cancel(); return;
      case 'F2': e.preventDefault(); E.mode = E.mode === 'edit' ? 'enter' : 'edit'; L.bus.emit('mode'); return;
      case 'F4': if (/^=/.test(surface.value)) { e.preventDefault(); toggleAbs(); } return;
      case 'F3': e.preventDefault(); L.ui.exec('pasteName'); return;
      case 'ArrowUp': case 'ArrowDown': case 'ArrowLeft': case 'ArrowRight': {
        const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
        if (E.mode === 'point') { e.preventDefault(); pointKey(d[0], d[1], shift, ctrl); return; }
        if (E.mode === 'enter' && surface === cell) { e.preventDefault(); moveAfter(d[0], d[1]); return; }
        return; /* edit mode: caret moves */
      }
      default:
    }
    if (ctrl && k === ';' && !shift) { e.preventDefault(); E.insertText(NF.text('m/d/yyyy', Math.floor(NF.jsDateToSerial(new Date(), false)))); return; }
    if (ctrl && (k === ':' || (k === ';' && shift))) { e.preventDefault(); E.insertText(NF.text('h:mm AM/PM', NF.jsDateToSerial(new Date(), false) % 1)); return; }
    if (ctrl && (k === 'a' || k === 'A') && shift && /^=/.test(surface.value)) { e.preventDefault(); insertArgNames(); return; }
    if (ctrl && (k === 'a' || k === 'A') && /^=/.test(surface.value)) { const fn = currentFunction(); if (fn && fn.complete) { e.preventDefault(); L.ui.exec('functionArgs'); return; } }
    if (ctrl && (k === 'z' || k === 'Z')) return; /* textarea undo */
    if (ctrl && (k === 'b' || k === 'i' || k === 'u')) { e.preventDefault(); return; }
    /* any other key leaves point mode */
    if (E.mode === 'point' && !ctrl && !alt && k.length === 1) point = null;
  };
  /* arrow keys while pointing move (or extend) the inserted reference */
  function pointKey(dr, dc, shift, ctrl) {
    const g = G();
    const sh = g.sheet();
    if (!point) {
      const s = g.sel();
      const base = sh === E.sheet ? { r: E.r, c: E.c } : { r: s.r, c: s.c };
      point = { start: surface.selectionStart, len: 0, anchor: base, cur: base, sheet: sh };
    }
    let cur = shift ? (point.ext || point.cur) : point.cur;
    let r = cur.r, c = cur.c;
    if (ctrl) [r, c] = g.edge(r, c, dr, dc); else { r = Math.max(0, Math.min(M.MAXR - 1, r + dr)); c = Math.max(0, Math.min(M.MAXC - 1, c + dc)); }
    if (shift) { point.ext = { r, c }; } else { point.cur = { r, c }; point.ext = null; }
    const a = point.cur, b = point.ext || point.cur;
    const rg = { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
    putRef(rg);
    g.reveal(b.r, b.c);
  }
  /** replace the current point reference with a reference to rg */
  function putRef(rg, sheetOverride) {
    const g = G();
    const sh = sheetOverride || g.sheet();
    const m = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? LY.mergeAt(sh, rg.r1, rg.c1) : null;
    const x = m || rg;
    let ref = x.r1 === 0 && x.r2 >= M.MAXR - 1 ? F.colName(x.c1) + ':' + F.colName(x.c2) : x.c1 === 0 && x.c2 >= M.MAXC - 1 ? (x.r1 + 1) + ':' + (x.r2 + 1) : F.rangeName(x);
    if (sh !== E.sheet) ref = F.quoteSheet(sh.name) + '!' + ref;
    const t = surface.value;
    const start = point.start;
    const nt = t.slice(0, start) + ref + t.slice(start + point.len);
    point.len = ref.length;
    cell.value = nt; fbar.value = nt;
    try { surface.setSelectionRange(start + ref.length, start + ref.length); } catch (e) { /* ignore */ }
    E.place(); E.highlight(); E.tip();
    L.bus.emit('mode');
  }
  /* mouse pointing (grid calls these) */
  E.pointStart = function (rg, add) {
    if (!point || add) point = { start: surface.selectionStart, len: 0 };
    if (add && !/[,(=]$/.test(surface.value.slice(0, surface.selectionStart))) { E.insertText(','); point = { start: surface.selectionStart, len: 0 }; }
    point.cur = { r: rg.r1, c: rg.c1 }; point.ext = null;
    putRef(rg);
  };
  E.pointMove = function (rg) { if (point) putRef(rg); };
  E.pointEnd = function () { surface.focus(); };
  /** the user switched sheets while pointing: references gain the sheet name */
  E.sheetChanged = function () { if (E.active) { E.place(); E.highlight(); } };

  /* F4: cycle $A$1 → A$1 → $A1 → A1 for the reference at the caret */
  function toggleAbs() {
    const t = surface.value, a = surface.selectionStart, b = surface.selectionEnd;
    const res = F.toggleAbs(t, a, b);
    if (!res) return;
    setText(res.text, res.caret);
    try { surface.setSelectionRange(res.start, res.caret); } catch (e) { /* ignore */ }
  }

  /* ------------------------------------------------------------ reference highlighting */
  E.highlight = function () {
    const g = G();
    const t = surface ? surface.value : '';
    if (!/^[=+-]/.test(t)) { g.state.hl = null; g.paint(); return; }
    let toks;
    try { toks = F.tokenize(t.slice(1)); } catch (e) { g.state.hl = null; g.paint(); return; }
    const hl = [];
    const seen = new Map();
    let k = 0;
    for (const tk of toks) {
      if (tk.t !== 'ref' || tk.book != null || tk.kind === 'err') continue;
      const sh = tk.sheet != null ? E.sheet.wb.sheetByName(tk.sheet) : E.sheet;
      if (!sh) continue;
      const rg = tk.kind === 'cell' ? { r1: tk.r, c1: tk.c, r2: tk.r, c2: tk.c } : tk.kind === 'cols' ? { r1: 0, c1: tk.c1, r2: M.MAXR - 1, c2: tk.c2 } : tk.kind === 'rows' ? { r1: tk.r1, c1: 0, r2: tk.r2, c2: M.MAXC - 1 } : { r1: Math.min(tk.r1, tk.r2), c1: Math.min(tk.c1, tk.c2), r2: Math.max(tk.r1, tk.r2), c2: Math.max(tk.c1, tk.c2) };
      const key = sh.id + F.rangeName(rg);
      if (!seen.has(key)) { seen.set(key, RD.HL[k++ % RD.HL.length]); }
      hl.push({ sh, range: rg, color: seen.get(key) });
    }
    g.state.hl = hl.length ? hl : null;
    g.paint();
  };

  /* ------------------------------------------------------------ function tips */
  /** function name and argument index at the caret */
  function currentFunction() {
    const t = surface.value;
    if (!/^=/.test(t)) return null;
    const pos = surface.selectionStart;
    const s = t.slice(1, pos);
    let depth = 0, arg = 0, q = false;
    const stack = [];
    let cur = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '"') { q = !q; continue; }
      if (q) continue;
      if (ch === '(') { const m = /([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(s.slice(0, i)); stack.push({ name: m ? m[1].toUpperCase().replace(/^_XLFN\./, '') : null, arg: 0 }); depth++; continue; }
      if (ch === ')') { stack.pop(); depth--; continue; }
      if ((ch === ',' || ch === ';') && stack.length && !inBraces(s, i)) stack[stack.length - 1].arg++;
      void cur; void arg;
    }
    const top = stack.filter((x) => x.name).pop();
    if (!top) return null;
    return { name: top.name, arg: top.arg, complete: true };
  }
  function inBraces(s, i) { const before = s.slice(0, i); return (before.match(/\{/g) || []).length > (before.match(/\}/g) || []).length; }
  E.currentFunction = () => (E.active ? currentFunction() : null);
  E.tip = function () {
    if (tipEl) { tipEl.remove(); tipEl = null; }
    if (!E.active || !L.fninfo) return;
    const fn = currentFunction();
    if (!fn) return;
    const info = L.fninfo.get(fn.name);
    if (!info) return;
    tipEl = h('div', { class: 'ftip' });
    const args = info.args || [];
    tipEl.appendChild(h('a', { text: fn.name, onclick: () => L.ui.exec('functionArgs') }));
    tipEl.appendChild(document.createTextNode('('));
    let shown = 0;
    args.forEach((a, i) => {
      if (shown) tipEl.appendChild(document.createTextNode(', '));
      const repeat = /\.\.\.$/.test(a);
      const active = i === fn.arg || (repeat && fn.arg >= i);
      const el = h(active ? 'b' : 'span', { text: a });
      tipEl.appendChild(el);
      shown++;
    });
    tipEl.appendChild(document.createTextNode(')'));
    const wrap = L.$('#gridwrap');
    if (surface === fbar) {
      const r = fbar.getBoundingClientRect(), w = wrap.getBoundingClientRect();
      Object.assign(tipEl.style, { left: Math.max(2, r.left - w.left + 4) + 'px', top: '2px' });
    } else {
      const q = G().cellRect(E.r, E.c);
      Object.assign(tipEl.style, { left: q.x + 'px', top: (q.y + Math.max(q.h, parseFloat(cell.style.height) || q.h) + 2) + 'px' });
    }
    wrap.appendChild(tipEl);
  };
  function insertArgNames() {
    const fn = currentFunction();
    if (!fn || !L.fninfo) return;
    const info = L.fninfo.get(fn.name);
    if (!info) return;
    const t = surface.value, pos = surface.selectionStart;
    if (t[pos - 1] !== '(') return;
    const names = info.args.filter((a) => !/^\[/.test(a)).map((a) => a.replace(/\.\.\.$/, '')).join(',');
    E.insertText(names + ')');
  }

  /* ------------------------------------------------------------ AutoComplete */
  E.autoComplete = function () {
    if (!L.app.opts.autoComplete || surface !== cell) return;
    const t = cell.value;
    if (!t || /^[=+\-\d.]/.test(t) || cell.selectionStart !== t.length || E.mode === 'edit') return;
    if (E._lastAC != null && t.length < E._lastAC) { E._lastAC = t.length; return; } /* deleting: no completion */
    E._lastAC = t.length;
    const sh = E.sheet, c = E.c;
    /* contiguous column block around the cell */
    const vals = new Map();
    const scan = (dir) => { let r = E.r + dir, n = 0; while (r >= 0 && r < M.MAXR && n < 5000) { const cl = sh.get(r, c); if (!cl || cl.v == null || cl.v === '') break; if (typeof cl.v === 'string' && cl.f == null) { const k = cl.v.toLowerCase(); if (!vals.has(k)) vals.set(k, cl.v); } r += dir; n++; } };
    scan(-1); scan(1);
    const low = t.toLowerCase();
    const hits = Array.from(vals.entries()).filter(([k]) => k.startsWith(low) && k.length > low.length);
    if (hits.length !== 1) return;
    const full = hits[0][1];
    const nt = t + full.slice(t.length);
    cell.value = nt; fbar.value = nt;
    cell.setSelectionRange(t.length, nt.length);
    E.place();
  };
  /** Alt+Down: pick from the column's entries */
  E.pickList = function (items, onPick) {
    hideAux();
    const q = G().cellRect(G().sel().r, G().sel().c);
    acList = h('div', { class: 'acl', role: 'listbox' });
    items.forEach((v, i) => {
      const d = h('div', { text: String(v), role: 'option' });
      d.addEventListener('pointerdown', (e) => { e.preventDefault(); e.stopPropagation(); const f = onPick; hideAux(); f(v); });
      acList.appendChild(d);
      d._i = i;
    });
    acList._items = items; acList._sel = null; acList._pick = onPick;
    Object.assign(acList.style, { left: q.x + 'px', top: (q.y + q.h) + 'px', minWidth: Math.max(q.w, 80) + 'px' });
    L.$('#gridwrap').appendChild(acList);
    const close = (e) => { if (acList && !acList.contains(e.target)) { hideAux(); document.removeEventListener('pointerdown', close, true); } };
    setTimeout(() => document.addEventListener('pointerdown', close, true), 0);
    E.listOpen = true;
  };
  E.listKey = function (e) { if (!acList) return false; acKey(e); return true; };
  function acKey(e) {
    e.preventDefault();
    const n = acList._items.length;
    if (e.key === 'Escape') { hideAux(); return; }
    if (e.key === 'Enter') { const v = acList._items[acList._sel]; const f = acList._pick; hideAux(); if (v != null) f(v); return; }
    acList._sel = acList._sel == null ? 0 : Math.max(0, Math.min(n - 1, acList._sel + (e.key === 'ArrowDown' ? 1 : -1)));
    Array.from(acList.children).forEach((d, i) => d.classList.toggle('on', i === acList._sel));
    const el = acList.children[acList._sel];
    if (el) el.scrollIntoView({ block: 'nearest' });
  }
  E.hasList = () => !!acList;
})(typeof window !== 'undefined' ? window : globalThis);
