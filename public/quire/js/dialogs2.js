/* Quire — dialogs, part 2: Insert menu (Break, Page Numbers, Date and Time, Field, Symbol,
 * Footnote, Caption, Cross-reference, Index and Tables, Bookmark, Hyperlink, Object, Diagram) and
 * the Table menu dialogs.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const G = L.dlg;
  const doc = () => D.doc;
  const A = () => L.app;
  const F = () => L.fields;

  /* ================= Break ================= */
  G.breakDlg = function () {
    const r = G.radios('br', [['page', '&Page break'], ['column', '&Column break'], ['wrap', 'Text &wrapping break'], ['nextPage', '&Next page'], ['continuous', 'Con&tinuous'], ['evenPage', '&Even page'], ['oddPage', '&Odd page']], 'page');
    const els = r.querySelectorAll('.ck');
    const body = h('div', { class: 'col' }, ui.group('Break types', els[0], els[1], els[2]), ui.group('Section break types', els[3], els[4], els[5], els[6]));
    ui.dialog({ title: 'Break', width: 260, body, buttons: [{ label: 'OK', primary: true, onClick: () => {
      const v = r.get();
      if (v === 'page' || v === 'column') { E.enter(v); return; }
      if (v === 'wrap') { E.enter('line'); return; }
      A().insertSectionBreak(v);
    } }, { label: 'Cancel' }] });
  };
  A().insertSectionBreak = function (type) {
    const d = doc();
    if (!E.sel || D.storyOf(d, E.sel.f.p) !== d.main || D.cellOf(d, E.sel.f.p)) { A().status('Section breaks can only be inserted in the main text, outside tables.'); return; }
    E.edit('Section Break', () => {
      const pos = E.deleteSelection();
      const cur = D.sectFor(d, pos.p);
      const np = O.splitPara(pos);
      D.touch(pos.p);
      pos.p.sect = L.clone(cur);
      /* the new section (after the break) carries the requested start type */
      const after = A().curSect ? null : null; void after;
      const tb = D.topBlock(d, np.p);
      const sec = D.sectOfBlock(d, tb.i);
      const holder = sec.endPara || d;
      D.touchKey(holder, 'sect');
      (holder === d ? d.sect : holder.sect).type = type;
      return np;
    });
  };

  /* ================= Page Numbers ================= */
  G.pageNumbers = function (mode) {
    const d = doc();
    const s = A().curSect();
    if (mode === 'format') { formatPageNumbers(s); return; }
    const pos = ui.select([['bottom', 'Bottom of page (Footer)'], ['top', 'Top of page (Header)']], 'bottom', () => draw(), { id: 'pn-pos' });
    const al = ui.select([['right', 'Right'], ['left', 'Left'], ['center', 'Center'], ['inside', 'Inside'], ['outside', 'Outside']], 'right', () => draw(), { id: 'pn-al' });
    const first = ui.check('&Show number on first page', true, null);
    const pv = h('div', { class: 'ps-page', style: 'width:80px;height:104px;margin:4px auto' });
    const draw = () => { L.clear(pv); const y = pos.value === 'top' ? 'top:6px' : 'bottom:6px'; const x = { left: 'left:8px', right: 'right:8px', center: 'left:36px', inside: 'right:8px', outside: 'right:8px' }[al.value]; pv.appendChild(h('span', { style: `position:absolute;${y};${x};width:8px;height:8px;border:1px solid #000;background:#fff` })); };
    draw();
    ui.dialog({
      title: 'Page Numbers', width: 360, body: G.row(h('div', { class: 'col' }, G.f('&Position:', pos), G.f('&Alignment:', al), first), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv)),
      buttons: [{ label: '&Format...', onClick: () => { formatPageNumbers(s); return false; } }, { label: 'OK', primary: true, onClick: () => {
        const kind = pos.value === 'top' ? 'hdr' : 'ftr';
        E.edit('Page Numbers', () => {
          D.touchKey(d, 'hf');
          const holder = s.holder;
          D.touchKey(holder, 'sect');
          const sect = holder === d ? d.sect : holder.sect;
          sect.refs = sect.refs || { hdr: {}, ftr: {} };
          sect.refs[kind] = Object.assign({}, sect.refs[kind]);
          const add = (ty, jc) => {
            let story = sect.refs[kind][ty] && d.hf[sect.refs[kind][ty]];
            if (!story) { const id = (kind === 'hdr' ? 'h' : 'f') + D.nid(); story = d.hf[id] = { kind, id, blocks: [D.para([], { style: kind === 'hdr' ? 'Header' : 'Footer' })] }; sect.refs[kind][ty] = id; }
            const fid = D.newFid();
            /* Word puts the number in its own paragraph aligned as chosen */
            story.blocks.push(D.para([D.item('fb', { fid, instr: ' PAGE ' }, { style: 'PageNumber' }), D.item('fs', { fid }, { style: 'PageNumber' }), D.text('1', { style: 'PageNumber' }), D.item('fe', { fid }, { style: 'PageNumber' })], { style: kind === 'hdr' ? 'Header' : 'Footer', jc }));
            if (!d.styles.PageNumber) { D.touchKey(d, 'styles'); d.styles.PageNumber = D.builtinStyles().PageNumber; }
          };
          const jc = (v, even) => (v === 'inside' ? (even ? 'right' : 'left') : v === 'outside' ? (even ? 'left' : 'right') : v);
          add('default', jc(al.value, false));
          if (d.settings.evenOdd) add('even', jc(al.value, true));
          sect.titlePg = !first.input.checked || sect.titlePg;
          if (!first.input.checked) { /* first page gets an empty header/footer */ if (!sect.refs[kind].first) { const id = (kind === 'hdr' ? 'h' : 'f') + D.nid(); d.hf[id] = { kind, id, blocks: [D.para([], { style: kind === 'hdr' ? 'Header' : 'Footer' })] }; sect.refs[kind].first = id; } }
          else if (sect.titlePg && sect.refs[kind].first) add('first', jc(al.value, false));
          d._idxDirty = true;
          return E.sel;
        });
      } }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
  };
  function formatPageNumbers(s) {
    const d = doc();
    const sect = s.sect;
    const pn = Object.assign({ fmt: 'decimal', start: null }, sect.pgNum || {});
    const fmt = ui.select([['decimal', '1, 2, 3, ...'], ['numberInDash', '- 1 -, - 2 -, - 3 -, ...'], ['lowerLetter', 'a, b, c, ...'], ['upperLetter', 'A, B, C, ...'], ['lowerRoman', 'i, ii, iii, ...'], ['upperRoman', 'I, II, III, ...']], pn.fmt, null, { id: 'pf-fmt' });
    const chap = ui.check('Include chapter &number', !!pn.chapStyle, null);
    const mode = G.radios('pf-m', [['cont', '&Continue from previous section'], ['start', 'Start &at:']], pn.start != null ? 'start' : 'cont');
    const startAt = G.num(pn.start != null ? pn.start : 1, { min: 0, id: 'pf-start' });
    ui.dialog({ title: 'Page Number Format', width: 330, body: h('div', { class: 'col' }, G.f('Number &format:', fmt, 'wide'), chap, ui.group('Page numbering', mode, startAt)), buttons: [{ label: 'OK', primary: true, onClick: () => {
      E.edit('Page Number Format', () => { D.touchKey(s.holder, 'sect'); const sc = s.holder === d ? d.sect : s.holder.sect; sc.pgNum = { fmt: fmt.value, start: mode.get() === 'start' ? startAt.get() : null, chapStyle: chap.input.checked ? 1 : undefined }; return E.sel; });
    } }, { label: 'Cancel' }] });
  }

  /* ================= Date and Time ================= */
  G.dateTime = function () {
    const now = new Date();
    const pics = F().DATE_PICTURES;
    const def = L.store.get('defDatePic', pics[0]);
    const lb = G.listBox(pics.map((p) => ({ value: p, label: F().formatDate(now, p) })), def, { height: 220, onDbl: () => ok.click() });
    const lang = ui.select([['en-US', 'English (U.S.)'], ['en-GB', 'English (U.K.)']], 'en-US', null, { id: 'dt-lang' });
    const auto = ui.check('&Update automatically', !!L.store.get('defDateAuto', false), null);
    const dlg = ui.dialog({ title: 'Date and Time', width: 360, body: h('div', { class: 'col' }, h('label', { text: 'Available formats:' }), lb, G.f('&Language:', lang), auto), buttons: [
      { label: '&Default...', onClick: () => { L.store.set('defDatePic', lb.get()); L.store.set('defDateAuto', auto.input.checked); A().status('Default date format saved.'); return false; } },
      { label: 'OK', primary: true, onClick: () => { const pic = lb.get(); if (auto.input.checked) A().insertFieldNow(`${/[hHms]/.test(pic) && !/[dMy]/.test(pic) ? 'TIME' : 'DATE'} \\@ "${pic}"`, 'Insert Date and Time'); else E.insertTextAt(F().formatDate(new Date(), pic), 'Insert Date and Time'); } },
      { label: 'Cancel' }], onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left') });
    const ok = dlg.buttons[1];
  };

  /* ================= Field ================= */
  const FIELD_CATS = {
    'Date and Time': ['CREATEDATE', 'DATE', 'EDITTIME', 'PRINTDATE', 'SAVEDATE', 'TIME'],
    'Document Information': ['AUTHOR', 'COMMENTS', 'DOCPROPERTY', 'FILENAME', 'FILESIZE', 'INFO', 'KEYWORDS', 'LASTSAVEDBY', 'NUMCHARS', 'NUMPAGES', 'NUMWORDS', 'SUBJECT', 'TEMPLATE', 'TITLE'],
    'Equations and Formulas': ['=', 'ADVANCE', 'EQ', 'SYMBOL'],
    'Index and Tables': ['INDEX', 'RD', 'TA', 'TC', 'TOA', 'TOC', 'XE'],
    'Links and References': ['AUTOTEXT', 'HYPERLINK', 'INCLUDEPICTURE', 'INCLUDETEXT', 'NOTEREF', 'PAGEREF', 'QUOTE', 'REF', 'STYLEREF'],
    'Mail Merge': ['ASK', 'COMPARE', 'DATABASE', 'FILLIN', 'IF', 'MERGEFIELD', 'MERGEREC', 'MERGESEQ', 'NEXT', 'NEXTIF', 'SET', 'SKIPIF'],
    Numbering: ['AUTONUM', 'AUTONUMLGL', 'AUTONUMOUT', 'BARCODE', 'LISTNUM', 'PAGE', 'REVNUM', 'SECTION', 'SECTIONPAGES', 'SEQ'],
    'User Information': ['USERADDRESS', 'USERINITIALS', 'USERNAME'],
  };
  const FIELD_DESC = { DATE: "Today's date", TIME: 'The current time', CREATEDATE: 'The date the document was created', SAVEDATE: 'The date the document was last saved', PRINTDATE: 'The date the document was last printed', EDITTIME: 'The total document editing time', AUTHOR: 'The name of the document author from Document Properties', TITLE: "The document's title from Document Properties", SUBJECT: "The document's subject", KEYWORDS: "The document's keywords", COMMENTS: 'The comments from Document Properties', FILENAME: "The document's name", NUMPAGES: 'The number of pages in the document', NUMWORDS: 'The number of words in the document', NUMCHARS: 'The number of characters in the document', PAGE: 'The number of the current page', SECTION: 'The number of the current section', SECTIONPAGES: 'The number of pages in the section', SEQ: 'Inserts an automatic sequence number', REF: 'Inserts the text marked by a bookmark', PAGEREF: 'Inserts the number of the page containing a bookmark', NOTEREF: 'Inserts the mark of a footnote or endnote', STYLEREF: 'Inserts the text from a like-style paragraph', TOC: 'Builds a table of contents', INDEX: 'Builds an index', XE: 'Marks an index entry', TC: 'Marks a table of contents entry', '=': 'Calculates the result of an expression', MERGEFIELD: 'Inserts a mail merge field', USERNAME: 'Your name from User Information', USERINITIALS: 'Your initials from User Information', USERADDRESS: 'Your postal address from User Information', LASTSAVEDBY: 'Name of the person who last saved the document', HYPERLINK: 'Opens and jumps to the specified file', QUOTE: 'Inserts literal text', SYMBOL: 'Inserts a symbol', IF: 'Evaluates arguments conditionally', REVNUM: "The document's revision number", TEMPLATE: "The document's template name", DOCPROPERTY: 'The value of a document property', MACROBUTTON: 'Runs a macro' };
  G.field = function (fieldRef, opts) {
    opts = opts || {};
    const editing = !!fieldRef;
    const cat = ui.select([['all', '(All)']].concat(Object.keys(FIELD_CATS).map((k) => [k, k])), 'all', () => fillNames(), { id: 'fl-cat' });
    const names = G.listBox([], null, { height: 200, onChange: (v) => { desc.textContent = FIELD_DESC[v] || ''; code.value = v === '=' ? '= ' : v + ' '; preview(); } });
    const fillNames = () => { const list = cat.value === 'all' ? Array.from(new Set(Object.values(FIELD_CATS).flat())).sort() : FIELD_CATS[cat.value]; names.setItems(list); };
    fillNames();
    const code = h('input', { type: 'text', id: 'fl-code', style: 'width:100%' });
    const desc = h('div', { class: 'tp-note', style: 'min-height:28px' });
    const res = h('div', { class: 'tp-box', style: 'min-height:22px' });
    const fmtSel = ui.select([['', '(none)'], ['Upper', 'Uppercase'], ['Lower', 'Lowercase'], ['FirstCap', 'First capital'], ['Caps', 'Title case'], ['Arabic', '1, 2, 3, ...'], ['roman', 'i, ii, iii, ...'], ['ROMAN', 'I, II, III, ...'], ['alphabetic', 'a, b, c, ...'], ['ALPHABETIC', 'A, B, C, ...'], ['Ordinal', '1st, 2nd, 3rd ...'], ['CardText', 'One, Two, Three'], ['OrdText', 'First, Second, Third']], '', () => preview(), { id: 'fl-fmt' });
    const datePic = ui.select([['', '(default)']].concat(F().DATE_PICTURES.map((p) => [p, p])), '', () => preview(), { id: 'fl-date' });
    const keep = ui.check('Preserve formatting during updates', true, null);
    const fullCode = () => { let c = code.value.trim(); if (fmtSel.value) c += ' \\* ' + fmtSel.value; if (datePic.value && /^(DATE|TIME|CREATEDATE|SAVEDATE|PRINTDATE)\b/i.test(c) && !/\\@/.test(c)) c += ` \\@ "${datePic.value}"`; if (keep.input.checked && !/MERGEFORMAT/i.test(c) && !/^(TOC|INDEX|=)/i.test(c)) c += ' \\* MERGEFORMAT'; return c; };
    const preview = () => { try { const r = F().evaluate(fullCode(), E.sel ? E.sel.f : null); res.textContent = r == null ? '(the result is produced when the field is inserted or updated)' : Array.isArray(r) ? '(generated paragraphs)' : String(r); } catch (e) { res.textContent = 'Error!'; } };
    if (editing) { code.value = fieldRef.it.instr.trim().replace(/\s*\\\*\s*MERGEFORMAT/i, ''); const t = F().parse(fieldRef.it.instr).type; names.set(t); desc.textContent = FIELD_DESC[t] || ''; }
    else if (opts.codes) code.value = '';
    code.addEventListener('input', preview);
    preview();
    ui.dialog({
      title: 'Field', width: 560, focus: opts.codes ? code : null,
      body: G.row(h('div', { class: 'col', style: 'width:190px' }, G.f('&Categories:', cat), h('label', { text: 'Field names:' }), names), h('div', { class: 'col', style: 'flex:1;min-width:220px' }, h('div', { class: 'cd-cap', text: 'Field codes:' }), code, desc, ui.group('Field options', G.f('&Format:', fmtSel), G.f('Date &picture:', datePic)), h('div', { class: 'cd-cap', text: 'Result preview' }), res, keep)),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const c = fullCode();
        if (!c) return false;
        const t = F().parse(c).type;
        if (editing) { F().setInstr(fieldRef, c); return; }
        if (t === 'TOC' || t === 'INDEX') { F().insertGenerated(c); return; }
        A().insertFieldNow(c);
      } }, { label: 'Cancel' }],
    });
  };

  /* ================= Symbol ================= */
  const SPECIALS = [['—', 'Em Dash', 'Alt+Ctrl+Num -'], ['–', 'En Dash', 'Ctrl+Num -'], ['‑', 'Nonbreaking Hyphen', 'Ctrl+_'], ['­', 'Optional Hyphen', 'Ctrl+-'], [' ', 'Em Space', ''], [' ', 'En Space', ''], [' ', '1/4 Em Space', ''], [' ', 'Nonbreaking Space', 'Ctrl+Shift+Space'], ['©', 'Copyright', 'Alt+Ctrl+C'], ['®', 'Registered', 'Alt+Ctrl+R'], ['™', 'Trademark', 'Alt+Ctrl+T'], ['§', 'Section', ''], ['¶', 'Paragraph', ''], ['…', 'Ellipsis', 'Alt+Ctrl+.'], ['‘', 'Single Opening Quote', "Ctrl+`,`"], ['’', 'Single Closing Quote', "Ctrl+','"], ['“', 'Double Opening Quote', 'Ctrl+`,"'], ['”', 'Double Closing Quote', "Ctrl+',\""], ['​', 'No-Width Optional Break', ''], ['⁠', 'No-Width Non Break', '']];
  const SUBSETS = [['Basic Latin', 0x20, 0x7e], ['Latin-1 Supplement', 0xa0, 0xff], ['Latin Extended-A', 0x100, 0x17f], ['Greek and Coptic', 0x370, 0x3ff], ['Cyrillic', 0x400, 0x4ff], ['General Punctuation', 0x2010, 0x205e], ['Currency Symbols', 0x20a0, 0x20bf], ['Letterlike Symbols', 0x2100, 0x214f], ['Number Forms', 0x2150, 0x218b], ['Arrows', 0x2190, 0x21ff], ['Mathematical Operators', 0x2200, 0x22ff], ['Miscellaneous Technical', 0x2300, 0x23ff], ['Box Drawing', 0x2500, 0x257f], ['Block Elements', 0x2580, 0x259f], ['Geometric Shapes', 0x25a0, 0x25ff], ['Miscellaneous Symbols', 0x2600, 0x26ff], ['Dingbats', 0x2700, 0x27bf]];
  G.symbol = function (opts) {
    opts = opts || {};
    let cur = null;
    const recent = L.store.get('recentSymbols', ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β']);
    const font = ui.select([['', '(normal text)'], ['Symbol', 'Symbol'], ['Wingdings', 'Wingdings'], ['Webdings', 'Webdings']].concat(G.FONTS().map((f) => [f, f])), '', () => fill(), { id: 'sy-font' });
    const subset = ui.select(SUBSETS.map((s, i) => [i, s[0]]), 1, () => fill(), { id: 'sy-sub' });
    const grid = h('div', { class: 'sym-grid' });
    const codeIn = h('input', { type: 'text', id: 'sy-code', style: 'width:70px' });
    const name = h('div', { class: 'tp-note', style: 'min-height:16px' });
    const fill = () => {
      L.clear(grid);
      const f = font.value;
      const chars = [];
      if (f === 'Symbol' || f === 'Wingdings' || f === 'Webdings') for (let c = 0x21; c <= 0xff; c++) { if (c >= 0x7f && c < 0xa1) continue; const u = L.mapSymbolChar(String.fromCharCode(c), f); if (u) chars.push({ ch: u, raw: String.fromCharCode(c), font: f }); }
      else { const [, a, b] = SUBSETS[+subset.value]; for (let c = a; c <= b; c++) chars.push({ ch: String.fromCodePoint(c), font: f }); }
      for (const c of chars) {
        const b = h('button', { type: 'button', text: c.ch, style: c.font && !/Symbol|Wingdings|Webdings/.test(c.font) ? `font-family:${L.fontStack(c.font)}` : '' });
        b.addEventListener('click', () => { L.$$('button.on', grid).forEach((x) => x.classList.remove('on')); b.classList.add('on'); cur = c; codeIn.value = c.ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); name.textContent = 'Unicode character ' + codeIn.value; });
        b.addEventListener('dblclick', () => { cur = c; insertBtn.click(); });
        grid.appendChild(b);
      }
    };
    fill();
    codeIn.addEventListener('change', () => { const n = parseInt(codeIn.value, 16); if (!isNaN(n)) { cur = { ch: String.fromCodePoint(n), font: '' }; name.textContent = 'Unicode character ' + codeIn.value.toUpperCase(); } });
    const rec = h('div', { class: 'sym-recent' });
    for (const c of recent.slice(0, 16)) { const b = h('button', { type: 'button', text: c }); b.addEventListener('click', () => { cur = { ch: c, font: '' }; codeIn.value = c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); }); b.addEventListener('dblclick', () => { cur = { ch: c, font: '' }; insertBtn.click(); }); rec.appendChild(b); }
    const spec = G.listBox(SPECIALS.map(([c, n, k]) => ({ value: c, label: `${n}${k ? '   ' + k : ''}` })), SPECIALS[0][0], { height: 220, onDbl: () => insertBtn.click() });
    const tabs = ui.tabs([{ label: '&Symbols', body: h('div', { class: 'col' }, G.row(G.f('&Font:', font), G.f('Subset:', subset)), grid, h('div', { class: 'cd-cap', text: 'Recently used symbols:' }), rec, G.row(name, G.f('Character code:', codeIn))) }, { label: 'S&pecial Characters', body: spec }]);
    let tabIdx = 0;
    tabs.querySelectorAll('.tab').forEach((t, i) => t.addEventListener('click', () => { tabIdx = i; }));
    const doInsert = () => {
      const c = tabIdx === 1 ? { ch: spec.get(), font: '' } : cur;
      if (!c) return;
      if (opts.pick) { opts.pick(c.ch, c.font); return; }
      const list = [c.ch].concat(recent.filter((x) => x !== c.ch)).slice(0, 16);
      L.store.set('recentSymbols', list);
      if (c.raw && c.font) E.insertItem(D.item('sym', { font: c.font, char: c.raw, code: c.raw.charCodeAt(0).toString(16).toUpperCase() }, O.inheritRPr(E.sel.f)), 'Insert Symbol');
      else E.insertTextAt(c.ch, 'Insert Symbol');
    };
    const dlg = ui.dialog({ title: 'Symbol', width: 520, body: tabs, buttons: [{ label: '&Insert', primary: true, onClick: () => { doInsert(); if (opts.pick) return; return false; } }, { label: opts.pick ? 'Cancel' : 'Close' }] });
    const insertBtn = dlg.buttons[0];
  };

  /* ================= Footnote and Endnote ================= */
  G.footnote = function () {
    const d = doc();
    const kind = G.radios('fn-k', [['fn', '&Footnotes:'], ['en', '&Endnotes:']], 'fn', () => sync());
    const fnLoc = ui.select([['pageBottom', 'Bottom of page'], ['beneathText', 'Below text']], (d.settings.fnPr && d.settings.fnPr.pos) || 'pageBottom', null, { id: 'fn-floc' });
    const enLoc = ui.select([['docEnd', 'End of document'], ['sectEnd', 'End of section']], (d.settings.enPr && d.settings.enPr.pos) || 'docEnd', null, { id: 'fn-eloc' });
    const fmt = ui.select([['decimal', '1, 2, 3, ...'], ['lowerLetter', 'a, b, c, ...'], ['upperLetter', 'A, B, C, ...'], ['lowerRoman', 'i, ii, iii, ...'], ['upperRoman', 'I, II, III, ...'], ['chicago', '*, †, ‡, §, ...']], 'decimal', null, { id: 'fn-fmt' });
    const custom = h('input', { type: 'text', id: 'fn-custom', style: 'width:60px' });
    const startAt = G.num(1, { min: 1, id: 'fn-start' });
    const numbering = ui.select([['continuous', 'Continuous'], ['eachSect', 'Restart each section'], ['eachPage', 'Restart each page']], 'continuous', null, { id: 'fn-num' });
    const sync = () => { const k = kind.get(); const pr = (k === 'fn' ? d.settings.fnPr : d.settings.enPr) || {}; fmt.value = pr.fmt || (k === 'fn' ? 'decimal' : 'lowerRoman'); startAt.set(pr.start || 1); numbering.value = pr.restart || 'continuous'; };
    sync();
    const settings = () => ({ fmt: fmt.value, start: startAt.get(), restart: numbering.value, pos: kind.get() === 'fn' ? fnLoc.value : enLoc.value });
    const kEls = kind.querySelectorAll('.ck');
    ui.dialog({
      title: 'Footnote and Endnote', width: 380,
      body: h('div', { class: 'col' }, ui.group('Location', G.row(kEls[0], fnLoc), G.row(kEls[1], enLoc), ui.button('C&onvert...', () => convert(), { class: 'btn small' })), ui.group('Format', G.f('Number f&ormat:', fmt, 'wide'), G.row(G.f('Custom &mark:', custom, 'wide'), ui.button('S&ymbol...', () => G.symbol({ pick: (ch) => { custom.value = ch; } }), { class: 'btn small' })), G.f('&Start at:', startAt, 'wide'), G.f('Numbering:', numbering, 'wide'))),
      buttons: [{ label: '&Insert', primary: true, onClick: () => { L.notes.setOptions(kind.get(), settings()); L.notes.insert(kind.get(), { mark: custom.value.trim() || null }); } }, { label: 'Cancel' }, { label: '&Apply', onClick: () => { L.notes.setOptions(kind.get(), settings()); return false; } }],
    });
    function convert() {
      const r = G.radios('fn-conv', [['fn', 'Convert all &footnotes to endnotes'], ['en', 'Convert all &endnotes to footnotes']], 'fn');
      ui.dialog({ title: 'Convert Notes', width: 300, body: r, buttons: [{ label: 'OK', primary: true, onClick: () => L.notes.convertAll(r.get()) }, { label: 'Cancel' }] });
    }
  };

  /* ================= Caption ================= */
  G.caption = function () {
    const labels = Array.from(new Set(F().CAPTION_LABELS.concat(L.store.get('captionLabels', []))));
    const ci = L.tables && L.tables.inTable();
    const lab = ui.select(labels.map((x) => [x, x]), ci ? 'Table' : 'Figure', () => upd(), { id: 'cp-lab' });
    const txt = h('input', { type: 'text', id: 'cp-text', style: 'width:100%' });
    const pos = ui.select([['below', 'Below selected item'], ['above', 'Above selected item']], ci ? 'above' : 'below', null, { id: 'cp-pos' });
    const excl = ui.check('&Exclude label from caption', false, () => upd());
    let num = { fmt: 'ARABIC', chapter: false, chapterStyle: 1, sep: '-' };
    const upd = () => { const n = L.D.fmtNum(1, { ARABIC: 'decimal', alphabetic: 'lowerLetter', ALPHABETIC: 'upperLetter', roman: 'lowerRoman', ROMAN: 'upperRoman' }[num.fmt]); txt.value = (excl.input.checked ? '' : lab.value + ' ') + (num.chapter ? '1' + num.sep : '') + n; };
    upd();
    ui.dialog({
      title: 'Caption', width: 360, focus: txt,
      body: h('div', { class: 'col' }, h('label', { for: 'cp-text', html: '<u>C</u>aption:' }), txt, ui.group('Options', G.f('&Label:', lab), G.f('&Position:', pos), excl), h('div', { class: 'tp-btns' }, ui.button('&New Label...', async () => { const n = await ui.prompt('Label:', '', 'New Label'); if (!n) return; const list = L.store.get('captionLabels', []); if (!list.includes(n)) { list.push(n); L.store.set('captionLabels', list); } lab.appendChild(h('option', { value: n, text: n })); lab.value = n; upd(); }, { class: 'btn small' }), ui.button('&Delete Label', () => { if (F().CAPTION_LABELS.includes(lab.value)) return; L.store.set('captionLabels', L.store.get('captionLabels', []).filter((x) => x !== lab.value)); lab.querySelector(`option[value="${lab.value}"]`).remove(); upd(); }, { class: 'btn small' }), ui.button('N&umbering...', () => numbering(), { class: 'btn small' }))),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const prefix = (excl.input.checked ? '' : lab.value + ' ');
        let rest = txt.value;
        const m = new RegExp('^' + prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:[\\dA-Za-z]+' + (num.chapter ? '[-.:—–]?[\\dA-Za-z]+' : '') + ')').exec(rest);
        rest = m ? rest.slice(m[0].length) : '';
        F().insertCaption({ label: lab.value, position: pos.value, text: rest, fmt: num.fmt, chapter: num.chapter, chapterStyle: num.chapterStyle, chapterSep: num.sep, noLabel: excl.input.checked });
      } }, { label: 'Cancel' }],
    });
    function numbering() {
      const fmt = ui.select([['ARABIC', '1, 2, 3, ...'], ['alphabetic', 'a, b, c, ...'], ['ALPHABETIC', 'A, B, C, ...'], ['roman', 'i, ii, iii, ...'], ['ROMAN', 'I, II, III, ...']], num.fmt, null, { id: 'cn-fmt' });
      const chap = ui.check('Include &chapter number', num.chapter, null);
      const style = ui.select([1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => [i, 'Heading ' + i]), num.chapterStyle, null, { id: 'cn-st' });
      const sep = ui.select([['-', '- (hyphen)'], ['.', '. (period)'], [':', ': (colon)'], ['—', '— (em dash)'], ['–', '– (en dash)']], num.sep, null, { id: 'cn-sep' });
      ui.dialog({ title: 'Caption Numbering', width: 320, body: h('div', { class: 'col' }, G.f('F&ormat:', fmt), chap, G.f('Chapter starts with style:', style, 'wide'), G.f('Use separator:', sep, 'wide')), buttons: [{ label: 'OK', primary: true, onClick: () => { num = { fmt: fmt.value, chapter: chap.input.checked, chapterStyle: +style.value, sep: sep.value }; upd(); } }, { label: 'Cancel' }] });
    }
  };

  /* ================= Cross-reference ================= */
  G.crossRef = function () {
    const types = ['Numbered item', 'Heading', 'Bookmark', 'Footnote', 'Endnote', 'Equation', 'Figure', 'Table'].concat(L.store.get('captionLabels', []));
    const what = { 'Numbered item': [['number', 'Paragraph number'], ['numberFull', 'Paragraph number (full context)'], ['text', 'Paragraph text'], ['page', 'Page number'], ['abovebelow', 'Above/below']], Heading: [['text', 'Heading text'], ['page', 'Page number'], ['number', 'Heading number'], ['abovebelow', 'Above/below']], Bookmark: [['text', 'Bookmark text'], ['page', 'Page number'], ['number', 'Paragraph number'], ['abovebelow', 'Above/below']], Footnote: [['notenum', 'Footnote number'], ['page', 'Page number'], ['abovebelow', 'Above/below']], Endnote: [['notenum', 'Endnote number'], ['page', 'Page number'], ['abovebelow', 'Above/below']] };
    const capWhat = [['text', 'Entire caption'], ['labelnum', 'Only label and number'], ['page', 'Page number'], ['abovebelow', 'Above/below']];
    const type = ui.select(types.map((t) => [t, t]), 'Heading', () => refill(), { id: 'xr-type' });
    const ref = ui.select([], '', null, { id: 'xr-ref' });
    const link = ui.check('Insert as &hyperlink', true, null);
    const ab = ui.check('Include above/&below', false, null);
    const list = G.listBox([], null, { height: 200 });
    let items = [];
    const refill = () => {
      L.clear(ref);
      for (const [v, l] of what[type.value] || capWhat) ref.appendChild(h('option', { value: v, text: l }));
      items = F().xrefItems(type.value);
      list.setItems(items.map((x, i) => ({ value: String(i), label: x.text || '(empty)' })));
      if (items.length) list.set('0');
    };
    refill();
    const dlg = ui.dialog({
      title: 'Cross-reference', width: 460, modeless: true,
      body: h('div', { class: 'col' }, G.row(G.f('Reference type:', type), G.f('Insert &reference to:', ref)), G.row(link, ab), h('div', { class: 'cd-cap', text: 'For which item:' }), list),
      buttons: [{ label: '&Insert', primary: true, onClick: () => { const it = items[+list.get()]; if (!it) return false; F().insertXref(type.value, it, ref.value, { link: link.input.checked, abovebelow: ab.input.checked }); return false; } }, { label: 'Close' }],
    });
    void dlg;
  };

  /* ================= Index and Tables ================= */
  G.indexTables = function (tab) {
    const leaders = [['dot', '.......'], ['hyphen', '-------'], ['underscore', '_______'], ['none', '(none)']];
    /* TOC */
    const tShow = ui.check('&Show page numbers', true, () => pvToc());
    const tRight = ui.check('&Right align page numbers', true, () => pvToc());
    const tLead = ui.select(leaders, 'dot', () => pvToc(), { id: 'it-tl' });
    const tLinks = ui.check('&Use hyperlinks instead of page numbers (Web)', true, null);
    const tLevels = G.num(3, { min: 1, max: 9, id: 'it-lv', onChange: () => pvToc() });
    const tFormat = ui.select([['template', 'From template'], ['classic', 'Classic'], ['distinctive', 'Distinctive'], ['formal', 'Formal'], ['simple', 'Simple']], 'template', null, { id: 'it-tf' });
    let tStyles = null;
    const tPv = h('div', { class: 'preview-box', style: 'height:150px;flex-direction:column;align-items:stretch;justify-content:flex-start;padding:6px;font:11px Times New Roman,Tinos,serif' });
    const pvToc = () => { L.clear(tPv); for (let i = 1; i <= Math.min(3, tLevels.get()); i++) tPv.appendChild(h('div', { style: `display:flex;padding-left:${(i - 1) * 12}px;${i === 1 ? 'font-weight:bold' : ''}` }, h('span', { text: 'Heading ' + i }), h('span', { style: `flex:1;${tRight.input.checked && tLead.value !== 'none' ? 'border-bottom:1px dotted #000;margin:0 2px 3px' : ''}` }), h('span', { text: tShow.input.checked ? String(i * 2 - 1) : '' }))); };
    pvToc();
    const tocOptions = () => {
      const d = doc();
      const rows = h('div', { style: 'max-height:240px;overflow:auto;border:1px solid #919b9c;background:#fff' });
      const ids = Object.keys(d.styles).filter((k) => d.styles[k].type === 'paragraph').sort((a, b) => D.styleDisplayName(d.styles[a]).localeCompare(D.styleDisplayName(d.styles[b])));
      const inputs = {};
      for (const k of ids) { const lv = D.styleProps(d, k).pPr.outline; const inp = h('input', { type: 'text', style: 'width:36px', value: tStyles ? tStyles[k] || '' : lv != null && lv < tLevels.get() && /^Heading\d$/.test(k) ? String(lv + 1) : '' }); inputs[k] = inp; rows.appendChild(h('div', { class: 'row', style: 'justify-content:space-between;padding:1px 4px' }, h('span', { text: D.styleDisplayName(d.styles[k]) }), inp)); }
      const outl = ui.check('&Outline levels', true, null), fields = ui.check('Table &entry fields', false, null);
      ui.dialog({ title: 'Table of Contents Options', width: 340, body: h('div', { class: 'col' }, h('div', { text: 'Build table of contents from: Styles — TOC level:' }), rows, outl, fields), buttons: [{ label: 'OK', primary: true, onClick: () => { tStyles = {}; for (const k in inputs) if (inputs[k].value.trim()) tStyles[k] = +inputs[k].value; tStyles.__outline = outl.input.checked; tStyles.__fields = fields.input.checked; } }, { label: 'Cancel' }] });
    };
    const tocBody = h('div', { class: 'col' }, G.row(h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: 'Print Preview' }), tPv, tShow, tRight, G.f('Ta&b leader:', tLead)), h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: 'Web Preview' }), h('div', { class: 'preview-box', style: 'height:150px;flex-direction:column;align-items:flex-start;justify-content:flex-start;padding:6px;gap:3px' }, h('u', { text: 'Heading 1', style: 'color:#00f' }), h('u', { text: 'Heading 2', style: 'color:#00f;margin-left:10px' }), h('u', { text: 'Heading 3', style: 'color:#00f;margin-left:20px' })), tLinks)), ui.group('General', G.row(G.f('Forma&ts:', tFormat), G.f('Show &levels:', tLevels))), h('div', { class: 'tp-btns' }, ui.button('&Options...', tocOptions, { class: 'btn small' }), ui.button('&Modify...', () => G.style('TOC1'), { class: 'btn small' })));
    /* Table of figures */
    const fLabel = ui.select(['Figure', 'Table', 'Equation'].concat(L.store.get('captionLabels', [])).map((x) => [x, x]), 'Figure', null, { id: 'it-fl' });
    const fIncl = ui.check('Include label and &number', true, null);
    const fShow = ui.check('&Show page numbers', true, null);
    const fRight = ui.check('&Right align page numbers', true, null);
    const fLead = ui.select(leaders, 'dot', null, { id: 'it-fle' });
    const tofBody = h('div', { class: 'col' }, G.f('Caption &label:', fLabel), fIncl, fShow, fRight, G.f('Ta&b leader:', fLead), h('div', { class: 'tp-note', text: 'Captions inserted with Insert ▸ Reference ▸ Caption are collected here.' }));
    /* Index */
    const iType = G.radios('it-it', [['indent', 'Inde&nted'], ['runin', 'Ru&n-in']], 'indent');
    const iCols = G.num(2, { min: 1, max: 4, id: 'it-ic' });
    const iRight = ui.check('&Right align page numbers', false, null);
    const iLead = ui.select(leaders, 'dot', null, { id: 'it-il' });
    const iHead = ui.check('Show letter &headings', false, null);
    const idxBody = h('div', { class: 'col' }, G.row(ui.group('Type', iType), h('div', { class: 'col' }, G.f('C&olumns:', iCols), iRight, G.f('Ta&b leader:', iLead), iHead)), h('div', { class: 'tp-btns' }, ui.button('Mar&k Entry...', () => G.markEntry(), { class: 'btn small' })), h('div', { class: 'tp-note', text: 'Mark entries first (Alt+Shift+X), then insert the index where it should appear.' }));
    const toaBody = h('div', { class: 'tp-note', text: 'Tables of authorities need citations marked with TA fields. Use Insert ▸ Field ▸ TA to mark citations, then Field ▸ TOA to build the table.' });
    let cur = tab === 'toc' || tab == null ? 1 : tab === 'tof' ? 2 : tab === 'index' ? 0 : 1;
    const tabs = ui.tabs([{ label: 'Inde&x', body: idxBody, onShow: () => { cur = 0; } }, { label: 'Table of &Contents', body: tocBody, onShow: () => { cur = 1; } }, { label: 'Table of &Figures', body: tofBody, onShow: () => { cur = 2; } }, { label: 'Table of &Authorities', body: toaBody, onShow: () => { cur = 3; } }], cur);
    ui.dialog({
      title: 'Index and Tables', width: 520, body: tabs,
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (cur === 1) {
          let instr = `TOC \\o "1-${tLevels.get()}"`;
          if (tStyles) { const st = Object.entries(tStyles).filter(([k]) => !k.startsWith('__') && !/^Heading\d$/.test(k)).map(([k, v]) => `${D.styleDisplayName(doc().styles[k])},${v}`).join(','); if (!tStyles.__outline) instr = 'TOC'; if (st) instr += ` \\t "${st}"`; if (tStyles.__fields) instr += ' \\f'; }
          if (tLinks.input.checked) instr += ' \\h';
          instr += ' \\z \\u';
          if (!tShow.input.checked) instr += ' \\n';
          else if (!tRight.input.checked) instr += ' \\p " "';
          F().insertGenerated(instr);
          if (tFormat.value !== 'template') applyTocFormat(tFormat.value, tLevels.get());
        } else if (cur === 2) {
          let instr = `TOC \\h \\z \\c "${fLabel.value}"`;
          if (!fIncl.input.checked) instr += ` \\a "${fLabel.value}"`;
          if (!fShow.input.checked) instr += ' \\n';
          F().insertGenerated(instr);
        } else if (cur === 0) {
          let instr = `INDEX \\c "${iCols.get()}"`;
          if (iType.get() === 'runin') instr += ' \\r';
          if (iHead.input.checked) instr += ' \\h "A"';
          if (iRight.input.checked) instr += ' \\e "\t"';
          F().insertGenerated(instr, { columns: iCols.get() });
        }
      } }, { label: 'Cancel' }],
    });
  };
  function applyTocFormat(kind, levels) {
    const d = doc();
    const look = { classic: [{ b: true, caps: true }, { smallCaps: true }, { i: true }], distinctive: [{ b: true, font: 'Arial', sz: 12 }, { font: 'Arial', sz: 11 }, { sz: 10 }], formal: [{ caps: true, b: true }, { caps: true }, { i: true }], simple: [{}, {}, {}] }[kind];
    E.edit('TOC Format', () => { D.touchKey(d, 'styles'); for (let i = 1; i <= levels; i++) { const id = 'TOC' + i; d.styles[id] = d.styles[id] || D.builtinStyles()[id]; d.styles[id].rPr = Object.assign({}, look[Math.min(i - 1, 2)]); } D.stylesChanged(); return E.sel; });
  }
  G.markEntry = function () {
    const sel = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim() : '';
    const main = h('input', { type: 'text', id: 'me-main', value: sel, style: 'width:100%' });
    const sub = h('input', { type: 'text', id: 'me-sub', style: 'width:100%' });
    const opt = G.radios('me-o', [['cur', '&Current page'], ['xref', 'Cross-&reference:']], 'cur');
    const xref = h('input', { type: 'text', id: 'me-x', value: 'See ', style: 'width:100%' });
    const bold = ui.check('&Bold', false, null), ital = ui.check('&Italic', false, null);
    const mark = (all) => {
      const m = main.value.trim();
      if (!m) return;
      const entry = (m + (sub.value.trim() ? ':' + sub.value.trim() : '')).replace(/"/g, '\\"');
      let instr = `XE "${entry}"`;
      if (opt.get() === 'xref') instr += ` \\t "${xref.value.replace(/"/g, '\\"')}"`;
      if (bold.input.checked) instr += ' \\b';
      if (ital.input.checked) instr += ' \\i';
      if (!all) { E.edit('Mark Index Entry', () => { const [, b] = E.range(); return O.insertField(b, instr, '', {}); }); return; }
      /* Mark All: every occurrence of the text in the main story */
      const d = doc();
      D.reindex(d);
      let n = 0;
      E.edit('Mark Index Entries', () => {
        for (const p of d.main.paras.slice()) {
          const t = D.ptext(p);
          let i = t.toLowerCase().indexOf(m.toLowerCase());
          const hits = [];
          while (i >= 0) { hits.push(i + m.length); i = t.toLowerCase().indexOf(m.toLowerCase(), i + m.length); }
          for (const at of hits.reverse()) { O.insertField(D.pos(p, at), instr, '', {}); n++; }
        }
        return E.sel;
      });
      A().status(`${n} entr${n === 1 ? 'y' : 'ies'} marked.`);
    };
    ui.dialog({ title: 'Mark Index Entry', width: 360, modeless: true, body: h('div', { class: 'col' }, G.f('Main &entry:', main, 'wide'), G.f('&Subentry:', sub, 'wide'), ui.group('Options', opt, xref), ui.group('Page number format', bold, ital)), buttons: [{ label: '&Mark', primary: true, onClick: () => { mark(false); return false; } }, { label: 'Mark &All', onClick: () => { mark(true); return false; } }, { label: 'Close' }] });
  };

  /* ================= Bookmark ================= */
  G.bookmark = function () {
    const sortBy = G.radios('bm-s', [['name', '&Name'], ['loc', '&Location']], 'name', () => fill());
    const hidden = ui.check('&Hidden bookmarks', false, () => fill());
    const name = h('input', { type: 'text', id: 'bm-name', style: 'width:100%' });
    const list = G.listBox([], null, { height: 160, onChange: (v) => { name.value = v; }, onDbl: () => go() });
    const fill = () => {
      const all = Array.from(F().bookmarks().values()).filter((b) => hidden.input.checked || !b.name.startsWith('_'));
      if (sortBy.get() === 'name') all.sort((a, b) => a.name.localeCompare(b.name)); else all.sort((a, b) => (a.p._o || 0) - (b.p._o || 0) || a.o - b.o);
      list.setItems(all.map((b) => b.name));
    };
    fill();
    if (E.sel && !E.collapsed()) { const t = O.textRange(...E.range(), ' ').trim().replace(/[^\p{L}\p{N}_]/gu, '_').replace(/^[^\p{L}]+/u, '').slice(0, 40); name.value = t; }
    const valid = (n) => /^[\p{L}][\p{L}\p{N}_]{0,39}$/u.test(n);
    const go = () => { const n = name.value.trim(); if (n) A().gotoBookmark(n); };
    ui.dialog({
      title: 'Bookmark', width: 320, focus: name,
      body: h('div', { class: 'col' }, h('label', { for: 'bm-name', html: '<u>B</u>ookmark name:' }), name, list, ui.group('Sort by:', sortBy), hidden),
      buttons: [{ label: '&Add', primary: true, onClick: () => { const n = name.value.trim(); if (!valid(n)) { ui.msg('Bookmark names must begin with a letter and contain only letters, numbers and underscores.', { icon: 'warn' }); return false; } const [a, b] = E.range(); E.edit('Bookmark', () => { F().addBookmark(n, a, b); return E.sel; }); } }, { label: '&Delete', onClick: () => { const bm = F().bookmarks().get(name.value.trim()); if (bm) E.edit('Delete Bookmark', () => { F().removeBookmarkItems(bm); return E.sel; }); fill(); name.value = ''; return false; } }, { label: '&Go To', onClick: () => { go(); return false; } }, { label: 'Cancel' }],
    });
  };

  /* ================= Hyperlink ================= */
  G.hyperlink = function () {
    const d = doc();
    let cur = null;
    if (E.sel) {
      const it = D.itemAfter(E.sel.f.p, E.sel.f.o) || D.itemBefore(E.sel.f.p, E.sel.f.o);
      if (it && it.rPr && it.rPr.link) cur = it.rPr.link;
    }
    /* the whole link text when the caret sits in a link */
    let range = E.sel ? E.range() : null;
    if (cur && E.collapsed()) {
      const p = E.sel.f.p; let pos = 0, s = -1, e = -1;
      for (const it of p.runs) { const n = D.ilen(it); if (it.rPr && it.rPr.link && JSON.stringify(it.rPr.link) === JSON.stringify(cur) && pos <= E.sel.f.o + 1) { if (s < 0) s = pos; e = pos + n; } else if (s >= 0 && pos >= E.sel.f.o) break; pos += n; }
      if (s >= 0) range = [D.pos(p, s), D.pos(p, e)];
    }
    const selText = range && !D.eqPos(range[0], range[1]) ? O.textRange(range[0], range[1], ' ') : '';
    const disp = h('input', { type: 'text', id: 'hl-disp', value: selText || '<<Selection in Document>>', style: 'width:100%' });
    if (!selText) disp.value = '';
    const addr = h('input', { type: 'text', id: 'hl-addr', value: cur && cur.url && !/^mailto:/.test(cur.url) ? cur.url : '', style: 'width:100%' });
    const tip = cur && cur.tip ? cur.tip : '';
    const place = G.listBox([], cur && cur.anchor ? cur.anchor : null, { height: 170 });
    const places = [{ value: '', label: 'Top of the Document' }];
    D.reindex(d);
    places.push({ value: '', label: 'Headings', group: true });
    for (const p of d.main.paras) { const lv = D.headingLevel(d, p); if (lv) places.push({ value: '§' + p.id, label: '    '.repeat(lv - 1) + D.plainText(p).trim().slice(0, 60) }); }
    places.push({ value: '', label: 'Bookmarks', group: true });
    for (const [n] of F().bookmarks()) if (!n.startsWith('_')) places.push({ value: n, label: n });
    place.setItems(places);
    const mail = h('input', { type: 'text', id: 'hl-mail', value: cur && /^mailto:/.test(cur.url || '') ? cur.url.slice(7).split('?')[0] : '', style: 'width:100%' });
    const subj = h('input', { type: 'text', id: 'hl-subj', style: 'width:100%' });
    let kind = cur && cur.anchor && !cur.url ? 1 : cur && /^mailto:/.test(cur.url || '') ? 2 : 0;
    const tabs = ui.tabs([
      { label: 'E&xisting File or Web Page', body: h('div', { class: 'col' }, h('label', { for: 'hl-addr', html: 'Addr<u>e</u>ss:' }), addr, h('div', { class: 'tp-note', text: 'Type a web address (https://…) or the name of a file.' })), onShow: () => { kind = 0; } },
      { label: 'Place in This &Document', body: h('div', { class: 'col' }, h('div', { text: 'Select a place in this document:' }), place), onShow: () => { kind = 1; } },
      { label: 'E-&mail Address', body: h('div', { class: 'col' }, h('label', { for: 'hl-mail', text: 'E-mail address:' }), mail, h('label', { for: 'hl-subj', text: 'Subject:' }), subj), onShow: () => { kind = 2; } },
    ], kind);
    let screenTip = tip;
    ui.dialog({
      title: cur ? 'Edit Hyperlink' : 'Insert Hyperlink', width: 520, focus: kind === 0 ? addr : null,
      body: h('div', { class: 'col' }, G.row(h('label', { for: 'hl-disp', html: '<u>T</u>ext to display:', style: 'width:110px' }), h('div', { style: 'flex:1' }, disp), ui.button('ScreenTi&p...', async () => { const t = await ui.prompt('ScreenTip text:', screenTip, 'Set Hyperlink ScreenTip'); if (t != null) screenTip = t; }, { class: 'btn small' })), tabs),
      buttons: (cur ? [{ label: '&Remove Link', onClick: () => { A().removeHyperlink(); } }] : []).concat([{ label: 'OK', primary: true, onClick: () => {
        let link;
        if (kind === 0) { let u = addr.value.trim(); if (!u) return false; if (/^www\./i.test(u)) u = 'http://' + u; link = { url: u }; }
        else if (kind === 1) {
          let a = place.get();
          if (a && a.startsWith('§')) { const p = D.byId(d, +a.slice(1)); E.edit('Bookmark', () => { a = F().ensureParaBookmark(p, '_Toc'); return E.sel; }); }
          link = { anchor: a || '_top' };
        } else { const m = mail.value.trim(); if (!m) return false; link = { url: 'mailto:' + m + (subj.value ? '?subject=' + encodeURIComponent(subj.value) : '') }; }
        if (screenTip) link.tip = screenTip;
        const text = disp.value || (kind === 0 ? addr.value : kind === 2 ? mail.value : place.get() || 'Top of the Document');
        if (!d.styles.Hyperlink) E.edit('Style', () => { D.touchKey(d, 'styles'); d.styles.Hyperlink = D.builtinStyles().Hyperlink; D.stylesChanged(); return E.sel; });
        E.edit(cur ? 'Edit Hyperlink' : 'Insert Hyperlink', () => {
          if (range && !D.eqPos(range[0], range[1]) && text === selText) { O.formatRange(range[0], range[1], (r) => Object.assign(r, { link, style: 'Hyperlink' })); return { a: range[0], f: range[1] }; }
          E.sel = { a: range[0], f: range[1] };
          const pos = E.deleteSelection();
          const base = O.inheritRPr(pos);
          delete base.link;
          return O.insertText(pos, text, Object.assign(base, { link, style: 'Hyperlink' }));
        });
        E.pending = { link: undefined, style: undefined };
      } }, { label: 'Cancel' }]),
    });
  };

  /* ================= Insert Object ================= */
  G.insertObject = function () {
    const types = [['chart', 'Chart'], ['equation', 'Equation'], ['wordart', 'WordArt'], ['excel', 'Spreadsheet (table)'], ['document', 'Document (inserted as text)'], ['picture', 'Bitmap Image']];
    const lb = G.listBox(types.map(([v, l]) => ({ value: v, label: l })), 'chart', { height: 160, onDbl: () => ok.click() });
    const dlg = ui.dialog({ title: 'Object', width: 360, body: ui.tabs([{ label: '&Create New', body: h('div', { class: 'col' }, h('label', { text: 'Object type:' }), lb) }, { label: 'Create from &File', body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Insert the contents of a .docx, .txt or .htm file, or a picture.' }), ui.button('&Browse...', () => { dlg.close(null); A().insertFileDialog(); })) }]), buttons: [{ label: 'OK', primary: true, onClick: () => {
      const v = lb.get();
      if (v === 'chart') A().insertChart();
      else if (v === 'equation') G.equation();
      else if (v === 'wordart') G.wordart();
      else if (v === 'excel') L.tables.insert(4, 5, { worksheet: true });
      else if (v === 'document') A().insertFileDialog();
      else A().insertPictureDialog();
    } }, { label: 'Cancel' }] });
    const ok = dlg.buttons[0];
  };
  /** a simple equation editor: linear format with Unicode math symbols */
  G.equation = function () {
    const inp = h('textarea', { rows: 3, id: 'eq-in', style: 'width:100%;font:16px "Cambria Math",Caladea,serif' });
    const pal = h('div', { class: 'sym-recent', style: 'grid-template-columns:repeat(16,26px)' });
    for (const c of '±×÷≠≈≤≥∞√∛∑∏∫∂∆∇πθαβγδλμσφωΩ→⇒∈∉⊂⊃∪∩∀∃¬∧∨²³⁴ⁿ₀₁₂ᵢ½'.split('')) { const b = h('button', { type: 'button', text: c }); b.addEventListener('click', () => { const s = inp.selectionStart; inp.value = inp.value.slice(0, s) + c + inp.value.slice(inp.selectionEnd); inp.focus(); inp.setSelectionRange(s + 1, s + 1); }); pal.appendChild(b); }
    ui.dialog({ title: 'Equation', width: 470, focus: inp, body: h('div', { class: 'col' }, pal, inp, h('div', { class: 'tp-note', text: 'Type the equation; it is inserted in Cambria Math italics.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { if (!inp.value.trim()) return; E.edit('Insert Equation', () => { const pos = E.deleteSelection(); return O.insertText(pos, inp.value.trim(), Object.assign(O.inheritRPr(pos), { font: 'Cambria Math', i: true })); }); } }, { label: 'Cancel' }] });
  };

  /* ================= Diagram Gallery ================= */
  G.diagram = function () {
    const kinds = [['org', 'Organization Chart'], ['cycle', 'Cycle Diagram'], ['radial', 'Radial Diagram'], ['pyramid', 'Pyramid Diagram'], ['venn', 'Venn Diagram'], ['target', 'Target Diagram']];
    const svg = { org: '<svg width="56" height="44"><rect x="20" y="2" width="16" height="10" fill="#fdd" stroke="#333"/><path d="M28 12v6M10 18h36M10 18v4M28 18v4M46 18v4" stroke="#333" fill="none"/><rect x="2" y="22" width="16" height="10" fill="#dfd" stroke="#333"/><rect x="20" y="22" width="16" height="10" fill="#dfd" stroke="#333"/><rect x="38" y="22" width="16" height="10" fill="#dfd" stroke="#333"/></svg>', cycle: '<svg width="56" height="44"><circle cx="28" cy="22" r="16" fill="none" stroke="#888" stroke-dasharray="3 2"/><circle cx="28" cy="6" r="5" fill="#9cf" stroke="#333"/><circle cx="42" cy="30" r="5" fill="#fc9" stroke="#333"/><circle cx="14" cy="30" r="5" fill="#cf9" stroke="#333"/></svg>', radial: '<svg width="56" height="44"><path d="M28 22L28 6M28 22L44 32M28 22L12 32" stroke="#333"/><circle cx="28" cy="22" r="7" fill="#fc9" stroke="#333"/><circle cx="28" cy="6" r="5" fill="#9cf" stroke="#333"/><circle cx="44" cy="32" r="5" fill="#9cf" stroke="#333"/><circle cx="12" cy="32" r="5" fill="#9cf" stroke="#333"/></svg>', pyramid: '<svg width="56" height="44"><path d="M28 2L50 42H6z" fill="#fc9" stroke="#333"/><path d="M19 18h18M12 30h32" stroke="#333"/></svg>', venn: '<svg width="56" height="44"><circle cx="22" cy="18" r="12" fill="#9cf" fill-opacity=".6" stroke="#333"/><circle cx="34" cy="18" r="12" fill="#fc9" fill-opacity=".6" stroke="#333"/><circle cx="28" cy="28" r="12" fill="#cf9" fill-opacity=".6" stroke="#333"/></svg>', target: '<svg width="56" height="44"><circle cx="24" cy="24" r="18" fill="#9cf" stroke="#333"/><circle cx="24" cy="24" r="12" fill="#fc9" stroke="#333"/><circle cx="24" cy="24" r="6" fill="#f99" stroke="#333"/></svg>' };
    let pick = 'org';
    const grid = h('div', { class: 'preset-grid' });
    for (const [k, l] of kinds) { const b = h('button', { type: 'button', class: k === pick ? 'on' : '', 'data-tip': l, 'aria-label': l, html: svg[k] }); b.addEventListener('click', () => { pick = k; L.$$('button', grid).forEach((x) => x.classList.toggle('on', x === b)); name.textContent = l; }); b.addEventListener('dblclick', () => ok.click()); grid.appendChild(b); }
    const name = h('div', { class: 'cd-cap', text: kinds[0][1] });
    const dlg = ui.dialog({ title: 'Diagram Gallery', width: 340, body: h('div', { class: 'col' }, h('div', { text: 'Select a diagram type:' }), grid, name), buttons: [{ label: 'OK', primary: true, onClick: () => insertDiagram(pick) }, { label: 'Cancel' }] });
    const ok = dlg.buttons[0];
  };
  function insertDiagram(kind) {
    const S = L.drawing.newShape;
    const tb = (sh, text) => { sh.tb = { kind: 'tb', blocks: [D.para(text ? [D.text(text, { sz: 10 })] : [], { jc: 'center' })] }; sh.ins = { l: 3.6, t: 3.6, r: 3.6, b: 3.6 }; sh.anchor = 'ctr'; return sh; };
    const kid = (sh, x, y) => Object.assign(sh, { x, y });
    const W = 324, H = 216;
    const kids = [];
    const fills = ['#99CCFF', '#FFCC99', '#CCFF99', '#FF9999', '#CC99FF', '#FFFF99'];
    if (kind === 'org') {
      kids.push(kid(tb(S('rect', 90, 40), 'Manager'), 117, 8));
      for (let i = 0; i < 3; i++) kids.push(kid(tb(S('rect', 90, 40), 'Employee ' + (i + 1)), 9 + i * 108, 120));
      kids.push(kid(S('line', 0, 40), 162, 48), kid(S('line', 216, 0), 54, 88));
      for (let i = 0; i < 3; i++) kids.push(kid(S('line', 0, 32), 54 + i * 108, 88));
    } else if (kind === 'pyramid') {
      for (let i = 0; i < 3; i++) { const sh = S('trapezoid', 120 + i * 80, 56); sh.fill = { t: 'solid', c: fills[i], a: 1 }; tb(sh, 'Level ' + (i + 1)); kids.push(kid(sh, (W - (120 + i * 80)) / 2, 16 + i * 62)); }
      kids[0].geom = 'triangle'; kids[0].w = 120;
    } else {
      const n = kind === 'target' ? 3 : kind === 'venn' ? 3 : 4;
      for (let i = 0; i < n; i++) {
        const a = (i / n) * Math.PI * 2 - Math.PI / 2;
        let sh;
        if (kind === 'target') { const r = 100 - i * 30; sh = S('ellipse', r * 2, r * 2); kids.push(kid(Object.assign(sh, { fill: { t: 'solid', c: fills[i], a: 1 } }), W / 2 - r, H / 2 - r)); continue; }
        if (kind === 'venn') { sh = S('ellipse', 120, 120); sh.fill = { t: 'solid', c: fills[i], a: 0.5 }; kids.push(kid(sh, W / 2 - 60 + Math.cos(a) * 40, H / 2 - 60 + Math.sin(a) * 40)); continue; }
        sh = tb(S('ellipse', 72, 52), kind === 'cycle' ? 'Step ' + (i + 1) : 'Item ' + (i + 1));
        sh.fill = { t: 'solid', c: fills[i], a: 1 };
        kids.push(kid(sh, W / 2 - 36 + Math.cos(a) * 110, H / 2 - 26 + Math.sin(a) * 75));
      }
      if (kind === 'radial') { const c = tb(S('ellipse', 80, 56), 'Center'); c.fill = { t: 'solid', c: '#FFCC99', a: 1 }; kids.unshift(kid(c, W / 2 - 40, H / 2 - 28)); }
    }
    const g = { t: 'group', w: W, h: H, kids, name: 'Diagram ' + D.nid() };
    E.edit('Insert Diagram', () => O.insertItem(E.deleteSelection(), g));
  }

  /* ================= WordArt ================= */
  const WA_STYLES = [
    { fill: { t: 'solid', c: '#000000' }, warp: 'textPlain' }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#FFFFFF' }, { p: 1, c: '#999999' }] }, line: { c: '#000000', w: 0.75 }, warp: 'textPlain' }, { fill: { t: 'solid', c: '#C0C0C0' }, line: { c: '#000000', w: 0.75 }, warp: 'textPlain', i: true }, { fill: { t: 'solid', c: '#000000' }, warp: 'textSlantUp' }, { fill: { t: 'grad', ang: 0, stops: [{ p: 0, c: '#FF9900' }, { p: 1, c: '#FFFF00' }] }, warp: 'textArchUp' },
    { fill: { t: 'solid', c: '#3366FF' }, line: { c: '#000080', w: 0.75 }, warp: 'textPlain', shadow: { dx: 3, dy: 3 } }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#FF0000' }, { p: 0.5, c: '#FFFF00' }, { p: 1, c: '#00CC00' }] }, warp: 'textWave1' }, { fill: { t: 'solid', c: '#FFFFFF' }, line: { c: '#3366CC', w: 1.5 }, warp: 'textPlain', b: true }, { fill: { t: 'solid', c: '#993300' }, warp: 'textChevron' }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#C0C0C0' }, { p: 1, c: '#333333' }] }, warp: 'textCanUp' },
    { fill: { t: 'solid', c: '#CC0000' }, line: { c: '#660000', w: 0.75 }, warp: 'textArchDown' }, { fill: { t: 'solid', c: '#336633' }, warp: 'textTriangle' }, { fill: { t: 'grad', ang: 45, stops: [{ p: 0, c: '#0000FF' }, { p: 1, c: '#00FFFF' }] }, warp: 'textPlain', b: true }, { fill: { t: 'solid', c: '#FFCC00' }, line: { c: '#996600', w: 0.75 }, warp: 'textSlantDown', shadow: { dx: 2, dy: 2 } }, { fill: { t: 'solid', c: '#660099' }, warp: 'textCircle' },
  ];
  G.wordart = function (mode) {
    const it = mode && L.drawing.selectedItem();
    if (mode === true && it && it.wordart) { editText(it, (wa) => L.drawing.modifySelected('Edit WordArt', (x) => { x.wordart = Object.assign({}, x.wordart, wa); })); return; }
    let pick = 0;
    const grid = h('div', { class: 'wa-grid' });
    WA_STYLES.forEach((s, i) => {
      const b = h('button', { type: 'button', class: i === 0 ? 'on' : '', 'aria-label': 'WordArt style ' + (i + 1) });
      const sh = { t: 'shape', w: 66, h: 40, fill: Object.assign({ a: 1 }, s.fill), line: s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }, warp: s.warp, shadow: s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined, wordart: { text: 'WordArt', font: 'Arial Black', b: s.b, i: s.i } };
      b.appendChild(L.drawing.wordartSVG(sh));
      b.addEventListener('click', () => { pick = i; L.$$('button', grid).forEach((x) => x.classList.toggle('on', x === b)); });
      b.addEventListener('dblclick', () => ok.click());
      grid.appendChild(b);
    });
    const dlg = ui.dialog({ title: 'WordArt Gallery', width: 470, body: h('div', { class: 'col' }, h('div', { text: 'Select a WordArt style:' }), grid), buttons: [{ label: 'OK', primary: true, onClick: () => {
      const s = WA_STYLES[pick];
      if (mode === 'gallery' && it) { L.drawing.modifySelected('WordArt Gallery', (x) => { x.fill = Object.assign({ a: 1 }, s.fill); x.line = s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }; x.warp = s.warp; x.shadow = s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined; }); return; }
      const init = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim().slice(0, 200) : 'Your Text Here';
      setTimeout(() => editText({ wordart: { text: init, font: 'Arial Black', sz: 36, b: s.b, i: s.i } }, (wa) => {
        const lines = wa.text.split('\n');
        const longest = Math.max(...lines.map((l) => l.length), 1);
        const w = Math.min(468, longest * wa.sz * 0.62 + 10), hh = lines.length * wa.sz * 1.15 + 6;
        const shape = { t: 'shape', geom: 'rect', w: L.round(w, 2), h: L.round(hh, 2), fill: Object.assign({ a: 1 }, s.fill), line: s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }, warp: s.warp, shadow: s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined, wordart: wa, name: 'WordArt ' + D.nid() };
        E.edit('Insert WordArt', () => O.insertItem(E.deleteSelection(), shape));
      }), 0);
    } }, { label: 'Cancel' }] });
    const ok = dlg.buttons[0];
  };
  function editText(it, done) {
    const wa = Object.assign({ text: 'Your Text Here', font: 'Arial Black', sz: 36 }, it.wordart || {});
    const font = ui.select(G.FONTS().map((f) => [f, f]), wa.font, () => upd(), { id: 'wa-font' });
    const size = ui.select([8, 10, 12, 14, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96].map((x) => [x, String(x)]), wa.sz || 36, null, { id: 'wa-size' });
    const b = ui.check('&Bold', !!wa.b, () => upd()), i = ui.check('&Italic', !!wa.i, () => upd());
    const ta = h('textarea', { rows: 5, id: 'wa-text', style: 'width:100%;font-size:20px' });
    ta.value = wa.text;
    const upd = () => { ta.style.fontFamily = L.fontStack(font.value); ta.style.fontWeight = b.input.checked ? 'bold' : 'normal'; ta.style.fontStyle = i.input.checked ? 'italic' : 'normal'; };
    upd();
    ui.dialog({ title: 'Edit WordArt Text', width: 440, focus: ta, body: h('div', { class: 'col' }, G.row(G.f('&Font:', font), G.f('&Size:', size), b, i), h('label', { for: 'wa-text', html: '<u>T</u>ext:' }), ta), buttons: [{ label: 'OK', primary: true, onClick: () => done({ text: ta.value || ' ', font: font.value, sz: +size.value, b: b.input.checked || undefined, i: i.input.checked || undefined }) }, { label: 'Cancel' }] });
  }

  /* ================= Table dialogs ================= */
  G.insertTable = function () {
    const last = L.store.get('tableDims', { c: 5, r: 2 });
    const cols = G.num(last.c, { min: 1, max: 63, id: 'it-c' }), rows = G.num(last.r, { min: 1, max: 32767, id: 'it-r' });
    const fit = G.radios('it-fit', [['fixed', 'Fi&xed column width:'], ['contents', 'Auto&Fit to contents'], ['window', 'AutoFit to win&dow']], 'fixed');
    const cw = G.len(null, { id: 'it-cw', autoLabel: 'Auto', min: 9 });
    let style = null, look = null;
    const styleNote = h('div', { class: 'tp-note', text: 'Table style: Table Grid' });
    const remember = ui.check('Remember dime&nsions for new tables', false, null);
    const fEls = fit.querySelectorAll('.ck');
    ui.dialog({
      title: 'Insert Table', width: 330, focus: cols.input,
      body: h('div', { class: 'col' }, ui.group('Table size', G.f('Number of &columns:', cols, 'wide'), G.f('Number of &rows:', rows, 'wide')), ui.group('AutoFit behavior', G.row(fEls[0], cw), fEls[1], fEls[2]), G.row(styleNote, ui.button('&AutoFormat...', () => G.tableAutoFormat({ pick: (s, lk) => { style = s; look = lk; styleNote.textContent = 'Table style: ' + s; } }), { class: 'btn small' })), remember),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (remember.input.checked) L.store.set('tableDims', { c: cols.get(), r: rows.get() });
        const o = { autofit: fit.get() };
        const w = cw.get();
        if (fit.get() === 'fixed' && w) o.colWidth = w;
        if (style) { o.style = L.tables.styleId(style); o.look = look; }
        L.tables.insert(rows.get(), cols.get(), o);
        if (style) L.tables.applyStyle(style, look);
      } }, { label: 'Cancel' }],
    });
  };
  G.insertCells = function () {
    const r = G.radios('ic', [['right', 'Shift cells r&ight'], ['down', 'Shift cells &down'], ['row', 'Insert entire &row'], ['col', 'Insert entire &column']], 'down');
    ui.dialog({ title: 'Insert Cells', width: 240, body: r, buttons: [{ label: 'OK', primary: true, onClick: () => L.tables.insertCells(r.get()) }, { label: 'Cancel' }] });
  };
  G.deleteCells = function () {
    const r = G.radios('dc', [['left', 'Shift cells &left'], ['up', 'Shift cells &up'], ['row', 'Delete entire &row'], ['col', 'Delete entire &column']], 'left');
    ui.dialog({ title: 'Delete Cells', width: 240, body: r, buttons: [{ label: 'OK', primary: true, onClick: () => L.tables.deleteCells(r.get()) }, { label: 'Cancel' }] });
  };
  G.splitCells = function () {
    const c = G.num(2, { min: 1, max: 63, id: 'sc-c' }), r = G.num(1, { min: 1, max: 100, id: 'sc-r' });
    const merge = ui.check('&Merge cells before split', !!E.cellSel, null);
    ui.dialog({ title: 'Split Cells', width: 260, focus: c.input, body: h('div', { class: 'col' }, G.f('Number of &columns:', c, 'wide'), G.f('Number of &rows:', r, 'wide'), merge), buttons: [{ label: 'OK', primary: true, onClick: () => L.tables.split(r.get(), c.get(), merge.input.checked) }, { label: 'Cancel' }] });
  };
  G.tableAutoFormat = function (opts) {
    opts = opts || {};
    const names = ['Table Normal', 'Table Grid'].concat(Object.keys(L.tables.STYLES).filter((n) => n !== 'Table Grid'));
    const ci = L.tables.cur();
    const curStyle = ci ? (doc().styles[ci.tbl.tblPr.style] || {}).name || 'Table Grid' : 'Table Grid';
    const look = Object.assign({ firstRow: true, firstCol: true, lastRow: false, lastCol: false }, ci ? ci.tbl.tblPr.look || {} : {});
    const lb = G.listBox(names, names.includes(curStyle) ? curStyle : 'Table Grid', { height: 200, onChange: () => draw() });
    const pv = h('div', { class: 'preview-box', style: 'width:220px;height:170px' });
    const chk = (k, l) => ui.check(l, !!look[k], (v) => { look[k] = v; draw(); });
    const checks = [chk('firstRow', '&Heading rows'), chk('firstCol', 'Fir&st column'), chk('lastRow', '&Last row'), chk('lastCol', 'Last co&lumn')];
    function draw() {
      L.clear(pv);
      const name = lb.get();
      const def = L.tables.STYLES[name];
      const grid = [['', 'Jan', 'Feb', 'Mar', 'Total'], ['East', '7', '7', '5', '19'], ['West', '6', '4', '7', '17'], ['South', '8', '7', '9', '24'], ['Total', '21', '18', '21', '60']];
      const t = h('table', { style: 'border-collapse:collapse;font:10px Arial,Arimo,sans-serif;width:200px' });
      const bd = (b) => (b && b.val !== 'nil' ? `${Math.max(1, b.sz || 0.5)}px ${b.val === 'double' ? 'double' : b.val === 'dotted' ? 'dotted' : 'solid'} ${b.color && b.color !== 'auto' ? '#' + b.color : '#000'}` : 'none');
      grid.forEach((row, ri) => {
        const tr = h('tr');
        row.forEach((v, cix) => {
          const conds = R.cellConds({ tblPr: {} }, ri, cix, grid.length, row.length, Object.assign({ noVBand: !/Columns/.test(name), noHBand: !/List|Contemporary|Subtle|3D/.test(name) }, look));
          let rr = Object.assign({}, def && def.rPr), tc = Object.assign({}, def && def.tcPr), tb = Object.assign({}, def && def.tblPr && def.tblPr.borders);
          if (name === 'Table Grid') tb = { top: { val: 'single' }, bottom: { val: 'single' }, left: { val: 'single' }, right: { val: 'single' }, insideH: { val: 'single' }, insideV: { val: 'single' } };
          let cb = {};
          for (const c of conds) { const cf = def && def.cond && def.cond[c]; if (!cf) continue; if (cf.rPr) rr = Object.assign(rr, cf.rPr); if (cf.tcPr) { tc = Object.assign(tc, cf.tcPr); if (cf.tcPr.borders) cb = Object.assign(cb, cf.tcPr.borders); } }
          const side = (k, outer, inner, isOuter) => cb[k] || (isOuter ? tb[outer] : tb[inner]);
          const css = [`border-top:${bd(side('top', 'top', 'insideH', ri === 0))}`, `border-bottom:${bd(side('bottom', 'bottom', 'insideH', ri === grid.length - 1))}`, `border-left:${bd(side('left', 'left', 'insideV', cix === 0))}`, `border-right:${bd(side('right', 'right', 'insideV', cix === row.length - 1))}`, 'padding:2px 4px'];
          if (tc.shd) css.push('background:' + R.shadeCSS(tc.shd));
          if (rr.b) css.push('font-weight:bold'); if (rr.i) css.push('font-style:italic'); if (rr.color) css.push('color:#' + rr.color); if (rr.caps) css.push('text-transform:uppercase');
          tr.appendChild(h('td', { text: v, style: css.join(';') }));
        });
        t.appendChild(tr);
      });
      pv.appendChild(t);
    }
    draw();
    ui.dialog({
      title: 'Table AutoFormat', width: 480, body: h('div', { class: 'col' }, h('div', { text: 'Category: All table styles' }), G.row(h('div', { class: 'col', style: 'width:200px' }, h('label', { text: 'Table styles:' }), lb), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv)), ui.group('Apply special formats to', G.row(...checks))),
      buttons: [{ label: '&Apply', primary: true, onClick: () => { if (opts.pick) { opts.pick(lb.get(), Object.assign({}, look, { noVBand: true })); return; } L.tables.applyStyle(lb.get(), Object.assign({}, look, { noVBand: !/Columns/.test(lb.get()) }), { clearDirect: true }); } }, { label: 'Cancel' }],
    });
  };
  G.tableProps = function (tab) {
    const ci = L.tables.cur();
    if (!ci) return;
    const t = ci.tbl;
    const d = doc();
    const { tp } = R.tblProps(d, t);
    const rg = L.tables.range();
    const total = t.grid.reduce((a, b) => a + b, 0);
    /* Table tab */
    const tW = ui.check('&Preferred width:', !!(t.tblPr.w && t.tblPr.w.type !== 'auto' && t.tblPr.w.v), null);
    const tWv = G.len(t.tblPr.w && t.tblPr.w.type === 'dxa' ? t.tblPr.w.v : total, { id: 'tp-w', min: 9 });
    const tWu = ui.select([['dxa', 'Inches/cm'], ['pct', 'Percent']], t.tblPr.w && t.tblPr.w.type === 'pct' ? 'pct' : 'dxa', null, { id: 'tp-wu' });
    const align = G.radios('tp-al', [['left', '&Left'], ['center', '&Center'], ['right', 'Rig&ht']], tp.jc === 'center' ? 'center' : tp.jc === 'right' || tp.jc === 'end' ? 'right' : 'left');
    const indent = G.len(tp.ind || 0, { id: 'tp-ind' });
    const wrap = G.radios('tp-wr', [['none', '&None'], ['around', 'Aro&und']], t.tblPr.float ? 'around' : 'none');
    const mar = { t: G.len(tp.cellMar.t || 0, { id: 'tp-mt', min: 0 }), b: G.len(tp.cellMar.b || 0, { id: 'tp-mb', min: 0 }), l: G.len(tp.cellMar.l || 0, { id: 'tp-ml', min: 0 }), r: G.len(tp.cellMar.r || 0, { id: 'tp-mr', min: 0 }) };
    const spacing = G.len(t.tblPr.cellSpacing || 0, { id: 'tp-sp', min: 0 });
    const autoResize = ui.check('Automatically resi&ze to fit contents', t.tblPr.layout !== 'fixed', null);
    const optsBtn = ui.button('&Options...', () => ui.dialog({ title: 'Table Options', width: 320, body: h('div', { class: 'col' }, ui.group('Default cell margins', G.row(G.f('&Top:', mar.t, 'narrow'), G.f('&Left:', mar.l, 'narrow')), G.row(G.f('&Bottom:', mar.b, 'narrow'), G.f('&Right:', mar.r, 'narrow'))), ui.group('Default cell spacing', G.f('Allow spacing between cells', spacing, 'wide')), ui.group('Options', autoResize)), buttons: [{ label: 'OK', primary: true }] }), { class: 'btn small' });
    /* Row tab */
    const row = t.rows[rg.r0];
    const rH = ui.check('&Specify height:', !!row.trPr.h, null);
    const rHv = G.len(row.trPr.h || 14.4, { id: 'tp-rh', min: 1 });
    const rRule = ui.select([['atLeast', 'At least'], ['exact', 'Exactly']], row.trPr.hRule === 'exact' ? 'exact' : 'atLeast', null, { id: 'tp-rr' });
    const rSplit = ui.check('Allow row to brea&k across pages', !row.trPr.cantSplit, null);
    const rHead = ui.check('Repeat as &header row at the top of each page', !!row.trPr.header, null);
    /* Column tab */
    const cW = ui.check('Preferred &width:', true, null);
    const cWv = G.len(t.grid.slice(rg.c0, rg.c1 + 1).reduce((a, b) => a + b, 0) / (rg.c1 - rg.c0 + 1), { id: 'tp-cw', min: 6 });
    /* Cell tab */
    const cell = ci.cell;
    const ceW = ui.check('Preferred &width:', true, null);
    const ceWv = G.len(t.grid.slice(rg.map[ci.ri].find((m) => m.cell === cell).c0, rg.map[ci.ri].find((m) => m.cell === cell).c0 + (cell.tcPr.span || 1)).reduce((a, b) => a + b, 0), { id: 'tp-cew', min: 6 });
    const va = G.radios('tp-va', [['top', 'To&p'], ['center', '&Center'], ['bottom', '&Bottom']], cell.tcPr.vAlign || 'top');
    const cm = cell.tcPr.mar || {};
    const cmSame = ui.check('Same as the whole &table', !cell.tcPr.mar, null);
    const cmar = { t: G.len(cm.t != null ? cm.t : tp.cellMar.t || 0, { id: 'tp-cmt', min: 0 }), b: G.len(cm.b != null ? cm.b : tp.cellMar.b || 0, { id: 'tp-cmb', min: 0 }), l: G.len(cm.l != null ? cm.l : tp.cellMar.l || 0, { id: 'tp-cml', min: 0 }), r: G.len(cm.r != null ? cm.r : tp.cellMar.r || 0, { id: 'tp-cmr', min: 0 }) };
    const noWrap = ui.check('&Wrap text', !cell.tcPr.noWrap, null);
    const fitText = ui.check('&Fit text', !!cell.tcPr.fitText, null);
    const cellOpts = ui.button('&Options...', () => ui.dialog({ title: 'Cell Options', width: 320, body: h('div', { class: 'col' }, ui.group('Cell margins', cmSame, G.row(G.f('Top:', cmar.t, 'narrow'), G.f('Left:', cmar.l, 'narrow')), G.row(G.f('Bottom:', cmar.b, 'narrow'), G.f('Right:', cmar.r, 'narrow'))), ui.group('Options', noWrap, fitText)), buttons: [{ label: 'OK', primary: true }] }), { class: 'btn small' });
    const tabs = ui.tabs([
      { label: '&Table', body: h('div', { class: 'col' }, ui.group('Size', G.row(tW, tWv, G.f('Measure in:', tWu))), ui.group('Alignment', G.row(align, G.f('&Indent from left:', indent, 'wide'))), ui.group('Text wrapping', wrap), h('div', { class: 'tp-btns' }, ui.button('&Borders and Shading...', () => G.borders(), { class: 'btn small' }), optsBtn)) },
      { label: '&Row', body: h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: `Rows ${rg.r0 + 1}${rg.r1 > rg.r0 ? '-' + (rg.r1 + 1) : ''}:` }), ui.group('Size', G.row(rH, rHv, G.f('Row height &is:', rRule))), ui.group('Options', rSplit, rHead)) },
      { label: 'Col&umn', body: h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: `Columns ${rg.c0 + 1}${rg.c1 > rg.c0 ? '-' + (rg.c1 + 1) : ''}:` }), ui.group('Size', G.row(cW, cWv))) },
      { label: 'C&ell', body: h('div', { class: 'col' }, ui.group('Size', G.row(ceW, ceWv)), ui.group('Vertical alignment', va), h('div', { class: 'tp-btns' }, cellOpts)) },
    ], tab || 0);
    ui.dialog({
      title: 'Table Properties', width: 430, body: tabs,
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        E.edit('Table Properties', () => {
          D.touchTbl(t);
          const tpp = t.tblPr;
          if (tW.input.checked) {
            if (tWu.value === 'pct') tpp.w = { type: 'pct', v: 5000 };
            else { const w = tWv.get(); tpp.w = { type: 'dxa', v: w }; const tot = t.grid.reduce((a, b) => a + b, 0) || 1; t.grid = t.grid.map((g) => L.round((g * w) / tot, 2)); }
          } else tpp.w = { type: 'auto', v: 0 };
          tpp.jc = align.get() === 'left' ? undefined : align.get();
          if (!tpp.jc) delete tpp.jc;
          const ind = indent.get();
          if (align.get() === 'left' && ind) tpp.ind = ind; else delete tpp.ind;
          tpp.cellMar = { t: mar.t.get(), b: mar.b.get(), l: mar.l.get(), r: mar.r.get() };
          const sp = spacing.get(); if (sp) tpp.cellSpacing = sp; else delete tpp.cellSpacing;
          tpp.layout = autoResize.input.checked ? 'autofit' : 'fixed';
          for (let r = rg.r0; r <= rg.r1; r++) {
            const tr = t.rows[r].trPr;
            if (rH.input.checked) { tr.h = rHv.get(); tr.hRule = rRule.value; } else { delete tr.h; delete tr.hRule; }
            if (rSplit.input.checked) delete tr.cantSplit; else tr.cantSplit = true;
          }
          if (rHead.input.checked) for (let r = 0; r <= rg.r1; r++) t.rows[r].trPr.header = true; else for (let r = rg.r0; r <= rg.r1; r++) delete t.rows[r].trPr.header;
          if (cW.input.checked) { const w = cWv.get(); for (let i = rg.c0; i <= rg.c1; i++) t.grid[i] = L.round(w, 2); }
          for (const c of rg.cells) {
            c.tcPr.vAlign = va.get() === 'top' ? undefined : va.get();
            if (!c.tcPr.vAlign) delete c.tcPr.vAlign;
            if (cmSame.input.checked) delete c.tcPr.mar; else c.tcPr.mar = { t: cmar.t.get(), b: cmar.b.get(), l: cmar.l.get(), r: cmar.r.get() };
            if (noWrap.input.checked) delete c.tcPr.noWrap; else c.tcPr.noWrap = true;
            if (fitText.input.checked) c.tcPr.fitText = true; else delete c.tcPr.fitText;
          }
          if (ceW.input.checked && rg.cells.length === 1 && !cW.input.checked) {
            const m = rg.map[ci.ri].find((x) => x.cell === cell);
            const want = ceWv.get(), have = t.grid.slice(m.c0, m.c0 + m.span).reduce((a, b) => a + b, 0);
            const k = want / (have || 1);
            for (let i = m.c0; i < m.c0 + m.span; i++) t.grid[i] = L.round(t.grid[i] * k, 2);
          }
          L.tables.fixWidths(t);
          return E.sel;
        });
      } }, { label: 'Cancel' }],
    });
  };
  G.sort = function () {
    const inTbl = L.tables.inTable();
    const ci = inTbl ? L.tables.cur() : null;
    const fields = inTbl ? ci.tbl.grid.map((_, i) => ['' + i, 'Column ' + (i + 1)]) : [['0', 'Paragraphs']].concat([1, 2, 3, 4, 5].map((i) => ['' + i, 'Field ' + i]));
    const mk = (n, def) => ({ by: ui.select(n ? [['', '(none)']].concat(fields) : fields, def, null, { id: 'so-b' + n }), type: ui.select([['text', 'Text'], ['number', 'Number'], ['date', 'Date']], 'text', null, { id: 'so-t' + n }), dir: G.radios('so-d' + n, [['asc', '&Ascending'], ['desc', '&Descending']], 'asc') });
    const k1 = mk(0, fields[0][0]), k2 = mk(1, ''), k3 = mk(2, '');
    const hdr = G.radios('so-h', [['yes', 'Header &row'], ['no', 'No header ro&w']], inTbl && ci.tbl.rows[0].trPr.header ? 'yes' : 'no');
    const sepSel = ui.select([['\t', 'Tabs'], [',', 'Commas'], ['other', 'Other']], '\t', null, { id: 'so-sep' });
    const caseS = ui.check('Case &sensitive', false, null);
    const row = (lbl, k) => ui.group(lbl, G.row(k.by, G.f('Type:', k.type), k.dir));
    ui.dialog({
      title: inTbl ? 'Sort' : 'Sort Text', width: 470,
      body: h('div', { class: 'col' }, row('&Sort by', k1), row('&Then by', k2), row('Then &by', k3), ui.group('My list has', hdr), G.row(G.f('Separate fields at:', sepSel), caseS)),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const keys = [k1, k2, k3].filter((k) => k.by.value !== '').map((k) => ({ col: +k.by.value, field: +k.by.value, type: k.type.value, desc: k.dir.get() === 'desc' }));
        if (inTbl) L.tables.sortTable(keys, { header: hdr.get() === 'yes', caseSensitive: caseS.input.checked });
        else L.tables.sortParagraphs(keys, { header: hdr.get() === 'yes', sep: sepSel.value === 'other' ? ';' : sepSel.value, caseSensitive: caseS.input.checked });
      } }, { label: 'Cancel' }],
    });
  };
  G.formula = function () {
    let guess = '=SUM(ABOVE)';
    try { if (L.tables.inTable()) { const v = L.tables.evalFormula('SUM(ABOVE)', E.sel.f.p); const l = L.tables.evalFormula('SUM(LEFT)', E.sel.f.p); if (!v && l) guess = '=SUM(LEFT)'; } else guess = '='; } catch (e) { /* default */ }
    const inp = h('input', { type: 'text', id: 'fo-f', value: guess, style: 'width:100%' });
    const fmt = h('input', { type: 'text', id: 'fo-fmt', list: 'fo-fmts', style: 'width:100%' });
    const dl = h('datalist', { id: 'fo-fmts' }, ...['#,##0', '#,##0.00', '$#,##0.00;($#,##0.00)', '0', '0%', '0.00', '0.00%'].map((x) => h('option', { value: x })));
    const fn = ui.select([['', 'Paste function']].concat(['ABS', 'AND', 'AVERAGE', 'COUNT', 'DEFINED', 'FALSE', 'IF', 'INT', 'MAX', 'MIN', 'MOD', 'NOT', 'OR', 'PRODUCT', 'ROUND', 'SIGN', 'SUM', 'TRUE'].map((x) => [x, x])), '', (v) => { if (!v) return; inp.value += v + '()'; inp.focus(); inp.setSelectionRange(inp.value.length - 1, inp.value.length - 1); fn.value = ''; }, { id: 'fo-fn' });
    const bms = ui.select([['', 'Paste bookmark']].concat(Array.from(F().bookmarks().keys()).filter((n) => !n.startsWith('_')).map((n) => [n, n])), '', (v) => { if (v) { inp.value += v; bms.value = ''; } }, { id: 'fo-bm' });
    ui.dialog({ title: 'Formula', width: 340, focus: inp, body: h('div', { class: 'col' }, h('label', { for: 'fo-f', html: '<u>F</u>ormula:' }), inp, dl, h('label', { for: 'fo-fmt', html: '<u>N</u>umber format:' }), fmt, G.row(fn, bms)), buttons: [{ label: 'OK', primary: true, onClick: () => {
      let c = inp.value.trim();
      if (!c.startsWith('=')) c = '=' + c;
      if (fmt.value.trim()) c += ` \\# "${fmt.value.trim()}"`;
      A().insertFieldNow(c, 'Formula');
    } }, { label: 'Cancel' }] });
  };
  G.textToTable = function () {
    const paras = E.selectedParas();
    const sepR = G.radios('tt-s', [['para', '&Paragraphs'], [',', '&Commas'], ['\t', '&Tabs'], ['other', 'O&ther:']], paras.some((p) => p.runs.some((it) => it.t === 'tab')) ? '\t' : paras.some((p) => D.plainText(p).includes(',')) ? ',' : 'para', () => upd());
    const other = h('input', { type: 'text', maxlength: 1, value: '-', style: 'width:30px' });
    const cols = G.num(1, { min: 1, max: 63, id: 'tt-c', onChange: () => upd(true) });
    const rowsN = h('span');
    const fit = G.radios('tt-fit', [['fixed', 'Fi&xed column width'], ['contents', 'Auto&Fit to contents'], ['window', 'AutoFit to win&dow']], 'fixed');
    const upd = (fromCols) => {
      const sep = sepR.get() === 'other' ? other.value || '-' : sepR.get();
      if (!fromCols) cols.set(sep === 'para' ? 1 : Math.max(...paras.map((p) => (sep === '\t' ? p.runs.filter((it) => it.t === 'tab').length : (D.plainText(p).split(sep).length - 1)) + 1)));
      rowsN.textContent = String(sep === 'para' ? Math.ceil(paras.length / cols.get()) : paras.length);
    };
    upd();
    const sEls = sepR.querySelectorAll('.ck');
    ui.dialog({ title: 'Convert Text to Table', width: 330, body: h('div', { class: 'col' }, ui.group('Table size', G.f('Number of &columns:', cols, 'wide'), G.f('Number of rows:', rowsN, 'wide')), ui.group('AutoFit behavior', fit), ui.group('Separate text at', G.row(sEls[0], sEls[1]), G.row(sEls[2], sEls[3], other))), buttons: [{ label: 'OK', primary: true, onClick: () => L.tables.textToTable({ sep: sepR.get() === 'other' ? other.value || '-' : sepR.get(), cols: cols.get(), autofit: fit.get() }) }, { label: 'Cancel' }] });
  };
  G.tableToText = function () {
    const r = G.radios('tx-s', [['para', '&Paragraph marks'], ['\t', '&Tabs'], [',', '&Commas'], ['other', '&Other:']], '\t');
    const other = h('input', { type: 'text', maxlength: 1, value: '-', style: 'width:30px' });
    ui.dialog({ title: 'Convert Table to Text', width: 280, body: ui.group('Separate text with', r, other), buttons: [{ label: 'OK', primary: true, onClick: () => L.tables.tableToText(r.get() === 'other' ? other.value || '-' : r.get()) }, { label: 'Cancel' }] });
  };
})();
