/* Quire — task panes (Getting Started, New Document, Help, Search, Clip Art, Clipboard, Research,
 * Styles and Formatting, Reveal Formatting, Mail Merge, Protect Document), the Document Map,
 * page thumbnails, Print Preview and the Reading Layout toolbar.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const P = (L.panes = {});
  const doc = () => D.doc;
  const A = () => L.app;

  /* ================= task pane shell ================= */
  const TP = (P.task = { stack: [], pos: -1, registry: {} });
  TP.mount = function (el) {
    TP.el = el;
    TP.back = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Back', 'data-tip': 'Back', html: L.icons.get('back') });
    TP.fwd = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Forward', 'data-tip': 'Forward', html: L.icons.get('forward') });
    TP.home = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Home', 'data-tip': 'Home', html: L.icons.get('home') });
    TP.title = h('div', { class: 'tp-title', role: 'button', tabindex: '0', 'aria-haspopup': 'menu' }, h('span', { text: '' }), h('span', { class: 'dd-arrow' }));
    const close = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Close task pane', 'data-tip': 'Close', html: L.icons.get('close', 12) });
    TP.body = h('div', { class: 'tp-body' });
    el.append(h('div', { class: 'tp-head' }, TP.back, TP.fwd, TP.home, TP.title, close), TP.body);
    TP.back.addEventListener('click', () => TP.go(TP.pos - 1));
    TP.fwd.addEventListener('click', () => TP.go(TP.pos + 1));
    TP.home.addEventListener('click', () => TP.show('getting-started'));
    close.addEventListener('click', () => A().toggleTask(false));
    const openList = () => { const r = TP.title.getBoundingClientRect(); ui.openMenu(Object.values(TP.registry).filter((p) => !p.hidden).map((p) => ({ label: p.title, run: () => TP.show(p.id), checked: () => TP.current === p.id })), { left: r.left, bottom: r.bottom }); };
    TP.title.addEventListener('click', openList);
    TP.title.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openList(); } });
    L.bus.on('sel', L.debounce(() => TP.refresh('sel'), 150));
    L.bus.on('doc-changed', L.debounce(() => TP.refresh('doc'), 300));
    L.bus.on('doc-loaded', () => TP.refresh('load'));
  };
  TP.register = (p) => { TP.registry[p.id] = p; };
  TP.show = function (id) {
    if (!TP.registry[id]) return;
    if (L.$('#taskpane').hidden) A().toggleTask(true);
    if (TP.stack[TP.pos] !== id) { TP.stack = TP.stack.slice(0, TP.pos + 1); TP.stack.push(id); TP.pos = TP.stack.length - 1; }
    TP.render();
  };
  TP.go = function (pos) { if (pos < 0 || pos >= TP.stack.length) return; TP.pos = pos; TP.render(); };
  TP.render = function () {
    const id = TP.stack[TP.pos];
    TP.current = id;
    const p = TP.registry[id];
    if (!p) return;
    TP.title.firstChild.textContent = p.title;
    TP.back.disabled = TP.pos <= 0; TP.fwd.disabled = TP.pos >= TP.stack.length - 1;
    L.clear(TP.body);
    p.render(TP.body);
  };
  TP.refresh = function (reason) {
    const p = TP.registry[TP.current];
    if (!p || !TP.el || TP.el.hidden) return;
    if (p.refresh) p.refresh(TP.body, reason);
    else if (p.live && reason !== 'sel') { L.clear(TP.body); p.render(TP.body); }
  };
  const sec = (title, ...kids) => h('div', { class: 'tp-sec' }, title ? h('div', { class: 'tp-h' }, title) : null, ...kids);
  const link = (icon, label, fn, o) => h('button', Object.assign({ class: 'tp-link', type: 'button', onclick: fn }, o || {}), icon ? h('span', { html: L.icons.get(icon) }) : null, h('span', { text: label }));
  P.sec = sec; P.link = link;

  /* ---------- Getting Started ---------- */
  /* ---------- Document Recovery: unsaved versions kept by common/suite.js ---------- */
  TP.register({ id: 'recovery', title: 'Document Recovery', render(b) { if (window.VO) VO.recoveryPane(b); } });
  TP.register({
    id: 'getting-started', title: 'Getting Started',
    render(b) {
      const q = h('input', { type: 'text', id: 'gs-search', placeholder: 'Search help', style: 'flex:1;min-width:0' });
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { P.helpQuery = q.value; TP.show('help'); } e.stopPropagation(); });
      b.append(
        h('div', { class: 'tp-row' }, h('span', { html: L.icons.app ? L.icons.app(32) : '' }), h('b', { text: 'Quire 2003 Web Edition', style: 'color:var(--pane-ink)' })),
        sec('Search for:', h('div', { class: 'tp-row' }, q, ui.button('Go', () => { P.helpQuery = q.value; TP.show('help'); }, { class: 'btn small' }))),
        sec('Open', window.VO ? VO.recentLinks(4) : null, link('open', 'More...', () => ui.exec('open')), link('new', 'Create a new document...', () => TP.show('new')), link('sample', 'Open the welcome document', () => A().newDocument({ build: (d) => Object.assign(d, A().buildSample()) }))),
        sec('Learn', link('help', 'Keyboard shortcuts and help', () => { P.helpTopic = 'keys'; TP.show('help'); }), link('stylesPane', 'Styles and Formatting', () => TP.show('styles')), link('mailMerge', 'Mail Merge wizard', () => TP.show('mailmerge'))));
    },
  });
  /* ---------- New Document ---------- */
  TP.register({
    id: 'new', title: 'New Document',
    render(b) {
      const tpl = L.templates ? L.templates.list() : [];
      b.append(
        sec('New', link('new', 'Blank document', () => A().newDocument()), link('webPage', 'Web page', () => A().newDocument({ build: (d) => { d.webPage = true; } }).then(() => A().setView('web'))), link('mail', 'E-mail message', () => A().newDocument({ build: (d) => { d.main.blocks = [D.para([D.text('To: ')]), D.para([D.text('Subject: ')]), D.para()]; } })), link('open', 'From existing document...', () => A().newFromExisting())),
        sec('Templates', link('templates', 'On my computer...', () => L.dlg.newFromTemplate()), ...tpl.slice(0, 6).map((t) => link('new', t.name, () => L.templates.create(t.id)))));
    },
  });
  /* ---------- Help ---------- */
  const HELP = [
    ['Typing and selecting', 'Click to place the insertion point and type. Double-click selects a word, triple-click a paragraph; click in the left margin to select a line, double-click there for a paragraph, Ctrl+click for the whole document. Ctrl+click in text selects a sentence. F8 extends the selection step by step (word, sentence, paragraph, document); Esc stops.'],
    ['Formatting text', 'Use the Formatting toolbar or Format ▸ Font (Ctrl+D). Ctrl+B, Ctrl+I, Ctrl+U toggle bold, italic and underline; Ctrl+Space removes character formatting and Ctrl+Q paragraph formatting. Shift+F3 cycles UPPER, lower and Title case. The Format Painter copies formatting: click it once for one use, double-click to keep it on.'],
    ['Styles', 'Pick a style in the Style box at the left of the Formatting toolbar, or open Format ▸ Styles and Formatting for the full list with Modify, New Style and Select All Instances. Type a new name into the Style box to create a style from the current paragraph.'],
    ['Paragraphs, tabs and indents', 'Format ▸ Paragraph sets alignment, indents, spacing, line spacing and pagination options (keep with next, widow control, page break before). Drag the markers on the ruler to change indents; click the ruler to add a tab of the kind shown in the box at its left end; drag a tab off the ruler to remove it.'],
    ['Bullets and numbering', 'Use the toolbar buttons or Format ▸ Bullets and Numbering for bulleted, numbered and outline-numbered lists. Tab and Shift+Tab at the start of a list item change its level. Typing "1. " or "* " and text, then Enter, starts a list automatically.'],
    ['Tables', 'Table ▸ Insert ▸ Table, or drag in the Insert Table grid on the Standard toolbar. Tab moves to the next cell (and adds a row at the end). Drag cell borders to resize columns. The Tables and Borders toolbar draws, erases, merges and splits cells, sets borders, shading and alignment, sorts and sums.'],
    ['Pictures and drawing', 'Insert ▸ Picture ▸ From File, Clip Art, AutoShapes, WordArt or Chart. Choose a wrapping style from Format ▸ Picture ▸ Layout or the Text Wrapping button. Drag floating objects to move them; drag handles to resize; the green handle rotates shapes.'],
    ['Headers, footers and page numbers', 'View ▸ Header and Footer (or double-click the top or bottom of a page). The Header and Footer toolbar inserts page numbers, the number of pages, date and time, switches between header and footer and links or unlinks sections. Insert ▸ Page Numbers adds a numbered frame for you.'],
    ['Sections, columns and page setup', 'Insert ▸ Break adds page, column and section breaks. Each section can have its own margins, orientation, paper size, columns and headers. File ▸ Page Setup sets margins, paper, layout (different first page, odd and even, vertical alignment, line numbers) for the current section or the whole document.'],
    ['References', 'Insert ▸ Reference ▸ Footnote (Ctrl+Alt+F footnote, Ctrl+Alt+D endnote), Caption, Cross-reference and Index and Tables (table of contents, figures, index). Right-click a field and choose Update Field, or press F9. Alt+F9 shows field codes.'],
    ['Track changes and comments', 'Tools ▸ Track Changes (Ctrl+Shift+E) marks insertions, deletions and formatting changes. The Reviewing toolbar accepts or rejects changes and moves between them. Insert ▸ Comment adds a comment balloon in Print Layout. Tools ▸ Compare and Merge Documents marks the differences with another .docx.'],
    ['Find and replace', 'Edit ▸ Find (Ctrl+F), Replace (Ctrl+H) and Go To (Ctrl+G). More options: match case, whole words, wildcards (? any character, * any text, [abc], <word starts, word ends>), sounds like, all word forms, formatting and special characters (^p paragraph mark, ^t tab, ^m manual page break, ^l line break, ^? any character, ^# any digit, ^$ any letter).'],
    ['Mail merge', 'Tools ▸ Letters and Mailings ▸ Mail Merge opens a six-step wizard: choose letters, envelopes, labels or a directory; open a .csv, .txt or .docx data source or type a new list; insert the address block, greeting line and merge fields; preview each record; then merge to a new document or to a PDF.'],
    ['Files', 'File ▸ Open reads .docx, .docm, .dotx, .dotm, .htm, .txt, .rtf and Markdown (.md, or a .zip holding a .md and its pictures) files. File ▸ Save writes a .docx that opens in Word, LibreOffice, Pages and Google Docs, or keeps a Markdown file Markdown (what you did not edit stays exactly as it was). Save As also offers Word Template (.dotx), Web Page, Rich Text Format, Plain Text, Markdown and PDF. File ▸ Print renders every page to a PDF you can print anywhere.'],
    ['Views', 'Normal shows the text in one continuous column with page breaks as dotted lines. Web Layout wraps text to the window. Print Layout shows the pages exactly as they print, with headers, footers and balloons. Outline shows heading levels to promote, demote and rearrange. Reading Layout shows two pages side by side for reading. Print Preview (Ctrl+F2) shows one or several whole pages.'],
  ];
  const KEYS = [['Ctrl+N / O / S / P', 'New, Open, Save, Print'], ['F12 / Shift+F12', 'Save As / Save'], ['Ctrl+Z / Ctrl+Y', 'Undo / Redo (F4 repeats)'], ['Ctrl+X / C / V', 'Cut, Copy, Paste'], ['Ctrl+A', 'Select all'], ['Ctrl+F / H / G', 'Find, Replace, Go To'], ['Ctrl+B / I / U', 'Bold, Italic, Underline'], ['Ctrl+Shift+D / W', 'Double / word underline'], ['Ctrl+= / Ctrl+Shift++', 'Subscript / Superscript'], ['Ctrl+Shift+A / K', 'All caps / Small caps'], ['Ctrl+Shift+> / <', 'Grow / Shrink font'], ['Ctrl+] / [', 'Font size ±1 point'], ['Ctrl+D', 'Font dialog'], ['Ctrl+Space / Ctrl+Q', 'Remove character / paragraph formatting'], ['Shift+F3', 'Change case'], ['Ctrl+L / E / R / J', 'Left, Center, Right, Justify'], ['Ctrl+1 / 2 / 5', 'Single, double, 1.5 line spacing'], ['Ctrl+0', 'Add/remove 12 pt space before'], ['Ctrl+M / Ctrl+Shift+M', 'Indent / outdent'], ['Ctrl+T / Ctrl+Shift+T', 'Hanging indent / reduce'], ['Ctrl+Shift+N', 'Normal style'], ['Ctrl+Alt+1 / 2 / 3', 'Heading 1, 2, 3'], ['Ctrl+Shift+L', 'List Bullet style'], ['Ctrl+Shift+S', 'Style box'], ['Ctrl+Enter', 'Page break'], ['Ctrl+Shift+Enter', 'Column break'], ['Shift+Enter', 'Line break'], ['Ctrl+Shift+Space', 'Nonbreaking space'], ['Ctrl+- / Ctrl+Shift+-', 'Optional / nonbreaking hyphen'], ['Ctrl+Alt+- ', 'Em dash'], ['Ctrl+Alt+C / R / T', '©, ®, ™'], ['Ctrl+K', 'Hyperlink'], ['Ctrl+Alt+F / D', 'Footnote / Endnote'], ['Ctrl+Alt+M', 'Comment'], ['Ctrl+Shift+E', 'Track changes'], ['F7', 'Spelling and Grammar'], ['Shift+F7', 'Thesaurus'], ['F9 / Alt+F9', 'Update field / toggle field codes'], ['Ctrl+Shift+F9', 'Unlink fields'], ['Alt+Shift+D / T / P', 'Date, time, page number field'], ['F8', 'Extend selection'], ['Shift+F5', 'Go back to previous edit'], ['Alt+Shift+↑ / ↓', 'Move paragraph up / down'], ['Alt+Shift+← / →', 'Promote / demote'], ['Ctrl+Alt+P / N / O', 'Print Layout, Normal, Outline view'], ['Ctrl+*', 'Show/Hide ¶'], ['Ctrl+F1 / Shift+F1', 'Task pane / Reveal Formatting'], ['F1', 'Help'], ['Tab / Shift+Tab', 'Next / previous cell; list level']];
  TP.register({
    id: 'help', title: 'Quire Help',
    render(b) {
      const q = h('input', { type: 'text', id: 'hp-q', value: P.helpQuery || '', placeholder: 'Search', style: 'flex:1;min-width:0' });
      const res = h('div');
      const draw = () => {
        L.clear(res);
        const words = q.value.toLowerCase().split(/\s+/).filter(Boolean);
        const hit = (t) => words.every((w) => t.toLowerCase().includes(w));
        if (P.helpTopic !== 'keys' || words.length) for (const [t, d] of HELP.filter(([t, d]) => !words.length || hit(t + ' ' + d))) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: t }), h('p', { text: d })));
        const keys = KEYS.filter(([k, d]) => !words.length || hit(k + ' ' + d));
        if (keys.length) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: 'Keyboard shortcuts' }), h('table', { class: 'kbd-table' }, ...keys.map(([k, d]) => h('tr', null, h('td', { text: k }), h('td', { text: d }))))));
        if (!res.childNodes.length) res.appendChild(h('div', { class: 'tp-note', text: 'No help topics match your search.' }));
        if (P.helpTopic === 'keys' && !words.length) res.appendChild(link('help', 'Show all help topics', () => { P.helpTopic = null; draw(); }));
      };
      q.addEventListener('input', draw);
      q.addEventListener('keydown', (e) => e.stopPropagation());
      b.append(h('div', { class: 'tp-row' }, q), res);
      draw();
      P.helpQuery = '';
    },
  });
  /* ---------- Search (in this document) ---------- */
  TP.register({
    id: 'search', title: 'Basic File Search',
    render(b) {
      const q = h('input', { type: 'text', id: 'srch-q', value: P.searchQuery || '', style: 'flex:1;min-width:0' });
      const res = h('div');
      const go = () => {
        P.searchQuery = q.value;
        L.clear(res);
        const needle = q.value.trim().toLowerCase();
        if (!needle) return;
        let n = 0;
        D.reindex(doc());
        for (const st of D.stories(doc())) for (const p of st.paras || []) {
          const t = D.ptext(p);
          let i = t.toLowerCase().indexOf(needle);
          while (i >= 0 && n < 200) {
            n++;
            const a = Math.max(0, i - 30), e = Math.min(t.length, i + needle.length + 40);
            const hit = h('div', { class: 'search-hit' });
            hit.append(document.createTextNode((a > 0 ? '…' : '') + t.slice(a, i)), h('b', { text: t.slice(i, i + needle.length) }), document.createTextNode(t.slice(i + needle.length, e) + (e < t.length ? '…' : '')));
            const off = i;
            hit.addEventListener('click', () => A().gotoPos(D.pos(p, off), D.pos(p, off + needle.length)));
            res.appendChild(hit);
            i = t.toLowerCase().indexOf(needle, i + needle.length);
          }
        }
        res.prepend(h('div', { class: 'tp-note', text: n ? `${n} match${n === 1 ? '' : 'es'} in ${A().fileName}` : 'No matches were found in this document.' }));
      };
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); e.stopPropagation(); });
      b.append(sec('Search for', h('div', { class: 'tp-row' }, h('label', { for: 'srch-q', text: 'Search text:' })), h('div', { class: 'tp-row' }, q, ui.button('Go', go, { class: 'btn small' })), h('div', { class: 'tp-note', text: 'Searches the text of the open document, including headers, footers, notes and comments.' })), res);
      if (q.value) go();
    },
  });
  /* ---------- Clip Art ---------- */
  TP.register({
    id: 'clipart', title: 'Clip Art',
    render(b) {
      const q = h('input', { type: 'text', id: 'ca-q', value: P.clipQuery || '', style: 'flex:1;min-width:0' });
      const grid = h('div', { class: 'clip-grid' });
      const draw = () => {
        L.clear(grid);
        const items = L.clipart ? L.clipart.search(q.value) : [];
        for (const it of items) {
          const el = h('button', { class: 'clip-it', type: 'button', 'data-tip': it.name, 'aria-label': it.name }, h('img', { src: L.clipart.url(it), alt: '' }), h('span', { text: it.name }));
          el.addEventListener('click', () => P.insertClip(it));
          grid.appendChild(el);
        }
        if (!items.length) grid.appendChild(h('div', { class: 'tp-note', text: 'No clips match. Try "people", "chart", "nature" or "office".' }));
      };
      q.addEventListener('input', () => { P.clipQuery = q.value; draw(); });
      q.addEventListener('keydown', (e) => e.stopPropagation());
      b.append(sec('Search for:', h('div', { class: 'tp-row' }, q, ui.button('Go', draw, { class: 'btn small' }))), grid, h('div', { class: 'tp-note', text: 'Click a clip to insert it at the insertion point.' }));
      draw();
    },
  });
  P.insertClip = async function (it) {
    const blob = await L.clipart.toPNG(it, 512);
    const id = L.media.add(blob, it.name.replace(/\W+/g, '_') + '.png');
    E.edit('Insert Clip Art', () => O.insertItem(E.deleteSelection(), D.item('img', { media: id, w: 96, h: 96, name: it.name, alt: it.name })));
  };
  /* ---------- Office Clipboard ---------- */
  TP.register({
    id: 'clipboard', title: 'Clipboard', live: true,
    render(b) {
      const items = A().clipItems || [];
      b.append(h('div', { class: 'tp-btns' }, ui.button('Paste All', () => { for (const it of items.slice().reverse()) E.pasteBlocks(it.blocks); }, { class: 'btn small' }), ui.button('Clear All', () => { items.length = 0; TP.refresh('doc'); }, { class: 'btn small' })), h('div', { class: 'tp-note', text: `${items.length} of 24 - Clipboard. Click an item to paste.` }));
      for (const it of items) {
        const el = h('button', { class: 'cb-item', type: 'button' }, h('span', { html: L.icons.get(it.blocks.some((x) => x.t === 'tbl') ? 'table' : 'font') }), h('span', { class: 'cb-t', text: (it.text || '').slice(0, 160) || '[object]' }));
        el.addEventListener('click', () => E.pasteBlocks(it.blocks));
        b.appendChild(el);
      }
    },
  });
  /* ---------- Research: offline thesaurus (WordNet-based) and searching this document ---------- */
  const TH = (L.thesaurus = { data: null, keys: null, loading: null });
  TH.load = function () {
    if (TH.data) return Promise.resolve(TH.data);
    if (TH.loading) return TH.loading;
    TH.loading = fetch('../common/dict/en.thes').then((r) => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); }).then((txt) => {
      const map = new Map();
      for (const ln of txt.split('\n')) { if (!ln) continue; const k = ln.slice(0, ln.indexOf('|')); map.set(k.toLowerCase(), ln); }
      TH.data = map;
      TH.keys = Array.from(map.keys()).sort();
      return map;
    }).catch((e) => { TH.loading = null; throw e; });
    return TH.loading;
  };
  const POS = { noun: 'n.', verb: 'v.', adj: 'adj.', adv: 'adv.' };
  /** entry for a word or its base form: {head, senses:[{pos, syns}], ants} */
  TH.lookup = function (word) {
    const m = TH.data;
    if (!m) return null;
    const w = String(word).trim().toLowerCase().replace(/’/g, "'");
    const tries = [w];
    const add = (x) => { if (x && x.length > 1 && !tries.includes(x)) tries.push(x); };
    add(w.replace(/'s$/, ''));
    add(w.replace(/ies$/, 'y')); add(w.replace(/es$/, '')); add(w.replace(/s$/, ''));
    add(w.replace(/ied$/, 'y')); add(w.replace(/ed$/, '')); add(w.replace(/d$/, '')); add(w.replace(/([b-df-hj-np-tv-z])\1ed$/, '$1'));
    add(w.replace(/ing$/, '')); add(w.replace(/ing$/, 'e')); add(w.replace(/([b-df-hj-np-tv-z])\1ing$/, '$1'));
    add(w.replace(/ily$/, 'y')); add(w.replace(/ly$/, ''));
    add(w.replace(/ier$/, 'y')); add(w.replace(/er$/, '')); add(w.replace(/iest$/, 'y')); add(w.replace(/est$/, ''));
    for (const t of tries) {
      const ln = m.get(t);
      if (!ln) continue;
      const parts = ln.split('|');
      return {
        head: parts[0], base: t !== w,
        senses: parts[1].split(';').map((x) => { const k = x.indexOf(':'); return { pos: x.slice(0, k), syns: x.slice(k + 1).split(',').filter(Boolean) }; }),
        ants: parts[2] ? parts[2].split(',').filter(Boolean) : [],
      };
    }
    return null;
  };
  /** nearby headwords for the alphabetical list Word shows when nothing matches */
  TH.near = function (word, n) {
    const keys = TH.keys || [];
    const w = String(word).toLowerCase();
    let lo = 0, hi = keys.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (keys[mid] < w) lo = mid + 1; else hi = mid; }
    return keys.slice(Math.max(0, lo - Math.floor(n / 2)), Math.max(0, lo - Math.floor(n / 2)) + n);
  };
  const SERVICES = [['thesaurus', 'Thesaurus: English (U.S.)'], ['thesaurusUK', 'Thesaurus: English (U.K.)'], ['document', 'Search This Document'], ['all', 'All Reference Books']];
  TP.register({
    id: 'research', title: 'Research',
    render(b) {
      const q = h('input', { type: 'text', id: 'rs-q', value: P.researchQuery || '', style: 'flex:1;min-width:0', 'aria-label': 'Search for' });
      const svc = ui.select(SERVICES, SERVICES.some((x) => x[0] === P.researchService) ? P.researchService : 'thesaurus', (v) => { P.researchService = v; }, { id: 'rs-svc', style: 'flex:1;min-width:0' });
      const hist = P.researchHistory || (P.researchHistory = []);
      const out = h('div');
      const go = async (text, fromHistory) => {
        text = (text == null ? q.value : text).trim();
        if (!text) return;
        q.value = text;
        P.researchQuery = text;
        if (!fromHistory && hist[hist.length - 1] !== text) hist.push(text);
        L.clear(out);
        const service = svc.value;
        if (service === 'thesaurus' || service === 'thesaurusUK' || service === 'all') {
          const status = h('div', { class: 'tp-note', text: 'Searching...' });
          out.appendChild(status);
          try { await TH.load(); }
          catch (e) { status.textContent = 'The thesaurus could not be loaded (' + (e.message || e) + '). Serve Quire from a web server rather than opening index.html as a local file.'; return; }
          status.remove();
          renderThesaurus(out, text, (w) => go(w));
        }
        if (service === 'document' || service === 'all') renderDocMatches(out, text);
      };
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); e.stopPropagation(); });
      const back = ui.button('◄ Back', () => { if (hist.length > 1) { hist.pop(); go(hist[hist.length - 1], true); } }, { class: 'btn small' });
      b.append(sec('Search for:', h('div', { class: 'tp-row' }, q, ui.button('Go', () => go(), { class: 'btn small' })), h('div', { class: 'tp-row' }, svc), h('div', { class: 'tp-row' }, back)), out);
      if (q.value && P.researchAuto !== false) { P.researchAuto = false; setTimeout(() => go(), 0); }
    },
  });
  function wordMenuButton(w, onLookUp) {
    const btn = h('button', { class: 'tp-link th-word', type: 'button' }, h('span', { text: w }), h('span', { class: 'dd-arrow' }));
    btn.addEventListener('click', (e) => {
      ui.openMenu([
        { label: '&Insert', run: () => { if (E.sel && E.collapsed()) { const wd = O.wordAt(E.sel.f); if (wd) E.setSel(wd[0], wd[1]); } E.insertTextAt(w, 'Insert'); } },
        { label: '&Copy', run: () => { try { navigator.clipboard.writeText(w).catch(() => {}); } catch (err) { /* ignore */ } E.clip = { text: w, blocks: [D.para([D.text(w)])], html: w }; } },
        { label: '&Look Up', run: () => onLookUp(w) },
      ], btn.getBoundingClientRect());
      e.stopPropagation();
    });
    btn.addEventListener('dblclick', () => onLookUp(w));
    return btn;
  }
  function renderThesaurus(out, text, onLookUp) {
    const r = TH.lookup(text);
    const box = h('div', { class: 'help-topic' }, h('h4', { text: 'Thesaurus: English' }));
    out.appendChild(box);
    if (!r) {
      box.appendChild(h('p', { class: 'tp-note', text: `No results were found for "${text}". Alphabetical list:` }));
      for (const k of TH.near(text, 12)) box.appendChild(wordMenuButton(k, onLookUp));
      return;
    }
    if (r.base) box.appendChild(h('p', { class: 'tp-note', text: `Showing results for "${r.head}".` }));
    for (const s of r.senses) {
      const head = h('div', { class: 'th-sense' }, h('b', { text: `${s.syns[0]} (${POS[s.pos] || s.pos})` }));
      box.appendChild(head);
      for (const w of s.syns) box.appendChild(wordMenuButton(w, onLookUp));
    }
    if (r.ants.length) {
      box.appendChild(h('div', { class: 'th-sense' }, h('b', { text: `${r.ants[0]} (Antonym)` })));
      for (const w of r.ants) box.appendChild(wordMenuButton(w, onLookUp));
    }
  }
  function renderDocMatches(out, text) {
    const d = doc();
    D.reindex(d);
    const box = h('div', { class: 'help-topic' }, h('h4', { text: 'In this document' }));
    out.appendChild(box);
    const needle = text.toLowerCase();
    let n = 0;
    for (const st of D.stories(d)) for (const p of st.paras || []) {
      const t = D.ptext(p);
      let i = t.toLowerCase().indexOf(needle);
      while (i >= 0 && n < 50) {
        const a = i, bEnd = i + needle.length;
        const snip = (a > 30 ? '…' : '') + t.slice(Math.max(0, a - 30), a);
        const hit = h('button', { class: 'search-hit', type: 'button' }, document.createTextNode(snip.replace(/[￼\t]/g, ' ')), h('b', { text: t.slice(a, bEnd) }), document.createTextNode(t.slice(bEnd, bEnd + 40).replace(/[￼\t]/g, ' ') + (t.length > bEnd + 40 ? '…' : '')));
        hit.addEventListener('click', () => A().gotoPos(D.pos(p, a), D.pos(p, bEnd)));
        box.appendChild(hit);
        n++;
        i = t.toLowerCase().indexOf(needle, bEnd);
      }
    }
    if (!n) box.appendChild(h('p', { class: 'tp-note', text: 'The text was not found in this document.' }));
    else if (n >= 50) box.appendChild(h('p', { class: 'tp-note', text: 'Showing the first 50 matches.' }));
  }
  /* ---------- Styles and Formatting ---------- */
  TP.register({
    id: 'styles', title: 'Styles and Formatting',
    render(b) {
      const d = doc();
      const cur = h('div', { class: 'tp-box', style: 'min-height:22px;font-size:12px' });
      const list = h('div', { class: 'sf-list' });
      const show = ui.select([['avail', 'Available formatting'], ['inuse', 'Formatting in use'], ['styles', 'Available styles'], ['all', 'All styles']], P.styleShow || 'avail', (v) => { P.styleShow = v; draw(); }, { id: 'sf-show', style: 'flex:1' });
      const draw = () => {
        L.clear(list);
        const used = D.usedStyles(d);
        const mode = P.styleShow || 'avail';
        let ids = Object.keys(d.styles).filter((k) => ['paragraph', 'character', 'table', 'numbering'].includes(d.styles[k].type) && !d.styles[k].hidden);
        if (mode === 'inuse') ids = ids.filter((k) => used.has(k));
        else if (mode === 'avail' || mode === 'styles') ids = ids.filter((k) => used.has(k) || d.styles[k].q || d.styles[k].custom || /^Heading[1-3]$|^Normal$/.test(k));
        if (mode === 'all') { const b0 = D.builtinStyles(); for (const k in b0) if (!ids.includes(k)) ids.push(k); }
        ids.sort((a, b2) => D.styleDisplayName(d.styles[a] || D.builtinStyles()[a]).localeCompare(D.styleDisplayName(d.styles[b2] || D.builtinStyles()[b2])));
        list.appendChild(h('div', { class: 'sf-item', role: 'button', tabindex: '0', onclick: () => ui.exec('clearFormats') }, h('span', { class: 'sf-n', text: 'Clear Formatting' })));
        const curId = E.sel ? E.sel.f.p.pPr.style || 'Normal' : '';
        for (const id of ids) {
          const s = d.styles[id] || D.builtinStyles()[id];
          if (!s) continue;
          const r = s.type === 'paragraph' ? (d.styles[id] ? D.styleProps(d, id).rPr : s.rPr || {}) : s.type === 'character' ? D.mergeRPr(d.defaults.rPr, d.styles[id] ? D.charStyleProps(d, id) : s.rPr) : d.defaults.rPr;
          const css = `font-family:${L.fontStack(r.font)};font-size:${Math.min(16, Math.max(8, (r.sz || 12)))}px;${r.b ? 'font-weight:bold;' : ''}${r.i ? 'font-style:italic;' : ''}${r.color && r.color !== 'auto' ? 'color:#' + r.color + ';' : ''}${r.u && r.u !== 'none' ? 'text-decoration:underline;' : ''}`;
          const it = h('div', { class: 'sf-item' + (id === curId ? ' on' : ''), role: 'button', tabindex: '0', 'data-tip': styleDescription(d, id) }, h('span', { class: 'sf-n', text: D.styleDisplayName(s), style: css }), h('span', { class: 'sf-k', text: s.type === 'paragraph' ? '¶' : s.type === 'character' ? 'a' : s.type === 'table' ? '⊞' : '≡' }));
          it.addEventListener('click', () => { if (s.type === 'table') { if (L.tables.inTable()) L.tables.applyStyle(s.name || id); return; } if (s.type === 'numbering') return; A().applyStyle(id); E.refocus(); });
          it.addEventListener('contextmenu', (e) => { e.preventDefault(); styleMenu(id, { x: e.clientX, y: e.clientY }); });
          list.appendChild(it);
        }
      };
      P.refreshCur = () => { cur.textContent = A().styleState ? A().styleState() + describeDirect() : ''; };
      b.append(
        sec('Formatting of selected text', cur, h('div', { class: 'tp-btns' }, ui.button('Select All', () => selectAllLike(), { class: 'btn small' }), ui.button('New Style...', () => L.dlg.style(null), { class: 'btn small' }))),
        sec('Pick formatting to apply', list),
        h('div', { class: 'tp-row' }, h('label', { for: 'sf-show', text: 'Show:' }), show));
      draw();
      P.refreshCur();
      this._draw = draw;
    },
    refresh(b, reason) { if (P.refreshCur) P.refreshCur(); if (reason !== 'sel' && this._draw) this._draw(); else if (reason === 'sel') { const curId = E.sel ? E.sel.f.p.pPr.style || 'Normal' : ''; L.$$('.sf-item', b).forEach((x) => x.classList.toggle('on', x.firstChild && x.firstChild.textContent === D.styleDisplayName(doc().styles[curId]))); } },
  });
  function describeDirect() {
    if (!E.sel) return '';
    const it = D.itemAfter(E.sel.f.p, E.sel.f.o) || D.itemBefore(E.sel.f.p, E.sel.f.o);
    const r = it && it.rPr ? it.rPr : {};
    const parts = [];
    if (r.font) parts.push(r.font);
    if (r.sz) parts.push(r.sz + ' pt');
    if (r.b) parts.push('Bold');
    if (r.i) parts.push('Italic');
    if (r.u && r.u !== 'none') parts.push('Underline');
    if (r.color && r.color !== 'auto') parts.push('Font color: #' + r.color);
    const pp = E.sel.f.p.pPr;
    if (pp.jc) parts.push({ center: 'Centered', right: 'Right', both: 'Justified', left: 'Left' }[pp.jc] || pp.jc);
    if (pp.ind && (pp.ind.l || pp.ind.fl)) parts.push('Indent');
    return parts.length ? ' + ' + parts.join(', ') : '';
  }
  function styleDescription(d, id) {
    const s = d.styles[id] || D.builtinStyles()[id];
    if (!s) return id;
    const r = s.rPr || {};
    const p = s.pPr || {};
    const parts = [];
    if (s.basedOn) parts.push((d.styles[s.basedOn] ? D.styleDisplayName(d.styles[s.basedOn]) : s.basedOn) + ' +');
    if (r.font) parts.push(r.font);
    if (r.sz) parts.push(r.sz + ' pt');
    if (r.b) parts.push('Bold');
    if (r.i) parts.push('Italic');
    if (p.jc) parts.push(p.jc === 'both' ? 'Justified' : p.jc);
    if (p.sp && p.sp.b) parts.push('Before: ' + p.sp.b + ' pt');
    if (p.sp && p.sp.a) parts.push('After: ' + p.sp.a + ' pt');
    if (p.keepNext) parts.push('Keep with next');
    if (p.outline != null && p.outline < 9) parts.push('Level ' + (p.outline + 1));
    return parts.join(', ') || D.styleDisplayName(s);
  }
  function styleMenu(id, at) {
    const d = doc();
    const s = d.styles[id];
    ui.openMenu([
      { label: '&Select All ' + countUses(id) + ' Instance(s)', enabled: () => countUses(id) > 0, run: () => selectAllOfStyle(id) },
      { label: '&Clear Formatting of ' + countUses(id) + ' Instance(s)', enabled: () => countUses(id) > 0, run: () => { selectAllOfStyle(id); ui.exec('clearFormats'); } },
      '-',
      { label: '&Update to Match Selection', enabled: () => !!s && !!E.sel, run: () => updateStyleFromSel(id) },
      { label: '&Modify...', run: () => L.dlg.style(id) },
      { label: '&Delete...', enabled: () => !!s && !s.builtin && id !== 'Normal', run: async () => { const r = await ui.msg(`Do you want to delete style ${D.styleDisplayName(s)}?`, { icon: 'question', buttons: ['&Yes', '&No'] }); if (r === 0) deleteStyle(id); } },
    ], at);
  }
  function countUses(id) {
    let n = 0;
    D.walk(doc().main, (b) => { if (b.t === 'p') { if ((b.pPr.style || 'Normal') === id) n++; for (const it of b.runs) if (it.rPr && it.rPr.style === id) { n++; break; } } });
    return n;
  }
  function selectAllOfStyle(id) {
    const d = doc();
    D.reindex(d);
    const ps = d.main.paras.filter((p) => (p.pPr.style || 'Normal') === id || p.runs.some((it) => it.rPr && it.rPr.style === id));
    if (!ps.length) return;
    E.setSel(D.pos(ps[0], 0), D.pos(ps[ps.length - 1], D.plen(ps[ps.length - 1])));
    A().status(`${ps.length} paragraph(s) use ${D.styleDisplayName(d.styles[id])}; the selection spans all of them.`);
  }
  function selectAllLike() { if (!E.sel) return; selectAllOfStyle(E.sel.f.p.pPr.style || 'Normal'); }
  function updateStyleFromSel(id) {
    const d = doc();
    const p = E.sel.f.p;
    const it = D.itemAfter(p, E.sel.f.o) || D.itemBefore(p, E.sel.f.o);
    E.edit('Update Style', () => {
      D.touchKey(d, 'styles');
      const s = d.styles[id];
      if (s.type === 'paragraph') {
        const pp = L.clone(p.pPr); delete pp.style; delete pp.num; delete pp.sect; delete pp.chg;
        s.pPr = D.mergePPr(s.pPr || {}, pp);
        const rr = it && it.rPr ? L.clone(it.rPr) : {}; delete rr.style; delete rr.link; delete rr.ins; delete rr.del; delete rr.chg;
        s.rPr = D.mergeRPr(s.rPr || {}, rr);
        /* direct formatting now lives in the style */
        D.walk(d.main, (b) => { if (b.t === 'p' && (b.pPr.style || 'Normal') === id) { D.touch(b); for (const k in pp) delete b.pPr[k]; for (const x of b.runs) if (x.rPr) for (const k in rr) if (JSON.stringify(x.rPr[k]) === JSON.stringify(rr[k])) delete x.rPr[k]; D.normalize(b); } });
      } else if (s.type === 'character' && it && it.rPr) { const rr = L.clone(it.rPr); delete rr.style; s.rPr = D.mergeRPr(s.rPr || {}, rr); }
      D.stylesChanged();
      return E.sel;
    });
  }
  function deleteStyle(id) {
    const d = doc();
    E.edit('Delete Style', () => {
      D.touchKey(d, 'styles');
      const s = d.styles[id];
      delete d.styles[id];
      for (const k in d.styles) if (d.styles[k].basedOn === id) d.styles[k].basedOn = s.basedOn || (s.type === 'paragraph' ? 'Normal' : null);
      for (const st of D.stories(d)) D.walk(st, (b) => { if (b.t === 'p') { if (b.pPr.style === id) { D.touch(b); delete b.pPr.style; } for (const it of b.runs) if (it.rPr && it.rPr.style === id) { D.touch(b); delete it.rPr.style; } } if (b.t === 'tbl' && b.tblPr.style === id) { D.touchTbl(b); delete b.tblPr.style; } });
      D.stylesChanged();
      return E.sel;
    });
  }
  P.styleMenu = styleMenu;
  /* ---------- Reveal Formatting ---------- */
  TP.register({
    id: 'reveal', title: 'Reveal Formatting',
    render(b) {
      const box = h('div', { class: 'rf-tree' });
      b.append(sec('Selected text', h('div', { class: 'tp-box', text: E.sel ? (O.textRange(...E.range(), ' ').slice(0, 120) || 'Sample Text') : '' })), sec('Formatting of selected text', box), ui.check('Distinguish style source', !!P.distinguish, (v) => { P.distinguish = v; TP.render(); }), ui.check('Show all formatting marks', !!LY.opts.showMarks, () => ui.exec('showMarks')));
      if (!E.sel) return;
      const d = doc();
      const p = E.sel.f.p;
      const r = E.curRun();
      const pp = E.curPara();
      const sect = D.sectFor(d, p);
      const row = (k, v, dlgFn) => { const el = h('div', null, h('span', { text: k + ': ' }), h('span', { text: v })); if (dlgFn) { el.style.cursor = 'pointer'; el.addEventListener('click', dlgFn); } return el; };
      const fl = (v) => D.fmtLen(v || 0);
      const head = (t, fn) => { const el = h('h5', null, h('a', { href: '#', text: t, onclick: (e) => { e.preventDefault(); fn(); } })); return el; };
      box.append(head('Font', () => L.dlg.font()), row('FONT', `${r.font || ''} ${r.sz || 12} pt${r.b ? ', Bold' : ''}${r.i ? ', Italic' : ''}`), r.color && r.color !== 'auto' ? row('Font color', '#' + r.color) : null, r.u && r.u !== 'none' ? row('Underline', r.u) : null, r.hl ? row('Highlight', r.hl) : null, r.caps ? row('Effects', 'All caps') : null, r.smallCaps ? row('Effects', 'Small caps') : null, r.vert ? row('Position', r.vert) : null, r.lang ? row('Language', A().langName(r.lang)) : row('Language', 'English (U.S.)'),
        head('Paragraph', () => L.dlg.paragraph()), row('Paragraph Style', D.styleDisplayName(d.styles[p.pPr.style || 'Normal'])), row('Alignment', { center: 'Centered', right: 'Right', both: 'Justified', left: 'Left', distribute: 'Distributed' }[pp.jc || 'left'] || pp.jc), row('Indentation', `Left: ${fl(pp.ind && pp.ind.l)}, Right: ${fl(pp.ind && pp.ind.r)}${pp.ind && pp.ind.fl ? (pp.ind.fl < 0 ? ', Hanging: ' + fl(-pp.ind.fl) : ', First line: ' + fl(pp.ind.fl)) : ''}`),
        row('Spacing', `Before: ${L.round((pp.sp && pp.sp.b) || 0, 1)} pt, After: ${L.round((pp.sp && pp.sp.a) || 0, 1)} pt, Line spacing: ${pp.sp && pp.sp.line ? (pp.sp.rule === 'auto' || !pp.sp.rule ? (pp.sp.line === 1 ? 'single' : pp.sp.line === 2 ? 'double' : pp.sp.line + ' lines') : (pp.sp.rule === 'exact' ? 'Exactly ' : 'At least ') + pp.sp.line + ' pt') : 'single'}`),
        pp.keepNext || pp.keepLines || pp.pageBreakBefore ? row('Line and Page Breaks', [pp.keepNext && 'Keep with next', pp.keepLines && 'Keep lines together', pp.pageBreakBefore && 'Page break before'].filter(Boolean).join(', ')) : null,
        L.lists && L.lists.info(p) ? row('Bullets and Numbering', L.lists.info(p).fmt === 'bullet' ? 'Bulleted' : 'Numbered') : null,
        head('Section', () => L.dlg.pageSetup()), row('Margins', `Top: ${fl(sect.mt)}, Bottom: ${fl(sect.mb)}, Left: ${fl(sect.ml)}, Right: ${fl(sect.mr)}`), row('Layout', `Section start: ${({ nextPage: 'New page', continuous: 'Continuous', evenPage: 'Even page', oddPage: 'Odd page', nextColumn: 'New column' })[sect.type || 'nextPage']}`), row('Paper', `${D.paperName(sect)} (${fl(sect.pgW)} × ${fl(sect.pgH)}), ${sect.orient === 'landscape' ? 'Landscape' : 'Portrait'}`), sect.cols && sect.cols.n > 1 ? row('Columns', String(sect.cols.n)) : null);
      if (P.distinguish) box.appendChild(h('div', { class: 'tp-note', text: 'Paragraph style: ' + styleDescription(d, p.pPr.style || 'Normal') }));
    },
    refresh(b, reason) { if (reason === 'sel' || reason === 'doc') { L.clear(b); this.render(b); } },
  });
  /* ---------- Protect Document ---------- */
  TP.register({
    id: 'protect', title: 'Protect Document',
    render(b) {
      const d = doc();
      const pr = d.settings.protect || {};
      const fmt = ui.check('Limit formatting to a selection of styles', !!pr.formatting, (v) => { pr.formatting = v; });
      const ed = ui.check('Allow only this type of editing in the document:', pr.kind && pr.kind !== 'none', (v) => { kind.disabled = !v; });
      const kind = ui.select([['readOnly', 'No changes (Read only)'], ['tracked', 'Tracked changes'], ['comments', 'Comments'], ['forms', 'Filling in forms']], pr.kind || 'readOnly', null, { id: 'pr-kind', style: 'width:100%' });
      kind.disabled = !(pr.kind && pr.kind !== 'none');
      const on = !!pr.on;
      const start = ui.button(on ? 'Stop Protection' : 'Yes, Start Enforcing Protection', async () => {
        if (on) {
          if (pr.hash) { const pw = await ui.prompt('Password:', '', 'Unprotect Document'); if (pw == null) return; if (await hash(pw) !== pr.hash) { ui.msg('The password is incorrect.', { icon: 'warn' }); return; } }
          d.settings.protect = Object.assign({}, pr, { on: false });
        } else {
          if (!ed.input.checked && !fmt.input.checked) { ui.msg('Choose a restriction first.', { icon: 'info' }); return; }
          const pw = await ui.prompt('Enter new password (optional):', '', 'Start Enforcing Protection');
          if (pw == null) return;
          d.settings.protect = { on: true, kind: ed.input.checked ? kind.value : 'none', formatting: fmt.input.checked, hash: pw ? await hash(pw) : null };
          if (d.settings.protect.kind === 'tracked') d.settings.track = true;
        }
        d.dirty = true;
        E.readOnly = false;
        TP.render();
        A().updateStatus();
        ui.refresh();
      }, { class: 'btn', style: 'width:100%' });
      b.append(sec('1. Formatting restrictions', fmt), sec('2. Editing restrictions', ed, kind), sec('3. Start enforcement', h('div', { class: 'tp-note', text: on ? 'This document is protected from unintentional editing.' : 'Are you ready to apply these settings? (You can turn them off later)' }), start));
    },
  });
  async function hash(s) { try { const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)); return Array.from(new Uint8Array(buf)).map((x) => x.toString(16).padStart(2, '0')).join(''); } catch (e) { return String(s.length); } }
  /* ---------- Mail Merge wizard ---------- */
  TP.register({ id: 'mailmerge', title: 'Mail Merge', render(b) { if (L.mailmerge) L.mailmerge.renderWizard(b); else b.appendChild(h('div', { class: 'tp-note', text: 'Mail merge is not available.' })); }, refresh(b, reason) { if (reason === 'load') { L.clear(b); this.render(b); } } });

  /* ================= Document Map & thumbnails ================= */
  const DM = (P.docmap = { collapsed: new Set() });
  DM.mount = function (el) { DM.el = el; L.bus.on('doc-changed', L.debounce(() => { if (!el.hidden) (A().opts.thumbs ? P.thumbs.render() : DM.render()); }, 400)); L.bus.on('sel', L.debounce(() => { if (!el.hidden && !A().opts.thumbs) DM.mark(); }, 120)); L.bus.on('doc-loaded', () => { if (!el.hidden) DM.render(); }); };
  DM.render = function () {
    const el = DM.el;
    L.clear(el);
    const d = doc();
    D.reindex(d);
    const heads = [];
    for (const p of d.main.paras) { const lv = D.headingLevel(d, p); if (lv) heads.push({ p, lv }); }
    if (!heads.length) { el.appendChild(h('div', { class: 'dm-item', style: 'color:#666', text: '(No headings)' })); return; }
    let hideBelow = 0;
    heads.forEach((x, i) => {
      if (hideBelow && x.lv > hideBelow) return;
      hideBelow = 0;
      const hasKids = heads[i + 1] && heads[i + 1].lv > x.lv;
      const coll = DM.collapsed.has(x.p.id);
      const tog = h('span', { class: 'dm-tog', text: hasKids ? (coll ? '+' : '−') : '' });
      const it = h('div', { class: 'dm-item', 'data-pid': x.p.id, style: `padding-left:${4 + (x.lv - 1) * 12}px`, title: D.plainText(x.p) }, tog, h('span', { text: D.plainText(x.p).trim() || '(blank)' }));
      tog.addEventListener('click', (e) => { e.stopPropagation(); if (coll) DM.collapsed.delete(x.p.id); else DM.collapsed.add(x.p.id); DM.render(); });
      it.addEventListener('click', () => { A().gotoPos(D.pos(x.p, 0)); const f = LY.fragFor(x.p, 0); if (f) f.scrollIntoView({ block: 'start' }); });
      el.appendChild(it);
      if (coll) hideBelow = x.lv;
    });
    DM.mark();
  };
  DM.mark = function () {
    if (!E.sel || !DM.el) return;
    const d = doc();
    let p = E.sel.f.p;
    if (D.storyOf(d, p) !== d.main) return;
    while (p && !D.headingLevel(d, p)) p = D.prevPara(d, p);
    L.$$('.dm-item', DM.el).forEach((x) => x.classList.toggle('on', !!p && +x.dataset.pid === p.id));
  };
  P.thumbs = {
    render() {
      const el = DM.el;
      L.clear(el);
      const box = h('div', { class: 'thumbs' });
      const w = Math.max(60, el.clientWidth - 40);
      LY.pages.forEach((pg, i) => {
        const k = w / pg.el.offsetWidth;
        const frame = h('div', { class: 'thumb', style: `width:${w}px;height:${pg.el.offsetHeight * k}px` });
        const clone = pg.el.cloneNode(true);
        clone.removeAttribute('data-pg');
        for (const e2 of clone.querySelectorAll('[contenteditable]')) e2.removeAttribute('contenteditable');
        clone.style.cssText += `;transform:scale(${k});transform-origin:0 0;margin:0;pointer-events:none;box-shadow:none`;
        frame.appendChild(clone);
        frame.addEventListener('click', () => { pg.el.scrollIntoView({ block: 'start' }); L.$$('.thumb', box).forEach((t, j) => t.classList.toggle('on', j === i)); });
        box.append(frame, h('div', { class: 'thumb-n', text: String(i + 1) }));
      });
      if (!LY.pages.length) box.appendChild(h('div', { class: 'tp-note', text: 'Thumbnails show pages in Print Layout and Reading Layout.' }));
      el.appendChild(box);
    },
  };

  /* ================= Select Browse Object ================= */
  P.browseTo = function (kind, dir) {
    const d = doc();
    if (!E.sel) return;
    D.reindex(d);
    const targets = [];
    if (kind === 'section') { for (const s of D.sections(d)) { const b0 = d.main.blocks[s.from]; const p = b0 && (b0.t === 'p' ? b0 : D.firstPara([b0])); if (p) targets.push(D.pos(p, 0)); } }
    else if (kind === 'endnote' || kind === 'footnote' || kind === 'comment' || kind === 'field' || kind === 'graphic' || kind === 'edit') {
      for (const p of d.main.paras) {
        let o = 0;
        for (const it of p.runs) {
          const ok = (kind === 'footnote' && it.t === 'fn' && !it.self) || (kind === 'endnote' && it.t === 'en' && !it.self) || (kind === 'comment' && it.t === 'cs') || (kind === 'field' && it.t === 'fb') || (kind === 'graphic' && (it.t === 'img' || it.t === 'shape' || it.t === 'group' || it.t === 'chart')) || (kind === 'edit' && it.rPr && (it.rPr.ins || it.rPr.del));
          if (ok) targets.push(D.pos(p, o));
          o += D.ilen(it);
        }
      }
    } else if (kind === 'heading') { for (const p of d.main.paras) if (D.headingLevel(d, p)) targets.push(D.pos(p, 0)); }
    else if (kind === 'table') { D.walk(d.main, (b, c, i, st, parents) => { if (b.t === 'tbl' && !parents.length) { const p = D.firstPara([b]); if (p) targets.push(D.pos(p, 0)); } }); }
    if (!targets.length) { A().status('There are no items of that kind in this document.'); return; }
    const cur = E.sel.f;
    const t = dir > 0 ? targets.find((x) => D.cmp(d, x, cur) > 0) : targets.slice().reverse().find((x) => D.cmp(d, x, cur) < 0);
    if (!t) { A().status(dir > 0 ? 'Quire reached the end of the document.' : 'Quire reached the beginning of the document.'); return; }
    A().gotoPos(t);
  };

  /* ================= Print Preview ================= */
  const PV = (P.preview = { pages: 1 });
  PV.enter = function () {
    E.readOnly = true;
    LY.root.contentEditable = 'false';
    const bar = h('div', { class: 'toolbar preview-bar', id: 'tb-preview', role: 'toolbar', 'aria-label': 'Print Preview' });
    const btn = (icon, label, fn, tip, on) => { const b = h('button', { class: 'tb-btn' + (label ? ' with-label' : '') + (on ? ' on' : ''), type: 'button', 'data-tip': tip || label, 'aria-label': tip || label, html: (icon ? L.icons.get(icon) : '') + (label ? `<span class="tb-lbl">${label}</span>` : '') }); b.addEventListener('pointerdown', (e) => e.preventDefault()); b.addEventListener('click', fn); return b; };
    const zoom = ui.combo({ id: 'pv-zoom', width: 74, tip: 'Zoom', options: () => ['500%', '200%', '150%', '100%', '75%', '50%', '25%', '10%', 'Page Width', 'Text Width', 'Whole Page', 'Two Pages'], value: () => Math.round(LY.zoom * 100) + '%', onChange: (v) => A().zoomTo(v) });
    const magBtn = btn('zoom', '', () => { PV.mag = !PV.mag; magBtn.classList.toggle('on', PV.mag); E.readOnly = PV.mag; LY.root.contentEditable = PV.mag ? 'false' : 'true'; }, 'Magnifier', true);
    PV.mag = true;
    bar.append(h('span', { class: 'tb-grip' }), btn('print', '', () => A().printPDF(), 'Print'), magBtn, btn('onePage', '', () => { LY.setMultiPage(0); LY.setZoom(LY.zoomFor('wholePage')); }, 'One Page'), btn('multiPage', '', (e) => PV.multiMenu(e.currentTarget), 'Multiple Pages'), zoom, btn('ruler', '', () => ui.exec('ruler'), 'View Ruler'), btn('shrink', '', () => PV.shrink(), 'Shrink to Fit'), btn('fullScreen', '', () => ui.exec('fullScreen'), 'Full Screen'), btn('', 'Close Preview', () => A().setView(A().prevView || 'print'), 'Close Preview'), btn('help', '', () => ui.exec('help'), 'Help'));
    L.$('#toolbars').appendChild(bar);
    PV.bar = bar;
    PV.prevZoom = LY.zoom;
    LY.setZoom(LY.zoomFor('wholePage'));
    PV.click = (e) => {
      if (!PV.mag) return;
      const pg = e.target.closest('.pg');
      if (!pg) return;
      if (LY.zoom < 0.95) { LY.setMultiPage(0); LY.setZoom(1); pg.scrollIntoView({ block: 'center' }); L.$('#scroller').classList.add('zoomed'); }
      else { LY.setZoom(LY.zoomFor('wholePage')); pg.scrollIntoView({ block: 'start' }); L.$('#scroller').classList.remove('zoomed'); }
    };
    LY.scroller.addEventListener('click', PV.click);
    ui.refresh();
  };
  PV.exit = function () {
    if (PV.bar) PV.bar.remove();
    PV.bar = null;
    LY.scroller.removeEventListener('click', PV.click);
    L.$('#scroller').classList.remove('zoomed');
    E.readOnly = false;
    LY.root.contentEditable = 'true';
    LY.setMultiPage(0);
    if (PV.prevZoom) LY.setZoom(PV.prevZoom);
  };
  PV.multiMenu = function (anchor) {
    ui.openMenu([{
      custom: (close) => {
        const box = h('div', { class: 'tbl-grid-pick' });
        const cells = h('div', { class: 'tgp-cells', style: 'grid-template-columns:repeat(3,24px)' });
        const lbl = h('div', { class: 'tgp-label', text: 'Cancel' });
        for (let r = 0; r < 2; r++) for (let c = 0; c < 3; c++) {
          const s = h('span', { style: 'height:30px;width:24px' });
          s.addEventListener('pointerenter', () => { L.$$('span', cells).forEach((x, k) => x.classList.toggle('on', Math.floor(k / 3) <= r && k % 3 <= c)); lbl.textContent = `${r + 1} x ${c + 1} Pages`; });
          s.addEventListener('click', () => { close(); LY.setMultiPage(c + 1); const n = c + 1; const pg = LY.pages[0]; const avail = LY.scroller.clientWidth - 40, availH = LY.scroller.clientHeight - 30; if (pg) LY.setZoom(Math.min((avail - 16 * n) / (pg.W * n), (availH - 14 * (r + 1)) / (pg.H * (r + 1)))); });
          cells.appendChild(s);
        }
        box.append(cells, lbl);
        return box;
      },
    }], anchor.getBoundingClientRect(), { cls: 'grid-menu' });
  };
  /** Shrink One Page: reduce font sizes until the last page disappears */
  PV.shrink = function () {
    const d = doc();
    const n0 = LY.pages.length;
    if (n0 < 2) { A().status('This document is already one page long.'); return; }
    let ok = false;
    E.readOnly = false;
    E.edit('Shrink One Page', () => {
      D.touchAll();
      for (let step = 0; step < 6 && !ok; step++) {
        const scale = (rp) => { if (rp && rp.sz) rp.sz = Math.max(1, rp.sz - 0.5); };
        for (const k in d.styles) { const s = d.styles[k]; if (s.rPr && s.rPr.sz) scale(s.rPr); }
        d.defaults.rPr.sz = Math.max(1, (d.defaults.rPr.sz || 12) - 0.5);
        D.walk(d.main, (b) => { if (b.t === 'p') { for (const it of b.runs) scale(it.rPr); scale(b.rPr); } });
        D.stylesChanged();
        LY.render();
        ok = LY.pages.length < n0;
      }
      return E.sel;
    });
    E.readOnly = true;
    A().status(ok ? 'The document was shrunk by one page.' : 'Quire could not shrink this document by one page.');
  };

  /* ================= Reading Layout toolbar ================= */
  const RL = (P.reading = {});
  RL.enter = function () {
    const bar = h('div', { class: 'readbar', role: 'toolbar', 'aria-label': 'Reading Layout' });
    const btn = (icon, label, fn, tip) => { const b = h('button', { class: 'tb-btn' + (label ? ' with-label' : ''), type: 'button', 'data-tip': tip || label, 'aria-label': tip || label, html: (icon ? L.icons.get(icon) : '') + (label ? `<span class="tb-lbl">${label}</span>` : '') }); b.addEventListener('pointerdown', (e) => e.preventDefault()); b.addEventListener('click', fn); return b; };
    bar.append(btn('docMap', 'Document Map', () => A().toggleDocMap()), btn('thumbnails', 'Thumbnails', () => A().toggleThumbs()), btn('find', 'Find', () => ui.exec('find')), btn('research', 'Research', () => ui.exec('research')), h('span', { class: 'tb-sep' }), btn('growFont', '', () => { LY.setZoom(LY.zoom * 1.1); }, 'Increase Text Size'), btn('shrinkFont', '', () => { LY.setZoom(LY.zoom / 1.1); }, 'Decrease Text Size'), btn('onePage', '', () => { LY.setMultiPage(0); LY.setZoom(LY.zoomFor('wholePage')); }, 'Actual Page'), btn('multiPage', '', () => { LY.setMultiPage(2); LY.setZoom(LY.zoomFor('twoPages')); }, 'Allow Multiple Pages'), h('span', { class: 'tb-sep' }), btn('', 'Close', () => A().setView(A().prevView && A().prevView !== 'reading' ? A().prevView : 'print'), 'Close'));
    L.$('#toolbars').appendChild(bar);
    RL.bar = bar;
  };
  RL.exit = function () { if (RL.bar) RL.bar.remove(); RL.bar = null; };
  L.bus.on('view', (v) => { if (v === 'reading') { if (!RL.bar) RL.enter(); } else RL.exit(); });
})();
