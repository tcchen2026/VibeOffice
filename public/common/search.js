/* VibeOffice — Search commands: the box at the top right of every app, in every look (the menu row, or the
 * ribbon's title row). Typing lists the app's commands that match, each with its icon, shortcut and where it
 * lives (the menu, or the ribbon tab and group); Enter or a click runs it. Words match anywhere in a
 * command's name or tooltip, as the start of a word, through everyday names for Office's ones (image →
 * Picture, margins → Page Setup) and with one typo. Alt+Q goes to the box; with nothing typed it shows the
 * commands last run from it.
 *
 *   ui.searchBox() → a new box (each command-bar layout has its own; they share the recent list)
 *   ui.searchCommands(query, layout) → [{ id, label, where, enabled }], best first (tools/search.test.mjs)
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const { h } = L;
  const ui = L.ui;
  const MAX = 10, RECENT = 6;

  /* everyday words and spellings → the words Office's command names use ('|' between alternatives) */
  const SYN = {
    smartart: 'intelliart', wordart: 'lettersart', diagram: 'intelliart|diagram', diagrams: 'intelliart|diagram', orgchart: 'intelliart|organization chart', fontwork: 'lettersart',
    image: 'picture|from file', images: 'picture|from file', pdf: 'pdf|save as', png: 'save as', html: 'web page|save as', photo: 'picture|photo album', photos: 'picture|photo album', pic: 'picture', img: 'picture',
    graph: 'chart', graphs: 'chart', plot: 'chart', colour: 'color', colours: 'color', grey: 'gray', centre: 'center', centred: 'center',
    remove: 'delete|clear', erase: 'delete|clear|eraser', strikethrough: 'strike', crossed: 'strike',
    margin: 'page setup', margins: 'page setup', orientation: 'page setup', landscape: 'page setup', portrait: 'page setup', paper: 'page setup',
    note: 'comment', notes: 'comment|notes', annotate: 'comment', toc: 'table of contents|index and tables', contents: 'index and tables|table of contents',
    sum: 'autosum', total: 'autosum|subtotal', totals: 'subtotal', redline: 'track changes', revisions: 'track changes|changes',
    password: 'protect|permission|encrypt', encrypt: 'protect|permission', lock: 'protect|freeze',
    emoji: 'symbol', special: 'symbol', character: 'symbol', characters: 'symbol|word count', link: 'hyperlink', url: 'hyperlink', links: 'hyperlink|links',
    search: 'find', locate: 'find', substitute: 'replace', swap: 'replace',
    uppercase: 'change case', lowercase: 'change case', capital: 'change case', capitals: 'change case', caps: 'change case|all caps',
    bigger: 'grow font', larger: 'grow font', enlarge: 'grow font', smaller: 'shrink font', plain: 'clear formatting|clear formats',
    brush: 'format painter', background: 'background|page color|fill', height: 'spacing|row height', list: 'bullets|numbering',
    bullet: 'bullets', numbered: 'numbering', shading: 'fill|shading', marker: 'highlight', resize: 'size|autofit', fit: 'autofit|fit',
    present: 'slide show|view show|from beginning', presentation: 'slide show|view show', play: 'slide show|view show|from beginning', slideshow: 'slide show|view show', speaker: 'notes',
    money: 'currency', dollar: 'currency', euro: 'currency', percentage: 'percent', decimals: 'decimal',
    formula: 'function|formula', vlookup: 'function', lookup: 'function', average: 'function|autosum', count: 'function|word count',
    export: 'save as|export|pdf', download: 'save as', statistics: 'word count', stats: 'word count',
    dictionary: 'spelling', spellcheck: 'spelling', spell: 'spelling', typo: 'spelling', proofread: 'spelling', synonym: 'thesaurus', synonyms: 'thesaurus',
    translate: 'language', settings: 'options', preferences: 'options', appearance: 'options', theme: 'options|theme', themes: 'options|theme',
    shortcut: 'keyboard shortcuts', shortcuts: 'keyboard shortcuts', hotkey: 'keyboard shortcuts', hotkeys: 'keyboard shortcuts',
    revert: 'undo', magnify: 'zoom', fullscreen: 'full screen', citation: 'footnote', pin: 'freeze', combine: 'merge', join: 'merge',
    alphabetize: 'sort', order: 'sort|bring to front|send to back', clone: 'duplicate', textbox: 'text box', shape: 'autoshapes|shapes',
    arrow: 'autoshapes|arrow', rectangle: 'autoshapes', circle: 'autoshapes', oval: 'autoshapes', email: 'send|mail', mail: 'mail merge|send',
    vba: 'macros', script: 'macros', quit: 'exit', blank: 'new', recent: 'open', author: 'properties', metadata: 'properties',
    info: 'properties', details: 'properties', tab: 'tabs|sheet', worksheet: 'sheet', column: 'column|columns', row: 'row|rows',
  };

  const norm = (s) => ui.stripAmp(String(s || '')).replace(/(\.\.\.|…)$/, '').trim();
  /* words, with compound names split as well as whole (AutoSum → autosum, auto, sum) */
  const words = (s) => { const w = String(s).toLowerCase().split(/[^a-z0-9%$&+]+/).filter(Boolean); const parts = String(s).replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase().split(/[^a-z0-9%$&+]+/).filter(Boolean); return w.concat(parts.filter((x) => !w.includes(x))); };
  const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /* one typo: a letter wrong, missing, extra or two letters swapped */
  function near(a, b) {
    if (Math.abs(a.length - b.length) > 1) return false;
    let i = 0;
    while (i < a.length && a[i] === b[i]) i++;
    if (i === a.length && i === b.length) return true;
    if (a.length === b.length) return a.slice(i + 1) === b.slice(i + 1) || (a[i] === b[i + 1] && a[i + 1] === b[i] && a.slice(i + 2) === b.slice(i + 2));
    return a.length > b.length ? a.slice(i + 1) === b.slice(i) : a.slice(i) === b.slice(i + 1);
  }

  /* where each command lives: its menu path (menu bar) and its ribbon tab and group */
  function places() {
    const menu = new Map(), ribbon = new Map(), actions = [];
    const walkMenu = (items, path) => {
      let list;
      try { list = typeof items === 'function' ? items() : items; } catch (e) { return; }
      for (const x of list || []) {
        if (typeof x === 'string') { if (x !== '-' && !menu.has(x)) menu.set(x, path); continue; }
        if (!x || x.custom) continue;
        if (x.cmd) { if (!menu.has(x.cmd)) menu.set(x.cmd, path); continue; }
        if (Array.isArray(x.sub)) walkMenu(x.sub, path.concat(norm(x.label)));
        else if (x.run && x.label) actions.push({ label: norm(x.label), run: x.run, icon: x.icon, enabled: x.enabled, path });
      }
    };
    for (const m of ui.menus || []) walkMenu(m.items, [norm(m.label)]);
    const spec = L.ribbonSpec ? L.ribbonSpec() : null;
    if (spec) {
      const put = (id, where, label) => { if (id && !ribbon.has(id)) ribbon.set(id, { where, label: label && norm(label) }); };
      const walk = (it, where) => {
        if (!it || it === '|') return;
        if (typeof it === 'string') return put(it, where);
        if (it.row) return it.row.forEach((x) => walk(x, where));
        put(it.cmd || it.split, where, it.label);
      };
      const tabs = spec.tabs.map((t) => [t, '']).concat(...(spec.contextual || []).map((c) => c.tabs.map((t) => [t, c.set + ' ▸ '])));
      for (const [t, pre] of tabs) for (const g of t.groups) {
        const where = [pre + norm(t.label), g.label];
        g.items.forEach((it) => walk(it, where));
        put(g.launcher, where);
      }
      (spec.qat || []).forEach((x) => walk(x, ['Quick Access Toolbar']));
      for (const x of (spec.file ? spec.file() : [])) if (typeof x === 'string' && x !== '-') put(x, ['File']);
    }
    /* commands only on a toolbar: 'Formatting toolbar' */
    const toolbar = new Map();
    for (const tb of L.$$('.toolbar[aria-label]')) for (const b of L.$$('[data-cmd]', tb)) if (!toolbar.has(b.dataset.cmd)) toolbar.set(b.dataset.cmd, [tb.getAttribute('aria-label') + ' toolbar']);
    return { menu, ribbon, toolbar, actions };
  }

  /* the searchable entries for the current layout */
  function entries(layout) {
    const { menu, ribbon, toolbar, actions } = places();
    const out = [], seen = new Set();
    for (const id of Object.keys(ui.cmds)) {
      const c = ui.cmds[id];
      if (!c.label || typeof c.label !== 'string') continue;
      if (layout === 'ribbon' && id.startsWith('tb_')) continue;   // toolbars are not shown with the ribbon
      const rb = ribbon.get(id), mp = menu.get(id);
      /* the name the user sees: the ribbon's own label for the button if it has one; a command in a submenu
         also answers to the submenu's name (Insert ▸ Picture ▸ From File) */
      const label = layout === 'ribbon' && rb && rb.label ? rb.label : norm(c.label);
      const path = (layout === 'ribbon' ? rb && rb.where : mp || toolbar.get(id)) || null;
      /* c.keys: more words the command answers to */
      const alias = [norm(c.label), c.keys || '', layout !== 'ribbon' && mp && mp.length > 1 ? mp[mp.length - 1] : ''].join(' ');
      const key = label.toLowerCase() + '|' + (path || []).join('/');
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ id, label, alias, keys: (c.keys || '').toLowerCase(), path, tip: norm(c.tip || ''), icon: c.icon, placed: !!path, toggle: id.startsWith('tb_') });
    }
    if (layout !== 'ribbon') for (const a of actions) {
      const key = a.label.toLowerCase() + '|' + a.path.join('/');
      if (!seen.has(key)) { seen.add(key); out.push({ label: a.label, alias: a.path.length > 1 ? a.path[a.path.length - 1] : '', path: a.path, tip: '', icon: a.icon, run: a.run, enabledFn: a.enabled, placed: true }); }
    }
    return out;
  }

  /* how well one query word matches an entry: 0 = not at all. Its name counts most (whole word, start of a
     word, an everyday synonym, one typo), then its tooltip and location. */
  function wordScore(q, e) {
    const own = e.own || (e.own = words(e.label)), also = e.also || (e.also = words(e.alias).filter((w) => !own.includes(w)));
    const hit = (list) => {
      let best = 0;
      for (const w of list) {
        if (w === q) best = Math.max(best, 5);
        else if (w.startsWith(q)) best = Math.max(best, 4);
        else if (q.length >= 4 && w.length >= 4 && near(q, w)) best = Math.max(best, 2.5);
      }
      const text = ' ' + list.join(' ') + ' ';
      if (SYN[q] && SYN[q].split('|').some((alt) => text.includes(' ' + alt + ' ') || text.includes(' ' + alt + 's '))) best = Math.max(best, 3.5);
      return best;
    };
    const best = Math.max(hit(own), hit(also) - 0.5);   // its own name before the submenu's
    if (best > 0) return best;
    return words(e.tip + ' ' + (e.path || []).join(' ')).some((w) => w.startsWith(q)) ? 1 : 0;
  }

  const enabledOf = (e) => (e.id ? ui.enabled(e.id) : !e.enabledFn || !!e.enabledFn());

  ui.searchCommands = (query, layout) => {
    layout = layout || document.documentElement.dataset.layout || 'bars';
    const qs = words(query);
    if (!qs.length) return [];
    const recent = L.store ? L.store.get('searchRecent', []) : [];
    const full = qs.join(' ');
    const scored = [];
    for (const e of entries(layout)) {
      let s = 0;
      for (const q of qs) { const w = wordScore(q, e); if (!w) { s = 0; break; } s += w; }
      if (!s) continue;
      const lab = e.label.toLowerCase();
      if (lab === full) s += 6; else if (lab.startsWith(full) || (e.keys && (' ' + e.keys).includes(' ' + full))) s += 3;
      if (e.placed) s += 0.5;
      if (e.toggle) s -= 4;   // showing a toolbar is rarely what is meant
      if (e.run) s -= 1.5;    // menu entries that are not commands (open windows, recent files) after commands
      if (e.id && recent.includes(e.id)) s += 0.5;
      const on = enabledOf(e);
      if (on) s += 2;   // what can run now before what cannot
      scored.push({ e, s, on });
    }
    scored.sort((a, b) => b.s - a.s || a.e.label.length - b.e.label.length);
    return scored.slice(0, MAX).map(({ e, on }) => ({ id: e.id, label: e.label, where: e.path ? e.path.join(' ▸ ') : '', enabled: on, entry: e }));
  };

  function recentList(layout) {
    const ids = L.store ? L.store.get('searchRecent', []) : [];
    const all = entries(layout);
    return ids.map((id) => all.find((e) => e.id === id)).filter(Boolean).map((e) => ({ id: e.id, label: e.label, where: e.path ? e.path.join(' ▸ ') : '', enabled: enabledOf(e), entry: e }));
  }
  function remember(id) {
    if (!id || !L.store) return;
    L.store.set('searchRecent', [id].concat(L.store.get('searchRecent', []).filter((x) => x !== id)).slice(0, RECENT));
  }

  /* bold the parts of a label the query words found */
  function marked(label, qs) {
    const lw = label.toLowerCase(), on = new Array(label.length).fill(false);
    for (const q of qs) {
      const re = new RegExp('(^|[^a-z0-9])(' + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'g');
      let m;
      while ((m = re.exec(lw))) for (let i = m.index + m[1].length; i < m.index + m[0].length; i++) on[i] = true;
    }
    let out = '', open = false;
    for (let i = 0; i < label.length; i++) {
      if (on[i] !== open) { out += on[i] ? '<b>' : '</b>'; open = on[i]; }
      out += esc(label[i]);
    }
    return out + (open ? '</b>' : '');
  }

  const SEARCH_ICON = '<svg class="cs-glass" width="12" height="12" viewBox="0 0 12 12" aria-hidden="true"><circle cx="5" cy="5" r="3.6" fill="none" stroke="currentColor" stroke-width="1.3"/><path d="M7.7 7.7L11 11" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
  let uid = 0;

  ui.searchBox = function () {
    const listId = 'cs-list-' + ++uid;
    const inp = h('input', { type: 'search', placeholder: 'Search commands', 'aria-label': 'Search commands', 'data-tip': 'Search commands (Alt+Q)', spellcheck: 'false', autocomplete: 'off', role: 'combobox', 'aria-expanded': 'false', 'aria-controls': listId, 'aria-autocomplete': 'list' });
    const box = h('div', { class: 'cmd-search', role: 'search', html: SEARCH_ICON });
    box.append(inp);
    let pop = null, rows = [], act = -1;

    const close = () => { if (pop) { pop.remove(); pop = null; } rows = []; act = -1; inp.setAttribute('aria-expanded', 'false'); inp.removeAttribute('aria-activedescendant'); };
    const setAct = (i) => {
      if (!rows.length) return;
      act = (i + rows.length) % rows.length;
      rows.forEach((r, j) => r.el.classList.toggle('act', j === act));
      inp.setAttribute('aria-activedescendant', rows[act].el.id);
      rows[act].el.scrollIntoView({ block: 'nearest' });
    };
    const run = (r) => {
      if (!r || !r.enabled) return;
      close();
      inp.value = '';
      inp.blur();
      if (ui.hooks && ui.hooks.refocus) ui.hooks.refocus();
      if (r.id) { remember(r.id); ui.exec(r.id); } else if (r.entry.run) r.entry.run();
    };
    const show = () => {
      const q = inp.value.trim(), layout = document.documentElement.dataset.layout || 'bars';
      const list = q ? ui.searchCommands(q, layout) : recentList(layout);
      if (!q && !list.length) { close(); return; }
      ui.closeMenus();
      if (!pop) {
        pop = h('div', { class: 'menu cs-pop', id: listId, role: 'listbox', 'aria-label': 'Commands' });
        pop.addEventListener('pointerdown', (e) => e.preventDefault());   // the box keeps the focus
        document.body.append(pop);
      }
      L.clear(pop);
      rows = [];
      if (!q) pop.append(h('div', { class: 'm-head', text: 'Recently used' }));
      if (q && !list.length) pop.append(h('div', { class: 'm-item dis cs-none', text: `No commands match “${q}”` }));
      const qs = words(q);
      list.forEach((r, i) => {
        const c = r.id ? ui.cmds[r.id] : null, icon = r.entry.icon;
        const checked = !!(r.id && ui.checked(r.id));
        const el = h('div', { class: 'm-item' + (r.enabled ? '' : ' dis'), id: listId + '-' + i, role: 'option', 'aria-disabled': String(!r.enabled), 'data-cmd': r.id || '' });
        el.append(h('span', { class: 'm-ico' + (checked && icon ? ' on' : ''), html: icon ? L.icons.get(icon) : checked ? '<span class="m-check">✓</span>' : '' }),
          h('span', { class: 'm-lbl', html: marked(r.label, qs) + (r.where ? `<span class="cs-where">${esc(r.where)}</span>` : '') }),
          h('span', { class: 'm-key', text: (c && c.key) || '' }), h('span'));
        el.addEventListener('pointerenter', () => setAct(i));
        el.addEventListener('click', () => run(r));
        pop.append(el);
        rows.push({ el, ...r });
      });
      const b = inp.getBoundingClientRect(), w = Math.min(420, window.innerWidth - 8);
      pop.style.width = w + 'px';
      pop.style.left = Math.max(4, Math.min(b.right - w, window.innerWidth - w - 4)) + 'px';
      pop.style.top = b.bottom + 2 + 'px';
      pop.style.maxHeight = Math.max(120, window.innerHeight - b.bottom - 10) + 'px';
      inp.setAttribute('aria-expanded', 'true');
      act = -1;
      const first = rows.findIndex((r) => r.enabled);
      if (q && rows.length) setAct(first < 0 ? 0 : first);
    };

    inp.addEventListener('input', show);
    inp.addEventListener('focus', () => { inp.select(); show(); });
    inp.addEventListener('blur', () => setTimeout(() => { if (document.activeElement !== inp) close(); }, 0));
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); if (!pop) show(); else setAct(act + (e.key === 'ArrowDown' ? 1 : -1)); }
      else if (e.key === 'Enter') { e.preventDefault(); run(rows[act]); }
      else if (e.key === 'Escape') { e.preventDefault(); if (inp.value) { inp.value = ''; close(); } else { close(); inp.blur(); if (ui.hooks && ui.hooks.refocus) ui.hooks.refocus(); } }
      else if (e.key === 'Tab') close();
      e.stopPropagation();   // typing here is not a shortcut for the document
    });
    window.addEventListener('resize', close);
    return box;
  };

  /* Alt+Q, as in Office: to the box of the layout in use */
  if (typeof document !== 'undefined') document.addEventListener('keydown', (e) => {
    if (!e.altKey || e.ctrlKey || e.shiftKey || e.metaKey || (e.code !== 'KeyQ' && (e.key || '').toLowerCase() !== 'q')) return;
    const inp = L.$$('.cmd-search input').find((x) => x.getClientRects().length);
    if (!inp) return;
    e.preventDefault();
    e.stopPropagation();
    inp.focus();
  }, true);
})(typeof window !== 'undefined' ? window : globalThis);
