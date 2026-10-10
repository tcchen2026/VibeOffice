/* VibeOffice — the ribbon: the command bars of looks whose layout is 'ribbon' (Paper 2016; common/suite.js
 * VO.look). The menus and toolbars stay the commands' other presentation; both run the same ui.cmds.
 *
 * Each app declares its ribbon as data, L.ribbonSpec = () => ({
 *   qat: ['save', 'undo', 'redo'],                 Quick Access Toolbar
 *   file: () => [menu items],                      the File tab opens these as a menu (ui.openMenu)
 *   tabs: [{ id, label: '&Home', groups: [{ label, launcher: cmdId, items: [...] }] }],
 *   contextual: [{ id, set: 'Table Tools', when: () => bool, tabs: [...] }],   shown while when() holds
 *   icons: { cmdId: iconName },                    icons for commands that have none of their own
 * }), and calls L.ribbonUI.init() after its toolbars exist. Group items:
 *   'cmd' | { cmd, size: 'large'|'medium'|'small', label, icon, arg }     a command button
 *   { split: cmd, menu(rect, anchor), size, label, icon }               button plus menu arrow
 *   { drop: label, icon, menu(rect, anchor), size, tip }                a menu button
 *   { make: () => Node, size }                                           a combo box or other control
 *   { row: [items] }                                                     small items side by side (Font, Paragraph)
 *   '|'                                                                  start a new column
 * Large items take the group's height; medium items stack three to a column, rows two to a column.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const { h } = L;
  const ui = L.ui;
  const RB = (L.ribbonUI = {});
  let spec = null, el = null, tabsEl = null, panelEl = null, cur = 'home';
  const panels = new Map();   // tab id → its panel, built once (the buttons stay registered with ui.refresh)
  let back = null, ctxShown = '';

  const isRibbon = () => document.documentElement.dataset.layout === 'ribbon';
  const cmdLabel = (id) => ui.stripAmp((ui.cmds[id] || {}).label || id).replace(/(\.\.\.|…)$/, '');
  const LAUNCH = '<svg width="9" height="9" viewBox="0 0 9 9" aria-hidden="true"><path d="M1 1h4M1 1v4M3.5 3.5l4 4M8 4.5V8H4.5" fill="none" stroke="currentColor"/></svg>';

  /** one item → { el, size } */
  function item(it) {
    if (typeof it === 'string') it = { cmd: it };
    if (!it.icon && spec.icons) it = Object.assign({ icon: spec.icons[it.cmd || it.split] }, it);
    const size = it.size || 'medium';
    const label = size === 'small' ? '' : it.label != null ? it.label : it.cmd || it.split ? cmdLabel(it.cmd || it.split) : '';
    let node;
    if (it.row) {
      node = h('div', { class: 'rb-row' });
      for (const x of it.row) node.append(x === '|' ? h('span', { class: 'rb-gap' }) : item(typeof x === 'string' ? { cmd: x, size: 'small' } : Object.assign({ size: 'small' }, x)).el);
      return { el: node, size: 'row' };
    }
    if (it.make) node = it.make();
    else if (it.split) node = ui.tbSplit(it.split, it.menu, { label, icon: it.icon });
    else if (it.drop) node = ui.tbDrop(size === 'small' ? '' : it.label != null ? it.label : it.drop, it.icon, it.menu, it.tip || ui.stripAmp(it.drop));
    else node = ui.tbButton(it.cmd, { label, icon: it.icon, arg: it.arg });
    node.classList.add('rb-' + size);
    if (size === 'large') {
      const icon = it.icon || (ui.cmds[it.cmd || it.split] || {}).icon;
      const svg = node.querySelector('svg.ico');
      if (svg && icon) svg.outerHTML = L.icons.get(icon, 32);
    }
    return { el: node, size };
  }

  function group(g) {
    const body = h('div', { class: 'rb-body' });
    let col = null, n = 0;
    for (const x of g.items) {
      if (x === '|') { col = null; continue; }
      const { el: node, size } = item(x);
      if (size === 'large') { body.append(node); col = null; continue; }
      const kind = size === 'row' ? 'rows' : 'stack';
      if (!col || col.dataset.kind !== kind || n >= (kind === 'rows' ? 2 : 3)) { col = h('div', { class: 'rb-col', 'data-kind': kind }); body.append(col); n = 0; }
      col.append(node);
      n++;
    }
    const label = h('div', { class: 'rb-label' }, g.label);
    if (g.launcher) {
      const b = h('button', { class: 'rb-launch', type: 'button', 'aria-label': cmdLabel(g.launcher), 'data-tip': cmdLabel(g.launcher), html: LAUNCH });
      b.addEventListener('pointerdown', (e) => e.preventDefault());
      b.addEventListener('click', (e) => ui.exec(g.launcher, undefined, e));
      label.append(b);
    }
    /* in a narrow window the group folds into one button that opens it below (fit()) */
    const first = g.items.flatMap((x) => (x && x.row) || [x]).find((x) => x !== '|' && !(x && x.make));
    const icon = g.icon || (first && (typeof first === 'string' ? (spec.icons || {})[first] || (ui.cmds[first] || {}).icon : first.icon || (spec.icons || {})[first.cmd || first.split] || (ui.cmds[first.cmd || first.split] || {}).icon));
    const fold = h('button', { class: 'tb-btn rb-large rb-fold tb-drop with-label', type: 'button', 'aria-label': g.label, 'data-tip': g.label, html: (icon ? L.icons.get(icon, 32) : '') + `<span class="tb-lbl">${L.esc(g.label)}</span><span class="dd-arrow"></span>` });
    const grp = h('div', { class: 'rb-group', role: 'group', 'aria-label': g.label }, fold, body, label);
    fold.addEventListener('pointerdown', (e) => e.preventDefault());
    fold.addEventListener('click', () => openFold(grp, fold));
    return grp;
  }
  /* a folded group's controls in a box below its button; they move back when the box closes */
  let foldOpen = null;
  function openFold(grp, btn) {
    closeFold();
    const r = btn.getBoundingClientRect();
    const pop = h('div', { class: 'rb-pop rb-group' });
    pop.style.left = Math.max(2, Math.min(r.left, window.innerWidth - 320)) + 'px';
    pop.style.top = r.bottom + 'px';
    pop.append(...grp.querySelectorAll(':scope > .rb-body, :scope > .rb-label'));
    document.body.append(pop);
    const away = (e) => { if (!pop.contains(e.target) && !e.target.closest('.menu')) closeFold(); };
    const done = (e) => { if (e.target.closest('.tb-btn:not(.tb-drop)')) setTimeout(closeFold, 0); };
    const esc = (e) => { if (e.key === 'Escape' && !document.querySelector('.menu')) { e.stopPropagation(); closeFold(); btn.focus(); } };
    foldOpen = { grp, pop, away, esc };
    document.addEventListener('pointerdown', away, true);
    document.addEventListener('keydown', esc, true);
    pop.addEventListener('click', done);
    ui.refresh();
  }
  function closeFold() {
    if (!foldOpen) return;
    const { grp, pop, away, esc } = foldOpen;
    grp.append(...pop.querySelectorAll(':scope > .rb-body, :scope > .rb-label'));
    pop.remove();
    document.removeEventListener('pointerdown', away, true);
    document.removeEventListener('keydown', esc, true);
    foldOpen = null;
  }
  /** fold groups from the right until the tab fits the window */
  function fit() {
    if (!el || !panelEl.firstChild) return;
    closeFold();
    const groups = [...panelEl.querySelectorAll(':scope > .rb-groups > .rb-group')];
    groups.forEach((g) => g.classList.remove('folded'));
    for (let i = groups.length - 1; i > 0 && panelEl.scrollWidth > panelEl.clientWidth + 1; i--) groups[i].classList.add('folded');
  }

  const allTabs = () => spec.tabs.concat(...(spec.contextual || []).map((c) => c.tabs.map((t) => Object.assign({ ctx: c }, t))));
  function panelFor(id) {
    if (!panels.has(id)) {
      const t = allTabs().find((x) => x.id === id);
      panels.set(id, h('div', { class: 'rb-groups', role: 'tabpanel' }, ...(t ? t.groups.map(group) : [])));
      ui.refresh();
    }
    return panels.get(id);
  }
  function tabButton(t, cls) {
    const b = h('button', { class: 'rb-tab' + (cls ? ' ' + cls : ''), type: 'button', role: 'tab', 'data-tab': t.id, html: ui.ampHTML(t.label) });
    b.addEventListener('pointerdown', (e) => e.preventDefault());
    b.addEventListener('click', () => (t.id === 'file' ? openFile(b) : RB.select(t.id)));
    if (t.id !== 'file') b.addEventListener('dblclick', () => RB.collapse());
    return b;
  }
  function openFile(b) {
    closeFold();
    const r = b.getBoundingClientRect();
    ui.openMenu(spec.file(), { left: r.left, bottom: r.bottom });
  }
  /** the tab strip: File, the app's tabs, then the contextual sets that apply now */
  function drawTabs() {
    tabsEl.textContent = '';
    tabsEl.append(tabButton({ id: 'file', label: '&File' }, 'file'));
    for (const t of spec.tabs) tabsEl.append(tabButton(t));
    for (const c of active()) tabsEl.append(h('div', { class: 'rb-ctx' }, h('span', { class: 'rb-ctx-name' }, c.set), h('div', { class: 'rb-ctx-tabs' }, ...c.tabs.map((t) => tabButton(t, 'ctx')))));
    for (const b of tabsEl.querySelectorAll('.rb-tab')) b.classList.toggle('on', b.dataset.tab === cur);
  }
  const active = () => (spec.contextual || []).filter((c) => { try { return c.when(); } catch (e) { return false; } });

  /** show a tab ('home', 'insert', … or a contextual tab's id) */
  RB.select = function (id) {
    if (!el) return;
    if (!allTabs().some((t) => t.id === id)) id = spec.tabs[0].id;
    cur = id;
    panelEl.textContent = '';
    panelEl.append(panelFor(id));
    for (const b of tabsEl.querySelectorAll('.rb-tab')) b.classList.toggle('on', b.dataset.tab === cur);
    if (el.classList.contains('collapsed')) peek(true);
    fit();
    ui.refresh();   // a panel shown again: its buttons and boxes were not refreshed while detached
  };
  /* collapsed (Ctrl+F1, a double-click on a tab, the pin): only the tabs show; a tab opens its panel over the
     document until a command runs or the pointer goes elsewhere */
  RB.collapse = function (on) {
    if (!el) return;
    on = on == null ? !el.classList.contains('collapsed') : !!on;
    el.classList.toggle('collapsed', on);
    peek(false);
    L.store.set('ribbonCollapsed', on);
    fit();
  };
  let peekAway = null;
  function peek(on) {
    el.classList.toggle('peek', on);
    if (peekAway) { document.removeEventListener('pointerdown', peekAway, true); peekAway = null; }
    if (on) {
      peekAway = (e) => { if (!el.contains(e.target) && !e.target.closest('.menu, .rb-pop')) peek(false); };
      document.addEventListener('pointerdown', peekAway, true);
    }
  }
  /* contextual tabs follow the selection: re-checked after every command bar refresh */
  function updateContext() {
    if (!el) return;
    const sig = active().map((c) => c.id).join(' ');
    if (sig === ctxShown) return;
    const before = ctxShown.split(' ');
    ctxShown = sig;
    drawTabs();
    /* as in Office, a tool set opens when its object was just inserted (from the Insert tab) or is a view
       of its own (open: true); when it goes, the tab used before it comes back */
    const main = spec.tabs.some((t) => t.id === cur);
    const take = active().find((c) => !before.includes(c.id) && (c.open || cur === 'insert'));
    if (take) { if (main) back = cur; RB.select(take.tabs[0].id); }
    else if (!main && !active().some((c) => c.tabs.some((t) => t.id === cur))) RB.select(back || spec.tabs[0].id);
    else fit();
  }

  /* Search commands: the labels of every command; Enter runs the first match */
  function searchBox() {
    const inp = h('input', { type: 'search', class: 'rb-search', placeholder: 'Search commands', 'aria-label': 'Search commands', spellcheck: 'false', autocomplete: 'off' });
    const matches = () => {
      const q = inp.value.trim().toLowerCase();
      if (!q) return [];
      const seen = new Set();
      return Object.keys(ui.cmds).filter((id) => {
        const c = ui.cmds[id], lab = cmdLabel(id).toLowerCase();
        if (!c.label || id.startsWith('tb_') || seen.has(lab) || !lab.includes(q)) return false;
        seen.add(lab);
        return true;
      }).sort((a, b) => cmdLabel(a).toLowerCase().indexOf(q) - cmdLabel(b).toLowerCase().indexOf(q)).slice(0, 12);
    };
    const show = () => {
      const m = matches();
      ui.closeMenus();
      if (!m.length) return;
      const r = inp.getBoundingClientRect();
      ui.openMenu(m.map((id) => ({ cmd: id, run: () => { inp.value = ''; ui.exec(id); } })), { left: r.left, bottom: r.bottom });
    };
    inp.addEventListener('input', show);
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { const m = matches(); if (m.length) { ui.closeMenus(); inp.value = ''; ui.hooks.refocus(); ui.exec(m[0]); } e.preventDefault(); }
      if (e.key === 'Escape') { inp.value = ''; ui.closeMenus(); ui.hooks.refocus(); }
      e.stopPropagation();
    });
    return inp;
  }

  /* KeyTips: Alt or F10 shows a letter on each tab (and a number on the Quick Access Toolbar); the letter opens
     the tab and shows letters on its commands. Esc goes back a level. */
  let tips = null;
  const ACT = '.rb-groups .tb-btn:not(.tb-dd), .rb-groups .combo, .rb-launch';
  /* One letter where the command's access key or initial is free, otherwise two letters (as Office's
     "FF", "AN"); a two-letter tip never starts with a one-letter tip, so typing is unambiguous. */
  function assign(items) {
    const used = new Set(), out = new Map(), rest = [];
    const letters = (it) => [...(it.label || '').toUpperCase()].filter((ch) => /[A-Z]/.test(ch));
    const take = (k, it) => { used.add(k); out.set(k, it); };
    for (const it of items) if (it.key && !used.has(it.key)) take(it.key, it); else rest.push(it);
    const later = [];
    for (const it of rest) { const ch = letters(it)[0]; if (ch && !used.has(ch) && !rest.some((o) => o !== it && letters(o)[0] === ch)) take(ch, it); else later.push(it); }
    const singles = new Set([...used].filter((k) => k.length === 1));
    const ABC = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    for (const it of later) {
      const ls = letters(it), words = (it.label || '').toUpperCase().split(/[^A-Z]+/).filter(Boolean);
      const firsts = [...ls, ...ABC].filter((ch) => !singles.has(ch) && !/[0-9]/.test(ch));
      let k = null;
      for (const a of firsts) {
        for (const c of [...words.slice(1).map((w) => w[0]), ...ls.slice(1), ...ABC]) if (!used.has(a + c)) { k = a + c; break; }
        if (k) break;
      }
      take(k, it);
    }
    return out;
  }
  function badge(target, key) {
    const r = target.getBoundingClientRect();
    const b = h('span', { class: 'rb-keytip' }, key);
    b.style.left = r.left + r.width / 2 + 'px';
    b.style.top = r.bottom - 7 + 'px';
    document.body.append(b);
    return b;
  }
  const accessKey = (label) => { const m = /&([^&])/.exec(label || ''); return m ? m[1].toUpperCase() : null; };
  function showTips(level) {
    clearTips(true);
    const items = [];
    if (level === 'tabs') {
      L.$$('.rb-qat > *', el).forEach((x, i) => items.push({ el: x.matches('.tb-split') ? x.firstChild : x, key: String(i + 1) }));
      L.$$('.rb-tab', tabsEl).forEach((x) => { const t = x.dataset.tab === 'file' ? { label: '&File' } : allTabs().find((y) => y.id === x.dataset.tab); items.push({ el: x, label: ui.stripAmp(t.label), key: accessKey(t.label), tab: x.dataset.tab }); });
    } else {
      L.$$(ACT, panelEl).forEach((x) => { if (!x.getClientRects().length) return; const c = ui.cmds[x.dataset.cmd] || {}; items.push({ el: x, label: x.getAttribute('aria-label') || x.textContent || x.dataset.tip || ui.stripAmp(c.label || "") }); });
    }
    const map = assign(items);
    tips = { level, map, buf: '', els: [...map].map(([k, it]) => badge(it.el, k)) };
    document.addEventListener('keydown', tipKey, true);
  }
  function clearTips(keepListener) {
    if (!tips) return;
    tips.els.forEach((x) => x.remove());
    tips = null;
    if (!keepListener) document.removeEventListener('keydown', tipKey, true);
  }
  function tipKey(e) {
    if (!tips || ['Alt', 'Shift', 'Control', 'Meta'].includes(e.key)) return;
    e.preventDefault();
    e.stopPropagation();
    if (e.key === 'Escape') { if (tips.level === 'cmds') showTips('tabs'); else { clearTips(); peek(false); } return; }
    if (e.key.length !== 1) { clearTips(); return; }
    tips.buf += e.key.toUpperCase();
    const hit = tips.map.get(tips.buf);
    if (!hit) { if (![...tips.map.keys()].some((k) => k.startsWith(tips.buf))) tips.buf = ''; return; }
    const level = tips.level;
    clearTips();
    if (level === 'tabs' && hit.tab) {
      if (hit.tab === 'file') { openFile(hit.el); return; }
      RB.select(hit.tab);
      showTips('cmds');
      return;
    }
    if (hit.el.matches('.combo')) { const inp = hit.el.querySelector('input'); inp.focus(); return; }
    hit.el.click();
    peek(false);
  }
  ui.barKeys = {
    active: () => !!el && isRibbon(),
    byKey: (ch) => { showTips('tabs'); tipKey(new KeyboardEvent('keydown', { key: ch })); return true; },
    index: () => showTips('tabs'),
  };

  function build() {
    spec = L.ribbonSpec();
    const qat = h('div', { class: 'rb-qat', role: 'toolbar', 'aria-label': 'Quick Access Toolbar' }, ...(spec.qat || []).map((x) => item(typeof x === 'string' ? { cmd: x, size: 'small' } : Object.assign({ size: 'small' }, x)).el));
    tabsEl = h('div', { class: 'rb-tabs', role: 'tablist' });
    panelEl = h('div', { class: 'rb-panel' });
    const pin = h('button', { class: 'rb-pin', type: 'button', 'aria-label': 'Collapse the Ribbon (Ctrl+F1)', 'data-tip': 'Collapse the Ribbon (Ctrl+F1)', html: '<svg width="10" height="6" viewBox="0 0 10 6" aria-hidden="true"><path d="M1 5l4-4 4 4" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>' });
    pin.addEventListener('click', () => RB.collapse());
    el = h('div', { class: 'ribbon' }, h('div', { class: 'rb-bar' }, qat, tabsEl, searchBox()), h('div', { class: 'rb-panelwrap' }, panelEl, pin));
    L.$('.cmdbars').prepend(el);
    drawTabs();
    if (L.store.get('ribbonCollapsed', false)) el.classList.add('collapsed');
    RB.select(cur);
    L.bus.on('ui-refresh', updateContext);
    window.addEventListener('resize', L.rafThrottle(fit));
    /* a click on a command closes a panel opened over the document */
    panelEl.addEventListener('click', (e) => { if (el.classList.contains('peek') && e.target.closest('.tb-btn:not(.tb-drop)')) setTimeout(() => peek(false), 0); });
    document.addEventListener('keydown', (e) => {
      if (!isRibbon()) return;
      if (e.key === 'F1' && e.ctrlKey && !e.altKey && !e.shiftKey) { e.preventDefault(); e.stopPropagation(); RB.collapse(); }
    }, true);
  }
  /** after the app's toolbars exist; the ribbon is built the first time a ribbon look is in use */
  RB.init = function () {
    if (!L.ribbonSpec) return;
    if (isRibbon()) build();
    L.onLook(() => { if (isRibbon() && !el) build(); else if (isRibbon()) fit(); clearTips(); ui.refresh(); });
  };
  RB.built = () => !!el;
})(typeof window !== 'undefined' ? window : globalThis);
