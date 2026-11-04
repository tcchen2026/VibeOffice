/* VibeOffice — UI toolkit shared by Quire, Ledger and Lectern: commands, menus, toolbars, dialogs, pickers.
 * Styled after the Office 2003 "Luna" command bars.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const { h } = L;
  const ui = (L.ui = {});

  /* What the app's editor does around commands, toolbar boxes and dialogs; each app fills these in (app.js).
   * beforeExec: bring the selection model up to date; refocus: give the keyboard back to the document;
   * commitEdit: finish an edit in progress before a toolbar box takes the focus; ownsText(el): el is the document's
   * own editable text, which refocus gives back with the app's selection (other editable text gets its own back). */
  ui.hooks = { beforeExec() {}, refocus() {}, commitEdit() {}, ownsText() { return false; } };
  /* Theme colours for the colour menus (Lectern): schemeColors(design) -> [{ ref, hex, tip }], currentDesign() */
  ui.schemeColors = null;
  ui.currentDesign = null;

  /* ---------- command registry ---------- */
  /* what the user is doing, for whoever follows along (the assistant, common/scout/assistant.js), as events on
   * window: vo-command { id, result } after a command ran, vo-busy { on, text } when work that keeps the user
   * waiting starts or ends, vo-message { icon } when a message box opens */
  const tell = (type, detail) => { if (typeof window !== 'undefined' && window.dispatchEvent) window.dispatchEvent(new CustomEvent(type, { detail })); };

  ui.cmds = {};
  ui.cmd = (id, def) => { const o = { id }; Object.defineProperties(o, Object.getOwnPropertyDescriptors(def)); ui.cmds[id] = o; return o; };
  ui.exec = (id, arg, ev) => {
    const c = ui.cmds[id];
    if (!c) { console.warn('unknown command', id); return; }
    if (c.enabled && !c.enabled()) return;
    ui.hooks.beforeExec();
    let r;
    try { r = c.run(arg, ev); } catch (e) { console.error(e); ui.msg(t('That action could not be completed: {0}', e && e.message ? e.message : e), { icon: 'warn' }); }
    ui.refresh();
    tell('vo-command', { id, result: r });
    return r;
  };
  ui.enabled = (id) => { const c = ui.cmds[id]; return !!c && (!c.enabled || !!c.enabled()); };
  ui.checked = (id) => { const c = ui.cmds[id]; return !!c && !!c.checked && !!c.checked(); };
  /* UI text is English in the code and shown in the UI's language (common/i18n.js; docs/i18n.md): labels through
     stripAmp / ampHTML, the rest through t */
  const t = (s, ...args) => (L.t ? L.t(s, ...args) : s);
  /* "&x" marks the access key; "&&" and "& " stay a literal ampersand; "|chart" after the text is its context
     (common/i18n.js), not shown */
  const plainAmp = (s) => (root.VO && root.VO.i18n ? root.VO.i18n.plain(String(s || '')) : String(s || '')).replace(/&&/g, '\u0000').replace(/&(\S)/g, '$1').replace(/\u0000/g, '&');
  /** a label as shown, without its access key mark */
  const stripAmp = (s) => plainAmp(t(s));
  /** text as HTML, its access key underlined */
  const ampMark = (s) => L.esc(s || '').replace(/&amp;&amp;/g, '\u0000').replace(/&amp;(\S)/g, '<u>$1</u>').replace(/\u0000/g, '&amp;');
  /** a label as shown, its access key underlined */
  const ampHTML = (s) => ampMark(t(s));
  const keyOf = (s) => { const m = /&([^&\s])/.exec(s || ''); return m ? m[1].toLowerCase() : null; };
  /** the access key of a label as shown */
  const accessKey = (s) => keyOf(t(s));
  /** a shortcut as shown: its key names in the UI's language (Ctrl+Shift+F12 is Strg+Umschalt+F12 in German), each
      in the context "key" (a key's Home, not a Home button's). On a Mac, where ⌘ does what Ctrl does, the modifiers
      are Apple's symbols in Apple's order, without "+": Ctrl+Shift+S is ⇧⌘S */
  const MAC_MODS = [['Alt', '⌥'], ['Shift', '⇧'], ['Ctrl', '⌘']];
  const keyName = (p) => ({ Ctrl: t('Ctrl|key'), Shift: t('Shift|key'), Alt: t('Alt|key'), Del: t('Del|key'), Ins: t('Ins|key'), Esc: t('Esc|key'), Home: t('Home|key'), End: t('End|key'), Enter: t('Enter|key'), Tab: t('Tab|key'), Space: t('Space|key'), Backspace: t('Backspace|key'), 'Page Up': t('Page Up|key'), 'Page Down': t('Page Down|key'), PgUp: t('PgUp|key'), PgDn: t('PgDn|key') })[p] || p;
  const keyText = (k) => {
    if (!k) return '';
    const parts = String(k).split('+');
    if (!L.isMac) return parts.map(keyName).join('+');
    return MAC_MODS.filter(([m]) => parts.includes(m)).map(([, s]) => s).join('') + parts.filter((p) => !MAC_MODS.some(([m]) => m === p)).map(keyName).join('+');
  };
  ui.stripAmp = stripAmp;
  ui.plainAmp = plainAmp;
  ui.ampHTML = ampHTML;
  ui.accessKey = accessKey;
  ui.keyText = keyText;

  /* ---------- tooltips ---------- */
  let tipEl = null, tipTimer = null, tipTarget = null;
  function showTip(target, x, y) {
    const t = target.getAttribute('data-tip');
    if (!t) return;
    if (!tipEl) { tipEl = h('div', { class: 'tooltip', role: 'tooltip' }); document.body.appendChild(tipEl); }
    const key = target.getAttribute('data-tip-key');
    tipEl.textContent = (target.closest('[translate="no"]') ? t : (L.t ? L.t(t) : t)) + (key && !document.body.classList.contains('notipkeys') ? ` (${key})` : '');   // data (a font's name) stays as it is
    tipEl.hidden = false;
    const r = tipEl.getBoundingClientRect();
    let left = x + 2, top = y + 20;
    if (left + r.width > window.innerWidth - 4) left = window.innerWidth - r.width - 4;
    if (top + r.height > window.innerHeight - 4) top = y - r.height - 6;
    tipEl.style.left = Math.max(2, left) + 'px';
    tipEl.style.top = Math.max(2, top) + 'px';
  }
  function hideTip() { clearTimeout(tipTimer); tipTarget = null; if (tipEl) tipEl.hidden = true; }
  if (typeof document !== 'undefined') {
    document.addEventListener('pointerover', (e) => {
      const t = e.target.closest && e.target.closest('[data-tip]');
      if (t === tipTarget) return;
      hideTip();
      if (!t || e.pointerType === 'touch') return;
      tipTarget = t;
      const x = e.clientX, y = e.clientY;
      tipTimer = setTimeout(() => showTip(t, x, y), 650);
    });
    document.addEventListener('pointerdown', hideTip, true);
    document.addEventListener('keydown', hideTip, true);
  }
  ui.hideTip = hideTip;

  /* ---------- menus ---------- */
  const openStack = []; /* stack of open menu panels */
  let menuBarState = null; /* {bar, index} */
  function resolveItem(it) {
    if (typeof it === 'string') {
      if (it === '-' || it === '|') return { sep: true };
      const c = ui.cmds[it];
      if (!c) return { label: it, disabled: true };
      return { cmd: it, label: c.menuLabel || c.label, icon: c.icon, key: c.key, enabled: c.enabled, checked: c.checked, radio: c.radio };
    }
    if (it.cmd && ui.cmds[it.cmd]) { const c = ui.cmds[it.cmd]; return Object.assign({ label: c.menuLabel || c.label, icon: c.icon, key: c.key, enabled: c.enabled, checked: c.checked, radio: c.radio }, it); }
    return it;
  }
  ui.closeMenus = (keepDepth) => {
    keepDepth = keepDepth || 0;
    while (openStack.length > keepDepth) {
      const m = openStack.pop();
      m.el.remove();
      if (m.opts && m.opts.onClose) m.opts.onClose();
    }
    if (!keepDepth && menuBarState) {
      L.$$('.mb-item.open', menuBarState.bar).forEach((x) => x.classList.remove('open'));
      menuBarState = null;
    }
  };
  /**
   * Open a dropdown menu. items: array of item specs or a function returning them.
   * at: {x,y} or DOMRect (opens below) ; opts: {depth, onClose, fromBar, side}
   */
  ui.openMenu = function (items, at, opts) {
    opts = opts || {};
    const depth = opts.depth || 0;
    ui.closeMenus(depth);
    const list = (typeof items === 'function' ? items() : items).map(resolveItem).filter((x) => x && !x.hidden);
    /* a menu from inside a dialog (a drop-down button there) goes above the dialogs */
    const el = h('div', { class: 'menu' + (opts.cls ? ' ' + opts.cls : '') + (ui.dialogOpen() ? ' over-dlg' : ''), role: 'menu', tabindex: '-1' });
    const rows = [];
    list.forEach((it) => {
      if (it.sep) { el.appendChild(h('div', { class: 'm-sep' })); return; }
      if (it.heading) { el.appendChild(h('div', { class: 'm-head', text: it.heading })); return; }
      if (it.custom) { const c = it.custom(() => ui.closeMenus()); el.appendChild(c); return; }
      const disabled = it.disabled || (it.enabled && !it.enabled());
      const checked = it.checked && it.checked();
      const row = h('div', { class: 'm-item' + (disabled ? ' dis' : '') + (checked ? ' chk' : '') + (it.sub ? ' has-sub' : '') + (it.bold ? ' bold' : ''), role: 'menuitem', 'aria-disabled': disabled ? 'true' : null, title: it.tip || null });
      const icoHTML = it.html ? it.html : it.icon ? L.icons.get(it.icon) : checked ? (it.radio ? '<span class="m-radio">●</span>' : '<span class="m-check">✓</span>') : '';
      row.appendChild(h('span', { class: 'm-ico' + (checked && it.icon ? ' on' : ''), html: icoHTML }));
      /* it.raw: data shown as it is (a font's name in a font list, a file's in Recent Files), not UI text; its
         access key, if any, is the label's own */
      row.appendChild(h('span', { class: 'm-lbl', html: it.raw ? ampMark(it.label) : ampHTML(it.label), translate: it.raw ? 'no' : null }));
      row.appendChild(h('span', { class: 'm-key', text: it.raw ? it.key || '' : keyText(it.key), translate: it.raw ? 'no' : null }));
      row.appendChild(h('span', { class: 'm-sub', html: it.sub ? '▶' : '' }));
      row._it = it;
      row._mn = it.raw ? keyOf(it.label) : accessKey(it.label);
      row.addEventListener('pointerenter', () => { setActive(row); if (it.sub && !disabled) openSub(row); else ui.closeMenus(depth + 1); });
      row.addEventListener('click', (e) => { e.stopPropagation(); activate(row); });
      rows.push(row);
      el.appendChild(row);
    });
    let active = null;
    function setActive(r) { if (active) active.classList.remove('act'); active = r; if (r) r.classList.add('act'); }
    function openSub(row) {
      const r = row.getBoundingClientRect();
      ui.openMenu(row._it.sub, { x: r.right - 3, y: r.top - 3, side: r.left }, { depth: depth + 1 });
    }
    function activate(row) {
      const it = row._it;
      if (it.disabled || (it.enabled && !it.enabled())) return;
      if (it.sub) { openSub(row); return; }
      ui.closeMenus();
      if (it.run) it.run();
      else if (it.cmd) ui.exec(it.cmd, it.arg);
      if (!ui.dialogOpen() && keyboardLost()) ui.hooks.refocus();
    }
    el._keys = (e) => {
      const idx = rows.indexOf(active);
      const usable = rows.filter((r) => !r.classList.contains('dis'));
      if (e.key === 'ArrowDown') { const n = usable[(usable.indexOf(active) + 1) % usable.length]; setActive(n); return true; }
      if (e.key === 'ArrowUp') { const i = usable.indexOf(active); setActive(usable[(i - 1 + usable.length) % usable.length]); return true; }
      if (e.key === 'ArrowRight' && active && active._it.sub) { openSub(active); const sub = openStack[openStack.length - 1]; if (sub && sub.first) sub.first(); return true; }
      if (e.key === 'Enter' || e.key === ' ') { if (active) activate(active); return true; }
      if (e.key.length === 1) {
        const k = e.key.toLowerCase();
        const hit = usable.find((r) => r._mn === k);
        if (hit) { setActive(hit); activate(hit); return true; }
      }
      void idx;
      return false;
    };
    (document.fullscreenElement || document.body).appendChild(el);   // over a slide show, as dialogs
    /* position */
    const r = el.getBoundingClientRect();
    let x, y;
    if (at instanceof DOMRect || (at && at.bottom != null && at.left != null && at.x == null)) { x = at.left; y = at.bottom; }
    else if (at && at.bottom != null) { x = at.left; y = at.bottom; }
    else { x = at.x; y = at.y; }
    if (x + r.width > window.innerWidth - 2) x = at && at.side != null ? at.side - r.width + 3 : window.innerWidth - r.width - 2;
    if (y + r.height > window.innerHeight - 2) y = Math.max(2, window.innerHeight - r.height - 2);
    el.style.left = Math.max(2, x) + 'px';
    el.style.top = Math.max(2, y) + 'px';
    const entry = { el, opts, first: () => setActive(rows.find((r2) => !r2.classList.contains('dis'))) };
    openStack.push(entry);
    if (opts.selectFirst) entry.first();
    return el;
  };
  ui.contextMenu = (e, items) => { e.preventDefault(); ui.openMenu(items, { x: e.clientX, y: e.clientY }); };
  if (typeof document !== 'undefined') {
    document.addEventListener('pointerdown', (e) => {
      if (!openStack.length) return;
      if (e.target.closest('.menu') || e.target.closest('.mb-item')) return;
      ui.closeMenus();
    }, true);
    window.addEventListener('blur', () => ui.closeMenus());
    window.addEventListener('resize', () => ui.closeMenus());
    /* The keyboard stays with the document when chrome is pressed: a toolbar's, the ribbon's or a pane's background, a
       grip, the status bar, a tab strip, a menu, the menu bar or a button outside a dialog. The browser would give it
       to <body> or to what was pressed, and the arrows, Delete and typing would stop working on the document until it
       is clicked. A control of its own (a box, a list, the document itself, anything in a dialog) still takes it. */
    document.addEventListener('mousedown', (e) => {
      const t = e.target;
      if (!t || !t.closest || t.closest('.dlg')) return;
      const f = t.closest('input, select, textarea, [contenteditable=""], [contenteditable="true"], a[href], label, [tabindex], button');
      if (f && !f.matches('button, .menu, .mb-item')) return;
      e.preventDefault();
    }, true);
  }
  /** after chrome that took the keyboard is done with it (a menu command, a dialog), the document gets it back */
  const keyboardLost = () => !document.activeElement || document.activeElement === document.body || !document.activeElement.isConnected;
  ui.menuOpen = () => openStack.length > 0;
  /* menu keyboard routing; returns true if handled */
  ui.menuKey = (e) => {
    if (!openStack.length) return false;
    const top = openStack[openStack.length - 1];
    if (e.key === 'Escape') {
      if (openStack.length > 1) ui.closeMenus(openStack.length - 1); else ui.closeMenus();
      return true;
    }
    if (e.key === 'ArrowLeft') {
      if (openStack.length > 1) { ui.closeMenus(openStack.length - 1); return true; }
      if (menuBarState) { menuBarState.step(-1); return true; }
    }
    if (e.key === 'ArrowRight' && menuBarState && !(top.el._keys && top.el.querySelector('.m-item.act.has-sub'))) { menuBarState.step(1); return true; }
    if (top.el._keys && top.el._keys(e)) return true;
    return true; /* swallow keys while a menu is open */
  };

  /* the keyboard of another command-bar layout (the ribbon, common/ribbon.js) while it is shown instead of the menus:
     { active(), byKey(ch) → handled, index(i) } */
  ui.barKeys = null;
  /** Menu bar: menus = [{label, items}] */
  ui.menuBar = function (container, menus) {
    const bar = h('div', { class: 'menubar', role: 'menubar' });
    const els = [];
    menus.forEach((m, i) => {
      const it = h('div', { class: 'mb-item', role: 'menuitem', tabindex: '-1', html: ampHTML(m.label) });
      it.addEventListener('pointerdown', (e) => { e.preventDefault(); if (it.classList.contains('open')) ui.closeMenus(); else open(i); });
      it.addEventListener('pointerenter', () => { if (menuBarState && menuBarState.index !== i) open(i); });
      els.push(it);
      bar.appendChild(it);
    });
    function open(i, selectFirst) {
      ui.closeMenus();
      const it = els[i];
      it.classList.add('open');
      const r = it.getBoundingClientRect();
      ui.openMenu(menus[i].items, { left: r.left, bottom: r.bottom - 1 }, { cls: 'from-bar', selectFirst });
      menuBarState = { bar, index: i, step: (d) => open((i + d + els.length) % els.length, true) };
    }
    bar.openByKey = (ch) => {
      if (ui.barKeys && ui.barKeys.active()) return ui.barKeys.byKey(ch);
      const i = menus.findIndex((m) => accessKey(m.label) === ch.toLowerCase());
      if (i >= 0) { open(i, true); return true; }
      return false;
    };
    bar.openIndex = (i) => { if (ui.barKeys && ui.barKeys.active()) ui.barKeys.index(i); else open(i, true); };
    bar.menus = ui.menus = menus;   // also read by Search commands (common/search.js)
    container.appendChild(bar);
    return bar;
  };

  /* ---------- toolbars ---------- */
  const tbButtons = [];
  ui.tbButton = function (id, opts) {
    opts = opts || {};
    const c = ui.cmds[id] || {};
    const label = opts.label != null ? opts.label : c.tbLabel;
    const b = h('button', { class: 'tb-btn' + (label ? ' with-label' : ''), type: 'button', 'data-cmd': id, 'aria-label': stripAmp(c.label || id), 'data-tip': stripAmp(c.tip || c.label || id), 'data-tip-key': c.key ? keyText(c.key) : null });
    b.innerHTML = (c.icon || opts.icon ? L.icons.get(opts.icon || c.icon) : '') + (label ? `<span class="tb-lbl">${ampHTML(label)}</span>` : '');
    b.addEventListener('pointerdown', (e) => e.preventDefault()); /* keep the focus and selection where they are */
    b.addEventListener('click', (e) => ui.exec(id, opts.arg, e));
    tbButtons.push(b);
    return b;
  };
  /** split button: main action + dropdown arrow */
  ui.tbSplit = function (id, menuFn, opts) {
    opts = opts || {};
    const wrap = h('div', { class: 'tb-split' });
    const main = ui.tbButton(id, opts);
    const dd = h('button', { class: 'tb-dd', type: 'button', 'aria-label': 'More options', 'data-tip': (ui.cmds[id] && stripAmp(ui.cmds[id].label)) || '' });
    dd.innerHTML = '<span class="dd-arrow"></span>';
    dd.addEventListener('pointerdown', (e) => e.preventDefault());
    dd.addEventListener('click', () => { const r = wrap.getBoundingClientRect(); menuFn({ left: r.left, bottom: r.bottom }, wrap); });
    wrap.append(main, dd);
    wrap.dataset.cmd = id;
    return wrap;
  };
  /** dropdown button (icon/label + arrow), opens a menu */
  ui.tbDrop = function (label, icon, menuFn, tip) {
    const b = h('button', { class: 'tb-btn tb-drop' + (label ? ' with-label' : ''), type: 'button', 'data-tip': tip || stripAmp(label), 'aria-label': tip || stripAmp(label) });
    b.innerHTML = (icon ? L.icons.get(icon) : '') + (label ? `<span class="tb-lbl">${ampHTML(label)}</span>` : '') + '<span class="dd-arrow"></span>';
    b.addEventListener('pointerdown', (e) => e.preventDefault());
    b.addEventListener('click', () => { const r = b.getBoundingClientRect(); menuFn({ left: r.left, bottom: r.bottom }, b); });
    return b;
  };
  /** editable combo box (font name / size / zoom); o.options() may include { heading } rows, { sep }, { value, label }
   * and { label, run } actions (the list reopens after the action);
   * o.note(v): grey text beside an option in the list, or { text, tip } with a tooltip for its row */
  ui.combo = function (o) {
    const wrap = h('div', { class: 'combo', style: `width:${o.width || 120}px`, 'data-tip': o.tip || '' });
    const inp = h('input', { type: 'text', class: 'combo-in', id: o.id, 'aria-label': o.tip || o.id, autocomplete: 'off', spellcheck: 'false' });
    const btn = h('button', { class: 'combo-btn', type: 'button', tabindex: '-1', 'aria-label': 'Open list', html: '<span class="dd-arrow"></span>' });
    wrap.append(inp, btn);
    const commit = () => { o.onChange(inp.value); inp.blur(); };
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); ui.hooks.refocus(); }
      if (e.key === 'Escape') { e.preventDefault(); inp.value = o.value(); inp.blur(); ui.hooks.refocus(); }
      e.stopPropagation();
    });
    inp.addEventListener('focus', () => { inp.select(); wrap.classList.add('focus'); ui.hooks.commitEdit(); });
    inp.addEventListener('blur', () => wrap.classList.remove('focus'));
    btn.addEventListener('pointerdown', (e) => e.preventDefault());
    btn.addEventListener('click', () => {
      ui.hooks.commitEdit();
      const r = wrap.getBoundingClientRect();
      // options: strings, { heading } section rows, or { value, label }
      ui.openMenu(o.options().map((x) => {
        if (x && x.heading) return { heading: x.heading };
        if (x && x.sep) return { sep: true };
        if (x && x.run) return { label: x.label, run: async () => { await x.run(); btn.click(); } };   // an action, then the list again
        const v = x && typeof x === 'object' ? x.value : x;
        const note = o.note ? o.note(v) : '';
        return {
          label: String(x && x.label || v).replace(/&/g, '&&'), value: v, raw: true,
          html: '', run: () => { inp.value = v; o.onChange(String(v)); ui.hooks.refocus(); },
          checked: () => String(v) === String(o.value()),
          key: note && typeof note === 'object' ? note.text : note, tip: note && note.tip,
        };
      }), { left: r.left, bottom: r.bottom }, { cls: 'combo-list' + (o.listCls ? ' ' + o.listCls : '') });
      if (o.preview) {
        const menu = document.querySelector('.menu.combo-list');
        // each row in its own font once scrolled into view: a long list (the computer's fonts) opens at once
        const show = (row) => { const st = o.preview(row._it.value); if (st) L.$('.m-lbl', row).style.cssText = st; };
        const io = menu && 'IntersectionObserver' in window ? new IntersectionObserver((es) => es.forEach((e) => { if (e.isIntersecting) { show(e.target); io.unobserve(e.target); } }), { root: menu }) : null;
        if (menu) L.$$('.m-item', menu).forEach((row) => row._it.value == null ? null : io ? io.observe(row) : show(row));
      }
      const m = document.querySelector('.menu.combo-list .m-item.chk');
      if (m) m.scrollIntoView({ block: 'nearest' });
    });
    wrap._refresh = () => { if (document.activeElement !== inp) inp.value = o.value(); const en = o.enabled ? o.enabled() : true; inp.disabled = !en; wrap.classList.toggle('dis', !en); };
    tbButtons.push(wrap);
    return wrap;
  };
  /** Build a toolbar row */
  ui.toolbar = function (id, title, items) {
    const tb = h('div', { class: 'toolbar', id: 'tb-' + id, role: 'toolbar', 'aria-label': title });
    tb.appendChild(h('span', { class: 'tb-grip', 'aria-hidden': 'true' }));
    for (const it of items) {
      if (it === '|') tb.appendChild(h('span', { class: 'tb-sep' }));
      else if (typeof it === 'string') tb.appendChild(ui.tbButton(it));
      else if (it instanceof Node) tb.appendChild(it);
    }
    tb.appendChild(h('span', { class: 'tb-end' }));
    toolbarOptions(tb);
    return tb;
  };
  /* Toolbar Options (») at the end of a toolbar cut short by the window or by a shared row, as in Office 2003: it shows
     the controls that don't fit in a pop-up, the controls themselves (moved there while it is open), so their menus,
     lists and states work as on the toolbar */
  function toolbarOptions(tb) {
    const more = h('button', { class: 'tb-more', type: 'button', tabindex: '-1', 'aria-label': 'Toolbar Options', 'data-tip': 'Toolbar Options', 'aria-haspopup': 'true', hidden: true, html: '<span aria-hidden="true">»</span>' });
    const end = tb.lastChild;
    let pop = null;
    tb.appendChild(more);
    const update = () => {
      if (pop) return;
      const need = tb.scrollWidth - (more.hidden ? 0 : more.offsetWidth);
      more.hidden = !(need > tb.clientWidth + 1);
    };
    if (typeof ResizeObserver !== 'undefined') new ResizeObserver(update).observe(tb);
    const close = () => {
      if (!pop) return;
      for (const n of [...pop.children]) tb.insertBefore(n, end);
      pop.remove(); pop = null;
      tb.style.width = tb.style.flex = '';
      document.removeEventListener('pointerdown', outside, true);
      document.removeEventListener('keydown', esc, true);
      more.classList.remove('open');
      update();
    };
    const outside = (e) => { if (!pop.contains(e.target) && !more.contains(e.target) && !e.target.closest('.menu')) close(); };
    const esc = (e) => { if (e.key === 'Escape' && !ui.menuOpen()) { e.preventDefault(); e.stopPropagation(); close(); } };
    more.addEventListener('pointerdown', (e) => e.preventDefault());
    more.addEventListener('click', () => {
      if (pop) { close(); return; }
      const edge = more.getBoundingClientRect().left;
      const kids = [...tb.children];
      const from = kids.findIndex((n) => n !== more && n !== end && n.getClientRects().length && n.getBoundingClientRect().right > edge + 1);
      if (from < 0) return;
      /* the toolbar keeps its width while its controls are away */
      tb.style.width = tb.getBoundingClientRect().width + 'px'; tb.style.flex = 'none';
      pop = h('div', { class: 'toolbar tb-pop', role: 'toolbar', 'aria-label': tb.getAttribute('aria-label') });
      for (const n of kids.slice(from)) if (n !== more && n !== end) pop.appendChild(n);
      document.body.appendChild(pop);
      const r = more.getBoundingClientRect(), pr = pop.getBoundingClientRect();
      pop.style.left = Math.max(2, Math.min(r.right - pr.width, window.innerWidth - pr.width - 2)) + 'px';
      pop.style.top = r.bottom + 1 + 'px';
      more.classList.add('open');
      document.addEventListener('pointerdown', outside, true);
      document.addEventListener('keydown', esc, true);
      /* a button that runs its command closes the pop-up; a drop-down's arrow or list keeps it until the choice is made */
      pop.addEventListener('click', (e) => { const b = e.target.closest('.tb-btn'); if (b && !b.classList.contains('tb-drop')) setTimeout(close); });
      ui.refresh();
    });
    window.addEventListener('resize', close);
  }
  /** the Undo or Redo drop-down: labels are the steps, newest first; choosing one undoes (or redoes) it and every step
   * after it, as in Office */
  ui.historyMenu = function (r, kind, labels) {
    if (!labels || !labels.length) return;
    ui.openMenu([{ custom: (close) => {
      const el = h('div', { class: 'style-list hist-list' });
      const actions = (n) => (kind === 'undo' ? (n === 1 ? L.t('Undo {0} Action', n) : L.t('Undo {0} Actions', n)) : n === 1 ? L.t('Redo {0} Action', n) : L.t('Redo {0} Actions', n));
      const lbl = h('div', { class: 'tgp-label', translate: 'no', text: actions(1) });
      const btns = labels.slice(0, 40).map((label, i) => {
        const b = h('button', { type: 'button', text: label });
        b.addEventListener('pointerenter', () => { btns.forEach((x, k) => x.classList.toggle('on', k <= i)); lbl.textContent = actions(i + 1); });
        b.addEventListener('click', () => { close(); for (let k = 0; k <= i; k++) ui.exec(kind); });
        el.appendChild(b);
        return b;
      });
      btns[0].classList.add('on');
      return h('div', { class: 'hist-menu' }, el, lbl);
    } }], r);
  };
  /* ---------- the status bar's right end: view buttons and the zoom slider, the same in every app ---------- */
  /** a group of status-bar buttons (the view buttons): toolbar items, so a drop-down (tbDrop) fits too */
  ui.statusGroup = function (items) {
    const g = h('div', { class: 'sb-group', role: 'group' });
    for (const it of items) {
      if (it === '|') g.appendChild(h('span', { class: 'sb-gap' }));
      else if (typeof it === 'string') g.appendChild(ui.tbButton(it));
      else if (it instanceof Node) g.appendChild(it);
    }
    return g;
  };
  /** The zoom slider: −, a track with 100 % at its centre (min…100 % on the left half, 100 %…max on the right, as in
   * Office), +, and the percentage, which opens the Zoom dialog. o = { min, max, value: () => percent,
   * set(percent, final): final is false while the thumb is dragged, true when it is let go and for −/+;
   * dialog: command id; fit: command id of a fit-to-window button; enabled(): false greys it out }.
   * Like a toolbar button it never takes the keyboard: pressing it leaves typing, the selection and shortcuts with
   * the document. Esc during a drag puts the zoom back. */
  ui.zoomSlider = function (o) {
    const min = o.min || 10, max = o.max || 400;
    const toPos = (p) => (p <= 100 ? (Math.max(min, p) - min) / (100 - min) / 2 : 0.5 + (Math.min(max, p) - 100) / (max - 100) / 2);
    const toPct = (x) => { if (Math.abs(x - 0.5) < 0.025) return 100; return Math.round(x < 0.5 ? min + (100 - min) * x * 2 : 100 + (max - 100) * (x - 0.5) * 2); };
    const step = (dir) => { const p = o.value(); const n = dir > 0 ? Math.floor(p / 10) * 10 + 10 : Math.ceil(p / 10) * 10 - 10; o.set(L.clamp(n, min, max), true); ui.refresh(); };
    const glyph = (plus) => `<svg width="9" height="9" viewBox="0 0 9 9" aria-hidden="true"><path d="M0 4.5h9${plus ? 'M4.5 0v9' : ''}" stroke="currentColor" stroke-width="1.4"/></svg>`;
    const btn = (cls, tip, html, run) => {
      const b = h('button', { class: cls, type: 'button', 'aria-label': tip, 'data-tip': tip });
      b.innerHTML = html;
      b.addEventListener('pointerdown', (e) => e.preventDefault());
      b.addEventListener('click', run);
      return b;
    };
    const out = btn('zs-step', L.N('Zoom Out'), glyph(false), () => step(-1));
    const inn = btn('zs-step', L.N('Zoom In'), glyph(true), () => step(1));
    const thumb = h('span', { class: 'zs-thumb' });
    const track = h('div', { class: 'zs-track', role: 'slider', 'aria-label': 'Zoom', 'aria-valuemin': String(min), 'aria-valuemax': String(max) }, thumb);
    const pct = btn('zs-pct', L.N('Zoom level'), '', () => ui.exec(o.dialog));
    const wrap = h('div', { class: 'sb-zoom' }, out, track, inn, pct);
    if (o.fit) wrap.appendChild(ui.tbButton(o.fit));
    const show = (p) => { thumb.style.left = toPos(p) * 100 + '%'; pct.textContent = p + '%'; track.setAttribute('aria-valuenow', String(p)); track.setAttribute('aria-valuetext', p + '%'); };
    let drag = null;
    const at = (e) => { const r = track.getBoundingClientRect(); return toPct(L.clamp((e.clientX - r.left) / r.width, 0, 1)); };
    const live = L.rafThrottle(() => { if (drag) o.set(drag.pct, false); });
    const end = (keep) => {
      if (!drag) return;
      const d = drag; drag = null;
      document.removeEventListener('keydown', onKey, true);
      wrap.classList.remove('dragging');
      o.set(keep ? d.pct : d.start, true);
      ui.refresh();
    };
    const onKey = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); end(false); } };
    track.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      if (e.button !== 0 || wrap.classList.contains('dis')) return;
      if (ui.hooks.commitEdit) ui.hooks.commitEdit();
      drag = { start: o.value(), pct: at(e) };
      track.setPointerCapture(e.pointerId);
      wrap.classList.add('dragging');
      document.addEventListener('keydown', onKey, true);
      show(drag.pct); live();
    });
    track.addEventListener('pointermove', (e) => { if (!drag) return; drag.pct = at(e); show(drag.pct); live(); });
    track.addEventListener('pointerup', () => end(true));
    track.addEventListener('lostpointercapture', () => end(true));
    wrap._refresh = () => {
      const en = o.enabled ? o.enabled() : true;
      wrap.classList.toggle('dis', !en);
      for (const b of [out, inn, pct]) b.disabled = !en;
      track.setAttribute('aria-disabled', String(!en));
      if (!drag) show(o.value());
    };
    tbButtons.push(wrap);
    wrap._refresh();
    return wrap;
  };
  ui.refresh = L.rafThrottle(() => {
    for (const b of tbButtons) {
      if (!b.isConnected) continue;
      if (b._refresh) { b._refresh(); continue; }
      const id = b.dataset.cmd;
      const c = ui.cmds[id];
      if (!c) continue;
      const en = !c.enabled || c.enabled();
      b.disabled = !en;
      b.classList.toggle('dis', !en);
      b.classList.toggle('on', !!(c.checked && c.checked()));
      b.setAttribute('aria-pressed', c.checked ? String(!!c.checked()) : null);
      if (c.swatch) { const sw = c.swatch(); if (sw) b.style.setProperty('--icon-swatch', sw); }
    }
    L.bus.emit('ui-refresh');
  });

  /* ---------- dialogs ---------- */
  const dlgStack = [];
  /* a tabbed dialog keeps one size whichever tab is shown, as Office's property sheets: its pages share the tallest
     one's height, so OK and Cancel stay where they are */
  function fitTabs(root) {
    for (const bodies of L.$$('.tabs-body', root)) {
      const panes = [...bodies.children], was = panes.map((p) => p.hidden);
      if (panes.length < 2) continue;
      let tallest = 0;
      for (const p of panes) { panes.forEach((q) => { q.hidden = q !== p; }); tallest = Math.max(tallest, bodies.getBoundingClientRect().height); }
      panes.forEach((q, i) => { q.hidden = was[i]; });
      const cs = getComputedStyle(bodies);
      const frame = cs.boxSizing === 'border-box' ? 0 : ['paddingTop', 'paddingBottom', 'borderTopWidth', 'borderBottomWidth'].reduce((n, k) => n + (parseFloat(cs[k]) || 0), 0);
      bodies.style.minHeight = Math.ceil(tallest - frame) + 'px';
    }
  }
  ui.dialogOpen = () => dlgStack.length > 0;
  /**
   * opts: {title, body, width, buttons:[{label, primary, cancel, onClick(close, button)→false to keep open}], onClose, help}
   */
  ui.dialog = function (opts) {
    ui.closeMenus();
    hideTip();
    const prevFocus = document.activeElement;
    /* the selection in text being edited, for text the app doesn't give back itself (Lectern's speaker notes) */
    const s0 = window.getSelection();
    const prevRange = prevFocus && prevFocus.isContentEditable && s0.rangeCount && prevFocus.contains(s0.anchorNode) ? s0.getRangeAt(0).cloneRange() : null;
    const overlay = h('div', { class: 'dlg-overlay' });
    const dlg = h('div', { class: 'dlg', role: 'dialog', 'aria-modal': 'true', 'aria-label': opts.title });
    dlg.style.width = Math.min(opts.width || 380, window.innerWidth - 16) + 'px';
    const titleBar = h('div', { class: 'dlg-title' }, h('span', { class: 'dlg-ttl', text: opts.title }),
      opts.help ? h('button', { class: 'dlg-x dlg-help', type: 'button', 'aria-label': 'Help', text: '?', onclick: opts.help }) : null,
      h('button', { class: 'dlg-x', type: 'button', 'aria-label': 'Close', html: '<svg width="9" height="9" viewBox="0 0 9 9"><path d="M1 1l7 7M8 1L1 8" stroke="currentColor" stroke-width="1.5"/></svg>', onclick: () => close(null) }));
    const body = h('div', { class: 'dlg-body' });
    if (opts.body) body.appendChild(opts.body);
    const foot = h('div', { class: 'dlg-foot' });
    let result;
    const done = new Promise((res) => { result = res; });
    const btns = (opts.buttons || [{ label: 'OK', primary: true }]).map((b, i) => {
      const el = h('button', { class: 'btn' + (b.primary ? ' primary' : '') + (b.cls ? ' ' + b.cls : ''), type: 'button', html: ampHTML(b.label) });
      el.addEventListener('click', () => {
        if (b.onClick) { const r = b.onClick(close, el); if (r === false) return; }
        close(i);
      });
      foot.appendChild(el);
      return el;
    });
    dlg.append(titleBar, body, foot);
    overlay.appendChild(dlg);
    (document.fullscreenElement || document.body).appendChild(overlay);   // over a slide show, which is full screen
    fitTabs(dlg);
    /* center */
    const r = dlg.getBoundingClientRect();
    dlg.style.left = Math.max(4, (window.innerWidth - r.width) / 2) + 'px';
    dlg.style.top = Math.max(4, (window.innerHeight - r.height) / 2.4) + 'px';
    /* drag */
    titleBar.addEventListener('pointerdown', (e) => {
      if (e.target.closest('button')) return;
      const sx = e.clientX, sy = e.clientY, ox = dlg.offsetLeft, oy = dlg.offsetTop;
      titleBar.setPointerCapture(e.pointerId);
      const mv = (ev) => { dlg.style.left = L.clamp(ox + ev.clientX - sx, -r.width + 60, window.innerWidth - 60) + 'px'; dlg.style.top = L.clamp(oy + ev.clientY - sy, 0, window.innerHeight - 30) + 'px'; };
      const up = () => { titleBar.removeEventListener('pointermove', mv); titleBar.removeEventListener('pointerup', up); };
      titleBar.addEventListener('pointermove', mv);
      titleBar.addEventListener('pointerup', up);
    });
    const entry = { overlay, dlg, keys: null };
    if (opts.modeless) {
      overlay.classList.add('modeless');
      dlg.addEventListener('keydown', (e) => { entry.keys(e); e.stopPropagation(); });   // keys stay in a modeless dialog
    } else dlgStack.push(entry);
    entry.keys = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); const ci = (opts.buttons || []).findIndex((b) => b.cancel || /^cancel$|^close$/i.test(plainAmp(b.label))); if (ci >= 0) btns[ci].click(); else close(null); return true; }
      if (e.key === 'Enter' && !e.shiftKey && !(e.target && (e.target.tagName === 'TEXTAREA' || e.target.tagName === 'BUTTON' || e.target.isContentEditable || e.target.tagName === 'SELECT'))) {
        const pi = (opts.buttons || []).findIndex((b) => b.primary);
        if (pi >= 0) { e.preventDefault(); btns[pi].click(); return true; }
      }
      if (e.altKey && !e.ctrlKey && !e.metaKey && accessKeyIn(dlg, e)) { e.preventDefault(); return true; }
      if (e.key === 'Tab') {
        const f = L.$$('input,select,textarea,button,[tabindex="0"]', dlg).filter((x) => !x.disabled && x.offsetParent !== null);
        if (!f.length) return false;
        const i = f.indexOf(document.activeElement);
        if (e.shiftKey && (i <= 0)) { e.preventDefault(); f[f.length - 1].focus(); return true; }
        if (!e.shiftKey && i === f.length - 1) { e.preventDefault(); f[0].focus(); return true; }
      }
      return false;
    };
    function close(i) {
      const k = dlgStack.indexOf(entry);
      if (k >= 0) dlgStack.splice(k, 1);
      overlay.remove();
      if (opts.onClose) opts.onClose(i);
      result(i);
      /* the keyboard goes back to what had it: a box, a pane's list, the dialog under this one, the grid, text being
         edited with the selection it had (the page's own moved into the dialog); the document's own text the app gives
         back itself, with its selection (ui.hooks.ownsText); nothing (<body>, a menu gone) goes to the document */
      const live = prevFocus && prevFocus.isConnected && prevFocus !== document.body;
      if (live && !(prevFocus.isContentEditable && ui.hooks.ownsText(prevFocus))) {
        try { prevFocus.focus({ preventScroll: true }); } catch (e) { /* ignore */ }
        if (prevRange && prevRange.startContainer.isConnected && prevRange.endContainer.isConnected) { const s = window.getSelection(); s.removeAllRanges(); s.addRange(prevRange); }
      } else ui.hooks.refocus();
    }
    setTimeout(() => {
      const first = opts.focus ? (typeof opts.focus === 'string' ? L.$(opts.focus, dlg) : opts.focus) : L.$('input:not([type=checkbox]):not([type=radio]),select,textarea', body) || btns.find((b, i) => opts.buttons && opts.buttons[i] && opts.buttons[i].primary) || btns[0];
      if (first) { first.focus(); if (first.select && first.tagName === 'INPUT') first.select(); }
    }, 0);
    if (opts.onOpen) opts.onOpen(dlg);
    return { el: dlg, body, close, done, buttons: btns };
  };
  /* Alt+letter in a dialog, as in Office: the underlined letter's button is pressed, a tab chosen, a checkbox or radio
     ticked, a label's box focused; when several share the letter, each press moves to the next. The letter is the key's
     own (e.key), or its place on a US keyboard where Option or AltGr makes another character (e.code) */
  function accessKeyIn(dlg, e) {
    const want = new Set([e.key.length === 1 ? e.key.toLowerCase() : '', (/^Key([A-Z])$/.exec(e.code) || /^Digit(\d)$/.exec(e.code) || [])[1]].filter(Boolean).map((x) => x.toLowerCase()));
    const hits = L.$$('u', dlg).filter((u) => want.has(u.textContent.toLowerCase()) && u.getClientRects().length)
      .map((u) => u.closest('button, label')).filter((x) => x && !x.disabled && !(x.control && x.control.disabled));
    if (!hits.length) return false;
    const target = (x) => (x.tagName === 'BUTTON' ? x : x.control || L.$$('input,select,textarea,button,[tabindex="0"]', dlg).find((c) => (x.compareDocumentPosition(c) & 4) && !c.disabled && c.getClientRects().length));
    const at = hits.findIndex((x) => target(x) === document.activeElement);
    const x = hits[(at + 1) % hits.length], c = target(x);
    if (!c) return true;
    if (hits.length > 1) { c.focus(); return true; }
    if (c.tagName === 'BUTTON' || c.type === 'checkbox' || c.type === 'radio') { c.focus(); c.click(); }
    else { c.focus(); if (c.select && c.type !== 'file') c.select(); }
    return true;
  }
  ui.dialogKey = (e) => { const top = dlgStack[dlgStack.length - 1]; return top ? top.keys(e) || true : false; };

  /** A document's hyperlink, which the suite never opens itself: a link in a file can lead anywhere, so it shows
   *  the address with a warning and Copy, to paste into the browser's address bar (or, for mail, into a message).
   *  Every app's Open Hyperlink, Ctrl+click and slide show link come here; links within the document don't. */
  ui.linkDialog = function (url) {
    const mail = /^mailto:/i.test(url);
    const box = h('input', { type: 'text', class: 'link-url', readonly: true, value: url, spellcheck: 'false', 'aria-label': t('Address') });
    const copied = h('p', { class: 'link-copied', role: 'status' });
    const body = h('div', { class: 'col' },
      h('p', { text: t('This link goes to:') }), box,
      h('p', { text: mail ? t('To write to this address, copy it and paste it into your e-mail program.')
        : t('Links in documents can lead to harmful sites. If you trust this one, copy its address and paste it into your browser\'s address bar.') }),
      copied);
    const copy = async () => {
      try { await navigator.clipboard.writeText(url); } catch (e) { box.select(); document.execCommand('copy'); }
      copied.textContent = t('Copied.');
    };
    // Copy leaves the dialog open (false), so the address can be pasted while it shows
    const d = ui.dialog({ title: t('Hyperlink'), width: 440, body, buttons: [{ label: '&Copy Address', primary: true, onClick: () => { copy(); return false; } }, { label: 'Close' }] });
    box.focus();
    box.select();
    return d;
  };

  /** Message box. Returns Promise<buttonIndex|null> */
  ui.msg = function (text, o) {
    o = o || {};
    const icons = {
      info: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#2c64c8" stroke="#163b84"/><text x="12.5" y="24" font-family="Georgia,serif" font-size="19" font-weight="bold" fill="#fff">i</text></svg>',
      warn: '<svg width="32" height="32" viewBox="0 0 32 32"><path d="M16 2l14.5 26h-29z" fill="#f6cf3b" stroke="#a07c08"/><text x="13" y="25" font-family="Arial" font-size="17" font-weight="bold" fill="#222">!</text></svg>',
      question: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#2c64c8" stroke="#163b84"/><text x="10.5" y="23.5" font-family="Arial" font-size="18" font-weight="bold" fill="#fff">?</text></svg>',
      error: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#d33a2c" stroke="#8a1a10"/><path d="M10 10l12 12M22 10L10 22" stroke="#fff" stroke-width="3.4"/></svg>',
    };
    const body = h('div', { class: 'msgbox' }, h('div', { class: 'msg-ico', html: icons[o.icon || 'info'] }), h('div', { class: 'msg-text' }));
    const t = body.querySelector('.msg-text');
    String(L.t ? L.t(text) : text).split('\n').forEach((line, i) => { if (i) t.appendChild(h('br')); t.appendChild(document.createTextNode(line)); });
    const labels = o.buttons || ['OK'];
    tell('vo-message', { icon: o.icon || 'info' });
    /* Esc answers Cancel; No only where there is no Cancel (Yes/No), never in Yes/No/Cancel, where No discards */
    const plain = labels.map(plainAmp);
    const esc = plain.some((l) => /cancel/i.test(l)) ? plain.findIndex((l) => /cancel/i.test(l)) : plain.findIndex((l) => /^no$/i.test(l));
    const d = ui.dialog({ title: o.title || L.APP, body, width: o.width || 380, buttons: labels.map((l, i) => ({ label: l, primary: i === (o.def || 0), cancel: i === esc })) });
    return d.done;
  };
  /** Help ▸ About, the same in every app: icon, name and version, the app's own paragraphs, the link to the open
   *  source licenses, and the suite's license and font notes (as in the Start Center's About).
   *  Descriptions stay general, without naming another company's products (o: { name, text: [paragraphs] }). */
  ui.about = function (o) {
    const body = h('div', { class: 'col about' },
      h('div', { class: 'about-head' }, h('span', { html: L.icons.app ? L.icons.app(48) : '' }),
        h('div', null, h('div', { class: 'about-name', translate: 'no', text: root.VO ? root.VO.editionName(o.name) : o.name }), h('div', { text: t('Version {0}', L.VERSION) }))),
      ...o.text.map((t) => h('p', { text: t })),
      h('p', null, ...t('Made possible by open source software: {0}.').split('{0}').flatMap((part, i) => (i ? [h('a', { href: root.VO ? root.VO.LICENSES : '../oss-licenses.html', target: '_blank', rel: 'noopener',
        class: 'about-oss', text: 'Open source licenses', onclick: (e) => { if (root.VO) { e.preventDefault(); root.VO.showLicenses(); } } }), part] : [part]))),
      h('p', { text: 'Open source under the Apache License 2.0, provided as is, without warranty. VibeOffice.work is an independent project. Product names in its documentation are trademarks of their owners; VibeOffice.work is not affiliated with them.' }),
      h('p', { class: 'about-marks', text: 'Font names are used solely to identify the fonts specified in documents and to maintain interoperability with Office Open XML files and the applications that create them; such names may be trademarks of their respective owners. VibeOffice renders text with the fonts installed on your computer. Where a specified font is not installed, VibeOffice substitutes an open-source font supplied with it or another font available on your computer, and the document may therefore appear differently than in other applications. VibeOffice does not include or distribute any commercial font software.' }));
    return ui.dialog({ title: t('About {0}', o.name), width: 420, body, buttons: [{ label: 'OK', primary: true }] });
  };
  ui.prompt = function (label, value, title) {
    const inp = h('input', { type: 'text', id: 'prompt-in', value: value || '', style: 'width:100%' });
    let val = null;
    const d = ui.dialog({ title: title || L.APP, body: h('div', { class: 'col' }, h('label', { for: 'prompt-in', text: label }), inp), buttons: [{ label: 'OK', primary: true, onClick: () => { val = inp.value; } }, { label: 'Cancel' }] });
    return d.done.then(() => val);
  };

  /** Check a prepared Office save before handing it to the download host.
   * Returns entries to acknowledge after a successful save, or null on cancel.
   * Snapshot writers never call this UI or acknowledge a loss.
   */
  ui.compatibility = async function (doc, entries) {
    const pending = L.opc.noticeLosses(doc, entries);
    if (!pending.length) return [];
    // One line per message, with every place it applies to
    const groups = new Map();
    for (const e of pending) {
      if (!groups.has(e.what)) groups.set(e.what, []);
      if (e.place && !groups.get(e.what).includes(e.place)) groups.get(e.what).push(e.place);
    }
    const list = h('ul', { style: 'max-height:240px;overflow:auto;margin:8px 0;padding-left:22px' });
    for (const [what, places] of groups) list.appendChild(h('li', { style: 'margin-bottom:8px' },
      h('div', { text: what }), places.length ? h('div', { class: 'hint', text: places.length > 4 ? t('{0} and {1} more', places.slice(0, 4).join(', '), places.length - 4) : places.join(', ') }) : null));
    const body = h('div', { class: 'col' },
      h('div', { text: 'Some of this file can\'t be saved exactly as it was:' }), list,
      h('div', { text: 'Click Continue to save anyway, or Cancel to go back to your document without saving.' }));
    const result = await ui.dialog({ title: 'Compatibility Checker', width: 540, body,
      buttons: [{ label: '&Continue', primary: true }, { label: 'Cancel', cancel: true }] }).done;
    return result === 0 ? pending : null;
  };

  /** the "Password" box for opening a protected file; resolves to the password or null. wrong: after a wrong one */
  ui.password = function (fileName, wrong) {
    const inp = h('input', { type: 'password', id: 'pw-in', value: '', style: 'width:100%', autocomplete: 'off' });
    let val = null;
    const d = ui.dialog({ title: 'Password', width: 340, body: h('div', { class: 'col' }, h('div', { text: wrong ? 'The password is incorrect. Try again.' : t('\'{0}\' is protected.', fileName) }), h('label', { for: 'pw-in', text: 'Enter password to open file' }), inp), buttons: [{ label: 'OK', primary: true, onClick: () => { val = inp.value; } }, { label: 'Cancel' }] });
    setTimeout(() => inp.focus(), 0);
    return d.done.then(() => val);
  };

  /* ---------- form controls ---------- */
  let ctlId = 0;
  ui.button = (label, onClick, o) => { const b = h('button', Object.assign({ class: 'btn', type: 'button', html: ampHTML(label) }, o || {})); b.addEventListener('click', onClick); return b; };
  ui.check = (label, checked, onChange, o) => {
    const id = (o && o.id) || 'ck' + (++ctlId);
    const inp = h('input', { type: 'checkbox', id });
    inp.checked = !!checked;
    inp.addEventListener('change', () => onChange && onChange(inp.checked));
    const wrap = h('label', { class: 'ck', for: id }, inp, h('span', { html: ampHTML(label) }));
    wrap.input = inp;
    return wrap;
  };
  ui.radio = (name, label, checked, onChange, o) => {
    const id = (o && o.id) || 'rd' + (++ctlId);
    const inp = h('input', { type: 'radio', id, name });
    inp.checked = !!checked;
    inp.addEventListener('change', () => { if (inp.checked && onChange) onChange(); });
    const wrap = h('label', { class: 'ck', for: id }, inp, h('span', { html: ampHTML(label) }));
    wrap.input = inp;
    return wrap;
  };
  /** options: [value, label] pairs, whose labels are UI text, or plain values shown as they are (font names, sizes);
      o.translate: 'no' shows every label as it is */
  ui.select = (options, value, onChange, o) => {
    const sel = h('select', Object.assign({ id: (o && o.id) || 'sel' + (++ctlId) }, o || {}));
    for (const op of options) {
      const [v, l] = Array.isArray(op) ? op : [op, op];
      const opt = h('option', { value: v, text: l, translate: !Array.isArray(op) || (o && o.translate === 'no') ? 'no' : null });
      if (String(v) === String(value)) opt.selected = true;
      sel.appendChild(opt);
    }
    sel.addEventListener('change', () => onChange && onChange(sel.value));
    return sel;
  };
  /** numeric spinner with unit, e.g. {value: 1.5, unit:'"', step:0.1, min, max} */
  ui.spin = (o) => {
    const id = o.id || 'sp' + (++ctlId);
    const fmt = (v) => (o.fmt ? o.fmt(v) : L.round(v, o.dec == null ? 2 : o.dec) + (o.unit || ''));
    const inp = h('input', { type: 'text', id, class: 'spin-in', value: fmt(o.value), inputmode: 'decimal', autocomplete: 'off' });
    const parse = () => { const v = parseFloat(String(inp.value).replace(/[^0-9.\-]/g, '')); return isNaN(v) ? o.value : v; };
    const set = (v) => { v = L.clamp(v, o.min == null ? -1e9 : o.min, o.max == null ? 1e9 : o.max); o.value = v; inp.value = fmt(v); if (o.onChange) o.onChange(v); };
    const up = h('button', { type: 'button', class: 'spin-b up', tabindex: '-1', 'aria-label': 'Increase', html: '<span></span>' });
    const dn = h('button', { type: 'button', class: 'spin-b dn', tabindex: '-1', 'aria-label': 'Decrease', html: '<span></span>' });
    up.addEventListener('click', () => set(L.round(parse() + (o.step || 1), 4)));
    dn.addEventListener('click', () => set(L.round(parse() - (o.step || 1), 4)));
    inp.addEventListener('change', () => set(parse()));
    inp.addEventListener('keydown', (e) => { if (e.key === 'ArrowUp') { e.preventDefault(); up.click(); } if (e.key === 'ArrowDown') { e.preventDefault(); dn.click(); } });
    const wrap = h('span', { class: 'spin' }, inp, h('span', { class: 'spin-bs' }, up, dn));
    wrap.input = inp;
    wrap.get = () => parse();
    wrap.set = (v) => { o.value = v; inp.value = fmt(v); };
    wrap.setDisabled = (d) => { inp.disabled = d; up.disabled = d; dn.disabled = d; };
    return wrap;
  };
  /** A list box, as Office's dialogs have (Font, Size, Go To, Date and Time…). items: values, [value, label] or
   *  { value, label, note (shown at the right: a shortcut key), style, group (a heading), dim (a line that can't be
   *  chosen) }; labels are UI text shown in the UI's language unless o.translate === 'no' (data:
   *  fonts, sizes, names). o: { input (a box above that picks as you type, as Word's Font box), height, id (the
   *  input's, else the list's), cls, translate, onChange(v, i), onDbl(v, i) }. Arrow keys, Home, End, Page Up/Down
   *  and typed letters move the selection; the input's Up and Down move it too. Items are matched by value, never by
   *  the text shown. Returns the element with get(), set(v), setItems(items), index(), .input and .list */
  ui.listBox = function (items, value, o) {
    o = o || {};
    const data = o.translate === 'no';
    const val = (it) => (Array.isArray(it) ? it[0] : it && typeof it === 'object' ? it.value : it);
    const shown = (it) => { const s = String(Array.isArray(it) ? it[1] : it && typeof it === 'object' ? it.label : it); return data ? s : t(s); };
    /* a group's heading, or a dimmed line that can't be chosen ("No Suggestions") */
    const isGroup = (it) => !!(it && typeof it === 'object' && !Array.isArray(it) && (it.group || it.dim));
    const box = h('div', { class: 'listbox' + (o.cls ? ' ' + o.cls : ''), role: 'listbox', tabindex: '0', id: o.input ? null : o.id || null, style: o.height ? `height:${o.height}px` : null, translate: data ? 'no' : null });
    let list = items.slice(), cur = value;
    const label = (v) => { const it = list.find((x) => !isGroup(x) && String(val(x)) === String(v)); return it ? shown(it) : v == null ? '' : String(v); };
    const inp = o.input ? h('input', { type: 'text', class: 'listbox-in', id: o.id || 'lb' + (++ctlId), value: label(value), autocomplete: 'off', translate: data ? 'no' : null }) : null;
    const rows = () => [...box.querySelectorAll('.li')];
    const indexOf = (v) => rows().findIndex((r) => String(r._v) === String(v));
    const mark = (v, scroll) => {
      for (const r of rows()) { const on = v !== undefined && String(r._v) === String(v); r.classList.toggle('on', on); r.setAttribute('aria-selected', String(on)); if (on && scroll) r.scrollIntoView({ block: 'nearest' }); }
    };
    const pick = (v, fromInput) => { cur = v; if (inp && !fromInput) inp.value = label(v); mark(v, true); if (o.onChange) o.onChange(v, indexOf(v)); };
    const render = () => {
      L.clear(box);
      for (const it of list) {
        const el = h('div', { class: isGroup(it) ? (it.dim ? 'li-dim' : 'li-group') : 'li', role: isGroup(it) ? 'presentation' : 'option' });
        el.textContent = shown(it);
        /* a second column on the right: a shortcut key, a count */
        if (it && typeof it === 'object' && it.note) el.appendChild(h('span', { class: 'li-note', translate: 'no', text: it.note }));
        if (it && typeof it === 'object' && it.style) el.style.cssText += it.style;
        box.appendChild(el);
        if (isGroup(it)) continue;
        const v = val(it);
        el._v = v;
        el.addEventListener('click', () => pick(v));
        el.addEventListener('dblclick', () => { pick(v); if (o.onDbl) o.onDbl(v, indexOf(v)); });
      }
      mark(cur);
      const on = box.querySelector('.li.on');
      if (on) setTimeout(() => on.scrollIntoView({ block: 'nearest' }), 0);
    };
    /* the keys that move the selection: in the list all of them, from the input only Up, Down, Page Up and Page Down */
    const move = (e, all) => {
      const rs = rows();
      if (!rs.length) return false;
      let i = rs.findIndex((r) => r.classList.contains('on'));
      const page = Math.max(1, Math.floor(box.clientHeight / (rs[0].offsetHeight || 16)) - 1);
      const k = e.key;
      if (k === 'ArrowDown') i = Math.min(rs.length - 1, i + 1);
      else if (k === 'ArrowUp') i = Math.max(0, i - 1);
      else if (k === 'PageDown') i = Math.min(rs.length - 1, Math.max(0, i) + page);
      else if (k === 'PageUp') i = Math.max(0, i - page);
      else if (all && k === 'Home') i = 0;
      else if (all && k === 'End') i = rs.length - 1;
      else return false;
      e.preventDefault(); e.stopPropagation();
      pick(rs[i]._v);
      return true;
    };
    let typed = '', typedAt = 0;
    box.addEventListener('keydown', (e) => {
      if (e.altKey || e.ctrlKey || e.metaKey || move(e, true)) return;
      if (e.key.length !== 1 || e.key === ' ') return;
      /* letters typed in quick succession find the first item that starts with them */
      typed = (Date.now() - typedAt < 800 ? typed : '') + e.key.toLowerCase(); typedAt = Date.now();
      const rs = rows(), from = Math.max(0, rs.findIndex((r) => r.classList.contains('on')));
      const order = typed.length > 1 ? rs.slice(from).concat(rs.slice(0, from)) : rs.slice(from + 1).concat(rs.slice(0, from + 1));
      const hit = order.find((r) => r.textContent.toLowerCase().startsWith(typed));
      if (hit) { e.preventDefault(); e.stopPropagation(); pick(hit._v); }
    });
    if (inp) {
      inp.addEventListener('keydown', (e) => { if (!e.altKey && !e.ctrlKey && !e.metaKey) move(e, false); });
      inp.addEventListener('input', () => {
        const v = inp.value, lc = v.toLowerCase();
        const hit = list.find((x) => !isGroup(x) && shown(x).toLowerCase().startsWith(lc));
        /* a label typed in the UI's language stands for its value */
        const exact = list.find((x) => !isGroup(x) && shown(x).toLowerCase() === lc);
        cur = exact && shown(exact) !== String(val(exact)) ? val(exact) : v;
        mark(hit ? val(hit) : undefined, true);
        if (o.onChange) o.onChange(cur, exact ? indexOf(val(exact)) : -1);
      });
    }
    render();
    const wrap = inp ? h('div', { class: 'col listbox-wrap', style: 'gap:2px' }, inp, box) : box;
    wrap.get = () => (inp ? (cur != null && label(cur) === inp.value ? String(cur) : inp.value) : cur);
    wrap.set = (v) => { cur = v; if (inp) inp.value = label(v); mark(v, true); };
    wrap.setItems = (it) => { list = it.slice(); render(); };
    wrap.index = () => indexOf(cur);
    wrap.input = inp;
    wrap.list = box;
    return wrap;
  };
  /** A colour button in a dialog (Font color, Border color, Fill…): a swatch, the colour's name and an arrow that
   *  opens ui.colorMenu. Values: 'auto', null (none), '#RRGGBB' or what the app's colour menu gives (a theme colour,
   *  a fill). o: { id, mode ('font' | 'fill' | 'line' | 'plain'), automatic (offer Automatic; default for 'font'),
   *  none (offer No Color in 'plain' and 'font'), autoText, noneText, autoColor (the swatch for Automatic),
   *  paint(v) → CSS background, or null for the none pattern (theme colours, fills), name(v) → its label,
   *  convert(pick, current) → the new value from the menu's pick (a fill keeping its transparency), effects(current)
   *  → a promise of the value from Fill Effects… (undefined: cancelled), grid, design, onChange(v) }. Returns the
   *  button with get() and set(v) */
  ui.colorButton = function (value, o) {
    o = o || {};
    let val = value;
    const sw = h('span', { class: 'cpk-sw' });
    const lbl = h('span', { class: 'cpk-l' });
    const btn = h('button', { type: 'button', class: 'cpk', id: o.id || 'cb' + (++ctlId), 'aria-haspopup': 'true' }, sw, lbl, h('span', { class: 'dd-arrow' }));
    const show = () => {
      const custom = o.paint && val && val !== 'auto' ? o.paint(val) : undefined;
      const bg = custom !== undefined ? custom : val === 'auto' ? o.autoColor || '#000' : typeof val === 'string' && val[0] === '#' ? val : null;
      sw.style.background = bg || '';
      sw.classList.toggle('none', !bg);
      const name = val === 'auto' ? t(o.autoText || 'Automatic') : !val ? t(o.noneText || (o.mode === 'fill' ? 'No Fill' : o.mode === 'line' ? 'No Line' : 'No Color')) : o.name ? o.name(val) : typeof val === 'string' && val[0] === '#' ? ui.colorName(val) : '';
      lbl.textContent = name;
      btn.setAttribute('aria-label', name);
    };
    show();
    btn.addEventListener('click', () => ui.colorMenu(btn, { mode: o.mode || 'font', automatic: o.automatic, noneLabel: o.none ? o.noneText || 'No Color' : null, effects: !!o.effects, grid: o.grid, design: o.design, current: typeof val === 'string' && val[0] === '#' ? val : null }, async (v) => {
      let next;
      if (v && v.effects) { next = await o.effects(val); if (next === undefined) return; }
      else if (v && v.auto) next = 'auto';
      else if (v && v.none) next = null;
      else if (typeof v === 'string') next = v[0] === '#' ? v.toUpperCase() : v;
      else return;
      val = o.convert && !(v && v.effects) ? o.convert(next, val) : next;
      show();
      if (o.onChange) o.onChange(val);
    }));
    btn.get = () => val;
    btn.set = (v) => { val = v; show(); };
    return btn;
  };
  /** Tabs: [{label, body}] */
  ui.tabs = (tabs, active) => {
    const head = h('div', { class: 'tabs-head', role: 'tablist' });
    const bodies = h('div', { class: 'tabs-body' });
    const btns = tabs.map((t, i) => {
      const b = h('button', { type: 'button', class: 'tab', role: 'tab', html: ampHTML(t.label) });
      b.addEventListener('click', () => sel(i));
      head.appendChild(b);
      t.body.classList.add('tab-pane');
      bodies.appendChild(t.body);
      return b;
    });
    function sel(i) { btns.forEach((b, k) => { b.classList.toggle('on', k === i); b.setAttribute('aria-selected', String(k === i)); tabs[k].body.hidden = k !== i; }); if (tabs[i].onShow) tabs[i].onShow(); }
    sel(active || 0);
    const el = h('div', { class: 'tabs' }, head, bodies);
    el.select = sel;
    return el;
  };
  ui.field = (label, ctl, o) => {
    const id = ctl.id || (ctl.input && ctl.input.id) || (ctl.querySelector && ctl.querySelector('input,select') && ctl.querySelector('input,select').id);
    return h('div', { class: 'field' + (o && o.cls ? ' ' + o.cls : '') }, h('label', { for: id || null, html: ampHTML(label) }), ctl);
  };
  ui.group = (title, ...kids) => h('fieldset', { class: 'grp' }, h('legend', { html: ampHTML(title) }), ...kids);
  /** Tools ▸ Options ▸ General ▸ Theme: the suite's look (common/looks.css). { el, apply() }; apply on OK
   *  changes every open window of the suite. null outside the suite (no common/suite.js). */
  /** the suite's own choices for Tools ▸ Options: the look (Theme) and the UI's language (common/i18n.js) */
  ui.lookField = () => {
    if (!window.VO || !VO.look) return null;
    let pick = VO.look.get();
    const sel = ui.select(VO.look.list().map((l) => [l.id, l.name]), pick, (v) => { pick = v; }, { id: 'op-look' });
    const el = h('div', { class: 'col' }, ui.field('T&heme:', sel));
    let lang = VO.lang && VO.lang.get();
    /* each language named in itself */
    if (VO.lang) el.appendChild(ui.field('L&anguage:', ui.select(VO.lang.list().map((l) => [l.id, l.id === 'auto' ? t('Auto (as the browser)') : l.name]), lang, (v) => { lang = v; }, { id: 'op-lang', translate: 'no' })));
    return {
      el,
      apply: () => {
        if (pick !== VO.look.get()) VO.look.set(pick);
        if (VO.lang && lang !== VO.lang.get()) { VO.lang.set(lang); ui.toast('The new language is used in the windows you open from now on.'); }
      },
    };
  };

  /* ---------- Tools ▸ Customize ---------- */
  /** the Customize options an app keeps in its saved options o, all on unless false: Standard and Formatting
   *  toolbars on two rows (twoRows; the two rows sit in a .tb-pair), ScreenTips (screenTips), shortcut keys in them (tipKeys) */
  ui.toolbarPrefs = (o) => {
    const host = document.getElementById('toolbars');
    if (host) host.classList.toggle('one-row', o.twoRows === false);
    document.body.classList.toggle('notips', o.screenTips === false);
    document.body.classList.toggle('notipkeys', o.tipKeys === false);
  };
  /** Tools ▸ Customize. o: { toolbars: ids, each with a 'tb_<id>' toggle command; opts: the app's saved options;
   *  save(); shortcuts: { key: command id } besides the commands' own keys }. Changes apply at once, as in Office */
  ui.customize = function (o) {
    const plain = (label) => h('span', { text: stripAmp(label), translate: 'no' });
    const tbs = h('div', { class: 'cz-list' });
    for (const id of o.toolbars) {
      const c = ui.cmds['tb_' + id];
      if (!c) continue;
      const inp = h('input', { type: 'checkbox', id: 'cz-tb-' + id });
      inp.checked = ui.checked('tb_' + id);
      inp.addEventListener('change', () => { ui.exec('tb_' + id); inp.checked = ui.checked('tb_' + id); });
      tbs.appendChild(h('label', { class: 'ck', for: inp.id }, inp, plain(c.label)));
    }
    /* every command with a name (not the toolbars' toggles or symbols), its shortcut keys beside it */
    const keys = {};
    for (const [k, id] of Object.entries(o.shortcuts || {})) (keys[id] || (keys[id] = new Set())).add(k);
    const seen = new Set(), rows = [];
    for (const c of Object.values(ui.cmds)) {
      /* the command's own name: a fixed menu label or label, not one that tells the moment ("Undo Typing") */
      const fixed = (k) => { const d = Object.getOwnPropertyDescriptor(c, k); return d && !d.get ? d.value : null; };
      const name = c.label && !c.id.startsWith('tb_') ? stripAmp(fixed('menuLabel') || fixed('label') || c.tip || c.label).replace(/\.\.\.$|…$/, '') : '';
      if (!/\p{L}/u.test(name)) continue;
      const ks = new Set(keys[c.id]); if (c.key) ks.add(c.key);
      const kt = [...ks].map(keyText).sort().join(', ');
      if (seen.has(name + '\u0000' + kt)) continue;
      seen.add(name + '\u0000' + kt);
      rows.push({ id: c.id, name, kt });
    }
    rows.sort((a, b) => a.name.localeCompare(b.name));
    const table = h('table', { class: 'kbd-table cz-cmds', translate: 'no' }, h('tr', null, h('th', { text: 'Command' }), h('th', { text: 'Shortcut keys' })));
    let on = null;
    const pick = (tr) => { if (on) on.classList.remove('on'); on = tr; if (tr) { tr.classList.add('on'); tr.scrollIntoView({ block: 'nearest' }); } };
    for (const r of rows) {
      const tr = h('tr', null, h('td', { text: r.name, translate: 'no' }), h('td', { text: r.kt, translate: 'no' }));
      tr.addEventListener('click', () => pick(tr));
      tr.addEventListener('dblclick', () => ui.exec(r.id));
      tr.cmd = r.id;
      table.appendChild(tr);
    }
    const box = h('div', { class: 'cz-box', tabindex: '0', role: 'grid', 'aria-label': 'Commands' }, table);
    box.addEventListener('keydown', (e) => {
      const all = [...table.rows].slice(1);
      const i = all.indexOf(on);
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); pick(all[L.clamp(i + (e.key === 'ArrowDown' ? 1 : -1), 0, all.length - 1)]); }
      else if (e.key === 'Enter' && on) { e.preventDefault(); e.stopPropagation(); ui.exec(on.cmd); }
    });
    const op = o.opts;
    const set = (k, v) => { op[k] = v; o.save(); ui.toolbarPrefs(op); };
    const keyTips = ui.check('Show shortcut &keys in ScreenTips', op.tipKeys !== false, (v) => set('tipKeys', v));
    keyTips.input.disabled = op.screenTips === false;
    const options = h('div', { class: 'col' },
      ui.check('Show Standard and Formatting toolbars on t&wo rows', op.twoRows !== false, (v) => set('twoRows', v)),
      ui.check('Show Screen&Tips on toolbars', op.screenTips !== false, (v) => { set('screenTips', v); keyTips.input.disabled = !v; }),
      keyTips);
    ui.dialog({ title: 'Customize', width: 440, body: ui.tabs([
      { label: 'Tool&bars', body: h('div', { class: 'col' }, h('label', { for: tbs.firstChild && tbs.firstChild.htmlFor, text: 'Toolbars:' }), tbs) },
      { label: '&Commands', body: h('div', { class: 'col' }, h('div', { class: 'hint', text: 'Double-click a command to run it.' }), box) },
      { label: '&Options', body: options },
    ]), buttons: [{ label: 'Close', primary: true }] });
  };

  /* ---------- colour pickers ---------- */
  const recent = [];
  ui.addRecentColor = (c) => { if (!c || c[0] !== '#') return; const i = recent.indexOf(c); if (i >= 0) recent.splice(i, 1); recent.unshift(c); recent.length = Math.min(recent.length, 8); };
  /**
   * Colour drop-down. o: {mode:'fill'|'line'|'font'|'plain', design, effects:bool, automatic:bool (false: no Automatic,
   *   even in 'font'), noneLabel (a No Color entry where the mode has none), current ('#hex' More Colors starts at),
   *   scheme:bool (theme colours; on when the app provides ui.schemeColors), grid:bool (the standard colours too)}
   * onPick(value) where value is a colour ('#hex' or scheme ref), {none:true}, {auto:true}, {effects:true}
   */
  ui.colorMenu = function (anchor, o, onPick) {
    o = o || {};
    const design = o.design || (ui.currentDesign ? ui.currentDesign() : null);
    const scheme = o.scheme != null ? !!o.scheme : !!ui.schemeColors;
    const r = anchor.getBoundingClientRect ? anchor.getBoundingClientRect() : anchor;
    const panel = (close) => {
      const el = h('div', { class: 'color-panel' });
      if (o.mode === 'fill' || o.mode === 'line') el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'fill' ? 'No Fill' : 'No Line', onclick: () => { close(); onPick({ none: true }); } }));
      if (o.automatic !== false && (o.mode === 'font' || o.automatic)) el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: 'Automatic', onclick: () => { close(); onPick({ auto: true }); } }));
      else if (o.noneLabel && o.mode !== 'fill' && o.mode !== 'line') el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.noneLabel, onclick: () => { close(); onPick({ none: true }); } }));
      el.appendChild(h('div', { class: 'cp-sep' }));
      if (scheme && ui.schemeColors) {
        const row = h('div', { class: 'cp-row' });
        for (const sc of ui.schemeColors(design)) {
          const b = h('button', { class: 'cp-sw', type: 'button', style: `background:${sc.hex}`, 'data-tip': sc.tip || null, 'aria-label': sc.tip || sc.ref });
          b.addEventListener('click', () => { close(); onPick(sc.ref); });
          row.appendChild(b);
        }
        el.appendChild(row);
      }
      if (recent.length) {
        el.appendChild(h('div', { class: 'cp-sep' }));
        const rr = h('div', { class: 'cp-row' });
        for (const c of recent) rr.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'aria-label': c, onclick: () => { close(); onPick(c); } }));
        el.appendChild(rr);
      }
      if (o.mode === 'plain' || o.grid || !scheme) {
        if (scheme) el.appendChild(h('div', { class: 'cp-sep' }));
        const g = h('div', { class: 'cp-grid' });
        for (const c of L.color.STANDARD) g.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'aria-label': c, 'data-tip': ui.colorName(c), onclick: () => { close(); ui.addRecentColor(c); onPick(c); } }));
        el.appendChild(g);
      }
      el.appendChild(h('div', { class: 'cp-sep' }));
      el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'fill' ? 'More Fill Colors...' : o.mode === 'line' ? 'More Line Colors...' : 'More Colors...', onclick: () => { close(); ui.moreColors(o.current || null, (c) => { if (c) { ui.addRecentColor(c); onPick(c); } }); } }));
      if (o.effects) el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'line' ? 'Patterned Lines...' : 'Fill Effects...', onclick: () => { close(); onPick({ effects: true }); } }));
      return el;
    };
    ui.openMenu([{ custom: panel }], { left: r.left, bottom: r.bottom }, { cls: 'color-menu' });
  };
  /* the colour names Office shows, in the UI's language where they are shown */
  const N = (x) => x;
  const CNAMES = { '#000000': N('Black'), '#993300': N('Brown'), '#333300': N('Olive Green'), '#003300': N('Dark Green'), '#003366': N('Dark Teal'), '#000080': N('Dark Blue'), '#333399': N('Indigo'), '#333333': N('Gray-80%'), '#800000': N('Dark Red'), '#FF6600': N('Orange'), '#808000': N('Dark Yellow'), '#008000': N('Green'), '#008080': N('Teal'), '#0000FF': N('Blue'), '#666699': N('Blue-Gray'), '#808080': N('Gray-50%'), '#FF0000': N('Red'), '#FF9900': N('Light Orange'), '#99CC00': N('Lime'), '#339966': N('Sea Green'), '#33CCCC': N('Aqua'), '#3366FF': N('Light Blue'), '#800080': N('Violet'), '#969696': N('Gray-40%'), '#FF00FF': N('Pink'), '#FFCC00': N('Gold'), '#FFFF00': N('Yellow'), '#00FF00': N('Bright Green'), '#00FFFF': N('Turquoise'), '#00CCFF': N('Sky Blue'), '#993366': N('Plum'), '#C0C0C0': N('Gray-25%'), '#FF99CC': N('Rose'), '#FFCC99': N('Tan'), '#FFFF99': N('Light Yellow'), '#CCFFCC': N('Light Green'), '#CCFFFF': N('Light Turquoise'), '#99CCFF': N('Pale Blue'), '#CC99FF': N('Lavender'), '#FFFFFF': N('White') };
  ui.colorName = (c) => t(CNAMES[String(c).toUpperCase()]) || String(c).toUpperCase();
  /** "Colors" dialog with Standard and Custom tabs. cb(hex|null) */
  ui.moreColors = function (initial, cb) {
    let cur = initial && initial[0] === '#' ? initial : '#3366FF';
    const prevNew = h('div', { class: 'cc-new', style: `background:${cur}` });
    const prevOld = h('div', { class: 'cc-cur', style: `background:${cur}` });
    const set = (c) => { cur = c.toUpperCase(); prevNew.style.background = cur; hex.value = cur; const [r2, g2, b2] = L.color.hexToRgb(cur); R.set(r2); Gs.set(g2); Bs.set(b2); };
    /* standard: hexagon-ish honeycomb approximated by staggered rows */
    const std = h('div', { class: 'cc-honey' });
    const rows = 13;
    for (let r2 = 0; r2 < rows; r2++) {
      const n = r2 < 7 ? 7 + r2 : 19 - r2;
      const row = h('div', { class: 'cc-hrow' });
      for (let i = 0; i < n; i++) {
        const ang = ((i / n) * 360 + r2 * 23) % 360;
        const dist = Math.abs(r2 - 6) / 6, light = 0.5 + (r2 < 6 ? 0.35 : -0.25) * dist;
        const [cr, cg, cb] = L.color.hslToRgb(ang / 360, 1 - dist * 0.4, L.clamp(light, 0.1, 0.92));
        const c = L.color.rgbToHex(cr, cg, cb);
        row.appendChild(h('button', { type: 'button', class: 'cc-hex', style: `background:${c}`, 'aria-label': c, onclick: () => set(c) }));
      }
      std.appendChild(row);
    }
    const grays = h('div', { class: 'cc-grays' });
    for (let i = 0; i <= 15; i++) { const v = Math.round((255 * i) / 15); const c = L.color.rgbToHex(v, v, v); grays.appendChild(h('button', { type: 'button', class: 'cc-hex', style: `background:${c}`, 'aria-label': c, onclick: () => set(c) })); }
    const nat = h('input', { type: 'color', id: 'cc-native', value: cur.toLowerCase() });
    nat.addEventListener('input', () => set(nat.value));
    const R = ui.spin({ value: 0, min: 0, max: 255, step: 1, dec: 0, onChange: () => upd() });
    const Gs = ui.spin({ value: 0, min: 0, max: 255, step: 1, dec: 0, onChange: () => upd() });
    const Bs = ui.spin({ value: 0, min: 0, max: 255, step: 1, dec: 0, onChange: () => upd() });
    const hex = h('input', { type: 'text', id: 'cc-hex', value: cur, maxlength: 7, style: 'width:80px' });
    hex.addEventListener('change', () => { if (/^#?[0-9a-f]{6}$/i.test(hex.value)) set(hex.value[0] === '#' ? hex.value : '#' + hex.value); });
    function upd() { cur = L.color.rgbToHex(R.get(), Gs.get(), Bs.get()); prevNew.style.background = cur; hex.value = cur; nat.value = cur.toLowerCase(); }
    const custom = h('div', { class: 'cc-custom' },
      h('div', { class: 'field' }, h('label', { for: 'cc-native', text: 'Pick:' }), nat),
      ui.field('&Red:', R), ui.field('&Green:', Gs), ui.field('&Blue:', Bs), ui.field('He&x:', hex));
    const tabs = ui.tabs([{ label: 'Standard', body: h('div', null, h('div', { class: 'cc-cap', text: 'Colors:' }), std, grays) }, { label: 'Custom', body: custom }]);
    const side = h('div', { class: 'cc-side' }, h('div', { class: 'cc-cap', text: 'New|color' }), prevNew, prevOld, h('div', { class: 'cc-cap', text: 'Current' }));
    set(cur);
    let ok = false;
    ui.dialog({ title: 'Colors', body: h('div', { class: 'cc-wrap' }, tabs, side), width: 440, buttons: [{ label: 'OK', primary: true, onClick: () => { ok = true; } }, { label: 'Cancel' }] })
      .done.then(() => cb(ok ? cur : null));
  };

  /* ---------- brushes ---------- */
  /* A brush (Format Painter, highlighter) is a mode that acts on what the user marks with the mouse, once,
   * when the mouse button is let go: never while a selection is still being dragged out. A click on its
   * button arms it for one use, or ends it; a double-click keeps it on until the button again or Esc. One
   * brush is on at a time. The app gives
   *   el()           the element that shows the brush cursor (class 'painting')
   *   pick()         what the brush carries, e.g. the formatting at the selection; null: nothing to carry
   *   apply(v, ...)  paint what the user marked with v; true when something was painted
   *   hint, status   optional: status(hint) while it is on, status('') after
   * and calls brush.release(...) when the user lets go of the mouse over what they marked. */
  let liveBrush = null;
  ui.brush = function (o) {
    const b = {
      value: null,
      sticky: false,
      on: () => liveBrush === b,
      /** arm it (value: carry this instead of pick()) */
      start(sticky, value) {
        const v = value !== undefined ? value : o.pick();
        if (v == null) return false;
        if (liveBrush && liveBrush !== b) liveBrush.stop();
        liveBrush = b;
        b.value = v;
        b.sticky = !!sticky;
        o.el().classList.add('painting');
        if (o.status) o.status(o.hint || '');
        ui.refresh();
        return true;
      },
      stop() {
        if (liveBrush !== b) return;
        liveBrush = null;
        b.value = null;
        b.sticky = false;
        o.el().classList.remove('painting');
        if (o.status) o.status('');
        ui.refresh();
      },
      /** its button: the second click of a double-click keeps it on */
      click(ev) {
        if (ev && ev.detail === 2) { if (b.on()) { b.sticky = true; if (o.status) o.status(o.hint || ''); } else b.start(true); }
        else if (b.on()) b.stop();
        else b.start(false);
      },
      release(...args) {
        if (!b.on()) return false;
        const done = !!o.apply(b.value, ...args);
        if (done && !b.sticky) b.stop();
        return done;
      },
    };
    return b;
  };
  ui.brushOn = () => liveBrush;
  if (typeof document !== 'undefined') document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape' || !liveBrush || ui.dialogOpen() || ui.menuOpen()) return;
    e.preventDefault();
    e.stopPropagation();
    liveBrush.stop();
  }, true);

  /* ---------- drags ---------- */
  /* A drag is what a press with the primary button starts and its release ends: moving a shape, a ruler marker or a
   * cell range, drawing a shape. One at a time, the same in every app (LibreOffice's vcl tracking does the same):
   *   ui.drag(e, { move(ev, d), up(ev, d), cancel(d), threshold = 4, capture })
   * on the press event e (a pointerdown, or a mousedown: no pointer capture then, the window listeners follow) returns
   * the drag d ({ x0, y0, x, y, dx, dy, moved, data: {} for the app }); null for a press of another button.
   * - move runs only once the pointer is `threshold` px from the press (Windows' drag rectangle), so a click that
   *   never moved changes nothing: up gets d.moved false.
   * - up runs on the release of the primary button. The pointer is held by `capture` (by default e.target; false: not
   *   held, the window listeners follow), so the click and dblclick that follow go to that element; if it is drawn
   *   again in the middle of the drag, the drag goes on. Don't hold it on a CSS-scaled ancestor of text the browser
   *   selects: its selection then stops short of the pointer.
   * - cancel(d, why) runs instead, and nothing may commit (the app puts back what it moved), on: Esc ('esc', which
   *   goes no further), a second button ('button', its context menu doesn't open), a move that reports the button no
   *   longer held ('lost': it was let go where the page couldn't see), a lost capture or pointercancel ('lost'), the
   *   window losing focus or the page hidden ('blur'), Ctrl+Z or Ctrl+Y ('undo': the drag is what gets undone), a new
   *   drag ('new'), ui.dragCancel() ('api': a tool or view switch). A tool stays armed unless why is 'esc'; Esc also
   *   puts a brush down (ui.brush). Other keys go nowhere while the button is held, except the modifiers.
   * - While it moves, the cursor shown at the press stays wherever the pointer goes (`cursor` sets another one;
   *   false: none is held).
   * A text selection is the browser's own drag; an app may watch it with { capture: false, cursor: false } (Quire
   * does, so its release, or how it ended otherwise, is known the same way). */
  let liveDrag = null, swallow = null, cursorStyle = null;
  /** the events of the button that cancelled a drag go no further: its press, release, auxclick and context menu */
  const swallowButton = (button) => { swallow = { button, until: Date.now() + 1500, seen: new Set() }; };
  const holdCursor = (c) => {
    if (!cursorStyle) { cursorStyle = document.createElement('style'); document.head.appendChild(cursorStyle); }
    cursorStyle.textContent = c ? `*, *::before, *::after { cursor: ${c} !important; }` : '';
  };
  ui.drag = function (e, o) {
    if (e.button != null && e.button !== 0) return null;
    if (liveDrag) liveDrag.end(false, null, 'new');
    const d = { x0: e.clientX, y0: e.clientY, x: e.clientX, y: e.clientY, dx: 0, dy: 0, moved: false, data: {} };
    const id = e.pointerId, threshold = o.threshold == null ? 4 : o.threshold;
    const cap = o.capture === false ? null : o.capture || e.target;
    let cursor = o.cursor || null;
    if (!cursor && o.cursor !== false) { try { const c = getComputedStyle(e.target).cursor; if (c && c !== 'auto') cursor = c; } catch (err) { /* no cursor kept */ } }
    try { if (cap && id != null && cap.setPointerCapture) cap.setPointerCapture(id); } catch (err) { /* window listeners still follow */ }
    const on = (t, type, fn, capture) => { t.addEventListener(type, fn, capture); return () => t.removeEventListener(type, fn, capture); };
    const second = (button) => { swallowButton(button); end(false, null, 'button'); };
    const offs = [
      on(window, 'pointermove', (ev) => {
        if (id != null && ev.pointerId !== id) return;
        if (ev.buttons & 6) { second(ev.buttons & 2 ? 2 : 1); return; }
        if (!(ev.buttons & 1)) { end(false, null, 'lost'); return; }
        d.x = ev.clientX; d.y = ev.clientY; d.dx = d.x - d.x0; d.dy = d.y - d.y0;
        if (!d.moved && Math.hypot(d.dx, d.dy) < threshold) return;
        if (!d.moved && cursor) holdCursor(cursor);
        d.moved = true;
        if (o.move) o.move(ev, d);
      }, true),
      on(window, 'pointerup', (ev) => { if ((id == null || ev.pointerId === id) && ev.button === 0) end(true, ev); }, true),
      on(window, 'keydown', (ev) => {
        /* the modifiers shape the drag (Shift constrains it, Ctrl copies, Alt doesn't snap) */
        if (/^(Shift|Control|Alt|Meta|AltGraph|CapsLock)$/.test(ev.key)) return;
        ev.preventDefault(); ev.stopImmediatePropagation();
        /* Ctrl+Z (or Y) in the middle of a drag undoes the drag, not what came before it */
        if ((ev.ctrlKey || ev.metaKey) && /^[zy]$/i.test(ev.key)) { end(false, null, 'undo'); return; }
        if (ev.key === 'Escape') { end(false, null, 'esc'); if (liveBrush) liveBrush.stop(); }
        /* any other key waits for the release: an edit in the middle of a drag would come between its start and the
           step it records */
      }, true),
      on(window, 'pointercancel', () => end(false, null, 'lost'), true),
      on(window, 'blur', () => end(false, null, 'blur')),
      on(document, 'visibilitychange', () => { if (document.hidden) end(false, null, 'blur'); }),
    ];
    if (cap) offs.push(on(cap, 'lostpointercapture', () => { if (cap.isConnected) end(false, null, 'lost'); }));
    function end(ok, ev, why) {
      if (liveDrag !== live) return;
      liveDrag = null;
      offs.forEach((off) => off());
      if (cursorStyle) holdCursor(null);
      try { if (cap && id != null && cap.hasPointerCapture && cap.hasPointerCapture(id)) cap.releasePointerCapture(id); } catch (err) { /* ignore */ }
      if (ok) { if (o.up) o.up(ev, d); } else if (o.cancel) o.cancel(d, why);
    }
    const live = { end, d };
    liveDrag = live;
    return d;
  };
  ui.dragging = () => !!liveDrag;
  ui.dragCancel = () => { if (liveDrag) liveDrag.end(false, null, 'api'); };
  /* a second button cancels the drag; its own events then go no further (the pointer events report it as a move,
     before the mouse events come): its press, its release, and its auxclick or context menu, whichever ends that
     click, in whatever order the system sends them (or for 1.5 s at most) */
  if (typeof window !== 'undefined') for (const type of ['mousedown', 'mouseup', 'auxclick', 'contextmenu']) window.addEventListener(type, (e) => {
    if (e.button === 0) return;
    if (liveDrag) { swallowButton(e.button); liveDrag.end(false, null, 'button'); }
    if (!swallow || e.button !== swallow.button) return;
    if (Date.now() > swallow.until) { swallow = null; return; }
    e.preventDefault(); e.stopImmediatePropagation();
    swallow.seen.add(type);
    if (swallow.seen.has('mouseup') && swallow.seen.has(swallow.button === 2 ? 'contextmenu' : 'auxclick')) swallow = null;
  }, true);

  /* ---------- status bar & misc ---------- */
  /** a short notice at the bottom; action: { label, run } adds a button to it (e.g. Browse) */
  ui.toast = (text, ms, action) => {
    let t = L.$('.toast');
    if (!t) { t = h('div', { class: 'toast', role: 'status' }); document.body.appendChild(t); }
    t.textContent = L.t ? L.t(text) : text;
    if (action) t.append(' ', h('button', { class: 'toast-act', type: 'button', text: action.label, onclick: () => { t.hidden = true; action.run(); } }));
    t.hidden = false;
    clearTimeout(t._t);
    t._t = setTimeout(() => { t.hidden = true; }, ms || 2600);
  };
  ui.busy = (on, text) => {
    let b = L.$('.busy');
    if (!!on !== !!(b && !b.hidden)) tell('vo-busy', { on: !!on, text: on ? text || 'Working...' : '' });
    if (on) { if (!b) { b = h('div', { class: 'busy' }, h('div', { class: 'busy-box' }, h('span', { class: 'busy-spin' }), h('span', { class: 'busy-t' }))); document.body.appendChild(b); } b.querySelector('.busy-t').textContent = t(text || 'Working...'); b.hidden = false; }
    else if (b) b.hidden = true;
  };
})(typeof window !== 'undefined' ? window : globalThis);
