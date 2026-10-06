/* Lectern — UI toolkit: commands, menus, toolbars, dialogs, pickers.
 * Styled after the Office 2003 "Luna" command bars.
 */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;
  const ui = (L.ui = {});

  /* ---------- command registry ---------- */
  ui.cmds = {};
  ui.cmd = (id, def) => { const o = { id }; Object.defineProperties(o, Object.getOwnPropertyDescriptors(def)); ui.cmds[id] = o; return o; };
  ui.exec = (id, arg, ev) => {
    const c = ui.cmds[id];
    if (!c) { console.warn('unknown command', id); return; }
    if (c.enabled && !c.enabled()) return;
    try { c.run(arg, ev); } catch (e) { console.error(e); ui.msg('That action could not be completed: ' + (e && e.message ? e.message : e), { icon: 'warn' }); }
    ui.refresh();
  };
  ui.enabled = (id) => { const c = ui.cmds[id]; return !!c && (!c.enabled || !!c.enabled()); };
  ui.checked = (id) => { const c = ui.cmds[id]; return !!c && !!c.checked && !!c.checked(); };
  const stripAmp = (s) => String(s || '').replace(/&(.)/g, '$1');
  const ampHTML = (s) => L.esc(s || '').replace(/&amp;(.)/g, '<u>$1</u>');
  ui.stripAmp = stripAmp;

  /* ---------- tooltips ---------- */
  let tipEl = null, tipTimer = null, tipTarget = null;
  function showTip(target, x, y) {
    const t = target.getAttribute('data-tip');
    if (!t) return;
    if (!tipEl) { tipEl = h('div', { class: 'tooltip', role: 'tooltip' }); document.body.appendChild(tipEl); }
    tipEl.textContent = t;
    tipEl.hidden = false;
    const r = tipEl.getBoundingClientRect();
    let left = x + 2, top = y + 20;
    if (left + r.width > window.innerWidth - 4) left = window.innerWidth - r.width - 4;
    if (top + r.height > window.innerHeight - 4) top = y - r.height - 6;
    tipEl.style.left = Math.max(2, left) + 'px';
    tipEl.style.top = Math.max(2, top) + 'px';
  }
  function hideTip() { clearTimeout(tipTimer); tipTarget = null; if (tipEl) tipEl.hidden = true; }
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
    const el = h('div', { class: 'menu' + (opts.cls ? ' ' + opts.cls : ''), role: 'menu', tabindex: '-1' });
    const rows = [];
    list.forEach((it) => {
      if (it.sep) { el.appendChild(h('div', { class: 'm-sep' })); return; }
      if (it.custom) { const c = it.custom(() => ui.closeMenus()); el.appendChild(c); return; }
      const disabled = it.disabled || (it.enabled && !it.enabled());
      const checked = it.checked && it.checked();
      const row = h('div', { class: 'm-item' + (disabled ? ' dis' : '') + (checked ? ' chk' : '') + (it.sub ? ' has-sub' : ''), role: 'menuitem', 'aria-disabled': disabled ? 'true' : null });
      const icoHTML = it.html ? it.html : it.icon ? L.icons.get(it.icon) : checked ? (it.radio ? '<span class="m-radio">●</span>' : '<span class="m-check">✓</span>') : '';
      row.appendChild(h('span', { class: 'm-ico' + (checked && it.icon ? ' on' : ''), html: icoHTML }));
      row.appendChild(h('span', { class: 'm-lbl', html: ampHTML(it.label) }));
      row.appendChild(h('span', { class: 'm-key', text: it.key || '' }));
      row.appendChild(h('span', { class: 'm-sub', html: it.sub ? '▶' : '' }));
      row._it = it;
      const mn = /&(.)/.exec(it.label || '');
      row._mn = mn ? mn[1].toLowerCase() : null;
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
    document.body.appendChild(el);
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
  document.addEventListener('pointerdown', (e) => {
    if (!openStack.length) return;
    if (e.target.closest('.menu') || e.target.closest('.mb-item')) return;
    ui.closeMenus();
  }, true);
  window.addEventListener('blur', () => ui.closeMenus());
  window.addEventListener('resize', () => ui.closeMenus());
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
      const i = menus.findIndex((m) => /&(.)/.exec(m.label) && /&(.)/.exec(m.label)[1].toLowerCase() === ch.toLowerCase());
      if (i >= 0) { open(i, true); return true; }
      return false;
    };
    bar.openIndex = (i) => open(i, true);
    container.appendChild(bar);
    return bar;
  };

  /* ---------- toolbars ---------- */
  const tbButtons = [];
  ui.tbButton = function (id, opts) {
    opts = opts || {};
    const c = ui.cmds[id] || {};
    const label = opts.label != null ? opts.label : c.tbLabel;
    const b = h('button', { class: 'tb-btn' + (label ? ' with-label' : ''), type: 'button', 'data-cmd': id, 'aria-label': stripAmp(c.label || id), 'data-tip': stripAmp(c.tip || c.label || id) + (c.key ? ` (${c.key})` : '') });
    b.innerHTML = (c.icon || opts.icon ? L.icons.get(opts.icon || c.icon) : '') + (label ? `<span class="tb-lbl">${ampHTML(label)}</span>` : '');
    b.addEventListener('pointerdown', (e) => e.preventDefault()); /* keep text selection in the slide */
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
  /** editable combo box (font name / size / zoom) */
  ui.combo = function (o) {
    const wrap = h('div', { class: 'combo', style: `width:${o.width || 120}px`, 'data-tip': o.tip || '' });
    const inp = h('input', { type: 'text', class: 'combo-in', id: o.id, 'aria-label': o.tip || o.id, autocomplete: 'off', spellcheck: 'false' });
    const btn = h('button', { class: 'combo-btn', type: 'button', tabindex: '-1', 'aria-label': 'Open list', html: '<span class="dd-arrow"></span>' });
    wrap.append(inp, btn);
    const commit = () => { o.onChange(inp.value); inp.blur(); };
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); if (L.ed) L.ed.refocus(); }
      if (e.key === 'Escape') { e.preventDefault(); inp.value = o.value(); inp.blur(); if (L.ed) L.ed.refocus(); }
      e.stopPropagation();
    });
    inp.addEventListener('focus', () => { inp.select(); wrap.classList.add('focus'); if (L.ed && L.ed.textSave) L.ed.textSave(); });
    inp.addEventListener('blur', () => wrap.classList.remove('focus'));
    btn.addEventListener('pointerdown', (e) => e.preventDefault());
    btn.addEventListener('click', () => {
      if (L.ed && L.ed.textSave) L.ed.textSave();
      const r = wrap.getBoundingClientRect();
      const opts = o.options();
      ui.openMenu(opts.map((v) => ({
        label: String(v).replace(/&/g, '&&'),
        html: '', run: () => { inp.value = v; o.onChange(String(v)); if (L.ed) L.ed.refocus(); },
        checked: () => String(v) === String(o.value()),
        style: o.preview ? o.preview(v) : null,
      })), { left: r.left, bottom: r.bottom }, { cls: 'combo-list' + (o.listCls ? ' ' + o.listCls : '') });
      if (o.preview) {
        const menu = document.querySelector('.menu.combo-list');
        if (menu) L.$$('.m-item', menu).forEach((row, i) => { const st = o.preview(opts[i]); if (st) L.$('.m-lbl', row).style.cssText = st; });
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
    return tb;
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
  ui.dialogOpen = () => dlgStack.length > 0;
  /**
   * opts: {title, body, width, buttons:[{label, primary, cancel, onClick(close)→false to keep open}], onClose, help}
   */
  ui.dialog = function (opts) {
    ui.closeMenus();
    hideTip();
    const prevFocus = document.activeElement;
    const overlay = h('div', { class: 'dlg-overlay' });
    const dlg = h('div', { class: 'dlg', role: 'dialog', 'aria-modal': 'true', 'aria-label': opts.title });
    dlg.style.width = Math.min(opts.width || 380, window.innerWidth - 16) + 'px';
    const titleBar = h('div', { class: 'dlg-title' }, h('span', { class: 'dlg-ttl', text: opts.title }),
      opts.help ? h('button', { class: 'dlg-x dlg-help', type: 'button', 'aria-label': 'Help', text: '?', onclick: opts.help }) : null,
      h('button', { class: 'dlg-x', type: 'button', 'aria-label': 'Close', html: '<svg width="9" height="9" viewBox="0 0 9 9"><path d="M1 1l7 7M8 1L1 8" stroke="#fff" stroke-width="1.8"/></svg>', onclick: () => close(null) }));
    const body = h('div', { class: 'dlg-body' });
    if (opts.body) body.appendChild(opts.body);
    const foot = h('div', { class: 'dlg-foot' });
    let result;
    const done = new Promise((res) => { result = res; });
    const btns = (opts.buttons || [{ label: 'OK', primary: true }]).map((b, i) => {
      const el = h('button', { class: 'btn' + (b.primary ? ' primary' : ''), type: 'button', html: ampHTML(b.label) });
      el.addEventListener('click', () => {
        if (b.onClick) { const r = b.onClick(close); if (r === false) return; }
        close(i);
      });
      foot.appendChild(el);
      return el;
    });
    dlg.append(titleBar, body, foot);
    overlay.appendChild(dlg);
    document.body.appendChild(overlay);
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
    dlgStack.push(entry);
    entry.keys = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); const ci = (opts.buttons || []).findIndex((b) => b.cancel || /^cancel$|^close$/i.test(stripAmp(b.label))); if (ci >= 0) btns[ci].click(); else close(null); return true; }
      if (e.key === 'Enter' && !e.shiftKey && !(e.target && (e.target.tagName === 'TEXTAREA' || e.target.tagName === 'BUTTON' || e.target.isContentEditable || e.target.tagName === 'SELECT'))) {
        const pi = (opts.buttons || []).findIndex((b) => b.primary);
        if (pi >= 0) { e.preventDefault(); btns[pi].click(); return true; }
      }
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
      if (prevFocus && prevFocus.focus && document.contains(prevFocus)) { try { prevFocus.focus({ preventScroll: true }); } catch (e) { /* ignore */ } }
      else if (L.ed) L.ed.refocus();
    }
    setTimeout(() => {
      const first = opts.focus ? (typeof opts.focus === 'string' ? L.$(opts.focus, dlg) : opts.focus) : L.$('input:not([type=checkbox]):not([type=radio]),select,textarea', body) || btns.find((b, i) => opts.buttons && opts.buttons[i] && opts.buttons[i].primary) || btns[0];
      if (first) { first.focus(); if (first.select && first.tagName === 'INPUT') first.select(); }
    }, 0);
    if (opts.onOpen) opts.onOpen(dlg);
    return { el: dlg, body, close, done, buttons: btns };
  };
  ui.dialogKey = (e) => { const top = dlgStack[dlgStack.length - 1]; return top ? top.keys(e) || true : false; };

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
    String(text).split('\n').forEach((line, i) => { if (i) t.appendChild(h('br')); t.appendChild(document.createTextNode(line)); });
    const labels = o.buttons || ['OK'];
    const d = ui.dialog({ title: o.title || L.APP, body, width: o.width || 380, buttons: labels.map((l, i) => ({ label: l, primary: i === (o.def || 0), cancel: /cancel|^no$/i.test(stripAmp(l)) })) });
    return d.done;
  };
  ui.prompt = function (label, value, title) {
    const inp = h('input', { type: 'text', id: 'prompt-in', value: value || '', style: 'width:100%' });
    let val = null;
    const d = ui.dialog({ title: title || L.APP, body: h('div', { class: 'col' }, h('label', { for: 'prompt-in', text: label }), inp), buttons: [{ label: 'OK', primary: true, onClick: () => { val = inp.value; } }, { label: 'Cancel' }] });
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
  ui.select = (options, value, onChange, o) => {
    const sel = h('select', Object.assign({ id: (o && o.id) || 'sel' + (++ctlId) }, o || {}));
    for (const op of options) {
      const [v, l] = Array.isArray(op) ? op : [op, op];
      const opt = h('option', { value: v, text: l });
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

  /* ---------- colour pickers ---------- */
  const recent = [];
  ui.addRecentColor = (c) => { if (!c || c[0] !== '#') return; const i = recent.indexOf(c); if (i >= 0) recent.splice(i, 1); recent.unshift(c); recent.length = Math.min(recent.length, 8); };
  /**
   * Colour drop-down. o: {mode:'fill'|'line'|'font'|'plain', design, noneLabel, effects:bool, automatic:bool}
   * onPick(value) where value is a colour ('#hex' or scheme ref), {none:true}, {auto:true}, {more:true}, {effects:true}
   */
  ui.colorMenu = function (anchor, o, onPick) {
    o = o || {};
    const design = o.design || (L.ed ? L.model.design(L.pres, L.ed.slide()) : null);
    const r = anchor.getBoundingClientRect ? anchor.getBoundingClientRect() : anchor;
    const panel = (close) => {
      const el = h('div', { class: 'color-panel' });
      if (o.mode === 'fill' || o.mode === 'line') el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'fill' ? 'No Fill' : 'No Line', onclick: () => { close(); onPick({ none: true }); } }));
      if (o.mode === 'font' || o.automatic) el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: 'Automatic', onclick: () => { close(); onPick({ auto: true }); } }));
      el.appendChild(h('div', { class: 'cp-sep' }));
      const row = h('div', { class: 'cp-row' });
      for (const sc of L.model.schemeRow(design)) {
        const b = h('button', { class: 'cp-sw', type: 'button', style: `background:${sc.hex}`, 'data-tip': L.model.SCHEME_SLOTS.find((x) => x[0] === ({ bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }[sc.ref] || sc.ref))[1], 'aria-label': sc.ref });
        b.addEventListener('click', () => { close(); onPick(sc.ref); });
        row.appendChild(b);
      }
      el.appendChild(row);
      if (recent.length) {
        el.appendChild(h('div', { class: 'cp-sep' }));
        const rr = h('div', { class: 'cp-row' });
        for (const c of recent) rr.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'aria-label': c, onclick: () => { close(); onPick(c); } }));
        el.appendChild(rr);
      }
      if (o.mode === 'plain' || o.grid) {
        el.appendChild(h('div', { class: 'cp-sep' }));
        const g = h('div', { class: 'cp-grid' });
        for (const c of L.color.STANDARD) g.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'aria-label': c, onclick: () => { close(); ui.addRecentColor(c); onPick(c); } }));
        el.appendChild(g);
      }
      el.appendChild(h('div', { class: 'cp-sep' }));
      el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'fill' ? 'More Fill Colors...' : o.mode === 'line' ? 'More Line Colors...' : 'More Colors...', onclick: () => { close(); ui.moreColors(null, (c) => { if (c) { ui.addRecentColor(c); onPick(c); } }); } }));
      if (o.effects) el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: o.mode === 'line' ? 'Patterned Lines...' : 'Fill Effects...', onclick: () => { close(); onPick({ effects: true }); } }));
      return el;
    };
    ui.openMenu([{ custom: panel }], { left: r.left, bottom: r.bottom }, { cls: 'color-menu' });
  };
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
    const side = h('div', { class: 'cc-side' }, h('div', { class: 'cc-cap', text: 'New' }), prevNew, prevOld, h('div', { class: 'cc-cap', text: 'Current' }));
    set(cur);
    let ok = false;
    ui.dialog({ title: 'Colors', body: h('div', { class: 'cc-wrap' }, tabs, side), width: 440, buttons: [{ label: 'OK', primary: true, onClick: () => { ok = true; } }, { label: 'Cancel' }] })
      .done.then(() => cb(ok ? cur : null));
  };

  /* ---------- status bar & misc ---------- */
  ui.toast = (text, ms) => {
    let t = L.$('.toast');
    if (!t) { t = h('div', { class: 'toast', role: 'status' }); document.body.appendChild(t); }
    t.textContent = text;
    t.hidden = false;
    clearTimeout(t._t);
    t._t = setTimeout(() => { t.hidden = true; }, ms || 2600);
  };
  ui.busy = (on, text) => {
    let b = L.$('.busy');
    if (on) { if (!b) { b = h('div', { class: 'busy' }, h('div', { class: 'busy-box' }, h('span', { class: 'busy-spin' }), h('span', { class: 'busy-t' }))); document.body.appendChild(b); } b.querySelector('.busy-t').textContent = text || 'Working...'; b.hidden = false; }
    else if (b) b.hidden = true;
  };
})();
