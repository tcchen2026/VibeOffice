/* Lectern — panes and alternate views. */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;
  const ui = L.ui;
  const M = L.model;
  const E = () => L.ed;
  const P = (L.panes = {});

  /* ================= Slides / Outline pane ================= */
  const LP = (P.left = { tab: 'slides' });
  LP.mount = function (el) {
    LP.el = el;
    LP.tabs = h('div', { class: 'lp-tabs', role: 'tablist' });
    LP.tabOutline = h('div', { class: 'lp-tab', role: 'tab', tabindex: '0', text: 'Outline' });
    LP.tabSlides = h('div', { class: 'lp-tab', role: 'tab', tabindex: '0', text: 'Slides' });
    const close = h('button', { class: 'lp-close', type: 'button', 'aria-label': 'Close pane', html: L.icons.get('close', 12) });
    close.addEventListener('click', () => L.app.toggleLeft(false));
    LP.tabs.append(LP.tabOutline, LP.tabSlides, close);
    LP.body = h('div', { class: 'lp-body', tabindex: '-1' });
    el.append(LP.tabs, LP.body);
    LP.tabOutline.addEventListener('click', () => LP.setTab('outline'));
    LP.tabSlides.addEventListener('click', () => LP.setTab('slides'));
    LP.body.addEventListener('focus', () => { if (L.app) L.app.focusArea = 'slides'; }, true);
    LP.body.addEventListener('blur', () => { if (L.app) L.app.focusArea = null; }, true);
    LP.setTab('slides');
  };
  LP.setTab = function (t) {
    if (LP.tab === 'outline' && t !== 'outline') LP.commitOutline();
    LP.tab = t;
    LP.tabOutline.classList.toggle('on', t === 'outline');
    LP.tabSlides.classList.toggle('on', t === 'slides');
    LP.render();
    L.ui.refresh();
  };
  LP.render = function () {
    if (!LP.body) return;
    if (E().view === 'master') { renderMasterThumbs(); return; }
    if (LP.tab === 'outline') renderOutline();
    else renderThumbs();
  };
  /* --- thumbnails --- */
  function thumbWidth() { return Math.max(60, (LP.body.clientWidth || 190) - 48); }
  function renderThumbs() {
    const pres = L.pres;
    const box = h('div', { class: 'thumbs', role: 'listbox', 'aria-label': 'Slides' });
    const w = thumbWidth();
    pres.slides.forEach((s, i) => box.appendChild(thumbItem(s, i, w)));
    L.clear(LP.body).appendChild(box);
    LP.list = box;
    scrollCur();
  }
  function thumbItem(s, i, w) {
    const sel = L.app.slideSel;
    const it = h('div', { class: 'th' + (i === E().idx ? ' cur' : '') + (sel.has(i) ? ' sel' : '') + (s.hidden ? ' hiddenslide' : ''), 'data-i': i, role: 'option', 'aria-selected': String(i === E().idx) });
    const meta = h('div', { class: 'th-meta' }, h('span', { class: 'th-num', text: String(i + (L.pres.firstNum || 1)) }));
    if ((s.trans && s.trans.type && s.trans.type !== 'none') || (s.anims && s.anims.length)) {
      const star = h('span', { class: 'th-star', 'data-tip': 'Play animations', html: '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M6 .8l1.5 3.2 3.4.4-2.5 2.3.7 3.4L6 8.4 2.9 10.1l.7-3.4L1.1 4.4l3.4-.4z" fill="none" stroke="#333"/></svg>' });
      star.addEventListener('click', (e) => { e.stopPropagation(); const fr = it.querySelector('.th-frame'); L.show.previewTransition(fr, L.pres, s, L.pres.slides[i - 1]); });
      meta.appendChild(star);
    }
    const frame = h('button', { class: 'th-frame', type: 'button', 'aria-label': `Slide ${i + 1}` });
    frame.appendChild(L.render.thumb(L.pres, s, w, { index: i }));
    it.append(meta, frame);
    it.addEventListener('pointerdown', (e) => onThumbDown(e, i));
    it.addEventListener('contextmenu', (e) => { e.preventDefault(); if (!L.app.slideSel.has(i)) { L.app.slideSel = new Set([i]); E().goto(i); } L.app.contextThumb(e); });
    it.addEventListener('dblclick', () => { if (L.app.view !== 'normal') L.app.setView('normal'); });
    return it;
  }
  function scrollCur() {
    const cur = LP.body.querySelector('.th.cur');
    if (cur) cur.scrollIntoView({ block: 'nearest' });
  }
  LP.refreshThumb = function (i) {
    if (LP.tab !== 'slides' || !LP.list || E().view === 'master') return;
    const s = L.pres.slides[i];
    const old = LP.list.querySelector(`.th[data-i="${i}"]`);
    if (!s || !old) { renderThumbs(); return; }
    old.replaceWith(thumbItem(s, i, thumbWidth()));
  };
  LP.markCurrent = function () {
    if (!LP.list) return;
    L.$$('.th', LP.list).forEach((el) => { const i = +el.dataset.i; el.classList.toggle('cur', i === E().idx); el.classList.toggle('sel', L.app.slideSel.has(i)); el.setAttribute('aria-selected', String(i === E().idx)); });
    scrollCur();
  };
  function onThumbDown(e, i) {
    if (e.button !== 0) return;
    LP.body.focus({ preventScroll: true });
    L.app.focusArea = 'slides';
    const sel = L.app.slideSel;
    if (e.shiftKey) { const a = Math.min(E().idx, i), b = Math.max(E().idx, i); L.app.slideSel = new Set(Array.from({ length: b - a + 1 }, (_, k) => a + k)); LP.markCurrent(); return; }
    if (e.ctrlKey || e.metaKey) { if (sel.has(i) && sel.size > 1) sel.delete(i); else sel.add(i); LP.markCurrent(); return; }
    if (!sel.has(i)) L.app.slideSel = new Set([i]);
    if (E().idx !== i || L.app.view !== 'normal') E().goto(i);
    LP.markCurrent();
    /* drag to reorder */
    const y0 = e.clientY;
    let line = null, target = null;
    const mv = (ev) => {
      if (!line && Math.abs(ev.clientY - y0) < 6) return;
      if (!line) { line = h('div', { class: 'drop-line' }); LP.body.appendChild(line); }
      const items = L.$$('.th', LP.list);
      const br = LP.body.getBoundingClientRect();
      target = items.length;
      for (let k = 0; k < items.length; k++) { const r = items[k].getBoundingClientRect(); if (ev.clientY < r.top + r.height / 2) { target = k; break; } }
      const ref = items[target] || items[items.length - 1];
      const rr = ref.getBoundingClientRect();
      line.style.top = (target < items.length ? rr.top - 5 : rr.bottom + 4) - br.top + LP.body.scrollTop + 'px';
      if (ev.clientY < br.top + 20) LP.body.scrollTop -= 8; else if (ev.clientY > br.bottom - 20) LP.body.scrollTop += 8;
    };
    const up = () => {
      window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up);
      if (line) { line.remove(); if (target != null) L.app.moveSlides(Array.from(L.app.slideSel).sort((a, b) => a - b), target); }
    };
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
  }
  function renderMasterThumbs() {
    const box = h('div', { class: 'thumbs' });
    const d = E().design();
    const kinds = [['slide', 'Slide Master']].concat(d.titleDeco ? [['title', 'Title Master']] : []);
    kinds.forEach(([k, label], i) => {
      E().masterKind = E().masterKind === 'title' && !d.titleDeco ? 'slide' : E().masterKind;
      const it = h('div', { class: 'th' + (E().masterKind === k ? ' cur' : '') });
      const frame = h('button', { class: 'th-frame', type: 'button', 'aria-label': label, 'data-tip': `${d.name} ${label}` });
      const saveKind = E().masterKind;
      E().masterKind = k; E().buildMaster();
      frame.appendChild(L.render.thumb(L.pres, E().masterSlide, thumbWidth(), { prompts: false }));
      E().masterKind = saveKind; E().buildMaster();
      frame.addEventListener('click', () => { E().masterKind = k; E().buildMaster(); E().sel = []; E().render(); renderMasterThumbs(); });
      it.append(h('div', { class: 'th-meta' }, h('span', { class: 'th-num', text: String(i + 1) })), frame);
      box.appendChild(it);
    });
    L.clear(LP.body).appendChild(box);
    LP.list = null;
  }

  /* --- outline --- */
  let olDirty = false, olHistPushed = false;
  function renderOutline() {
    const ol = h('div', { class: 'outline' + (L.app.opts.outlinePlain ? ' plain' : ''), contenteditable: 'true', spellcheck: String(!!L.app.opts.spell), role: 'textbox', 'aria-multiline': 'true', 'aria-label': 'Outline' });
    L.pres.slides.forEach((s, i) => {
      const title = s.shapes.find((x) => x.ph && (x.ph.type === 'title' || x.ph.type === 'ctrTitle'));
      const row = h('div', { class: 'ol-row title', 'data-slide': s.id });
      row.appendChild(document.createTextNode(title ? L.txt.plain(title.tx).replace(/\n/g, ' ') : ''));
      if (!row.textContent) row.appendChild(h('br'));
      ol.appendChild(row);
      const body = bodyShape(s);
      if (body) for (const p of body.tx.ps) {
        const t = L.txt.paraText(p).replace(/\n/g, ' ');
        const r = h('div', { class: 'ol-row body', 'data-lvl': p.lvl || 0 });
        r.appendChild(document.createTextNode(t));
        if (!t) r.appendChild(h('br'));
        ol.appendChild(r);
      }
      void i;
    });
    ol.addEventListener('keydown', onOutlineKey);
    ol.addEventListener('input', () => { if (!olHistPushed) { L.hist.push('Outline Edit'); olHistPushed = true; } olDirty = true; commitSoon(); });
    ol.addEventListener('paste', (e) => { e.preventDefault(); const t = (e.clipboardData && e.clipboardData.getData('text/plain')) || ''; document.execCommand('insertText', false, t.replace(/\r?\n/g, ' ')); });
    ol.addEventListener('focusout', () => LP.commitOutline());
    ol.addEventListener('keyup', syncOutlineSlide);
    ol.addEventListener('click', syncOutlineSlide);
    L.clear(LP.body).appendChild(ol);
    LP.outline = ol;
    styleOutline();
  }
  function styleOutline() {
    if (!LP.outline) return;
    let n = L.pres.firstNum || 1;
    for (const r of LP.outline.children) {
      if (!r.classList) continue;
      if (r.classList.contains('title')) { r.dataset.num = n++; r.style.cssText = 'font-weight:bold;margin-top:6px;padding-left:36px;position:relative;min-height:1.35em'; }
      else { const lv = +r.dataset.lvl || 0; r.style.cssText = `padding-left:${48 + lv * 18}px;position:relative;min-height:1.35em;font-weight:${L.app.opts.outlinePlain ? 'normal' : 'normal'}`; }
    }
    let style = document.getElementById('ol-style');
    if (!style) { style = h('style', { id: 'ol-style' }); document.head.appendChild(style); }
    style.textContent = '.outline .ol-row.title::before{content:attr(data-num);position:absolute;left:0;width:16px;text-align:right;font-weight:normal}.outline .ol-row.title::after{content:"";position:absolute;left:20px;top:3px;width:12px;height:9px;border:1px solid #555;background:linear-gradient(#cfdcf3 0 35%,#fff 35%)}.outline .ol-row.body::before{content:"•";position:absolute;margin-left:-12px}';
  }
  const commitSoon = L.debounce(() => LP.commitOutline(true), 700);
  function rowCaret() {
    const s = window.getSelection();
    if (!s.rangeCount) return null;
    let n = s.anchorNode;
    while (n && n.parentNode !== LP.outline) n = n.parentNode;
    return n;
  }
  function syncOutlineSlide() {
    const row = rowCaret();
    if (!row) return;
    let idx = -1;
    for (const r of LP.outline.children) { if (r.classList.contains('title')) idx++; if (r === row) break; }
    idx = Math.max(0, idx);
    if (idx !== E().idx && idx < L.pres.slides.length) { E().idx = idx; E().render(); L.bus.emit('slide-changed-silent'); L.app.updateStatus(); }
  }
  function onOutlineKey(e) {
    const row = rowCaret();
    if (!row) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      if (!olHistPushed) { L.hist.push('Outline Edit'); olHistPushed = true; }
      const s = window.getSelection();
      const r = s.getRangeAt(0);
      r.deleteContents();
      const after = document.createRange();
      after.setStart(r.endContainer, r.endOffset);
      after.setEndAfter(row.lastChild || row);
      const tail = after.extractContents();
      const nr = h('div', { class: row.className, 'data-lvl': row.dataset.lvl || '0' });
      nr.appendChild(tail);
      if (!nr.textContent) { L.clear(nr); nr.appendChild(h('br')); }
      if (!row.textContent) { L.clear(row); row.appendChild(h('br')); }
      if (e.ctrlKey && row.classList.contains('title')) { nr.className = 'ol-row body'; nr.dataset.lvl = '0'; }
      row.after(nr);
      const c = document.createRange(); c.setStart(nr, 0); c.collapse(true); s.removeAllRanges(); s.addRange(c);
      olDirty = true; styleOutline(); commitSoon();
      return;
    }
    if (e.key === 'Tab' || (e.altKey && e.shiftKey && (e.key === 'ArrowRight' || e.key === 'ArrowLeft'))) {
      e.preventDefault();
      const demote = e.key === 'Tab' ? !e.shiftKey : e.key === 'ArrowRight';
      if (!olHistPushed) { L.hist.push('Outline Edit'); olHistPushed = true; }
      if (demote) { if (row.classList.contains('title')) { if (row.previousElementSibling) { row.className = 'ol-row body'; row.dataset.lvl = '0'; } } else row.dataset.lvl = Math.min(4, (+row.dataset.lvl || 0) + 1); }
      else if (row.classList.contains('body')) { const lv = +row.dataset.lvl || 0; if (lv === 0) row.className = 'ol-row title'; else row.dataset.lvl = lv - 1; }
      olDirty = true; styleOutline(); commitSoon();
    }
  }
  function bodyShape(s) { return s.shapes.find((x) => x.ph && x.type === 'text' && ['body', 'obj', 'subTitle'].includes(x.ph.type)); }
  LP.commitOutline = function (keepFocus) {
    if (!olDirty || !LP.outline) return;
    olDirty = false;
    const rows = Array.from(LP.outline.children).map((r) => ({ title: r.classList.contains('title'), lvl: +r.dataset.lvl || 0, text: r.textContent.replace(/ /g, ' ') }));
    const groups = [];
    for (const r of rows) {
      if (r.title || !groups.length) groups.push({ title: r.title ? r.text : '', body: r.title ? [] : [r] });
      else groups[groups.length - 1].body.push(r);
    }
    const pres = L.pres;
    groups.forEach((g, i) => {
      let s = pres.slides[i];
      if (!s) { s = M.newSlide(pres, 'text', (pres.slides[i - 1] || pres.slides[0] || {}).design); M.insertSlides(pres, pres.slides.length, [s]); }
      let t = s.shapes.find((x) => x.ph && (x.ph.type === 'title' || x.ph.type === 'ctrTitle'));
      if (!t && g.title) { M.applyLayout(pres, s, s.layout === 'blank' || s.layout === 'contentOnly' ? 'titleOnly' : s.layout); t = s.shapes.find((x) => x.ph && (x.ph.type === 'title' || x.ph.type === 'ctrTitle')); }
      if (t && L.txt.plain(t.tx).replace(/\n/g, ' ') !== g.title) L.txt.setPlain(t.tx, g.title);
      let b = bodyShape(s);
      if (!b && g.body.some((x) => x.text)) { M.applyLayout(pres, s, s.layout === 'titleOnly' || s.layout === 'blank' ? 'text' : s.layout); b = bodyShape(s); }
      if (b) {
        const old = b.tx.ps;
        b.tx.ps = (g.body.length ? g.body : [{ lvl: 0, text: '' }]).map((r, k) => {
          const src = old[k] || old[old.length - 1] || L.txt.para('');
          if (old[k] && L.txt.paraText(old[k]).replace(/\n/g, ' ') === r.text && (old[k].lvl || 0) === r.lvl) return old[k];
          const rp = src.rs[0] ? L.txt.runProps(src.rs[0]) : src.end || {};
          return { lvl: r.lvl, pp: L.clone(src.pp || {}), rs: r.text ? [Object.assign({ t: r.text }, rp)] : [], end: L.clone(rp) };
        });
      }
    });
    /* titles removed: drop slides that only held text */
    for (let i = pres.slides.length - 1; i >= groups.length; i--) {
      const s = pres.slides[i];
      const onlyText = s.shapes.every((x) => x.ph && x.type === 'text');
      if (onlyText && pres.slides.length > 1) pres.slides.splice(i, 1);
    }
    if (E().idx >= pres.slides.length) E().idx = pres.slides.length - 1;
    olHistPushed = false;
    E().render();
    L.bus.emit('slides-changed', { keepOutline: true });
    if (!keepFocus) renderOutline();
    else styleOutline();
  };
  L.app = L.app || {};

  /* ================= notes pane ================= */
  const NP = (P.notes = {});
  NP.mount = function (el) {
    NP.el = el;
    NP.ed = h('div', { class: 'notes-in', contenteditable: 'true', spellcheck: 'true', role: 'textbox', 'aria-multiline': 'true', 'aria-label': 'Speaker notes', 'data-ph': 'Click to add notes' });
    el.appendChild(NP.ed);
    NP.ed.addEventListener('input', () => {
      const s = E().slide();
      if (!s || s.isMaster) return;
      L.hist.push('Notes', 'notes:' + s.id);
      s.notes = NP.ed.innerText.replace(/\n$/, '');
      if (!NP.ed.textContent) NP.ed.innerHTML = '';
    });
    NP.ed.addEventListener('paste', (e) => { e.preventDefault(); document.execCommand('insertText', false, (e.clipboardData && e.clipboardData.getData('text/plain')) || ''); });
    NP.ed.addEventListener('focus', () => { if (L.te.active()) L.te.end(); });
  };
  NP.render = function () {
    if (!NP.ed) return;
    const s = E().slide();
    NP.ed.contentEditable = s && !s.isMaster ? 'true' : 'false';
    NP.ed.textContent = s && !s.isMaster ? s.notes || '' : '';
  };

  /* ================= Slide Sorter ================= */
  const SO = (P.sorter = { zoom: 0.25 });
  SO.mount = function (parent) {
    SO.el = h('div', { class: 'sorter', tabindex: '-1', 'aria-label': 'Slide Sorter' });
    parent.appendChild(SO.el);
    SO.el.addEventListener('pointerdown', (e) => { if (e.target === SO.el || e.target.classList.contains('sorter-grid')) { L.app.slideSel = new Set([E().idx]); SO.mark(); } });
    SO.el.addEventListener('focus', () => { L.app.focusArea = 'slides'; });
    SO.el.addEventListener('blur', () => { L.app.focusArea = null; });
  };
  SO.render = function () {
    if (!SO.el) return;
    const grid = h('div', { class: 'sorter-grid' });
    const w = Math.max(90, Math.round(SO.el.clientWidth ? Math.min(260, (SO.el.clientWidth - 60) / 4) : 180));
    L.pres.slides.forEach((s, i) => {
      const it = h('div', { class: 'so-it' + (L.app.slideSel.has(i) ? ' sel' : '') + (s.hidden ? ' hiddenslide' : ''), 'data-i': i });
      const fr = h('button', { class: 'so-frame', type: 'button', 'aria-label': `Slide ${i + 1}` });
      fr.appendChild(L.render.thumb(L.pres, s, w, { index: i }));
      const foot = h('div', { class: 'so-foot' });
      if (s.trans && s.trans.type && s.trans.type !== 'none') {
        const star = h('span', { class: 'th-star', 'data-tip': L.anim.transName(s.trans), html: '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M6 .8l1.5 3.2 3.4.4-2.5 2.3.7 3.4L6 8.4 2.9 10.1l.7-3.4L1.1 4.4l3.4-.4z" fill="none" stroke="#333"/></svg>' });
        star.addEventListener('click', (e) => { e.stopPropagation(); L.show.previewTransition(fr, L.pres, s, L.pres.slides[i - 1]); });
        foot.appendChild(star);
      }
      if (s.trans && s.trans.after != null) foot.appendChild(h('span', { text: fmtSecs(s.trans.after) }));
      foot.appendChild(h('span', { class: 'so-num' + (s.hidden ? ' hid' : ''), text: String(i + (L.pres.firstNum || 1)) }));
      it.append(fr, foot);
      it.addEventListener('pointerdown', (e) => onSorterDown(e, i, it));
      it.addEventListener('dblclick', () => { E().goto(i); L.app.setView('normal'); });
      it.addEventListener('contextmenu', (e) => { e.preventDefault(); if (!L.app.slideSel.has(i)) { L.app.slideSel = new Set([i]); E().idx = i; SO.mark(); } L.app.contextThumb(e); });
      grid.appendChild(it);
    });
    L.clear(SO.el).appendChild(grid);
  };
  const fmtSecs = (ms) => { const s = Math.round(ms / 1000); return `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`; };
  SO.mark = function () { L.$$('.so-it', SO.el).forEach((el) => el.classList.toggle('sel', L.app.slideSel.has(+el.dataset.i))); L.app.updateStatus(); L.ui.refresh(); };
  function onSorterDown(e, i, it) {
    if (e.button !== 0) return;
    SO.el.focus({ preventScroll: true });
    const sel = L.app.slideSel;
    if (e.shiftKey) { const a = Math.min(E().idx, i), b = Math.max(E().idx, i); L.app.slideSel = new Set(Array.from({ length: b - a + 1 }, (_, k) => a + k)); SO.mark(); return; }
    if (e.ctrlKey || e.metaKey) { if (sel.has(i) && sel.size > 1) sel.delete(i); else sel.add(i); SO.mark(); return; }
    if (!sel.has(i)) L.app.slideSel = new Set([i]);
    E().idx = i;
    SO.mark();
    const x0 = e.clientX, y0 = e.clientY;
    let marker = null, target = null;
    const mv = (ev) => {
      if (!marker && Math.hypot(ev.clientX - x0, ev.clientY - y0) < 6) return;
      if (!marker) { marker = h('div', { style: 'position:absolute;width:3px;background:#000;pointer-events:none;z-index:5' }); SO.el.appendChild(marker); }
      const items = L.$$('.so-it', SO.el);
      target = items.length;
      let ref = null, before = true;
      for (let k = 0; k < items.length; k++) {
        const r = items[k].getBoundingClientRect();
        if (ev.clientY >= r.top - 12 && ev.clientY <= r.bottom + 12 && ev.clientX < r.left + r.width / 2) { target = k; ref = items[k]; before = true; break; }
        if (ev.clientY >= r.top - 12 && ev.clientY <= r.bottom + 12) { ref = items[k]; before = false; target = k + 1; }
      }
      if (!ref) return;
      const br = SO.el.getBoundingClientRect(), rr = ref.getBoundingClientRect();
      marker.style.left = (before ? rr.left - 12 : rr.right + 10) - br.left + SO.el.scrollLeft + 'px';
      marker.style.top = rr.top - br.top + SO.el.scrollTop + 'px';
      marker.style.height = rr.height + 'px';
    };
    const up = () => {
      window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up);
      if (marker) { marker.remove(); if (target != null) L.app.moveSlides(Array.from(L.app.slideSel).sort((a, b) => a - b), target); }
    };
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
    void it;
  }

  /* ================= Notes Page view ================= */
  const NV = (P.notesView = {});
  NV.mount = function (parent) { NV.el = h('div', { class: 'notespage', 'aria-label': 'Notes Page' }); parent.appendChild(NV.el); };
  NV.render = function () {
    if (!NV.el) return;
    const s = E().slide();
    L.clear(NV.el);
    if (!s) return;
    const avail = Math.max(300, NV.el.clientWidth - 60);
    const pw = Math.min(avail, 640), ph = pw * (10 / 7.5);
    const page = h('div', { class: 'np-page', style: `width:${pw}px;height:${ph}px` });
    const sw = pw * 0.67, shh = (sw * L.pres.H) / L.pres.W;
    const sb = h('div', { class: 'np-slide', style: `left:${(pw - sw) / 2}px;top:${ph * 0.075}px;width:${sw}px;height:${shh}px` });
    sb.appendChild(L.render.thumb(L.pres, s, sw, { index: E().idx }));
    const ny = ph * 0.075 + shh + ph * 0.05;
    const notes = h('div', { class: 'np-notes', contenteditable: 'true', spellcheck: 'true', style: `left:${pw * 0.1}px;top:${ny}px;width:${pw * 0.8}px;height:${ph - ny - ph * 0.08}px;font-size:${Math.max(11, pw / 48)}px`, role: 'textbox', 'aria-label': 'Notes' });
    notes.textContent = s.notes || '';
    notes.addEventListener('input', () => { L.hist.push('Notes', 'notes:' + s.id); s.notes = notes.innerText.replace(/\n$/, ''); });
    notes.addEventListener('paste', (e) => { e.preventDefault(); document.execCommand('insertText', false, (e.clipboardData && e.clipboardData.getData('text/plain')) || ''); });
    page.append(sb, notes, h('div', { class: 'pv-hdr', style: `right:${pw * 0.05}px;bottom:${ph * 0.03}px`, text: String(E().idx + (L.pres.firstNum || 1)) }));
    NV.el.appendChild(page);
  };

  /* ================= Print Preview ================= */
  const PV = (P.preview = {});
  PV.mount = function (parent) {
    PV.el = h('div', { class: 'preview', 'aria-label': 'Print Preview' });
    PV.bar = h('div', { class: 'pv-bar' });
    PV.scroll = h('div', { class: 'pv-scroll' });
    PV.el.append(PV.bar, PV.scroll);
    parent.appendChild(PV.el);
  };
  PV.render = function () {
    if (!PV.el) return;
    const o = L.app.printOpts;
    L.clear(PV.bar);
    const canPrint = window.self === window.top;
    if (canPrint) PV.bar.appendChild(ui.button('&Print...', () => { document.body.classList.add('printing'); window.print(); setTimeout(() => document.body.classList.remove('printing'), 500); }));
    PV.bar.append(
      h('label', { text: 'Print What:' }),
      ui.select([['slides', 'Slides'], ['handouts', 'Handouts'], ['notes', 'Notes Pages'], ['outline', 'Outline View']], o.what, (v) => { o.what = v; PV.render(); }),
      h('label', { text: 'Per page:' }),
      ui.select([[1, '1'], [2, '2'], [3, '3'], [4, '4'], [6, '6'], [9, '9']], o.per, (v) => { o.per = +v; PV.render(); }, { disabled: o.what !== 'handouts' }),
      ui.select([['color', 'Color'], ['gray', 'Grayscale'], ['bw', 'Pure Black and White']], o.color, (v) => { o.color = v; PV.render(); }),
      ui.check('Frame slides', o.frame, (v) => { o.frame = v; PV.render(); }),
      ui.button('Save as P&DF...', () => L.app.previewPDF()),
      ui.button('&Close', () => L.app.setView(L.app.prevView || 'normal')));
    if (!canPrint) PV.bar.appendChild(h('span', { class: 'tp-note', text: 'To print, save these pages as a PDF and print the PDF.' }));
    L.clear(PV.scroll);
    const slides = L.pres.slides.map((s, i) => [s, i]).filter(([s, i]) => (o.hidden || !s.hidden) && (!o.range || o.range.includes(i)));
    const landscape = o.what === 'slides';
    const pw = Math.min(PV.scroll.clientWidth - 40 || 700, landscape ? 760 : 560), ph = landscape ? pw * (8.5 / 11) : pw * (11 / 8.5);
    const filt = o.color === 'gray' ? 'grayscale(1)' : o.color === 'bw' ? 'grayscale(1) contrast(30)' : '';
    const page = () => { const p = h('div', { class: 'pv-page', style: `width:${pw}px;height:${ph}px` }); if (filt) p.style.filter = filt; PV.scroll.appendChild(p); return p; };
    const cell = (p, s, i, x, y, w) => {
      const hh = (w * L.pres.H) / L.pres.W;
      const c = h('div', { class: 'pv-cell', style: `left:${x}px;top:${y}px;width:${w}px;height:${hh}px;${o.frame || o.what !== 'slides' ? '' : 'border-color:transparent'}` });
      c.appendChild(L.render.thumb(L.pres, s, w, { index: i }));
      p.appendChild(c);
      return hh;
    };
    const footer = (p, n) => { p.appendChild(h('div', { class: 'pv-hdr', style: `right:${pw * 0.05}px;bottom:${ph * 0.025}px`, text: String(n) })); };
    if (o.what === 'slides') {
      slides.forEach(([s, i]) => { const p = page(); const w = Math.min(pw * 0.9, ((ph * 0.9) * L.pres.W) / L.pres.H); const hh = (w * L.pres.H) / L.pres.W; cell(p, s, i, (pw - w) / 2, (ph - hh) / 2, w); });
    } else if (o.what === 'notes') {
      slides.forEach(([s, i], n) => {
        const p = page();
        const w = pw * 0.67; const hh = cell(p, s, i, (pw - w) / 2, ph * 0.07, w);
        p.appendChild(h('div', { class: 'pv-text', style: `left:${pw * 0.1}px;top:${ph * 0.07 + hh + ph * 0.04}px;width:${pw * 0.8}px;height:${ph * 0.4}px;font-size:${pw / 50}px`, text: s.notes || '' }));
        footer(p, n + 1);
      });
    } else if (o.what === 'outline') {
      let p = page(), y = ph * 0.06, pn = 1;
      const lh = pw / 40;
      const line = (t, indent, bold) => { if (y > ph * 0.92) { footer(p, pn++); p = page(); y = ph * 0.06; } p.appendChild(h('div', { class: 'pv-text', style: `left:${pw * 0.08 + indent}px;top:${y}px;width:${pw * 0.84 - indent}px;font-size:${lh * 0.85}px;${bold ? 'font-weight:bold' : ''}`, text: t })); y += lh * 1.3; };
      slides.forEach(([s, i]) => {
        line(`${i + 1}  ${M.slideTitle(s)}`, 0, true);
        const b = bodyShape(s);
        if (b) for (const para of b.tx.ps) { const t = L.txt.paraText(para); if (t) line('• ' + t, 24 + (para.lvl || 0) * 18, false); }
      });
      footer(p, pn);
    } else {
      const per = o.per;
      const layouts = { 1: [[0.5, 0.5, 0.8]], 2: [[0.5, 0.28, 0.6], [0.5, 0.72, 0.6]], 3: [[0.3, 0.2, 0.42], [0.3, 0.5, 0.42], [0.3, 0.8, 0.42]], 4: [[0.28, 0.3, 0.42], [0.72, 0.3, 0.42], [0.28, 0.7, 0.42], [0.72, 0.7, 0.42]], 6: [[0.28, 0.2, 0.4], [0.72, 0.2, 0.4], [0.28, 0.5, 0.4], [0.72, 0.5, 0.4], [0.28, 0.8, 0.4], [0.72, 0.8, 0.4]], 9: [[0.2, 0.2, 0.28], [0.5, 0.2, 0.28], [0.8, 0.2, 0.28], [0.2, 0.5, 0.28], [0.5, 0.5, 0.28], [0.8, 0.5, 0.28], [0.2, 0.8, 0.28], [0.5, 0.8, 0.28], [0.8, 0.8, 0.28]] }[per];
      for (let k = 0, pn = 1; k < slides.length; k += per, pn++) {
        const p = page();
        layouts.forEach(([cx, cy, wf], j) => {
          const it = slides[k + j];
          if (!it) return;
          const w = pw * wf, hh = (w * L.pres.H) / L.pres.W;
          cell(p, it[0], it[1], pw * cx - w / 2, ph * cy - hh / 2, w);
          if (per === 3) { const lines = h('div', { class: 'pv-lines', style: `left:${pw * 0.56}px;top:${ph * cy - hh / 2}px;width:${pw * 0.36}px;height:${hh}px` }); for (let q = 0; q < 7; q++) lines.appendChild(h('span')); p.appendChild(lines); }
        });
        footer(p, pn);
      }
    }
    if (!slides.length) PV.scroll.appendChild(h('div', { text: 'Nothing to print.', style: 'color:#fff' }));
  };

  /* ================= task panes ================= */
  const TP = (P.task = { stack: [], pos: -1, registry: {} });
  TP.mount = function (el) {
    TP.el = el;
    TP.back = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Back', html: L.icons.get('back') });
    TP.fwd = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Forward', html: L.icons.get('forward') });
    TP.home = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Home', html: L.icons.get('home') });
    TP.title = h('div', { class: 'tp-title', role: 'button', tabindex: '0', 'aria-haspopup': 'menu' }, h('span', { text: '' }), h('span', { class: 'dd-arrow' }));
    const close = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Close task pane', html: L.icons.get('close', 12) });
    TP.body = h('div', { class: 'tp-body' });
    el.append(h('div', { class: 'tp-head' }, TP.back, TP.fwd, TP.home, TP.title, close), TP.body);
    TP.back.addEventListener('click', () => TP.go(TP.pos - 1));
    TP.fwd.addEventListener('click', () => TP.go(TP.pos + 1));
    TP.home.addEventListener('click', () => TP.show('getting-started'));
    close.addEventListener('click', () => L.app.toggleTask(false));
    const openList = () => { const r = TP.title.getBoundingClientRect(); ui.openMenu(Object.values(TP.registry).filter((p) => !p.hidden).map((p) => ({ label: p.title, run: () => TP.show(p.id), checked: () => TP.current === p.id })), { left: r.left, bottom: r.bottom }); };
    TP.title.addEventListener('click', openList);
    TP.title.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openList(); } });
  };
  TP.register = (p) => { TP.registry[p.id] = p; };
  TP.show = function (id) {
    if (!TP.registry[id]) return;
    L.app.toggleTask(true);
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
    else if (p.live) { L.clear(TP.body); p.render(TP.body); }
  };
  const sec = (title, ...kids) => h('div', { class: 'tp-sec' }, title ? h('div', { class: 'tp-h' }, title) : null, ...kids);
  const link = (icon, label, fn) => h('button', { class: 'tp-link', type: 'button', onclick: fn }, icon ? h('span', { html: L.icons.get(icon) }) : null, h('span', { text: label }));

  /* --- Getting Started --- */
  /* ---------- Document Recovery: unsaved versions kept by common/suite.js ---------- */
  TP.register({ id: 'recovery', title: 'Document Recovery', render(b) { if (window.VO) VO.recoveryPane(b); } });
  TP.register({
    id: 'getting-started', title: 'Getting Started',
    render(b) {
      const q = h('input', { type: 'text', id: 'gs-search', placeholder: 'Search help', style: 'flex:1;min-width:0' });
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { P.helpQuery = q.value; TP.show('help'); } });
      b.append(
        h('div', { class: 'tp-row' }, h('span', { html: L.icons.app(32) }), h('b', { text: 'Lectern 2003 Web Edition', style: 'color:var(--pane-ink)' })),
        sec('Search for:', h('div', { class: 'tp-row' }, q, ui.button('Go', () => { P.helpQuery = q.value; TP.show('help'); }, { class: 'btn small' }))),
        sec('Open', window.VO ? VO.recentLinks(4) : null, link('open', 'Open a presentation (.pptx)...', () => ui.exec('open')), link('new', 'Create a new presentation...', () => TP.show('new')), link('slideshow', 'View the sample presentation', () => L.app.loadSample())),
        sec('Learn', link('help', 'Keyboard shortcuts and help', () => TP.show('help')), link('design', 'Apply a design template', () => TP.show('design')), link('animation', 'Animate your slides', () => TP.show('customAnim'))));
    },
  });

  /* --- New Presentation --- */
  TP.register({
    id: 'new', title: 'New Presentation',
    render(b) {
      b.append(
        sec('New', link('new', 'Blank presentation', () => L.app.newPresentation()), link('design', 'From design template', () => { L.app.newPresentation(); TP.show('design'); }), link('properties', 'From AutoContent wizard...', () => L.app.autoContent()), link('open', 'From existing presentation...', () => ui.exec('open')), link('photoAlbum', 'Photo album...', () => ui.exec('photoAlbum'))),
        sec('Templates', link('slideshow', 'Sample: café expansion plan', () => L.app.loadSample())));
    },
  });

  /* --- Help --- */
  const HELP = [
    ['Get around', 'Use the Slides tab on the left to move between slides; drag thumbnails to reorder. The Outline tab lets you type titles and bullets — press Tab to demote and Shift+Tab to promote. Enter on a title line starts a new slide.'],
    ['Add text', 'Click a placeholder ("Click to add title") and type. Use the Text Box tool on the Drawing toolbar to add text anywhere. Click a box once to start typing; drag its border to move it.'],
    ['Format text', 'Select text, then use the Formatting toolbar or Format ▸ Font. Bullets and numbering live under Format ▸ Bullets and Numbering. Press Tab at the start of a bullet to demote it.'],
    ['Draw shapes', 'Choose AutoShapes on the Drawing toolbar, then drag on the slide. Hold Shift for perfect squares, circles and 15° angles; hold Ctrl to draw from the center. Yellow diamonds adjust a shape\'s proportions; the green dot rotates.'],
    ['Pictures, tables and charts', 'Use the Insert menu or the icons inside a content placeholder. Double-click a chart to edit its datasheet. Double-click a picture to crop or recolor it.'],
    ['Design templates', 'Format ▸ Slide Design shows the design gallery. Color Schemes and Animation Schemes are linked at the top of that pane.'],
    ['Animation', 'Slide Show ▸ Custom Animation. Select an object, choose Add Effect, then set Start, Direction and Speed. Use By 1st Level Paragraphs (Effect Options ▸ Text Animation) to build bullets one at a time.'],
    ['Slide show', 'Press F5 to start from the first slide or Shift+F5 from the current slide. Click, Space or → advances; ← goes back; B blacks the screen; Ctrl+P switches to the pen; Esc ends the show.'],
    ['Files', 'File ▸ Open reads .pptx, .ppsx and .potx files. File ▸ Save writes a .pptx that opens in PowerPoint, Keynote, Google Slides and LibreOffice. Save As also offers .ppsx, .potx, a single-file web page, a PNG of the current slide and a text outline.'],
  ];
  const KEYS = [['Ctrl+N / Ctrl+O / Ctrl+S', 'New, Open, Save'], ['Ctrl+M', 'New slide'], ['Ctrl+D', 'Duplicate'], ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['Ctrl+X / C / V', 'Cut, Copy, Paste'], ['Ctrl+A', 'Select all'], ['Ctrl+G / Ctrl+Shift+G', 'Group / Ungroup'], ['Ctrl+B / I / U', 'Bold, Italic, Underline'], ['Ctrl+E / L / R / J', 'Center, Left, Right, Justify'], ['Ctrl+Shift+> / <', 'Grow / Shrink font'], ['Ctrl+T', 'Font dialog'], ['Shift+F3', 'Change case'], ['Ctrl+K', 'Hyperlink'], ['Ctrl+F / Ctrl+H', 'Find / Replace'], ['Tab / Shift+Tab', 'Demote / Promote; next object'], ['Arrow keys', 'Nudge (Ctrl for fine nudge)'], ['F2', 'Toggle between text and object selection'], ['F5 / Shift+F5', 'Slide show from start / current'], ['Page Up / Page Down', 'Previous / next slide'], ['Ctrl+F1', 'Task pane']];
  TP.register({
    id: 'help', title: 'Help',
    render(b) {
      const q = h('input', { type: 'text', id: 'hp-q', value: P.helpQuery || '', placeholder: 'Search', style: 'flex:1;min-width:0' });
      const res = h('div');
      const draw = () => {
        L.clear(res);
        const words = q.value.toLowerCase().split(/\s+/).filter(Boolean);
        const hit = (t) => words.every((w) => t.toLowerCase().includes(w));
        const topics = HELP.filter(([t, d]) => !words.length || hit(t + ' ' + d));
        for (const [t, d] of topics) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: t }), h('p', { text: d })));
        const keys = KEYS.filter(([k, d]) => !words.length || hit(k + ' ' + d));
        if (keys.length) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: 'Keyboard shortcuts' }), h('table', { class: 'kbd-table' }, ...keys.map(([k, d]) => h('tr', null, h('td', { text: k }), h('td', { text: d }))))));
        if (!res.children.length) res.appendChild(h('div', { class: 'tp-note', text: 'No topics match. Try a different word.' }));
      };
      q.addEventListener('input', draw);
      b.append(h('div', { class: 'tp-row' }, q), res);
      draw();
    },
  });

  /* --- Clip Art --- */
  TP.register({
    id: 'clipart', title: 'Clip Art',
    render(b) {
      const q = h('input', { type: 'text', id: 'ca-q', placeholder: 'e.g. people, growth, time', style: 'flex:1;min-width:0' });
      const grid = h('div', { class: 'clip-grid' });
      const draw = () => {
        L.clear(grid);
        const list = L.clipart.search(q.value);
        for (const it of list) {
          const c = h('button', { class: 'clip-it', type: 'button', 'data-tip': it.name }, h('img', { src: L.clipart.url(it), alt: it.name }), h('span', { text: it.name }));
          c.addEventListener('click', () => L.app.insertClipArt(it));
          grid.appendChild(c);
        }
        if (!list.length) grid.appendChild(h('div', { class: 'tp-note', text: 'No clips found.' }));
      };
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') draw(); });
      b.append(sec('Search for:', h('div', { class: 'tp-row' }, q, ui.button('Go', draw, { class: 'btn small' }))), sec('Results', grid), h('div', { class: 'tp-note', text: 'Click a clip to insert it on the current slide.' }));
      draw();
    },
  });

  /* --- Clipboard --- */
  TP.register({
    id: 'clipboard', title: 'Clipboard', live: true,
    render(b) {
      const items = L.app.clipHistory;
      b.append(h('div', { class: 'tp-btns' }, ui.button('Paste All', () => { for (const it of items.slice().reverse()) L.app.pasteItem(it); }, { class: 'btn small' }), ui.button('Clear All', () => { items.length = 0; TP.refresh(); }, { class: 'btn small' })), h('div', { class: 'tp-note', text: `${items.length} of 24 — Click an item to paste.` }));
      for (const it of items) {
        const el = h('button', { class: 'cb-item', type: 'button' }, h('span', { class: 'cb-ico', html: L.icons.get(it.kind === 'slides' ? 'sorterView' : it.kind === 'text' ? 'font' : 'autoshapes') }), h('span', { class: 'cb-t', text: it.label }));
        el.addEventListener('click', () => L.app.pasteItem(it));
        b.appendChild(el);
      }
    },
  });

  /* --- Slide Layout --- */
  function layoutSVG(key, w) {
    const d = L.model.buildDesign('default', 720, 540);
    const frames = M.layoutFrames(key, d);
    const k = w / 720, hh = 540 * k;
    let body = '';
    for (const f of frames) {
      const x = f.x * k, y = f.y * k, fw = f.w * k, fh = f.h * k;
      if (f.type === 'title' || f.type === 'ctrTitle') body += f.vert ? `<rect x="${x + fw * 0.35}" y="${y + 3}" width="${fw * 0.3}" height="${fh - 6}" fill="#7a7a7a"/>` : `<rect x="${x + fw * 0.15}" y="${y + fh * 0.35}" width="${fw * 0.7}" height="${Math.max(3, fh * 0.3)}" fill="#7a7a7a"/>`;
      else if (f.type === 'subTitle') body += `<rect x="${x + fw * 0.25}" y="${y + fh * 0.3}" width="${fw * 0.5}" height="${Math.max(2, fh * 0.12)}" fill="#a8a8a8"/>`;
      else if (f.type === 'body') { if (f.vert) for (let i = 0; i < 4; i++) body += `<rect x="${x + fw * 0.85 - i * fw * 0.15}" y="${y + 4}" width="2" height="${fh * 0.7}" fill="#a8a8a8"/>`; else for (let i = 0; i < 4; i++) body += `<circle cx="${x + 5}" cy="${y + 6 + i * Math.min(fh / 5, 9)}" r="1.2" fill="#555"/><rect x="${x + 9}" y="${y + 5 + i * Math.min(fh / 5, 9)}" width="${fw * 0.75}" height="2" fill="#a8a8a8"/>`; }
      else body += `<rect x="${x + 2}" y="${y + 2}" width="${fw - 4}" height="${fh - 4}" fill="none" stroke="#b0b0b0" stroke-dasharray="2 1"/><g transform="translate(${x + fw / 2 - 9},${y + fh / 2 - 6})">${f.type === 'tbl' ? '<rect width="18" height="12" fill="#fff" stroke="#3a5a8c"/><path d="M0 4h18M0 8h18M6 0v12M12 0v12" stroke="#3a5a8c"/>' : f.type === 'chart' ? '<rect x="1" y="6" width="4" height="6" fill="#9999ff"/><rect x="7" y="2" width="4" height="10" fill="#993366"/><rect x="13" y="4" width="4" height="8" fill="#ffffcc" stroke="#999" stroke-width=".5"/>' : '<rect x="0" y="0" width="5" height="5" fill="#9cb7e3"/><rect x="6.5" y="0" width="5" height="5" fill="#9999ff"/><rect x="13" y="0" width="5" height="5" fill="#9fd39a"/><rect x="0" y="7" width="5" height="5" fill="#bfe1ff"/><rect x="6.5" y="7" width="5" height="5" fill="#cfe0fb"/><rect x="13" y="7" width="5" height="5" fill="#ccc"/>'}</g>`;
    }
    return `<svg width="${w}" height="${hh}" viewBox="0 0 ${w} ${hh}">${body}</svg>`;
  }
  P.layoutSVG = layoutSVG;
  TP.register({
    id: 'layout', title: 'Slide Layout',
    render(b) {
      b.appendChild(h('div', { class: 'tp-note', text: 'Apply slide layout:' }));
      const groups = {};
      for (const l of M.LAYOUTS) (groups[l.group] = groups[l.group] || []).push(l);
      const cur = E().slide() ? E().slide().layout : null;
      for (const g in groups) {
        const grid = h('div', { class: 'lay-grid' });
        for (const l of groups[g]) {
          const it = h('button', { class: 'lay-it' + (l.key === cur ? ' on' : ''), type: 'button', 'data-tip': l.name, 'aria-label': l.name, html: layoutSVG(l.key, 64) });
          it.addEventListener('click', () => L.app.applyLayoutToSel(l.key));
          it.addEventListener('contextmenu', (e) => { e.preventDefault(); ui.openMenu([{ label: 'Apply to &Selected Slides', run: () => L.app.applyLayoutToSel(l.key) }, { label: '&Reapply Master Style', run: () => L.app.applyLayoutToSel(l.key, true) }, '-', { label: 'Insert &New Slide', run: () => L.app.insertSlide(l.key) }], { x: e.clientX, y: e.clientY }); });
          grid.appendChild(it);
        }
        b.appendChild(sec(g + ':', grid));
      }
      b.appendChild(ui.check('Show when inserting new slides', !!L.app.opts.layoutPaneOnNew, (v) => { L.app.opts.layoutPaneOnNew = v; L.app.saveOpts(); }));
    },
    refresh(b) { L.$$('.lay-it', b).forEach((el) => el.classList.toggle('on', E().slide() && el.getAttribute('aria-label') === M.layoutInfo(E().slide().layout).name)); },
  });

  /* --- Slide Design --- */
  const previewCache = {};
  function designPreview(key, w) {
    const ck = key + '|' + w + '|' + L.pres.W + 'x' + L.pres.H;
    if (previewCache[ck]) return previewCache[ck].cloneNode(true);
    const pres = M.newPresentation({ design: key, w: L.pres.W, h: L.pres.H, empty: true });
    const did = Object.keys(pres.designs)[0];
    const s = M.newSlide(pres, 'text', did);
    const t = s.shapes.find((x) => x.ph.type === 'title');
    L.txt.setPlain(t.tx, 'Title');
    const b = s.shapes.find((x) => x.ph.type === 'body');
    b.tx.ps = [L.txt.para('Bullet text'), L.txt.para('Second line')];
    pres.slides.push(s);
    const el = L.render.thumb(pres, s, w);
    previewCache[ck] = el;
    return el.cloneNode(true);
  }
  TP.register({
    id: 'design', title: 'Slide Design',
    render(b) {
      b.append(h('div', { class: 'col' }, link('design', 'Design Templates', () => TP.show('design')), link('colorPic', 'Color Schemes', () => TP.show('colors')), link('animation', 'Animation Schemes', () => TP.show('animSchemes'))));
      const cur = E().design();
      const used = Array.from(new Set(L.pres.slides.map((s) => s.design))).map((id) => L.pres.designs[id]).filter(Boolean);
      const usedGrid = h('div', { class: 'dsn-grid' });
      for (const d of used) {
        const it = h('button', { class: 'dsn-it' + (d === cur ? ' on' : ''), type: 'button', 'data-tip': d.name });
        const tmp = L.pres.slides.find((s) => s.design === d.id);
        it.append(L.render.thumb(L.pres, Object.assign({}, tmp, { shapes: tmp.shapes.filter((x) => x.ph && x.type === 'text') }), 100, { prompts: false }), h('span', { class: 'dsn-name', text: d.name }));
        it.addEventListener('click', () => { const sel = L.app.selectedSlides(); L.hist.push('Apply Design Template'); for (const s of sel) s.design = d.id; E().render(); L.bus.emit('slides-changed'); });
        usedGrid.appendChild(it);
      }
      b.appendChild(sec('Used in This Presentation', usedGrid));
      const grid = h('div', { class: 'dsn-grid' });
      for (const t of M.DESIGNS) {
        const it = h('button', { class: 'dsn-it', type: 'button', 'data-tip': t.name, 'aria-label': t.name });
        it.append(designPreview(t.key, 100), h('span', { class: 'dsn-name', text: t.name }), h('span', { class: 'dsn-dd', html: '<span class="dd-arrow" style="margin:0"></span>' }));
        it.addEventListener('click', (e) => {
          if (e.target.closest('.dsn-dd')) { ui.openMenu([{ label: 'Apply to &All Slides', run: () => L.app.applyDesign(t.key, true) }, { label: 'Apply to &Selected Slides', run: () => L.app.applyDesign(t.key, false) }], { x: e.clientX, y: e.clientY }); return; }
          L.app.applyDesign(t.key, true);
        });
        it.addEventListener('contextmenu', (e) => { e.preventDefault(); ui.openMenu([{ label: 'Apply to &All Slides', run: () => L.app.applyDesign(t.key, true) }, { label: 'Apply to &Selected Slides', run: () => L.app.applyDesign(t.key, false) }], { x: e.clientX, y: e.clientY }); });
        grid.appendChild(it);
      }
      b.appendChild(sec('Available For Use', grid));
      b.appendChild(link('open', 'Browse... (use the design of another .pptx/.potx)', () => L.app.browseDesign()));
    },
    refresh(b, reason) { if (reason === 'slides' || reason === 'design') { L.clear(b); this.render(b); } },
  });

  /* --- Color Schemes --- */
  TP.register({
    id: 'colors', title: 'Slide Design - Color Schemes',
    render(b) {
      b.append(h('div', { class: 'col' }, link('design', 'Design Templates', () => TP.show('design')), link('colorPic', 'Color Schemes', () => TP.show('colors')), link('animation', 'Animation Schemes', () => TP.show('animSchemes'))));
      const grid = h('div', { class: 'scheme-grid' });
      const d0 = E().design();
      const schemes = [d0.colors].concat(M.COLOR_SCHEMES);
      schemes.forEach((cs, i) => {
        const it = h('button', { class: 'scheme-it', type: 'button', 'aria-label': i === 0 ? 'Current scheme' : 'Color scheme ' + i, 'data-tip': i === 0 ? 'Current color scheme' : 'Apply this color scheme' });
        for (const k of ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'hlink', 'folHlink']) it.appendChild(h('span', { style: `background:${cs[k]}` }));
        it.addEventListener('click', () => L.app.applyColorScheme(cs, true));
        it.addEventListener('contextmenu', (e) => { e.preventDefault(); ui.openMenu([{ label: 'Apply to &All Slides', run: () => L.app.applyColorScheme(cs, true) }, { label: 'Apply to &Selected Slides', run: () => L.app.applyColorScheme(cs, false) }], { x: e.clientX, y: e.clientY }); });
        grid.appendChild(it);
      });
      b.append(sec('Apply a color scheme:', grid), link('options', 'Edit Color Schemes...', () => L.app.editColorScheme()));
    },
    refresh(b, reason) { if (reason === 'design') { L.clear(b); this.render(b); } },
  });

  /* --- Animation Schemes --- */
  const AS = [
    ['No Animation', null], ['Subtle', 'group'],
    ['Appear', { tr: { type: 'cut' }, title: ['appear'], body: ['appear'] }], ['Appear and dim', { tr: { type: 'fade' }, title: ['appear'], body: ['appear'] }],
    ['Fade in all', { tr: { type: 'fade' }, title: ['fade'], body: ['fade', null, 'all'] }], ['Fade in one by one', { tr: { type: 'fade' }, title: ['fade'], body: ['fade'] }],
    ['Faded wipe', { tr: { type: 'fade' }, title: ['fade'], body: ['wipe', 'b'] }], ['Faded zoom', { tr: { type: 'fade' }, title: ['fade'], body: ['zoom'] }],
    ['Dissolve in', { tr: { type: 'dissolve' }, title: ['dissolve'], body: ['dissolve'] }], ['Random bars', { tr: { type: 'randomBar', dir: 'horz' }, title: ['randomBars', 'horz'], body: ['randomBars', 'horz'] }],
    ['Wipe', { tr: { type: 'wipe', dir: 'r' }, title: ['wipe', 'l'], body: ['wipe', 't'] }], ['Blinds', { tr: { type: 'blinds', dir: 'horz' }, title: ['blinds', 'horz'], body: ['blinds', 'horz'] }],
    ['Moderate', 'group'],
    ['Ascend', { tr: { type: 'fade' }, title: ['ascend'], body: ['ascend'] }], ['Descend', { tr: { type: 'fade' }, title: ['descend'], body: ['descend'] }],
    ['Zoom', { tr: { type: 'zoom', dir: 'out' }, title: ['zoom'], body: ['zoom'] }], ['Expand', { tr: { type: 'fade' }, title: ['expand'], body: ['expand'] }],
    ['Spinner', { tr: { type: 'fade' }, title: ['spinner'], body: ['spinner'] }], ['Peek in', { tr: { type: 'cover', dir: 'u' }, title: ['peekIn', 'b'], body: ['peekIn', 'b'] }],
    ['Exciting', 'group'],
    ['Fly in', { tr: { type: 'push', dir: 'u' }, title: ['flyIn', 't'], body: ['flyIn', 'l'] }], ['Checkerboard', { tr: { type: 'checker', dir: 'horz' }, title: ['checkerboard', 'across'], body: ['checkerboard', 'across'] }],
    ['Wheel', { tr: { type: 'wheel', spokes: 4 }, title: ['wheel', '4'], body: ['wheel', '1'] }], ['Plus', { tr: { type: 'plus' }, title: ['plus', 'out'], body: ['plus', 'out'] }],
    ['Diamond', { tr: { type: 'diamond' }, title: ['diamond', 'out'], body: ['diamond', 'out'] }], ['Box out', { tr: { type: 'zoom', dir: 'out' }, title: ['box', 'out'], body: ['box', 'out'] }],
  ];
  P.ANIM_SCHEMES = AS;
  TP.register({
    id: 'animSchemes', title: 'Slide Design - Animation Schemes',
    render(b) {
      b.append(h('div', { class: 'col' }, link('design', 'Design Templates', () => TP.show('design')), link('colorPic', 'Color Schemes', () => TP.show('colors')), link('animation', 'Animation Schemes', () => TP.show('animSchemes'))));
      const list = h('div', { class: 'list-box trans-list', style: 'max-height:300px' });
      let pick = null;
      AS.forEach(([name, def]) => {
        if (def === 'group') { list.appendChild(h('div', { class: 'li-group', text: name })); return; }
        const li = h('div', { class: 'li', text: name });
        li.addEventListener('click', () => { L.$$('.li', list).forEach((x) => x.classList.remove('on')); li.classList.add('on'); pick = def; L.app.applyAnimScheme(def, false); if (L.app.opts.autoPreview) P.playSlide(); });
        list.appendChild(li);
      });
      b.append(sec('Apply to selected slides:', list),
        h('div', { class: 'tp-btns' }, ui.button('Apply to All Slides', () => L.app.applyAnimScheme(pick, true), { class: 'btn tp-btn' })),
        h('div', { class: 'tp-btns' }, ui.button('▶ Play', () => P.playSlide(), { class: 'btn small' }), ui.button('Slide Show', () => ui.exec('showFromCurrent'), { class: 'btn small' })),
        ui.check('AutoPreview', !!L.app.opts.autoPreview, (v) => { L.app.opts.autoPreview = v; L.app.saveOpts(); }));
    },
  });

  /* --- Custom Animation --- */
  const CA = { selId: null };
  function effectMenu(cls, onPick) {
    return L.anim.list(cls).map(([k, n]) => ({ label: n, run: () => onPick(cls, k) }));
  }
  P.addEffectMenu = (onPick) => [
    { label: '&Entrance', html: '<span style="color:#2f9a3a;font-size:13px">★</span>', sub: effectMenu('entr', onPick) },
    { label: 'E&mphasis', html: '<span style="color:#c9a100;font-size:13px">★</span>', sub: effectMenu('emph', onPick) },
    { label: 'E&xit', html: '<span style="color:#d33a2c;font-size:13px">★</span>', sub: effectMenu('exit', onPick) },
    { label: 'Motion &Paths', html: '<span style="color:#555;font-size:13px">☆</span>', sub: effectMenu('path', onPick) },
  ];
  TP.register({
    id: 'customAnim', title: 'Custom Animation',
    render(b) { renderCA(b); },
    refresh(b) { renderCA(b); },
  });
  function renderCA(b) {
    L.clear(b);
    const slide = E().slide();
    if (!slide || slide.isMaster) { b.appendChild(h('div', { class: 'tp-note', text: 'Custom animation is available in Normal view.' })); return; }
    const anims = slide.anims || (slide.anims = []);
    const shapes = E().selected();
    const add = ui.tbDrop('Add Effect', 'animation', (r) => {
      if (!E().selected().length) { ui.msg('Select an object on the slide first, then choose an effect.'); return; }
      ui.openMenu(P.addEffectMenu((cls, eff) => L.app.addAnimation(cls, eff)), r);
    }, 'Add Effect');
    const sel = anims.find((a) => a.id === CA.selId) || null;
    const remove = ui.button('Remove', () => { if (!sel) return; L.hist.push('Remove Effect'); slide.anims = anims.filter((a) => a.id !== sel.id); CA.selId = null; renderCA(b); L.bus.emit('slide-modified', E().idx); }, { class: 'btn small', disabled: !sel });
    b.append(h('div', { class: 'tp-row' }, add, remove));
    const upd = (k, v) => { if (!sel) return; L.hist.push('Modify Effect'); sel[k] = v; renderCA(b); if (L.app.opts.autoPreview) P.playSlide(sel.id); };
    const info = sel ? L.anim.info(sel.cls, sel.eff) : null;
    const start = ui.select([['click', 'On Click'], ['with', 'With Previous'], ['after', 'After Previous']], sel ? sel.start : 'click', (v) => upd('start', v), { disabled: !sel });
    const dir = ui.select(info && info.dirs ? info.dirs.map((d) => [d[0], d[1]]) : [['', '']], sel && sel.dir, (v) => upd('dir', v), { disabled: !(info && info.dirs) });
    const spd = ui.select(L.anim.SPEEDS.map(([n, ms]) => [ms, n]), sel ? L.anim.SPEEDS.reduce((best, s) => (Math.abs(s[1] - sel.dur) < Math.abs(best - sel.dur) ? s[1] : best), 500) : 1000, (v) => upd('dur', +v), { disabled: !sel || (sel.eff === 'appear' || sel.eff === 'disappear') });
    b.append(sec(sel ? 'Modify: ' + L.anim.name(sel) : 'Modify effect', h('div', { class: 'tp-row' }, h('label', { text: 'Start:' }), start), h('div', { class: 'tp-row' }, h('label', { text: 'Direction:' }), dir), h('div', { class: 'tp-row' }, h('label', { text: 'Speed:' }), spd)));
    const list = h('div', { class: 'list-box', style: 'min-height:150px' });
    let step = 0;
    anims.forEach((a, i) => {
      const sh = M.shapeById(slide, a.sid);
      if (a.start === 'click' || i === 0) step += a.start === 'click' ? 1 : 0;
      const label = (sh ? sh.name : '?') + (sh && sh.tx ? ': ' + L.txt.plain(sh.tx).slice(0, 24) : '');
      const col = { entr: '#2f9a3a', emph: '#c9a100', exit: '#d33a2c', path: '#555' }[a.cls];
      const li = h('div', { class: 'li' + (a.id === CA.selId ? ' on' : ''), 'data-tip': `${L.anim.name(a)} — ${label}` },
        h('span', { class: 'li-num', text: a.start === 'click' ? String(step) : '' }),
        h('span', { class: 'li-ico', html: a.start === 'click' ? '<svg width="10" height="14" viewBox="0 0 10 14"><rect x="1" y="1" width="8" height="12" rx="4" fill="#fff" stroke="#555"/><path d="M5 1v5" stroke="#555"/></svg>' : a.start === 'after' ? '<svg width="12" height="12" viewBox="0 0 12 12"><circle cx="6" cy="6" r="5" fill="#fff" stroke="#555"/><path d="M6 3v3l2 1" stroke="#555" fill="none"/></svg>' : '' }),
        h('span', { class: 'li-ico', html: `<span style="color:${col}">★</span>` }),
        h('span', { class: 'li-n', text: `${L.anim.name(a)}${a.by === 'para' ? ' ¶' : ''} — ${label}` }));
      li.addEventListener('click', () => { CA.selId = a.id; if (sh && E().topOf(a.sid)) E().select([E().topOf(a.sid).id]); renderCA(b); });
      li.addEventListener('dblclick', () => openOptions(a));
      li.addEventListener('contextmenu', (e) => {
        e.preventDefault(); CA.selId = a.id; renderCA(b);
        ui.openMenu([
          { label: 'Start On &Click', checked: () => a.start === 'click', run: () => upd('start', 'click') },
          { label: 'Start &With Previous', checked: () => a.start === 'with', run: () => upd('start', 'with') },
          { label: 'Start &After Previous', checked: () => a.start === 'after', run: () => upd('start', 'after') },
          '-', { label: '&Effect Options...', run: () => openOptions(a) }, { label: '&Timing...', run: () => openOptions(a) }, '-', { label: '&Remove', run: () => { L.hist.push('Remove Effect'); slide.anims = slide.anims.filter((x) => x.id !== a.id); CA.selId = null; renderCA(b); } },
        ], { x: e.clientX, y: e.clientY });
      });
      list.appendChild(li);
    });
    if (!anims.length) list.appendChild(h('div', { class: 'tp-note', style: 'padding:8px', text: shapes.length ? 'Select an element of the slide, then click "Add Effect" to add animation.' : 'Select an element of the slide, then click "Add Effect" to add animation.' }));
    const move = (d) => { if (!sel) return; const i = anims.indexOf(sel), j = i + d; if (j < 0 || j >= anims.length) return; L.hist.push('Reorder Effects'); [anims[i], anims[j]] = [anims[j], anims[i]]; renderCA(b); };
    b.append(list,
      h('div', { class: 'tp-row' }, h('span', { text: 'Re-Order' }), ui.button('↑', () => move(-1), { class: 'btn small', 'aria-label': 'Move up' }), ui.button('↓', () => move(1), { class: 'btn small', 'aria-label': 'Move down' })),
      h('div', { class: 'tp-btns' }, ui.button('▶ Play', () => P.playSlide(), { class: 'btn small' }), ui.button('Slide Show', () => ui.exec('showFromCurrent'), { class: 'btn small' })),
      ui.check('AutoPreview', !!L.app.opts.autoPreview, (v) => { L.app.opts.autoPreview = v; L.app.saveOpts(); }));
    function openOptions(a) { L.dlg.effectOptions(a, (o) => { L.hist.push('Effect Options'); Object.assign(a, o); renderCA(b); if (L.app.opts.autoPreview) P.playSlide(a.id); }); }
  }
  P.selectAnim = (id) => { CA.selId = id; };

  /* --- Slide Transition --- */
  TP.register({
    id: 'transition', title: 'Slide Transition',
    render(b) { renderTR(b); },
    refresh(b, reason) { if (reason !== 'trans') renderTR(b); },
  });
  function renderTR(b) {
    L.clear(b);
    const slide = E().slide();
    if (!slide || slide.isMaster) { b.appendChild(h('div', { class: 'tp-note', text: 'Transitions are set in Normal or Slide Sorter view.' })); return; }
    const tr = slide.trans || {};
    const curName = L.anim.transName(tr);
    const list = h('div', { class: 'list-box trans-list', style: 'height:220px' });
    for (const t of L.anim.TRANSITIONS) {
      const li = h('div', { class: 'li' + (t.name === curName ? ' on' : ''), text: t.name });
      li.addEventListener('click', () => {
        L.$$('.li', list).forEach((x) => x.classList.remove('on')); li.classList.add('on');
        L.app.setTransition((s) => { const keep = { spd: s.trans.spd, click: s.trans.click, after: s.trans.after }; s.trans = Object.assign({}, keep, { type: t.type, dir: t.dir, thruBlk: t.thruBlk, spokes: t.spokes, orient: t.orient }); for (const k in s.trans) if (s.trans[k] === undefined) delete s.trans[k]; });
        if (L.app.opts.autoPreview !== false) P.previewTransition();
      });
      list.appendChild(li);
    }
    setTimeout(() => { const on = list.querySelector('.li.on'); if (on) on.scrollIntoView({ block: 'nearest' }); }, 0);
    const spd = ui.select([['slow', 'Slow'], ['med', 'Medium'], ['fast', 'Fast']], tr.spd || 'fast', (v) => L.app.setTransition((s) => { s.trans.spd = v; delete s.trans.dur; }));
    const snd = ui.select([['', '[No Sound]']], '', null, { disabled: true });
    const onClick = ui.check('On mouse click', tr.click !== false, (v) => L.app.setTransition((s) => { s.trans.click = v; }));
    const secs = ui.spin({ value: tr.after != null ? tr.after / 1000 : 0, min: 0, max: 3600, step: 1, dec: 1, unit: ' s', onChange: (v) => { if (auto.input.checked) L.app.setTransition((s) => { s.trans.after = v * 1000; }); } });
    const auto = ui.check('Automatically after', tr.after != null, (v) => L.app.setTransition((s) => { s.trans.after = v ? secs.get() * 1000 : null; }));
    b.append(sec('Apply to selected slides:', list),
      sec('Modify transition', h('div', { class: 'tp-row' }, h('label', { text: 'Speed:' }), spd), h('div', { class: 'tp-row' }, h('label', { text: 'Sound:' }), snd)),
      sec('Advance slide', onClick, h('div', { class: 'tp-row' }, auto, secs)),
      h('div', { class: 'tp-btns' }, ui.button('Apply to All Slides', () => L.app.transitionToAll(), { class: 'btn tp-btn' })),
      h('div', { class: 'tp-btns' }, ui.button('▶ Play', () => P.previewTransition(), { class: 'btn small' }), ui.button('Slide Show', () => ui.exec('showFromCurrent'), { class: 'btn small' })),
      ui.check('AutoPreview', L.app.opts.autoPreview !== false, (v) => { L.app.opts.autoPreview = v; L.app.saveOpts(); }));
  }

  /* --- previews over the editor --- */
  P.previewTransition = function () {
    const ed = E();
    if (L.app.view === 'sorter') { const it = document.querySelector(`.so-it[data-i="${ed.idx}"] .so-frame`); if (it) L.show.previewTransition(it, L.pres, ed.slide(), L.pres.slides[ed.idx - 1]); return; }
    const host = ed.host;
    if (!host || !ed.slide()) return;
    const ov = h('div', { style: `position:absolute;left:${host.style.left};top:${host.style.top};width:${host.style.width};height:${host.style.height};overflow:hidden;z-index:4;pointer-events:none` });
    ed.canvas.appendChild(ov);
    L.show.previewTransition(ov, L.pres, ed.slide(), L.pres.slides[ed.idx - 1]);
    setTimeout(() => ov.remove(), L.anim.transDur(ed.slide().trans) + 400);
  };
  P.playSlide = function (onlyId) {
    const ed = E();
    const host = ed.host;
    const slide = ed.slide();
    if (!host || !slide) return;
    const ov = h('div', { style: `position:absolute;left:${host.style.left};top:${host.style.top};width:${host.style.width};height:${host.style.height};overflow:hidden;z-index:4;background:#fff` });
    ed.canvas.appendChild(ov);
    const stg = h('div', { style: `position:absolute;left:0;top:0;width:${L.pres.W}px;height:${L.pres.H}px;transform:scale(${ed.scale});transform-origin:0 0` });
    ov.appendChild(stg);
    L.show.playInline(stg, slide, { only: onlyId, done: () => ov.remove() });
    ov.addEventListener('pointerdown', () => ov.remove());
  };
})();
