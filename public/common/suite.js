/* VibeOffice suite services, shared by the start page and the apps (classic script, window.VO).
 *
 * - Which app opens which file (VO.APPS, VO.appFor).
 * - The service worker (public/sw.js), registered from every page.
 * - Files the start page hands to an app, and the Recent Files list. Both live in IndexedDB
 *   ('vibeoffice', stores 'handoff' and 'recent') in this browser only; nothing is uploaded.
 *
 * An app page is opened as  <app>/?new           a blank document instead of the sample
 *                           <app>/?template=<id>  a new document from one of the app's templates
 *                           <app>/?open=<id>      a file handed over by VO.handoff (read once, then deleted)
 *                           <app>/?recent=<key>   a Recent Files entry
 * The app reads VO.launch while it starts, then calls VO.attach(app, { open, thumb }), and reports
 * VO.opened(file) after it opens a file and VO.saved(name, blob) after it saves in its own format.
 */
(function () {
  'use strict';
  const VO = (window.VO = {});
  const BASE = new URL('..', document.currentScript.src).href;   // public/
  VO.base = BASE;

  /* ------------------------------------------------------------ apps */
  VO.APPS = {
    quire: { name: 'Quire', kind: 'Document', ext: ['docx', 'docm', 'dotx', 'dotm', 'doc', 'rtf', 'htm', 'html', 'txt'] },
    ledger: { name: 'Ledger', kind: 'Spreadsheet', ext: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'csv', 'tsv', 'prn', 'xml'] },
    lectern: { name: 'Lectern', kind: 'Presentation', ext: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'ppt'] },
  };
  VO.ext = (name) => (/\.([^./\\]+)$/.exec(name || '') || ['', ''])[1].toLowerCase();
  /** the app for a file name, or null */
  VO.appFor = (name) => { const e = VO.ext(name); return Object.keys(VO.APPS).find((a) => VO.APPS[a].ext.includes(e)) || null; };
  VO.accept = () => Object.values(VO.APPS).flatMap((a) => a.ext.map((e) => '.' + e)).join(',');
  /** VO.url('quire', { new: true }) -> .../quire/?new */
  VO.url = (app, params) => {
    const qs = Object.keys(params || {}).map((k) => (params[k] === true ? k : k + '=' + encodeURIComponent(params[k]))).join('&');
    return new URL(app + '/', BASE).href + (qs ? '?' + qs : '');
  };

  /** the page's icon (browser tab) from an inline <svg> */
  VO.setIcon = (svg) => {
    if (!/xmlns=/.test(svg)) svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
    let link = document.querySelector('link[rel="icon"]');
    if (!link) link = document.head.appendChild(Object.assign(document.createElement('link'), { rel: 'icon' }));
    link.type = 'image/svg+xml';
    link.href = 'data:image/svg+xml,' + encodeURIComponent(svg);
  };

  /* ------------------------------------------------------------ service worker */
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
    navigator.serviceWorker.register(BASE + 'sw.js').catch((e) => console.warn('service worker:', e.message));
  }

  /* ------------------------------------------------------------ IndexedDB */
  let dbp = null;
  function db() {
    return dbp || (dbp = new Promise((resolve, reject) => {
      const r = indexedDB.open('vibeoffice', 1);
      r.onupgradeneeded = () => {
        r.result.createObjectStore('recent', { keyPath: 'key' });
        r.result.createObjectStore('handoff', { keyPath: 'id' });
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    }));
  }
  /** run fn(store) in one transaction; resolves with the result of the request fn returns, once committed */
  async function tx(name, mode, fn) {
    const t = (await db()).transaction(name, mode);
    return new Promise((resolve, reject) => {
      let out;
      const req = fn(t.objectStore(name));
      if (req) req.onsuccess = () => { out = req.result; };
      t.oncomplete = () => resolve(out);
      t.onerror = t.onabort = () => reject(t.error);
    });
  }
  const chan = 'BroadcastChannel' in window ? new BroadcastChannel('vibeoffice') : null;
  const changed = () => { if (chan) chan.postMessage('recent'); };
  /** call fn when another page changes the Recent Files list */
  VO.onRecentChange = (fn) => { if (chan) chan.addEventListener('message', (e) => { if (e.data === 'recent') fn(); }); };

  /* ------------------------------------------------------------ Recent Files */
  const MAX_RECENT = 30;
  const MAX_BYTES = 50 * 1024 * 1024;   // bigger files are not kept (open them from disk again)
  VO.recent = {
    /** newest first: { key, app, name, size, time, file: Blob, thumb: Blob|null } */
    async list() {
      const all = (await tx('recent', 'readonly', (s) => s.getAll())) || [];
      return all.sort((a, b) => b.time - a.time);
    },
    get: (key) => tx('recent', 'readonly', (s) => s.get(key)),
    /** add or refresh the entry for this file name (a newer copy replaces the older one, its thumbnail stays) */
    async put(app, name, blob) {
      if (!app || !blob || blob.size > MAX_BYTES) return null;
      const key = app + ':' + name;
      const old = await VO.recent.get(key);
      await tx('recent', 'readwrite', (s) => s.put({ key, app, name, size: blob.size, time: Date.now(), file: blob, thumb: old ? old.thumb : null }));
      const all = await VO.recent.list();
      if (all.length > MAX_RECENT) await tx('recent', 'readwrite', (s) => { for (const e of all.slice(MAX_RECENT)) s.delete(e.key); });
      changed();
      return key;
    },
    async setThumb(key, thumb) {
      const e = await VO.recent.get(key);
      if (!e) return;
      e.thumb = thumb;
      await tx('recent', 'readwrite', (s) => s.put(e));
      changed();
    },
    async remove(key) { await tx('recent', 'readwrite', (s) => s.delete(key)); changed(); },
    async clear() { await tx('recent', 'readwrite', (s) => s.clear()); changed(); },
  };

  /* ------------------------------------------------------------ handing files to apps */
  /** where an app page opens: a new window when installed (the start page stays, as in LibreOffice), else this tab */
  VO.go = (url) => {
    const standalone = matchMedia('(display-mode: standalone)').matches || matchMedia('(display-mode: window-controls-overlay)').matches;
    if (standalone && window.open(url, '_blank')) return;
    location.href = url;
  };
  /** open a file in the app made for it; false if no app opens this kind of file */
  VO.handoff = async (file) => {
    const app = VO.appFor(file.name);
    if (!app) return false;
    const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    const day = Date.now() - 86400000;
    await tx('handoff', 'readwrite', (s) => {
      s.put({ id, name: file.name, file, time: Date.now() });
      s.openCursor().onsuccess = (e) => { const c = e.target.result; if (!c) return; if (c.value.time < day) c.delete(); c.continue(); };
    });
    VO.go(VO.url(app, { open: id }));
    return true;
  };

  /* ------------------------------------------------------------ the app side */
  const q = new URLSearchParams(location.search);
  /** how this app page was opened: null (plain: show the sample), { new }, { template: id }, { open: id } or { recent: key } */
  VO.launch = q.has('open') ? { open: q.get('open') } : q.has('recent') ? { recent: q.get('recent') }
    : q.has('template') ? { template: q.get('template') } : q.has('new') ? { new: true } : null;

  let current = null;   // { app, hooks }
  /** an app is ready: open the file it was launched with, if any.
   *  hooks: open(file), thumb() -> a canvas (or image) of the first page, message(text),
   *  templates() -> [{ id, name }] (read by tools/templates.mjs) */
  VO.attach = async (app, hooks) => {
    current = { app, hooks: hooks || {} };
    const l = VO.launch;
    if (!l || l.new) return;
    if (history.replaceState) history.replaceState(null, '', location.pathname);   // a reload starts clean
    if (l.template) return;                                                          // built by the app as it started
    let rec = null;
    try {
      if (l.open) {
        rec = await tx('handoff', 'readonly', (s) => s.get(l.open));
        if (rec) await tx('handoff', 'readwrite', (s) => s.delete(l.open));
      } else rec = await VO.recent.get(l.recent);
    } catch (e) { console.error(e); }
    if (!rec || !rec.file) { (current.hooks.message || alert)('That file is no longer available here. Open it again from your computer.'); return; }
    const f = rec.file instanceof File ? rec.file : new File([rec.file], rec.name, { type: rec.file.type });
    await current.hooks.open(f.name === rec.name ? f : new File([f], rec.name, { type: f.type }));
  };
  /** File ▸ New… inside an app: Blank and the app's templates with their previews (the Start Center's
   *  Create File gallery, common/templates.json), in the app's own dialog. create(id): id '' is Blank. */
  VO.newDialog = async (app, create) => {
    const ui = window.L && window.L.ui;
    if (!ui || !ui.dialog) return create('');
    let list = null;
    try { list = (await (await fetch(BASE + 'common/templates.json')).json())[app]; } catch (e) { console.warn('templates:', e); }
    if (!list || !list.length) list = [{ id: '', name: 'Blank', img: null }];
    if (!document.getElementById('vo-gallery-css')) {
      const st = document.createElement('style');
      st.id = 'vo-gallery-css';
      st.textContent = '.vo-gal{display:grid;grid-template-columns:repeat(auto-fill,minmax(132px,1fr));gap:8px;max-height:min(62vh,540px);overflow:auto;padding:6px;background:#fff;border:1px solid #7F9DB9}'
        + '.vo-tpl{display:flex;flex-direction:column;gap:4px;padding:5px;border:1px solid transparent;border-radius:2px;background:none;text-align:left;cursor:default;font:inherit;color:#000}'
        + '.vo-tpl:hover,.vo-tpl:focus-visible{background:#FFEEC2;border-color:#E3A23B;outline:none}'
        + '.vo-tpl span{display:flex;align-items:center;justify-content:center;height:112px;background:#f3f6fb;border:1px solid #b8c7df;overflow:hidden}'
        + '.vo-tpl img{max-width:100%;max-height:100%;box-shadow:0 1px 2px rgba(0,0,0,.25)}'
        + '.vo-tpl b{font-weight:normal;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
        + '.vo-gal-note{margin:0 0 6px}';
      document.head.appendChild(st);
    }
    const gal = document.createElement('div');
    gal.className = 'vo-gal';
    let d = null;
    for (const t of list) {
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'vo-tpl'; b.title = t.name;
      const pic = document.createElement('span');
      if (t.img) { const img = document.createElement('img'); img.src = BASE + t.img; img.alt = ''; pic.append(img); }
      const name = document.createElement('b');
      name.textContent = t.name;
      b.append(pic, name);
      b.addEventListener('click', () => { if (d) d.close(-1); create(t.id); });
      gal.append(b);
    }
    const body = document.createElement('div');
    const note = document.createElement('p');
    note.className = 'vo-gal-note';
    note.textContent = 'Choose Blank or a template:';
    body.append(note, gal);
    d = ui.dialog({ title: 'New ' + VO.APPS[app].kind, width: 700, body, buttons: [{ label: 'Cancel', cancel: true }], focus: gal.firstChild });
  };
  /** the open app's templates: [{ id, name }] */
  VO.templates = () => (current && current.hooks.templates ? current.hooks.templates() : []);
  /** a JPEG (at most 320 px) of the open document's first page, or null */
  VO.thumbnail = async () => {
    const make = current && current.hooks.thumb;
    const src = make && (await make());
    if (!src) return null;
    const k = Math.min(1, 320 / src.width, 320 / src.height);
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(src.width * k)); c.height = Math.max(1, Math.round(src.height * k));
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    g.drawImage(src, 0, 0, c.width, c.height);
    return new Promise((r) => c.toBlob(r, 'image/jpeg', 0.88));
  };
  /** the thumbnail for a Recent Files entry, a little after the document has been drawn */
  function thumbLater(key) {
    if (!key) return;
    setTimeout(async () => {
      try { const blob = await VO.thumbnail(); if (blob) await VO.recent.setThumb(key, blob); }
      catch (e) { console.warn('thumbnail:', e); }
    }, 600);
  }
  /** the app opened a file (from any source) */
  VO.opened = (file) => { if (!current) return; VO.recent.put(current.app, file.name, file).then(thumbLater, (e) => console.warn('recent files:', e)); };
  /** the app saved a file in its own format */
  VO.saved = (name, blob) => { if (!current) return; VO.recent.put(current.app, name, blob).then(thumbLater, (e) => console.warn('recent files:', e)); };
})();
