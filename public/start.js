/* The Start Center (public/index.html): Open File, Recent Files, Create File, the apps, Settings, About, the update
   check. A file of its own, not inline in the page (the pages' Content-Security-Policy). */
/* once the UI's dictionary is here (common/i18n.js) */
VO.i18n.whenReady(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  /* Safari's Lockdown Mode takes IndexedDB away (the name is undefined), which Recent Files, unsaved
     versions and opening files from here need */
  const noStorage = VO.host.kind === 'web' && typeof indexedDB === 'undefined';   // (the desktop app keeps them itself)
  if (noStorage) alert(VO.t('VibeOffice.work keeps Recent Files and unsaved changes in this browser, and Safari\'s Lockdown Mode turns that off.\n\nTo allow it for this site only: tap aA in the address bar, tap Website Settings, and turn off Lockdown Mode. Safari reloads the page.'));
  /* text and the attributes the user reads are shown in the UI's language (common/i18n.js), as L.h does in the apps */
  const T = (s) => (typeof s === 'string' ? VO.t(s) : s);
  /* translate="no": data (a file's name), shown as it is */
  const h = (tag, attrs, ...kids) => {
    const el = document.createElement(tag), tr = attrs && attrs.translate === 'no' ? (s) => s : T;
    for (const k in attrs || {}) if (attrs[k] != null) el.setAttribute(k, k === 'title' || k === 'aria-label' || k === 'placeholder' ? tr(attrs[k]) : attrs[k]);
    el.append(...kids.filter((k) => k != null).map(tr));
    return el;
  };
  const COLOR = { quire: 'var(--quire)', ledger: 'var(--ledger)', lectern: 'var(--lectern)' };

  /* app pictures: a page with lines, a grid, a slide; small for Create, large when a file has no thumbnail */
  function picture(app, ext) {
    const c = { quire: '#2a64c8', ledger: '#1f8a4c', lectern: '#e0682d' }[app];
    let art;
    if (app === 'ledger') art = `<rect x="14" y="22" width="52" height="54" fill="#fff" stroke="${c}"/><path d="M14 31h52M14 40h52M14 49h52M14 58h52M14 67h52M30 22v54M48 22v54" stroke="${c}" stroke-opacity=".55"/><rect x="14" y="22" width="52" height="9" fill="${c}" fill-opacity=".25"/>`;
    else if (app === 'lectern') art = `<rect x="12" y="28" width="56" height="40" rx="2" fill="${c}" fill-opacity=".14" stroke="${c}"/><rect x="18" y="34" width="44" height="8" fill="${c}"/><path d="M20 50h40M20 57h30" stroke="${c}" stroke-width="3" stroke-linecap="round" stroke-opacity=".6"/>`;
    else art = `<path d="M18 26h44M18 34h44M18 42h44M18 50h44M18 58h30" stroke="${c}" stroke-width="3" stroke-linecap="round"/>`;
    const badge = ext ? `<rect x="40" y="76" width="38" height="16" rx="3" fill="${c}"/><text x="59" y="88" font-family="Tahoma, sans-serif" font-size="10" font-weight="bold" fill="#fff" text-anchor="middle">${ext.toUpperCase().slice(0, 4)}</text>` : '';
    return `<svg viewBox="0 0 80 100" width="${ext ? 80 : 24}" height="${ext ? 100 : 30}" aria-hidden="true"><path d="M4 2h54l18 18v76H4z" fill="#fff" stroke="#8ea3c4"/><path d="M58 2v18h18" fill="#e3ebf8" stroke="#8ea3c4"/>${art}${badge}</svg>`;
  }
  for (const el of document.querySelectorAll('[data-icon]')) el.innerHTML = picture(el.dataset.icon);

  const status = (text) => { $('status').textContent = VO.t(text || ''); };
  /* the desktop app keeps Recent Files itself, not a browser */
  if (VO.host.kind === 'desktop') document.querySelector('.note').firstChild.textContent = VO.t('Your files stay on this computer.') + ' ';

  /* links to apps open in a new window when installed, so this start page stays */
  document.addEventListener('click', (e) => {
    const a = e.target.closest('a[data-go]');
    if (!a || e.ctrlKey || e.metaKey || e.shiftKey || e.button) return;
    e.preventDefault();
    VO.go(a.href).catch((err) => status(String(err)));
  });

  /* ---------- opening files: Open File..., drop, the operating system (installed app) */
  async function openFile(file) {
    status('');
    if (!file) return;
    let opened;
    try { opened = await VO.handoff(file); } catch (e) { status(String(e)); return; }   // the desktop app: no window
    if (!opened) status(VO.t('VibeOffice.work can\'t open “{0}”. It opens documents (.docx, .rtf), spreadsheets (.xlsx, .csv) and presentations (.pptx).', file.name));
  }
  $('picker').accept = VO.accept();
  /* the desktop app: the system's Open dialog, each file in a window of its own */
  $('open').addEventListener('click', () => {
    if (VO.host.kind === 'desktop') { status(''); VO.host.open().catch((e) => status(String(e))); return; }
    $('picker').value = ''; $('picker').click();
  });
  $('picker').addEventListener('change', () => openFile($('picker').files[0]));
  let depth = 0;
  const hasFiles = (e) => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
  document.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; depth++; $('drop').hidden = false; });
  document.addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; $('drop').hidden = true; } });
  document.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  document.addEventListener('drop', (e) => { if (!hasFiles(e)) return; e.preventDefault(); depth = 0; $('drop').hidden = true; openFile(e.dataTransfer.files[0]); });
  VO.host.onFiles((files) => openFile(files[0]));   // Open with, a double-click on a file

  /* ---------- Recent Files */
  const day = (t) => new Date(t).toDateString();
  const when = (t) => {
    const time = new Date(t).toLocaleTimeString(VO.i18n.dates, { hour: '2-digit', minute: '2-digit' });
    if (day(t) === day(Date.now())) return VO.t('Today, {0}', time);
    if (day(t) === day(Date.now() - 86400000)) return VO.t('Yesterday, {0}', time);
    return new Date(t).toLocaleDateString(VO.i18n.dates, { year: 'numeric', month: 'short', day: 'numeric' });
  };
  let urls = [];
  async function render() {
    let list = [];
    if (!noStorage) try { list = await VO.recent.list(); } catch (e) { status(VO.t('Recent Files are not available in this browser ({0}).', e.message)); }
    if (view !== 'recent') return;   // switched to Create File meanwhile
    for (const u of urls) URL.revokeObjectURL(u);
    urls = [];
    const grid = $('grid');
    grid.replaceChildren(...list.map((e) => {
      const thumb = h('div', { class: 'thumb' });
      thumb.innerHTML = picture(e.app, VO.ext(e.name));
      VO.recent.preview(e).then((blob) => {
        if (!blob) return;
        const u = URL.createObjectURL(blob);
        urls.push(u);
        thumb.replaceChildren(h('img', { src: u, alt: '' }));
      }, () => {});
      /* a document with unsaved changes opens at its unsaved version; × then discards those changes */
      const rm = h('button', { class: 'rm', type: 'button', title: e.draft ? 'Discard the unsaved changes' : 'Remove from list' }, '×');
      rm.setAttribute('aria-label', e.draft ? VO.t('Discard the unsaved changes to {0}', e.name) : VO.t('Remove from the list: {0}', e.name));
      rm.addEventListener('click', async () => {
        try {
          if (e.draft) { if (confirm(e.saved ? VO.t('Discard the unsaved changes to “{0}”? The saved version stays in the list.', e.name) : VO.t('Discard the unsaved changes to “{0}”?', e.name))) await VO.recent.discardDraft(e.key); }
          else await VO.recent.remove(e.key);
        } catch (err) { status(String(err)); }
        render();
      });
      const app = VO.APPS[e.app] || { name: e.app };
      const dot = h('i'); dot.style.background = COLOR[e.app] || '#888';
      const card = h('div', { class: 'doc' },
        thumb,
        h('a', { class: 'name', href: VO.url(e.app, e.draft ? { recent: e.key, draft: true } : { recent: e.key }), 'data-go': '', translate: 'no' }, e.name),
        h('div', { class: 'meta', translate: 'no' }, dot, app.name + ' · ' + when(e.time)),
        e.draft ? h('span', { class: 'unsaved' }, e.saved ? 'Unsaved changes' : 'Never saved') : null,
        e.draft && e.saved ? h('a', { class: 'alt', href: VO.url(e.app, { recent: e.key }), 'data-go': '', title: 'Open the version saved before the changes' }, 'Saved version') : null,
        rm);
      card.title = `${e.name}\n${app.name}` + (e.folder ? ' · ' + e.folder : '') + (e.saved ? ' · ' + VO.t('{0} KB', (e.size / 1024).toFixed(0)) : '') + ` · ${new Date(e.time).toLocaleString(VO.i18n.dates)}` + (e.draft ? '\n' + VO.t('Unsaved changes') : '');
      return card;
    }));
    grid.hidden = !list.length;
    $('empty').hidden = !!list.length;
    $('clear').hidden = !list.length;
  }
  const clearRecent = async () => {
    if (!confirm('Remove all files from the Recent Files list? The files on your computer are not touched.')) return false;
    await VO.recent.clear(); render();
    return true;
  };
  $('clear').addEventListener('click', clearRecent);
  VO.onRecentChange(() => { if (view === 'recent') render(); });
  document.addEventListener('visibilitychange', () => { if (!document.hidden && view === 'recent') render(); });

  /* ---------- Create File: Blank and the templates of each app (common/templates.json, made by tools/templates.mjs) */
  const KIND = { quire: VO.N('Documents'), ledger: VO.N('Spreadsheets'), lectern: VO.N('Presentations') };
  const NEW = { quire: [VO.N('New blank document'), VO.N('New document from “{0}”')], ledger: [VO.N('New blank spreadsheet'), VO.N('New spreadsheet from “{0}”')], lectern: [VO.N('New blank presentation'), VO.N('New presentation from “{0}”')] };
  let gallery = null;
  async function renderGallery() {
    if (gallery) return;
    gallery = $('gallery');
    let cat = {};
    try { cat = await (await fetch('common/templates.json')).json(); } catch (e) { status('The templates could not be loaded.'); }
    for (const app of Object.keys(KIND)) {
      const list = cat[app] || [{ id: '', name: 'Blank', img: null }];
      gallery.append(h('h2', null, h('span', { 'data-icon': app }), KIND[app], h('small', null, '· ' + VO.APPS[app].name)),
        h('div', { class: 'grid' }, ...list.map((t) => {
          const thumb = h('div', { class: 'thumb' });
          if (t.img) thumb.append(h('img', { src: t.img, alt: '', loading: 'lazy' })); else thumb.innerHTML = picture(app);
          return h('div', { class: 'doc', title: t.id ? VO.t(NEW[app][1], VO.t(t.name)) : VO.t(NEW[app][0]) },
            thumb, h('a', { class: 'name', href: VO.url(app, t.id ? { template: t.id } : { new: true }), 'data-go': '' }, t.name));
        })));
    }
    for (const el of gallery.querySelectorAll('[data-icon]')) el.innerHTML = picture(el.dataset.icon);
  }

  /* ---------- the two views: Recent Files (default) and Create File (#create), so Back goes between them */
  let view = '';
  function show() {
    view = location.hash === '#create' ? 'create' : 'recent';
    $('viewtitle').textContent = view === 'create' ? VO.t('Create File') : VO.t('Recent Files');
    $('recentbtn').classList.toggle('on', view === 'recent'); $('recentbtn').setAttribute('aria-pressed', view === 'recent');
    $('createbtn').classList.toggle('on', view === 'create'); $('createbtn').setAttribute('aria-pressed', view === 'create');
    status('');
    if (view === 'create') { $('grid').hidden = $('empty').hidden = $('clear').hidden = true; $('gallery').hidden = false; renderGallery(); }
    else { $('gallery').hidden = true; render(); }
  }
  $('recentbtn').addEventListener('click', () => { if (location.hash) history.pushState(null, '', location.pathname + location.search); show(); });
  $('createbtn').addEventListener('click', () => { if (location.hash !== '#create') location.hash = 'create'; else show(); });
  window.addEventListener('hashchange', show);
  window.addEventListener('popstate', show);
  /* the desktop app: an app window's Recent Files brings this window forward on them (host_show_recent) */
  window.addEventListener('vo-show-recent', () => $('recentbtn').click());
  /* the Mac's menu bar: File ▸ Open… with the Start Center in front (desktop/src/main.rs, mac_menu_event) */
  window.addEventListener('vo-menu', (e) => { if (e.detail === 'open' && !VO.busyDialog()) $('open').click(); });
  show();

  /* ---------- install, about */
  /* the title says what the suite is, for search results and tabs; the title bar of the installed app or the desktop
     app shows the name alone. The welcome's offline and install line is true where the service worker runs and Install
     is offered (VO.offersInstall, common/suite.js): not in the desktop app or a copy of the site elsewhere. */
  if (VO.host.kind === 'desktop' || matchMedia('(display-mode: standalone), (display-mode: window-controls-overlay)').matches) document.title = 'VibeOffice.work';
  if (!VO.offersInstall) document.querySelectorAll('.webonly').forEach((el) => { el.hidden = true; });
  /* a page served over plain http (a copy of the site; localhost and 127.0.0.1 count as secure) has no WebCrypto's
     AES, async clipboard, Local Font Access, service worker or persistent storage: say so, so nobody expects them */
  if (VO.host.kind === 'web' && window.isSecureContext === false) $('insecure').hidden = false;
  let installEvent = null;
  if (VO.offersInstall) window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); installEvent = e; $('install').hidden = false; });
  window.addEventListener('appinstalled', () => { installEvent = null; $('install').hidden = true; });
  $('install').addEventListener('click', async () => { if (!installEvent) return; installEvent.prompt(); await installEvent.userChoice; installEvent = null; $('install').hidden = true; });
  document.querySelectorAll('.version').forEach((el) => { el.textContent = VO.VERSION; });
  $('updatethis').textContent = VO.t('A newer version of VibeOffice is available. This one is {0}.', VO.VERSION);
  $('aboutversion').textContent = VO.t('Version {0}', VO.VERSION);
  /* the sentence whole, the link where its {0} is */
  $('aboutoss').append(...VO.t('Made possible by open source software: {0}.').split('{0}').flatMap((part, i) => (i ? [h('a', { class: 'link', id: 'osslink', href: 'oss-licenses.html', target: '_blank', rel: 'noopener' }, 'Open source licenses'), part] : [part])));
  $('about').addEventListener('click', () => $('aboutdlg').showModal());
  $('osslink').addEventListener('click', (e) => { e.preventDefault(); VO.showLicenses(); });
  /* Themes: the suite's looks (common/looks.css); choosing one changes every open window of the suite */
  const BODY = '<div class="lk-body"><div class="lk-canvas"><div class="lk-page"><p></p><p></p><p></p><p></p><p></p><p></p></div></div>'
    + '<div class="lk-pane"><em></em><u></u><u></u><s></s><u></u></div></div><div class="lk-status"></div>';
  const MOCK = '<div class="lk-menu"><span>File</span><span>Edit</span><span>View</span><span>Insert</span></div>'
    + '<div class="lk-tb"><i></i><i></i><i></i><b></b><i></i><i></i><i></i></div>'
    + BODY;
  const grp = (icons) => '<div class="lk-rgrp"><div>' + icons + '</div><p></p></div>';
  const RIBBON = '<div class="lk-rbar"><span>File</span><span class="on">Home</span><span>Insert</span><span>View</span></div>'
    + '<div class="lk-rpanel">' + grp('<i class="big"></i><i></i><i></i>') + grp('<i></i><i></i><i></i><i></i><i></i>') + grp('<i></i><i></i><i></i>') + grp('<i class="big"></i>') + '</div>'
    + BODY;
  const drawLooks = () => {
    const list = $('looklist');
    list.textContent = '';
    for (const l of VO.look.list()) {
      const mock = h('div', { class: 'lk-mock', 'data-look': l.id, 'aria-hidden': 'true' });
      mock.innerHTML = l.layout === 'ribbon' ? RIBBON : MOCK;
      const card = h('button', { class: 'lk-card', type: 'button', 'aria-pressed': String(l.id === VO.look.get()) }, mock, h('b', null, l.name), h('small', null, l.note));
      card.addEventListener('click', () => VO.look.set(l.id));
      list.append(card);
    }
  };
  $('themes').addEventListener('click', () => { drawLooks(); $('themesdlg').showModal(); });

  /* The web edition's service worker caching a version of the site (sw.js, 'vo-cache' from common/suite.js):
     a progress bar above the GitHub button while it works, with the megabytes downloaded; after an update, a
     button to reload this page in the new version (a first visit's page is already the cached version) */
  const mb = (bytes) => (bytes / 1048576).toFixed(1);
  window.addEventListener('vo-cache', (e) => {
    const d = e.detail || {};
    $('caching').hidden = d.state !== 'progress';
    $('cachingreload').hidden = !(d.state === 'done' && d.update);
    if (d.state !== 'progress') return;
    $('cachingtext').textContent = d.update ? 'Updating cached resources' : 'Caching resources for offline use';
    $('cachingsize').textContent = d.total ? `${mb(d.received || 0)} of ${mb(d.total)} MB` : d.received ? `${mb(d.received)} MB` : '';
    $('cachingbar').value = d.progress || 0;
  });
  $('cachingreload').addEventListener('click', () => location.reload());

  /* Updates (desktop app): an Update button, or once a day a dialog, when updater/current.json asks for it. Everywhere
     (the site, the installed PWA, the desktop app), a version two months old or more shows Check for updates instead,
     with its age, even when current.json cannot be read; it waits up to 5 s for the check, so the two buttons don't
     swap places. Both open the updater page. The Mac App Store edition never reads current.json (VO.host.updates: no
     network, the store updates it), but its Check for updates opens the page for the App Store */
  const download = () => VO.host.openUrl(VO.updateUrl()).catch((e) => status(String(e)));
  $('update').addEventListener('click', download);
  $('updatecheck').addEventListener('click', download);
  $('updatedlg').addEventListener('close', () => { if ($('updatedlg').returnValue === 'download') download(); });
  const checked = VO.host.kind === 'desktop' && VO.host.updates ? VO.checkUpdate().then(({ ask, news }) => {
    if (!ask) return;
    $('updatenews').replaceChildren(...news.flatMap((n) => [h('b', { translate: 'no' }, n.version), h('p', { translate: 'no' }, n.text)]));
    $('updatenews').hidden = !news.length;
    $('update').hidden = false;
    $('updatecheck').hidden = true;
    const today = new Date().toDateString();
    if (ask === 'dialog' && VO.settings.get('vibeoffice.updateAsked') !== today) {
      VO.settings.set('vibeoffice.updateAsked', today);   // not kept: asks again next time
      $('updatedlg').showModal();
    }
  }).catch((e) => console.info('update check:', e.message)) : null;   // offline: nothing to show
  if (VO.isOutdated()) {
    $('updateage').textContent = VO.t('This version is {0} months old.', VO.monthsOld());
    Promise.race([checked, new Promise((r) => setTimeout(r, 5000))]).then(() => { $('updatecheck').hidden = !$('update').hidden; });
  }

  /* Settings: what applies to every app. UI size is the desktop app's zoom (VO.host); on the web the browser's own */
  const fontInfo = () => {
    const n = VO.localFonts.list().length;
    $('fontbtns').hidden = !VO.localFonts.supported();
    $('fontforget').hidden = !n;
    $('fontload').textContent = n ? VO.t('Reload fonts on this computer') : VO.t('Load fonts on this computer');
    $('fontinfo').textContent = n ? VO.t('The Font boxes also list the {0} fonts installed on this computer.', n)
      : VO.localFonts.supported() ? VO.t('The Font boxes list the frequently used Office fonts and the document\'s own. They can also list the fonts installed on this computer, once you allow it.')
        : VO.t('The Font boxes list the frequently used Office fonts and the document\'s own.');
  };
  const recentInfo = async () => {
    let list = [];
    if (!noStorage) try { list = await VO.recent.list(); } catch (e) { /* shown as none */ }
    const drafts = list.filter((e) => e.draft).length;
    let mb = '';
    const desk = VO.host.kind === 'desktop';   // its Recent Files aren't the browser's storage
    if (!desk) try { const est = navigator.storage && await navigator.storage.estimate(); if (est && est.usage) mb = (est.usage / 1048576).toFixed(1); } catch (e) { /* no estimate */ }
    /* how many files (and with unsaved changes), then how much room they take and where */
    const files = list.length === 1 ? (drafts ? VO.t('1 file, with unsaved changes') : VO.t('1 file')) : drafts ? VO.t('{0} files, {1} with unsaved changes', list.length, drafts) : VO.t('{0} files', list.length);
    const kept = mb ? (desk ? VO.t('About {0} MB, kept by VibeOffice on this computer.', mb) : VO.t('About {0} MB, kept in this browser.', mb)) : desk ? VO.t('Kept by VibeOffice on this computer.') : VO.t('Kept in this browser.');
    $('recentinfo').textContent = noStorage ? VO.t('Recent Files are not available here.')
      : !list.length ? VO.t('The list is empty.') : VO.t('{0}. {1}', files, kept);
    $('recentclear').hidden = !list.length;
  };
  $('settings').addEventListener('click', async () => {
    const scale = await VO.host.uiScale().catch(() => null);
    $('scalerow').hidden = scale === null;
    $('scaleweb').hidden = scale !== null;
    if (scale !== null) $('uiscale').value = String(scale);
    $('uilang').replaceChildren(...VO.lang.list().map((l) => (l.id === 'auto' ? h('option', { value: l.id }, 'Auto (as the browser)') : h('option', { value: l.id, translate: 'no' }, l.name))));
    $('uilang').value = VO.lang.get();
    fontInfo();
    await recentInfo();
    $('settingsdlg').showModal();
  });
  /* the Start Center shows the new language at once; the apps' windows when they next open */
  $('uilang').addEventListener('change', () => { VO.lang.set($('uilang').value); VO.settings.kept().then(() => location.reload()); });
  $('uiscale').addEventListener('change', () => VO.host.setUiScale(+$('uiscale').value).catch((e) => status(String(e))));
  $('fontload').addEventListener('click', async () => {
    try { await VO.localFonts.load(); } catch (e) { $('fontinfo').textContent = e.name === 'NotAllowedError' || e.name === 'SecurityError' ? VO.t('The fonts on this computer could not be listed: permission was not given.') : VO.t('The fonts on this computer could not be listed: {0}', e.message || e); return; }
    fontInfo();
  });
  $('fontforget').addEventListener('click', () => { VO.localFonts.forget(); fontInfo(); });
  $('recentclear').addEventListener('click', async () => { if (await clearRecent()) recentInfo(); });
  window.addEventListener('vo-look', drawLooks);
});
