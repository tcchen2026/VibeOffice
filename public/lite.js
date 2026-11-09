/* VibeOffice Lite's start page (public/lite.html; desktop/src/lite.rs): the documents in ~/VibeOffice/Documents,
   newest first, a new document, help; made for a keyboard and a 320×170 screen. The window goes to the app's page
   and comes back here (VO.host.lite). A file of its own, not inline in the page (the pages' Content-Security-Policy). */
/* once the UI's dictionary is here (common/i18n.js) */
VO.i18n.whenReady(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const host = VO.host.lite;
  const list = $('list'), status = $('status');
  /* the page outside the lite edition (the web site): the full suite instead */
  if (!host) {
    $('empty').hidden = false;
    $('empty').textContent = VO.t('This page is VibeOffice Lite\'s, for small devices. The full VibeOffice is at vibeoffice.work.');
    return;
  }
  $('version').textContent = VO.t('Version {0}', VO.VERSION);

  /* the footer's line, until the next one: a note, or what went wrong */
  const advert = status.innerHTML;
  let quiet = 0;
  const say = (text, error) => {
    clearTimeout(quiet);
    status.textContent = VO.t(text);
    status.classList.toggle('error', !!error);
    quiet = setTimeout(() => { status.innerHTML = advert; status.classList.remove('error'); }, 6000);
  };
  const failed = (e) => say(String(e && e.message || e), true);

  /* a sentence translated whole, its {0}, {1}… filled in as `tag` elements (a key, a name): the nodes */
  const filled = (text, values, tag) => VO.t(text).split(/\{(\d)\}/).map((part, i) => {
    if (i % 2 === 0) return part;
    const el = document.createElement(tag);
    el.translate = false;
    el.textContent = values[+part] || '';
    return el;
  }).filter((n) => n !== '');

  /* ---------------------------------------------------------------- the documents */
  let items = [], at = 0, where = null;
  const loc = VO.i18n ? VO.i18n.dates : [];
  const when = (ms) => {
    const d = new Date(ms);
    return new Date().toDateString() === d.toDateString() ? d.toLocaleTimeString(loc, { hour: '2-digit', minute: '2-digit' })
      : d.toLocaleDateString(loc, { year: d.getFullYear() === new Date().getFullYear() ? undefined : 'numeric', month: 'short', day: 'numeric' });
  };
  const extOf = (name) => (/\.([^.]+)$/.exec(name) || ['', ''])[1].toUpperCase().slice(0, 4);
  function draw() {
    list.replaceChildren(...items.map((it, i) => {
      const row = document.createElement('div');
      row.className = 'doc';
      row.setAttribute('role', 'option');
      row.dataset.i = i;
      const ext = document.createElement('span');
      ext.className = 'ext ' + it.app;
      ext.textContent = extOf(it.name);
      const name = document.createElement('span');
      name.className = 'name';
      name.translate = false;
      name.textContent = it.name;
      const time = document.createElement('span');
      time.className = 'when';
      time.textContent = when(it.modified);
      row.append(ext, name, time);
      return row;
    }));
    $('empty').hidden = items.length > 0;
    drawWhere();
    select(Math.min(at, items.length - 1));
  }
  /* where the documents are and, while an SSH server listens, how an SFTP app copies files there over Wi-Fi: only
     where the list leaves room for it */
  function drawWhere() {
    const p = $('where');
    p.hidden = !where;
    if (!where) return;
    const line = (nodes) => { const d = document.createElement('div'); d.append(...nodes); return d; };
    p.replaceChildren(line(filled(VO.N('Your documents: {0}'), [where.folder], 'b')));
    if (where.sftp) {
      p.append(line(where.sftp.length ? filled(VO.N('Copy files over Wi-Fi with SFTP: {0}'), [where.sftp.join(', ')], 'b')
        : [VO.t('Connect to Wi-Fi to copy files in with SFTP.')]));
    }
    if (list.scrollHeight > list.clientHeight) p.hidden = true;
  }
  function select(i) {
    const rows = list.children;
    if (rows[at]) rows[at].classList.remove('on');
    at = Math.max(0, i);
    const row = rows[at];
    if (!row) return;
    row.classList.add('on');
    row.setAttribute('aria-selected', 'true');
    if (row.offsetTop < list.scrollTop) list.scrollTop = row.offsetTop;
    else if (row.offsetTop + row.offsetHeight > list.scrollTop + list.clientHeight) list.scrollTop = row.offsetTop + row.offsetHeight - list.clientHeight;
  }
  const page = () => Math.max(1, Math.floor(list.clientHeight / 22) - 1);
  const open = (i) => { if (items[i]) host.open(items[i].name).catch(failed); };
  async function refresh() {
    try { ({ documents: items, ...where } = await host.list()); } catch (e) { items = []; where = null; failed(e); }
    draw();
  }
  list.addEventListener('click', (e) => { const row = e.target.closest('.doc'); if (row) open(+row.dataset.i); });

  /* ---------------------------------------------------------------- New and Help */
  /* Help's lines about keys: each sentence translated whole (VO.N puts it in the catalog), {0}, {1}… its keys, drawn
     as keys */
  const keyLine = (text, keys) => {
    const p = document.createElement('p');
    p.append(...filled(text, keys, 'kbd'));
    return p;
  };
  $('keys').replaceChildren(
    keyLine(VO.N('{0} {1} (or {2} {3}) choose a document, {4} {5} a page at a time, {6} opens it.'), ['F', 'X', '↑', '↓', 'Z', 'C', 'Enter']),
    keyLine(VO.N('{0} makes a new document. {1} quits.'), ['N', 'Esc']),
    keyLine(VO.N('In a document: {0} or {1}+letter for the menus, {2}+{3} saves, File ▸ Close or Exit comes back here.'), ['F10', 'Alt', 'Ctrl', 'S']),
    keyLine(VO.N('In a document the arrows are {0}+{1} {2} {3} {4}; in a menu {1} {2} {3} {4} alone.'), ['Fn', 'F', 'X', 'Z', 'C']),
    keyLine(VO.N('The pointer: {0}+{1} {2} {3} {4} move it; {0}+{5} clicks, {0}+{6} right-clicks. Tap {0} twice to keep it on.'), ['Sym', 'F', 'Z', 'X', 'C', 'Enter', 'Space']));
  let sheet = null;   // the open one
  const picks = () => [...sheet.querySelectorAll('.pick')];
  function show(el) {
    if (sheet) sheet.hidden = true;
    sheet = el;
    $('esc').textContent = VO.t(el ? 'Back' : 'Quit');
    if (!el) { list.focus(); return; }
    el.hidden = false;
    picks().forEach((p, i) => p.classList.toggle('on', i === 0 && el.id === 'new'));
    el.querySelector('.body').scrollTop = 0;
  }
  const create = (app) => host.create(app, { new: true }).catch(failed);
  for (const b of $('new').querySelectorAll('.pick')) b.addEventListener('click', () => create(b.dataset.app));
  $('licenses').addEventListener('click', () => host.licenses().catch(failed));

  /* the CardputerZero's arrow letters without Fn (F, X, Z, C: up, down, left, right): no other key here uses them */
  const ARROWS = { f: 'ArrowUp', x: 'ArrowDown', z: 'ArrowLeft', c: 'ArrowRight' };
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey || e.altKey || e.metaKey) return;
    const k = ARROWS[e.key.toLowerCase()] || e.key, lower = k.length === 1 ? k.toLowerCase() : k;
    let done = true;
    if (sheet && sheet.id === 'new') {
      const p = picks(), cur = p.findIndex((b) => b.classList.contains('on'));
      const move = (i) => p.forEach((b, j) => b.classList.toggle('on', j === (i + p.length) % p.length));
      if (k === 'ArrowDown' || k === 'ArrowRight') move(cur + 1);
      else if (k === 'ArrowUp' || k === 'ArrowLeft') move(cur - 1);
      else if (k === 'Enter') create(p[Math.max(0, cur)].dataset.app);
      else if (lower === 'd') create('quire');
      else if (lower === 's') create('ledger');
      else if (lower === 'p') create('lectern');
      else if (k === 'Escape' || lower === 'n') show(null);
      else done = false;
    } else if (sheet) {   // Help
      const body = sheet.querySelector('.body');
      if (k === 'ArrowDown') body.scrollTop += 18;
      else if (k === 'ArrowUp') body.scrollTop -= 18;
      else if (k === 'PageDown' || k === 'ArrowRight' || k === ' ') body.scrollTop += body.clientHeight - 18;
      else if (k === 'PageUp' || k === 'ArrowLeft') body.scrollTop -= body.clientHeight - 18;
      else if (lower === 'l') host.licenses().catch(failed);
      else if (k === 'Escape' || k === 'F1' || lower === 'h' || k === '?') show(null);
      else done = false;
    } else if (k === 'ArrowDown') select(Math.min(at + 1, items.length - 1));
    else if (k === 'ArrowUp') select(at - 1);
    else if (k === 'PageDown' || k === 'ArrowRight') select(Math.min(at + page(), items.length - 1));
    else if (k === 'PageUp' || k === 'ArrowLeft') select(at - page());
    else if (k === 'Home') select(0);
    else if (k === 'End') select(items.length - 1);
    else if (k === 'Enter') open(at);
    else if (lower === 'n') show($('new'));
    else if (k === 'F1' || lower === 'h' || k === '?') show($('help'));
    else if (k === 'F5') refresh();
    else if (k === 'Escape') host.quit().catch(failed);
    else done = false;
    if (done) e.preventDefault();
  });

  list.focus();
  refresh();
  /* files copied in while the page shows */
  document.addEventListener('visibilitychange', () => { if (!document.hidden) refresh(); });
});
