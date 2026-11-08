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

  /* ---------------------------------------------------------------- the documents */
  let items = [], at = 0;
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
    select(Math.min(at, items.length - 1));
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
    try { items = await host.list(); } catch (e) { items = []; failed(e); }
    draw();
  }
  list.addEventListener('click', (e) => { const row = e.target.closest('.doc'); if (row) open(+row.dataset.i); });

  /* ---------------------------------------------------------------- New and Help */
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

  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey || e.altKey || e.metaKey) return;
    const k = e.key, lower = k.length === 1 ? k.toLowerCase() : k;
    let done = true;
    if (sheet && sheet.id === 'new') {
      const p = picks(), cur = p.findIndex((b) => b.classList.contains('on'));
      const move = (i) => p.forEach((b, j) => b.classList.toggle('on', j === (i + p.length) % p.length));
      if (k === 'ArrowDown') move(cur + 1);
      else if (k === 'ArrowUp') move(cur - 1);
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
      else if (k === 'PageDown' || k === ' ') body.scrollTop += body.clientHeight - 18;
      else if (k === 'PageUp') body.scrollTop -= body.clientHeight - 18;
      else if (lower === 'l') host.licenses().catch(failed);
      else if (k === 'Escape' || k === 'F1' || lower === 'h' || k === '?') show(null);
      else done = false;
    } else if (k === 'ArrowDown') select(Math.min(at + 1, items.length - 1));
    else if (k === 'ArrowUp') select(at - 1);
    else if (k === 'PageDown') select(Math.min(at + page(), items.length - 1));
    else if (k === 'PageUp') select(at - page());
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
