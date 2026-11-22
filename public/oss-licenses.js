/* Open source licenses (oss-licenses.html, written by tools/oss-licenses.py): a file of its own, not inline in the
   page (the pages' Content-Security-Policy). */
/* a license is filled in when its entry is opened: from this page's texts (each kept once), or from its own file */
const TEXTS = JSON.parse(document.getElementById('texts').textContent);
for (const d of document.querySelectorAll('details')) {
  d.addEventListener('toggle', async () => {
    const pre = d.querySelector('pre');
    if (!d.open || pre.textContent) return;
    if (d.dataset.texts) { pre.textContent = d.dataset.texts.split(' ').map((k) => TEXTS[k]).join('\n\n' + '-'.repeat(60) + '\n\n'); return; }
    try {
      const res = await fetch(d.dataset.src);
      pre.textContent = res.ok ? await res.text() : 'See ' + d.dataset.src + '.';
    } catch (e) { pre.textContent = 'See ' + d.dataset.src + '.'; }
  });
}

/* In the desktop app this page's links open nothing (the app sends only its update page and its GitHub to the
   browser: docs/desktop.md, Links); a click shows the address, to copy into the browser, as a document's links do */
if (window.__VO_HOST__) {
  const dialog = document.createElement('dialog');
  dialog.setAttribute('aria-label', 'Address');
  dialog.style.cssText = 'max-width:480px;padding:16px 18px;border:1px solid var(--chrome-line);border-radius:6px;background:var(--surface);color:var(--ink)';
  const box = Object.assign(document.createElement('input'), { type: 'text', readOnly: true, spellcheck: false });
  box.setAttribute('aria-label', 'Address');
  box.style.cssText = 'width:100%;margin:6px 0 10px;font:inherit';
  const note = Object.assign(document.createElement('p'), { textContent: 'VibeOffice doesn\'t open web pages itself. Copy the address and paste it into your browser.' });
  const copied = Object.assign(document.createElement('p'), { role: 'status' });
  copied.style.minHeight = '1.5em';   // the buttons stay where they were clicked
  const copy = Object.assign(document.createElement('button'), { type: 'button', textContent: 'Copy Address' });
  const close = Object.assign(document.createElement('button'), { type: 'button', textContent: 'Close' });
  const buttons = document.createElement('p');
  buttons.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;margin:0';
  buttons.append(copy, close);
  dialog.append(Object.assign(document.createElement('p'), { textContent: 'This link goes to:' }), box, note, copied, buttons);
  document.body.append(dialog);
  copy.addEventListener('click', async () => {
    try { await navigator.clipboard.writeText(box.value); } catch (e) { box.select(); document.execCommand('copy'); }
    copied.textContent = 'Copied.';
  });
  close.addEventListener('click', () => dialog.close());
  document.addEventListener('click', (e) => {
    const a = e.target.closest('a[href]');
    if (!a || !/^https?:/.test(a.href)) return;
    e.preventDefault();
    box.value = a.href;
    copied.textContent = '';
    dialog.showModal();
    box.focus();
    box.select();
  });
}

/* The lite edition has one window (desktop/src/lite.rs): Esc goes back to its start page */
if (window.__VO_HOST__ && window.__VO_HOST__.edition === 'lite' && window.__TAURI_INTERNALS__) {
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !document.querySelector('dialog[open]')) window.__TAURI_INTERNALS__.invoke('host_lite_home').catch(() => {});
  });
}
