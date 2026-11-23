/* VibeOffice — cast the document surface from a native window. */
(function () {
  'use strict';
  const host = window.VO?.host?.cast;
  if (!host) return;
  const { h, ui } = L;
  const video = { width: 1280, height: 720, fps: 2 };
  let active, dialog, discovery;

  // A view supplies its viewport and ordered document layers in screen coordinates.
  // Canvas layers can be copied directly. Snapshot DOM layers while their styles and
  // animations are live, then draw a self-contained SVG at the output resolution.
  const properties = (
    'display position top right bottom left z-index box-sizing width height min-width max-width min-height max-height ' +
    'margin padding border border-radius border-collapse border-spacing table-layout vertical-align ' +
    'overflow overflow-x overflow-y clip-path opacity visibility background-color background-image background-size ' +
    'background-position background-repeat color direction writing-mode font-family font-size font-weight font-style ' +
    'font-stretch font-variant line-height letter-spacing word-spacing white-space word-break overflow-wrap ' +
    'text-align text-align-last text-indent text-decoration text-transform text-shadow tab-size list-style ' +
    'flex-direction flex-wrap align-items align-self align-content justify-content flex gap grid-template-columns ' +
    'grid-template-rows grid-column grid-row object-fit object-position transform transform-origin translate ' +
    'rotate scale filter box-shadow fill fill-opacity fill-rule stroke stroke-width stroke-opacity stroke-dasharray ' +
    'stroke-dashoffset stroke-linejoin stroke-linecap paint-order text-anchor dominant-baseline stop-color stop-opacity'
  ).split(' ');
  const serializer = new XMLSerializer();

  function framePainter() {
    const media = new Map(), fontCache = new Map();
    // Keep only last frame's visible DOM pictures. A static slide still gets a new
    // encoded frame on every tick, without rasterizing and embedding its fonts again.
    let cached = new Map(), revision = 0;
    const invalidate = () => { revision++; cached.clear(); };
    const changes = new MutationObserver(invalidate);
    changes.observe(document.documentElement, { subtree: true, childList: true, attributes: true, characterData: true });
    document.fonts?.addEventListener('loadingdone', invalidate);
    // Only document images from this webview can become part of the stream.
    function imageURL(value) {
      if (!value || value.startsWith('#') || value.startsWith('data:')) return value;
      const url = new URL(value, document.baseURI);
      if (url.hash && url.href.split('#')[0] === location.href.split('#')[0]) return url.hash;
      if (url.protocol !== 'blob:' && url.origin !== location.origin) return '';
      if (!media.has(url.href)) {
        if (media.size >= 128) media.clear();
        media.set(url.href, fetch(url.href).then(r => r.blob()).then(L.blobToDataURL).catch(() => ''));
      }
      return media.get(url.href);
    }

    async function picture(source, w, height, crop, pixelWidth, pixelHeight) {
      const pending = [], fonts = new Set();
      function copy(node) {
        if (node.nodeType !== Node.ELEMENT_NODE) return node.cloneNode(false);
        if (node.matches('script,style,link,iframe,object,embed')) return document.createTextNode('');
        const cs = getComputedStyle(node);
        if (cs.display === 'none') return document.createTextNode('');
        const isCanvas = node instanceof HTMLCanvasElement;
        const clone = isCanvas ? document.createElement('img') : node.cloneNode(false);
        if (isCanvas && node.width && node.height) clone.src = node.toDataURL();
        clone.removeAttribute('style');
        clone.removeAttribute('contenteditable');
        for (const prop of properties) {
          let value = cs.getPropertyValue(prop);
          if (!value) continue;
          if (value.includes('url(')) {
            const matches = Array.from(value.matchAll(/url\(["']?([^)"']+)["']?\)/g));
            pending.push(Promise.all(matches.map(m => imageURL(m[1]))).then(urls => {
              matches.forEach((m, i) => { value = value.replace(m[0], 'url("' + urls[i] + '")'); });
              clone.style.setProperty(prop, value);
            }));
          } else clone.style.setProperty(prop, value);
        }
        if (node.localName === 'img' || node.localName === 'image') {
          for (const attr of Array.from(node.attributes)) {
            if (attr.localName === 'src' || attr.localName === 'href') {
              pending.push(Promise.resolve(imageURL(attr.value)).then(url => clone.setAttributeNS(attr.namespaceURI, attr.name, url)));
            }
          }
          clone.removeAttribute('srcset');
        }
        if (node.childNodes.length) {
          fonts.add(cs.fontFamily + (Number(cs.fontWeight) >= 600 ? ' bold' : '') + ' ' + cs.fontStyle);
          for (const child of node.childNodes) clone.appendChild(copy(child));
        }
        return clone;
      }
      const clone = copy(source);
      Object.assign(clone.style, { position: 'relative', top: '0', left: '0', right: 'auto', bottom: 'auto',
        width: w + 'px', height: height + 'px', minWidth: '0', minHeight: '0', maxWidth: 'none', maxHeight: 'none',
        boxSizing: 'border-box', margin: '0', transform: 'none', translate: 'none', rotate: 'none', scale: 'none' });
      const families = Array.from(fonts).sort().join('\n');
      if (!fontCache.has(families)) {
        if (fontCache.size >= 8) fontCache.clear();
        fontCache.set(families, L.embedFonts(families));
      }
      const fontCSS = fontCache.get(families);
      await Promise.all(pending);
      const markup = serializer.serializeToString(clone);
      const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + pixelWidth + '" height="' + pixelHeight
        + '" viewBox="' + crop.join(' ') + '" preserveAspectRatio="none"><foreignObject width="' + w + '" height="' + height + '">'
        + '<div xmlns="http://www.w3.org/1999/xhtml"><style>' + (await fontCSS) + '</style>' + markup + '</div></foreignObject></svg>';
      return L.loadImage('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg));
    }

    async function paint(canvas) {
      const previous = cached, version = revision;
      cached = new Map();
      const view = await VO.castView(), ctx = canvas.getContext('2d');
      ctx.fillStyle = '#000'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      if (!view) return;
      const direct = view instanceof HTMLCanvasElement;
      const box = direct ? { width: view.width, height: view.height } : view.viewport.getBoundingClientRect();
      if (!box.width || !box.height) return;
      const k = Math.min(canvas.width / box.width, canvas.height / box.height);
      ctx.save();
      try {
        ctx.translate((canvas.width - box.width * k) / 2, (canvas.height - box.height * k) / 2);
        ctx.scale(k, k);
        ctx.beginPath(); ctx.rect(0, 0, box.width, box.height); ctx.clip();
        if (direct) { ctx.drawImage(view, 0, 0); return; }
        ctx.fillStyle = view.background || '#fff'; ctx.fillRect(0, 0, box.width, box.height);
        for (const el of view.layers) {
          if (!el) continue;
          const r = el.getBoundingClientRect();
          if (!r.width || !r.height || r.right <= box.left || r.left >= box.right || r.bottom <= box.top || r.top >= box.bottom) continue;
          if (el instanceof HTMLCanvasElement) {
            ctx.drawImage(el, r.left - box.left, r.top - box.top, r.width, r.height);
            continue;
          }
          const left = Math.max(r.left, box.left), top = Math.max(r.top, box.top);
          const width = Math.min(r.right, box.right) - left, height = Math.min(r.bottom, box.bottom) - top;
          const w = el.offsetWidth || r.width, h = el.offsetHeight || r.height;
          const crop = [(left - r.left) * w / r.width, (top - r.top) * h / r.height, width * w / r.width, height * h / r.height];
          const pw = Math.max(1, Math.ceil(width * k)), ph = Math.max(1, Math.ceil(height * k));
          const key = [w, h, pw, ph, ...crop].join(',');
          const reusable = !el.querySelector('canvas,video') && !el.getAnimations?.({ subtree: true }).some(a => a.playState === 'running');
          const prior = reusable && version === revision && previous.get(el);
          const img = prior?.key === key ? prior.img : await picture(el, w, h, crop, pw, ph);
          if (reusable && version === revision) cached.set(el, { key, img });
          ctx.drawImage(img, left - box.left, top - box.top, width, height);
        }
      } finally { ctx.restore(); }
    }
    return { paint, close() {
      changes.disconnect();
      document.fonts?.removeEventListener('loadingdone', invalidate);
      cached.clear();
    } };
  }

  async function stop(session) {
    if (active !== session) return;
    active = null;
    try { await session.capture?.stop(); }
    finally { session.painter?.close(); await host.stop(session.key); }
  }

  function choose() {
    if (dialog) return;
    if (active) {
      const session = active;
      dialog = ui.dialog({ title: L.t('Cast to TV'), body: h('p', { text: L.t('Casting to {0}', session.name) }),
        buttons: [{ label: L.t('Stop casting'), primary: true, onClick() {
          stop(session).catch(e => ui.msg(L.t('Could not stop casting: {0}', String(e)), { icon: 'warn' }));
        } }, { label: L.t('Close') }] });
      dialog.done.then(() => { dialog = null; });
      return;
    }
    const tv = h('select', { id: 'cast-tv' });
    const mac = VO.host.platform === 'macos';
    const status = h('div', { role: 'status' });
    const accessHelp = mac ? h('div', { hidden: true, text: L.t('On macOS 15 or later, check VibeOffice in System Settings → Privacy & Security → Local Network, then choose Search again.') }) : null;
    const body = h('div', { class: 'col' }, h('label', { for: 'cast-tv', text: L.t('TV') }), tv, status, accessHelp);
    let connecting = false, searching = false, retryOnFocus = false, regainedFocus = false, timer, supported;
    const d = ui.dialog({ title: L.t('Cast to TV'), body, buttons: [
      { label: L.t('Cast'), primary: true, onClick() { connect(); return false; } },
      { label: L.t('Search again'), onClick() { scan(); return false; } }, { label: L.t('Cancel'), cancel: true }
    ] });
    dialog = d;
    function refocus() {
      if (!retryOnFocus || !d.el.isConnected) return;
      if (searching) regainedFocus = true;
      else scan(false);
    }
    if (mac) window.addEventListener('focus', refocus);
    d.done.then(() => {
      if (dialog === d) dialog = null;
      clearTimeout(timer);
      if (mac) window.removeEventListener('focus', refocus);
    });
    async function scan(retry = true) {
      if (searching || connecting || !d.el.isConnected) return;
      searching = true;
      regainedFocus = retryOnFocus = false;
      clearTimeout(timer);
      if (accessHelp) accessHelp.hidden = true;
      d.buttons[0].disabled = d.buttons[1].disabled = true;
      tv.replaceChildren();
      try {
        // Check the actual output configuration before discovery can request LAN access.
        const sdk = window.DLNAnything;
        if (!supported) supported = sdk?.isSupported ? await sdk.isSupported(video) : !!(sdk && window.VideoEncoder && window.VideoFrame);
        if (!d.el.isConnected) return;
        if (!supported) { status.textContent = L.t('This webview cannot encode H.264 for casting.'); return; }
        retryOnFocus = mac;
        status.textContent = mac ? L.t('Searching for TVs. If macOS asks for Local Network access, choose Allow.') : L.t('Searching for TVs...');
        // Reopening the picker during an outstanding SSDP search shares its result.
        if (!discovery) discovery = host.scan().finally(() => { discovery = null; });
        const renderers = await discovery;
        if (!d.el.isConnected) return;
        for (const r of renderers) tv.add(new Option(r.name || r.id, r.id));
        status.textContent = tv.length ? '' : L.t('No TVs found on this network.');
      } catch (e) { status.textContent = L.t('Could not search for TVs: {0}', String(e)); }
      finally {
        searching = false;
        if (d.el.isConnected) {
          d.buttons[0].disabled = !tv.length;
          d.buttons[1].disabled = false;
          retryOnFocus = retryOnFocus && !tv.length;
          if (retryOnFocus) {
            accessHelp.hidden = false;
            // macOS can reject the first packet while its consent sheet is open.
            // Some WebKit versions deliver a focus event and some keep the window focused.
            if (regainedFocus || (retry && document.hasFocus()))
              timer = setTimeout(() => { if (document.hasFocus()) scan(false); }, regainedFocus ? 0 : 1500);
          }
        }
      }
    }
    async function connect() {
      if (connecting || !tv.value) return;
      connecting = true;
      d.buttons[0].disabled = d.buttons[1].disabled = tv.disabled = true;
      status.textContent = L.t('Starting cast...');
      let session;
      try {
        const key = await host.prepare(tv.value);
        session = { key, name: tv.selectedOptions[0].text, capture: null };
        const canvas = document.createElement('canvas');
        canvas.width = video.width; canvas.height = video.height;
        if (!d.el.isConnected) { await host.stop(key); return; }
        active = session;
        session.painter = framePainter();
        session.capture = await DLNAnything.startCanvas(canvas, { fps: video.fps, beforeFrame: session.painter.paint,
          sendFrame: bytes => host.frame(key, bytes), onError(e) {
            if (active !== session) return;
            stop(session).catch(() => {});
            ui.msg(L.t('Casting stopped: {0}', String(e)), { icon: 'warn' });
          } });
        if (active !== session) { await session.capture.stop(); return; }
        if (!d.el.isConnected) { await stop(session); return; }
        await host.play(key);
        if (!d.el.isConnected) { await stop(session); return; }
        d.close();
      } catch (e) {
        if (session) await stop(session).catch(() => {});
        status.textContent = L.t('Could not cast to this TV: {0}', String(e));
        connecting = false;
        d.buttons[0].disabled = d.buttons[1].disabled = tv.disabled = false;
      }
    }
    scan();
  }
  ui.cmd('castTV', { label: L.N('Cast to TV...'), icon: 'fullScreen', checked: () => !!active, run: choose });
})();
