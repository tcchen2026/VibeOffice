/* VibeOffice — cast the document surface from a native window. */
(function () {
  'use strict';
  const host = window.VO?.host?.cast;
  if (!host) return;
  const { h, ui } = L;
  const video = { width: 1280, height: 720, fps: 2 };
  let active, dialog, discovery;

  // Mount the empty live region before playback so its first message is announced.
  const editor = document.getElementById('app');
  const castStatus = h('div', { class: 'cast-status', role: 'status', 'aria-atomic': 'true' });
  editor.prepend(castStatus);
  window.addEventListener('vo-presentation', e => (e.detail || editor).prepend(castStatus));

  function updateBar(session) {
    session.bar.textContent = session.painter.frozen ? L.t('Casting to TV Now · Picture frozen') : L.t('Casting to TV Now');
  }

  function showBar(session, picker) {
    session.bar = h('button', { class: 'cast-bar', type: 'button', 'aria-haspopup': 'dialog',
      onclick() {
        if (picker.el.isConnected) picker.close(); // Leave the stream running while opening its controls.
        choose();
      },
      onkeydown(e) {
        if (e.key === 'Enter' || e.key === ' ') e.stopPropagation(); // The button owns activation in the editor.
      }
    });
    updateBar(session);
    castStatus.appendChild(session.bar);
    castStatus.classList.add('waiting');
    picker.done.then(() => { if (active === session) castStatus.classList.remove('waiting'); });
  }

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

    async function render(canvas) {
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
    let frozen = false, generation = 0, buffer, rendering = false, releasePaint;
    return {
      get frozen() { return frozen; },
      set frozen(value) {
        frozen = value;
        generation++;
        if (frozen) releasePaint?.();
      },
      async paint(canvas) {
        if (frozen || rendering) return;
        const before = generation;
        if (!buffer) {
          buffer = document.createElement('canvas');
          buffer.width = canvas.width;
          buffer.height = canvas.height;
        }
        // A slow snapshot may finish after Freeze. Keep sending the published canvas
        // while that work drains, and only reuse its buffer once it has finished.
        rendering = true;
        const picture = render(buffer).then(() => {
          if (generation === before) canvas.getContext('2d').drawImage(buffer, 0, 0);
        }).finally(() => { rendering = false; });
        try { await Promise.race([picture, new Promise(resolve => { releasePaint = resolve; })]); }
        finally { releasePaint = null; }
      },
      close() {
        changes.disconnect();
        document.fonts?.removeEventListener('loadingdone', invalidate);
        cached.clear();
      }
    };
  }

  async function stop(session) {
    if (active !== session) return;
    active = null;
    castStatus.replaceChildren();
    castStatus.classList.remove('waiting');
    try { await session.capture?.stop(); }
    finally { session.painter?.close(); await host.stop(session.key); }
  }

  function errorMessage(error, fallback) {
    console.debug('Cast to TV', error?.detail || error);
    switch (error?.kind) {
      case 'notAnswering': return L.t('The TV did not answer. Check that it is on and connected to the same network, then try again.');
      case 'refused': return L.t('The TV declined to play this stream. Check that it supports HLS video, then try again.');
      case 'picture': return L.t('Casting stopped because the next picture could not be sent. Try casting again.');
      case 'gone': return L.t('This TV is no longer available. Choose Search again to find it.');
      case 'busy': return L.t('This TV is already in use. Stop the other cast or choose another TV.');
      default: return (!error?.kind && fallback) || L.t('VibeOffice could not reach the local network. Check the connection and try again.');
    }
  }

  function firewallHelp() {
    switch (VO.host.platform) {
      case 'windows': return L.t('Allow VibeOffice through Windows Defender Firewall on Private networks. Set your home or office network to Private in Windows Settings.');
      case 'macos': return L.t('In System Settings → Network → Firewall → Options, allow incoming connections to VibeOffice.');
      default: return L.t('Check that your firewall (such as ufw or firewalld) allows the TV to connect to VibeOffice.');
    }
  }

  function stopFromDialog(session) {
    stop(session).catch(e => {
      console.debug('Stop casting', e?.detail || e);
      ui.msg(L.t('VibeOffice stopped sending pictures. If the TV is still playing, stop playback with its remote.'), { icon: 'warn' });
    });
  }

  function choose() {
    if (dialog?.el.isConnected) return;
    if (active) {
      const session = active;
      if (!session.bar) return;
      const freeze = h('input', { type: 'checkbox', id: 'cast-freeze', checked: session.painter.frozen });
      freeze.addEventListener('change', () => { session.painter.frozen = freeze.checked; updateBar(session); });
      const controls = dialog = ui.dialog({ title: L.t('Cast to TV'), body: h('div', { class: 'col' },
        h('p', { text: L.t('Casting to {0}', session.name) }),
        h('label', { for: 'cast-freeze' }, freeze, L.t('Freeze picture'))),
        buttons: [{ label: L.t('Stop casting'), primary: true, onClick() {
          stopFromDialog(session);
        } }, { label: L.t('Close') }] });
      controls.done.then(() => { if (dialog === controls) dialog = null; });
      return;
    }
    const tv = h('select', { id: 'cast-tv' });
    const mac = VO.host.platform === 'macos';
    const status = h('div', { role: 'status' });
    const accessHelp = mac ? h('div', { hidden: true, text: L.t('On macOS 15 or later, check VibeOffice in System Settings → Privacy & Security → Local Network, then choose Search again.') }) : null;
    const body = h('div', { class: 'col' }, h('label', { for: 'cast-tv', text: L.t('TV') }), tv, status, accessHelp);
    const preferred = VO.settings.get('vibeoffice.castTV');
    let selectionChanged = false;
    tv.addEventListener('change', () => { selectionChanged = true; });
    let playingSession;
    let connecting = false, searching = false, retryOnFocus = false, regainedFocus = false, timer, supported;
    const d = ui.dialog({ title: L.t('Cast to TV'), body, buttons: [
      { label: L.t('Cast'), primary: true, onClick() {
        if (!playingSession) { connect(); return false; }
        // Keep waiting closes the dialog and leaves the stream running.
      } },
      { label: L.t('Search again'), onClick() {
        if (playingSession) stopFromDialog(playingSession);
        else { scan(); return false; }
      } }, { label: L.t('Cancel'), cancel: true }
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
      selectionChanged = false;
      try {
        // Check the actual output configuration before discovery can request LAN access.
        const sdk = window.DLNAnything;
        if (!supported) supported = sdk?.isSupported ? await sdk.isSupported(video) : !!(sdk && window.VideoEncoder && window.VideoFrame);
        if (!d.el.isConnected) return;
        if (!supported) { status.textContent = L.t('This webview cannot encode H.264 for casting.'); return; }
        retryOnFocus = mac;
        status.textContent = mac ? L.t('Searching for TVs. If macOS asks for Local Network access, choose Allow.') : L.t('Searching for TVs...');
        function show(devices) {
          if (!d.el.isConnected || connecting) return;
          for (const r of devices.slice(tv.length)) {
            tv.add(new Option(r.name || r.id, r.id));
            if (!selectionChanged && r.id === preferred) tv.value = r.id;
          }
          d.buttons[0].disabled = !tv.length;
        }
        // Reopening the picker can immediately show everything found so far.
        if (!discovery) {
          const current = discovery = { devices: [], show };
          current.done = host.scan(devices => {
            current.devices = devices;
            current.show(devices);
          }).finally(() => { discovery = null; });
        }
        discovery.show = show;
        show(discovery.devices);
        const devices = await discovery.done;
        if (!d.el.isConnected || connecting) return;
        show(devices);
        status.textContent = tv.length ? '' : L.t('No TVs found on this network.');
      } catch (e) { if (!connecting) status.textContent = errorMessage(e); }
      finally {
        searching = false;
        if (d.el.isConnected && !connecting) {
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
    async function waitForTV(session) {
      playingSession = session;
      body.replaceChildren(status);
      status.textContent = L.t('Waiting for {0}...', session.name);
      const [keep, end, cancel] = d.buttons;
      keep.textContent = L.t('Keep waiting');
      end.textContent = L.t('Stop casting');
      keep.disabled = end.disabled = false;
      cancel.hidden = true;
      keep.focus();
      let helpShown = false;
      try {
        while (d.el.isConnected && active === session) {
          const opened = await host.wait(session.key);
          if (!d.el.isConnected || active !== session) return;
          if (opened) { d.close(); return; }
          if (!helpShown) {
            status.textContent = L.t('The TV has not opened the stream. Check that the TV and this computer are on the same network. Guest Wi-Fi can keep devices apart.');
            body.append(h('p', { text: firewallHelp() }));
            helpShown = true;
          }
        }
      } catch (e) {
        if (d.el.isConnected && active === session) status.textContent = errorMessage(e);
      }
    }
    async function connect() {
      if (connecting || !tv.value) return;
      const id = tv.value;
      connecting = true;
      d.buttons[0].disabled = d.buttons[1].disabled = tv.disabled = true;
      status.textContent = L.t('Starting cast...');
      let session;
      try {
        const key = await host.prepare(id);
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
            if (playingSession === session && dialog?.el.isConnected) dialog.close();
            ui.msg(errorMessage(e, L.t('Casting stopped because the next picture could not be sent. Try casting again.')), { icon: 'warn' });
          } });
        if (active !== session) { await session.capture.stop(); return; }
        if (!d.el.isConnected) { await stop(session); return; }
        await host.play(key);
        if (active !== session) return;
        if (!d.el.isConnected) { await stop(session); return; }
        VO.settings.set('vibeoffice.castTV', id);
        showBar(session, d);
        await waitForTV(session);
      } catch (e) {
        if (session) await stop(session).catch(() => {});
        const pictureError = session && !session.capture && L.t('VibeOffice could not prepare the document picture. Try casting again.');
        status.textContent = errorMessage(e, pictureError);
        connecting = false;
        d.buttons[0].disabled = d.buttons[1].disabled = tv.disabled = false;
      }
    }
    scan();
  }
  ui.cmd('castTV', { label: L.N('Cast to TV...'), icon: 'cast', checked: () => !!active, run: choose });
})();
