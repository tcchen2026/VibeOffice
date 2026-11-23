(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const canvas = $('canvas'), ctx = canvas.getContext('2d');
  const palette = [['#143740','#f2c882'], ['#50284c','#fee4bc'], ['#174238','#a9e2d3']];
  let picture = 1, ticket, capture, busy = false, finding = false;
  const stamp = () => new Date().toISOString().slice(11, 23);
  function draw() {
    const [bg, fg] = palette[(picture - 1) % palette.length];
    ctx.fillStyle = bg; ctx.fillRect(0, 0, 960, 540);
    ctx.fillStyle = fg;
    ctx.fillRect(46, 62, 70, 5);
    ctx.font = '28px sans-serif'; ctx.fillText('DLNAnything', 46, 112);
    ctx.font = 'bold 98px sans-serif'; ctx.fillText('Picture ' + picture, 40, 300);
    ctx.font = '36px monospace'; ctx.fillText(stamp() + ' UTC', 46, 466);
  }
  draw();
  setInterval(draw, 250);
  $('change').onclick = () => { picture++; draw(); };
  function ui(status) {
    $('status').textContent = status;
    $('scan').disabled = !!ticket || busy || finding;
    $('receiver').disabled = !!ticket || busy || finding;
    $('cast').disabled = busy || finding;
    $('cast').textContent = ticket ? 'Stop cast' : 'Start cast';
    $('address').hidden = !ticket;
    $('url').value = ticket ? ticket.url : '';
  }
  const invoke = (cmd, data, options) => window.__TAURI_INTERNALS__.invoke(cmd, data, options);
  $('scan').onclick = async () => {
    finding = true; ui('Searching for TVs…');
    try {
      const devices = await invoke('scan');
      const options = devices.map(d => new Option(d.name || d.id, d.id));
      $('receiver').replaceChildren(new Option('Local HTTP test', ''), ...options);
      if (devices.length) $('receiver').value = devices[0].id;
      ui(devices.length ? 'Choose a receiver.' : 'No TVs found on this network.');
    } catch (e) { ui(String(e)); }
    finally { finding = false; ui($('status').textContent); }
  };
  async function end(reason) {
    const current = ticket, running = capture;
    capture = undefined; ticket = undefined;
    if (running) await running.stop();
    let message = reason;
    try { if (current) await invoke('stop', { key: current.key }); }
    catch (e) { message += ' · ' + String(e); }
    return message;
  }
  $('cast').onclick = async () => {
    if (busy) return;
    busy = true;
    if (ticket) {
      ui('Stopping…');
      const status = await end('Stopped');
      busy = false; ui(status); return;
    }
    ui('Starting…');
    try {
      const current = await invoke('prepare', { receiver: $('receiver').value || null });
      ticket = current;
      capture = await DLNAnything.startCanvas(canvas, {
        sendFrame: bytes => invoke('push_frame', bytes, { headers: { 'X-Cast-Session': current.key } }),
        onError: error => { queueMicrotask(async () => {
          if (ticket !== current) return;
          busy = true;
          const status = await end(String(error));
          busy = false; ui(status);
        }); }
      });
      await invoke('play', { key: current.key });
      if (ticket === current) ui('Casting at 2 fps');
    } catch (e) { ui(await end(String(e))); }
    finally { busy = false; ui($('status').textContent); }
  };
})();
