/* DLNAnything • Apache-2.0. A classic script for native webviews. */
(function (global) {
  'use strict';
  const defaultFps = 2;

  // sendFrame receives exactly one Annex B access unit per call, in order.
  // Awaiting the transport applies backpressure to capture as well as IPC.
  async function startCanvas(canvas, { sendFrame, beforeFrame, onError, bitrate = 2500000,
    codec, width, height, fps = defaultFps } = {}) {
    if (!Number.isInteger(fps) || fps < 2 || fps > 30) throw new RangeError('fps must be an integer from 2 to 30');
    const period = 1000 / fps;
    if (typeof sendFrame !== 'function') throw new TypeError('sendFrame is required');
    if (!canvas || !canvas.width || !canvas.height) throw new TypeError('A nonempty canvas is required');
    if (!global.VideoEncoder || !global.VideoFrame) throw new Error('H.264 WebCodecs encoding is unavailable in this webview.');
    width = Math.ceil((width || canvas.width) / 2) * 2;
    height = Math.ceil((height || canvas.height) / 2) * 2;
    const blocks = Math.ceil(width / 16) * Math.ceil(height / 16);
    const config = { codec: codec || (blocks <= 3600 ? 'avc1.42E01F' : blocks <= 8192 ? 'avc1.42E028' : 'avc1.42E033'),
      width, height, bitrate, framerate: fps,
      latencyMode: 'realtime', avc: { format: 'annexb' } };
    if (!(await VideoEncoder.isConfigSupported(config)).supported) {
      throw new Error('This webview does not support the requested H.264 encoding configuration.');
    }
    let stopped = false, failure, encoded, timer, wake, scaled, keyFrame;
    const encoder = new VideoEncoder({
      output(chunk) {
        if ((keyFrame && chunk.type !== 'key') || encoded) { failure = new Error('Expected one picture and an IDR at each segment boundary'); return; }
        encoded = new Uint8Array(chunk.byteLength);
        chunk.copyTo(encoded);
      },
      error(error) { failure = error; }
    });
    const closeEncoder = () => { if (encoder.state !== 'closed') encoder.close(); };
    async function capture(index) {
      if (beforeFrame) await beforeFrame(canvas);
      if (stopped) return;
      let source = canvas;
      if (canvas.width !== width || canvas.height !== height) {
        if (!scaled) {
          scaled = global.OffscreenCanvas ? new OffscreenCanvas(width, height) : document.createElement('canvas');
          scaled.width = width; scaled.height = height;
        }
        const ctx = scaled.getContext('2d'), fit = Math.min(width / canvas.width, height / canvas.height);
        ctx.fillStyle = '#000'; ctx.fillRect(0, 0, width, height);
        if (Number.isFinite(fit)) ctx.drawImage(canvas, (width-canvas.width*fit)/2, (height-canvas.height*fit)/2, canvas.width*fit, canvas.height*fit);
        source = scaled;
      }
      encoded = undefined;
      const timestamp = Math.round(index * 1000000 / fps);
      const duration = Math.round((index + 1) * 1000000 / fps) - timestamp;
      keyFrame = index === 0 || Math.floor(index * 2 / fps) !== Math.floor((index - 1) * 2 / fps);
      const frame = new VideoFrame(source, { timestamp, duration });
      try { encoder.encode(frame, { keyFrame }); } finally { frame.close(); }
      await encoder.flush();
      if (failure) throw failure;
      if (!encoded) throw new Error('The H.264 encoder returned no picture');
      if (!stopped) await sendFrame(encoded);
    }
    const start = performance.now();
    try {
      encoder.configure(config);
      await capture(0);
    } catch (e) { closeEncoder(); throw e; }

    const finished = (async () => {
      let index = 1, due = start + period;
      try {
        while (!stopped) {
          await new Promise(resolve => { wake = resolve; timer = setTimeout(resolve, Math.max(0, due - performance.now())); });
          wake = undefined;
          if (stopped) break;
          const capturedAt = performance.now();
          await capture(index++);
          // A suspended webview resumes with the current canvas and never emits a catch-up burst.
          due = Math.max(due + period, capturedAt + period - 1);
        }
      } catch (e) {
        if (!stopped) { failure = e; if (onError) onError(e); }
      } finally { stopped = true; closeEncoder(); }
      return failure;
    })();
    return {
      finished,
      async stop() { stopped = true; clearTimeout(timer); if (wake) wake(); return finished; }
    };
  }
  global.DLNAnything = Object.freeze({ startCanvas, defaultFps });
})(globalThis);
