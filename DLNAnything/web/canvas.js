/* DLNAnything • Apache-2.0. A classic script for native webviews. */
(function (global) {
  'use strict';
  const defaultFps = 2;

  // Prefer Annex B. Some WebCodecs encoders expose AVC samples with their
  // parameter sets in decoderConfig.description instead.
  async function encoderConfig({ width, height, fps = defaultFps, bitrate = 2500000, codec }) {
    if (!Number.isInteger(fps) || fps < 2 || fps > 30) throw new RangeError('fps must be an integer from 2 to 30');
    if (!global.VideoEncoder || !global.VideoFrame) return null;
    width = Math.ceil(width / 2) * 2;
    height = Math.ceil(height / 2) * 2;
    const blocks = Math.ceil(width / 16) * Math.ceil(height / 16);
    const profile = codec || (blocks <= 3600 ? 'avc1.42E01F' : blocks <= 8192 ? 'avc1.42E028' : 'avc1.42E033');
    for (const format of ['annexb', 'avc']) {
      const config = { codec: profile, width, height, bitrate, framerate: fps, latencyMode: 'realtime', avc: { format } };
      try {
        if ((await VideoEncoder.isConfigSupported(config)).supported) return config;
      } catch (e) {
        if (e.name !== 'TypeError' && e.name !== 'NotSupportedError') throw e;
      }
    }
    return null;
  }

  function avcDescription(description) {
    const data = ArrayBuffer.isView(description)
      ? new Uint8Array(description.buffer, description.byteOffset, description.byteLength) : new Uint8Array(description);
    const invalid = () => new Error('Invalid H.264 decoder configuration');
    if (data.length < 7 || data[0] !== 1) throw invalid();
    const lengthSize = (data[4] & 3) + 1, sets = [];
    if (lengthSize === 3) throw invalid();
    let pos = 6;
    function readSets(count) {
      for (let i = 0; i < count; i++) {
        if (pos + 2 > data.length) throw invalid();
        const size = data[pos] * 256 + data[pos + 1];
        pos += 2;
        if (!size || pos + size > data.length) throw invalid();
        sets.push(data.slice(pos, pos + size));
        pos += size;
      }
    }
    readSets(data[5] & 31);
    if (pos >= data.length) throw invalid();
    readSets(data[pos++]);
    return { lengthSize, sets };
  }

  function annexB(sample, { lengthSize, sets }, key) {
    const units = [];
    for (let pos = 0; pos < sample.length;) {
      if (pos + lengthSize > sample.length) throw new Error('Truncated H.264 length prefix');
      let size = 0;
      for (let i = 0; i < lengthSize; i++) size = size * 256 + sample[pos++];
      if (!size || pos + size > sample.length) throw new Error('Invalid H.264 NAL unit length');
      units.push(sample.subarray(pos, pos + size));
      pos += size;
    }
    if (!units.length) throw new Error('Empty H.264 picture');
    // An access unit delimiter, if present, must precede the parameter sets.
    if (key) units.splice((units[0][0] & 31) === 9 ? 1 : 0, 0, ...sets);
    const result = new Uint8Array(units.reduce((n, unit) => n + 4 + unit.length, 0));
    let offset = 0;
    for (const unit of units) {
      result[offset + 3] = 1;
      result.set(unit, offset + 4);
      offset += 4 + unit.length;
    }
    return result;
  }

  // sendFrame receives exactly one Annex B access unit per call, in order.
  // Awaiting the transport applies backpressure to capture as well as IPC.
  async function startCanvas(canvas, { sendFrame, beforeFrame, onError, bitrate = 2500000,
    codec, width, height, fps = defaultFps } = {}) {
    const period = 1000 / fps;
    if (typeof sendFrame !== 'function') throw new TypeError('sendFrame is required');
    if (!canvas || !canvas.width || !canvas.height) throw new TypeError('A nonempty canvas is required');
    if (!global.VideoEncoder || !global.VideoFrame) throw new Error('H.264 WebCodecs encoding is unavailable in this webview.');
    const config = await encoderConfig({ width: width || canvas.width, height: height || canvas.height, bitrate, fps, codec });
    if (!config) throw new Error('This webview does not support the requested H.264 encoding configuration.');
    ({ width, height } = config);
    let stopped = false, failure, encoded, timer, wake, scaled, keyFrame, avc;
    const encoder = new VideoEncoder({
      output(chunk, metadata) {
        try {
          if ((keyFrame && chunk.type !== 'key') || encoded) throw new Error('Expected one picture and an IDR at each segment boundary');
          if (metadata?.decoderConfig) {
            const desc = metadata.decoderConfig.description;
            avc = desc ? avcDescription(desc) : undefined;
          }
          if (config.avc.format === 'avc' && !avc) throw new Error('The H.264 encoder omitted its AVC configuration');
          const sample = new Uint8Array(chunk.byteLength);
          chunk.copyTo(sample);
          encoded = avc ? annexB(sample, avc, chunk.type === 'key') : sample;
        } catch (e) { failure = e; }
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
  global.DLNAnything = Object.freeze({ startCanvas, defaultFps, isSupported: async options => !!(await encoderConfig(options)) });
})(globalThis);
