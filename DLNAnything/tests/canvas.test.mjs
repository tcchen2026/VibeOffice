import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const source = readFileSync(new URL('../web/canvas.js', import.meta.url), 'utf8');
const canvas = { width: 64, height: 48 };

function webview({ probe = () => true, output = () => ({ data: [7] }) } = {}) {
  let now = 0, timer;
  const encoded = [];
  class Encoder {
    static async isConfigSupported(config) { return { supported: await probe(config.avc.format) }; }
    constructor(callbacks) { this.callbacks = callbacks; }
    configure(config) { this.config = config; this.state = 'configured'; }
    encode(frame, { keyFrame }) {
      encoded.push({ fps: this.config.framerate, keyFrame, timestamp: frame.timestamp, duration: frame.duration });
      const { data, description } = output(encoded.length - 1, keyFrame);
      this.callbacks.output({ type: keyFrame ? 'key' : 'delta', byteLength: data.length, copyTo(bytes) { bytes.set(data); } },
        description && { decoderConfig: { description } });
    }
    async flush() {}
    close() { this.state = 'closed'; }
  }
  class Frame {
    constructor(_, options) { Object.assign(this, options); }
    close() {}
  }
  const globals = { VideoEncoder: Encoder, VideoFrame: Frame, performance: { now: () => now },
    setTimeout(callback, ms) { timer = { callback, due: now + ms }; return 1; },
    clearTimeout() { timer = null; } };
  vm.runInNewContext(source, globals);
  return { api: globals.DLNAnything, encoded, get timer() { return timer; },
    async tick(delay = 0) {
      now = timer.due + delay;
      const { callback } = timer; timer = null;
      callback();
      await new Promise(resolve => setImmediate(resolve));
    } };
}

test('every supported fps rounds half-second boundaries to whole frames', async () => {
  for (let fps = 2; fps <= 30; fps++) {
    const view = webview();
    const opts = fps === 2 ? {} : { fps };
    const capture = await view.api.startCanvas(canvas, { ...opts, sendFrame(bytes) { assert.equal(bytes[0], 7); } });
    for (let i = 1; i < fps * 2; i++) await view.tick();
    await capture.stop();
    assert.equal(view.timer, null);
    assert.deepEqual(view.encoded.map((f,i) => f.keyFrame ? i : null).filter(i => i !== null),
      [0, Math.ceil(fps / 2), fps, Math.ceil(fps * 1.5)]);
    assert.equal(view.encoded.at(-1).timestamp + view.encoded.at(-1).duration, 2000000);
    assert.ok(view.encoded.every(f => f.fps === fps));
  }
  for (const fps of [1, 31, 2.5]) {
    await assert.rejects(webview().api.startCanvas(canvas, { fps, sendFrame() {} }), /fps/);
  }
});

test('a slow host has one outstanding frame and stop waits for its acknowledgement', async () => {
  const view = webview();
  let sent = 0, acknowledge;
  const capture = await view.api.startCanvas(canvas, {
    sendFrame() { if (++sent === 2) return new Promise(resolve => { acknowledge = resolve; }); }
  });
  await view.tick(1500);
  assert.equal(sent, 2);
  assert.equal(view.timer, null);
  const stopped = capture.stop();
  acknowledge();
  await stopped;
  assert.equal(view.timer, null);
  assert.equal(view.encoded.length, 2);
  assert.deepEqual(view.encoded.map(f => f.timestamp), [0, 500000]);
});

const sps = [0x67, 0x42], pps = [0x68, 0x80], aud = [9, 0xf0], idr = [0x65, 0x88], delta = [0x41, 0x9a];
function description(lengthSize) {
  return Uint8Array.from([1, 0x42, 0, 31, 0xfc | (lengthSize - 1), 0xe1, 0, sps.length, ...sps, 1, 0, pps.length, ...pps]);
}

test('AVC samples become Annex B with parameter sets on every keyframe', async () => {
  for (const lengthSize of [1, 2, 4]) {
    const probes = [], sent = [];
    const view = webview({
      probe(format) { probes.push(format); return format === 'avc'; },
      output(index, key) {
        const units = key ? [aud, idr] : [delta];
        return { data: units.flatMap(unit => [...Array(lengthSize - 1).fill(0), unit.length, ...unit]),
          description: index === 0 && description(lengthSize) };
      }
    });
    const capture = await view.api.startCanvas(canvas, { fps: 4, sendFrame(bytes) { sent.push([...bytes]); } });
    await view.tick(); await view.tick(); await capture.stop();
    assert.deepEqual(probes, ['annexb', 'avc']);
    const startCodes = units => units.flatMap(unit => [0, 0, 0, 1, ...unit]);
    const key = startCodes([aud, sps, pps, idr]);
    assert.deepEqual(sent, [key, startCodes([delta]), key]);
  }
});

test('the decoder description identifies AVC even if the encoder accepted the Annex B option', async () => {
  const view = webview({ output: () => ({ data: [0, 2, ...idr], description: description(2) }) });
  let sent;
  const capture = await view.api.startCanvas(canvas, { sendFrame(bytes) { sent = [...bytes]; } });
  await capture.stop();
  assert.deepEqual(sent, [0, 0, 0, 1, ...sps, 0, 0, 0, 1, ...pps, 0, 0, 0, 1, ...idr]);
});

test('capability probing and malformed AVC output leave the host untouched', async () => {
  const unsupported = webview({ probe() { throw new TypeError('unsupported format'); } });
  assert.equal(await unsupported.api.isSupported(canvas), false);
  assert.equal(unsupported.encoded.length, 0);
  for (const [config, data, message] of [
    [undefined, [0], /omitted its AVC configuration/],
    [description(3), [0], /Invalid H.264 decoder configuration/],
    [description(4).slice(0, -1), [0], /Invalid H.264 decoder configuration/],
    [description(4), [0, 1], /Truncated H.264 length prefix/],
    [description(4), [0, 0, 0, 3, 0x65], /Invalid H.264 NAL unit length/]
  ]) {
    const view = webview({ probe: format => format === 'avc', output: () => ({ description: config, data }) });
    await assert.rejects(view.api.startCanvas(canvas, { sendFrame() { assert.fail('sent a malformed frame'); } }), message);
    assert.equal(view.timer, undefined);
  }
});
