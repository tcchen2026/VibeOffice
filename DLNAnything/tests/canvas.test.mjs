import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const source = readFileSync(new URL('../web/canvas.js', import.meta.url), 'utf8');
const canvas = { width: 64, height: 48 };

function webview() {
  let now = 0, timer;
  const encoded = [];
  class Encoder {
    static async isConfigSupported() { return { supported: true }; }
    constructor(callbacks) { this.callbacks = callbacks; }
    configure(config) { this.config = config; this.state = 'configured'; }
    encode(frame, { keyFrame }) {
      encoded.push({ fps: this.config.framerate, keyFrame, timestamp: frame.timestamp, duration: frame.duration });
      this.callbacks.output({ type: keyFrame ? 'key' : 'delta', byteLength: 1, copyTo(bytes) { bytes[0] = 7; } });
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
