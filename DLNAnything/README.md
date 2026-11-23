# DLNAnything

Cast a canvas from a native app to a UPnP AV/DLNA MediaRenderer on the same network. A vanilla JavaScript producer encodes H.264 with WebCodecs. A Rust library serves the pictures as an HLS playlist with MPEG-TS segments. The receiver gets one URL and one Play command per cast. Scrolling, editing and changing slides only update the video.

## Embed in a native app

Use a local Cargo dependency pointing to this directory. The crate has its own manifest and lockfile. It can be copied out of the VibeOffice monorepo. The optional Tauri example lives under `demo/` and has a separate manifest.

```toml
[dependencies]
dlnanything = { path = "../DLNAnything" }
```

Start discovery when the user chooses to cast. Present the returned renderer names and keep the chosen `Renderer` in the native host. The host prepares a session on the local interface that reaches that renderer.

```rust
use dlnanything::{MediaServer, Renderer, StreamConfig};
use std::io;

async fn prepare(tv: &Renderer, fps: u8) -> io::Result<MediaServer> {
    let protocols = tv.sink_protocols().await.unwrap_or_default();
    tv.media_server_with_config(&protocols, StreamConfig::new(fps)?).await
}

async fn start(tv: &Renderer, server: &MediaServer) -> io::Result<()> {
    tv.play(server, "Document view").await
}
```

- `discover(Duration)` searches each active multicast-capable private or link-local IPv4 interface, sending three requests during the first second. `discover_with(duration, callback)` delivers each distinct TV as its description arrives. `discover_on(ip, duration)` selects the interface. `Renderer::from_location(url)` accepts a known device-description URL.
- Inject `dlnanything::CANVAS_JS` into your native webview, or load the same classic script from `web/canvas.js`. In a secure context, call `await DLNAnything.isSupported({width, height, fps})` to check the encoder before requesting local network access. The producer accepts both WebCodecs H.264 formats and delivers Annex B to the host.
- Bridge each binary `sendFrame` call to `server.frames().push_annex_b(bytes)`. Retain the server for the cast and propagate a failed push back to the producer. Capture order is decode order.
- Start capture before calling `play`. The producer resolves after its first encoded picture reaches the host. For a multi-picture segment the server can accept a receiver request while the remaining pictures arrive.
- To attach a local player during development, bind a specific address with `MediaServer::bind_with_config` and read `playlist_url()`.
- After Play, `server.wait_for_request(timeout).await` tells you whether the receiver has requested the playlist or a segment. It includes HEAD and remembers requests that arrived before Play returned.
- When the user ends a cast, stop the producer, call `tv.stop().await` and drop the server. Dropping the server also wakes waiting readers and invalidates existing frame handles.

On macOS, set `NSLocalNetworkUsageDescription` in the host app’s Info.plist. A sandboxed host needs the outgoing and incoming network entitlements for discovery, playback control and serving the stream. macOS 15 and later can prompt on the first discovery packet and reject that packet while the user decides. Let the user search again after granting access. [Apple’s local network privacy guide](https://developer.apple.com/documentation/technotes/tn3179-understanding-local-network-privacy)

```js
// The native host's frame method accepts a Uint8Array and resolves after the push.
const capture = await DLNAnything.startCanvas(canvas, {
  fps: 2,
  sendFrame: bytes => host.frame(session, bytes),
  onError: error => reportCaptureFailure(error)
});
await host.play(session);
// On Stop
await capture.stop();
await host.stop(session);
```

An existing canvas can keep drawing on its own schedule. To render a document on demand, pass an async `beforeFrame(canvas)` callback. Optional `width` and `height` fix the encoded size and fit a resized source into it. Optional `codec` and `bitrate` configure the encoder. A rejected startup reports an unsupported encoder or a failed first frame. Later failures call `onError` and resolve `capture.finished` with the error.

## Frame rate and segmentation

Set the same integer rate from 2 to 30 on the producer and the Rust session. Both default to 2. Each segment begins with SPS, PPS and an IDR and can be decoded on its own. Later pictures in the segment can use inter-frame prediction. Pictures are paced even when the canvas has not changed.

| Frames per second | Pictures in successive segments | Segment duration |
| --- | --- | --- |
| 2 | 1, 1 | 500 ms each |
| 3 | 2, 1 | about 667 ms, 333 ms |
| 25 | 13, 12 | 520 ms, 480 ms |
| 30 | 15, 15 | 500 ms each |

Boundary `n` starts at frame `ceil(n × fps / 2)`. PTS and PCR use a 90 kHz clock and each HLS `EXTINF` reports the segment’s actual duration. The webview waits for capture, encode and IPC to finish before submitting another picture. If a machine sleeps or a producer falls behind, it resumes with the current canvas.

## HTTP timing

Each cast gets a private, random URL path and an ephemeral local port. Requests for `live.m3u8` or a near-future numbered `.ts` receive HTTP headers immediately. The first segment becomes available once its last picture arrives. For each later segment the body waits for **both** completion and the previous segment’s presentation interval to elapse.

The playlist advertises complete, released segments. A client’s first ordinary GET receives the current playlist. An unchanged reload from that IP waits for the next release. Explicit `_HLS_msn` requests can also wait for a full segment. A stalled ordinary reload can return the last available playlist after three seconds. A segment that is still missing after the wait ends its response. Published segments are immediately available for retries and support byte ranges. The HLS target duration is one second, the smallest positive integer the tag accepts. The playlist shows up to eight segments and the server retains twenty.

A TV chooses its own startup and playback buffers. Approximately two seconds of glass-to-glass latency is a target for hardware testing.

## Checks

```sh
cargo test
node --test tests/canvas.test.mjs
```

The loopback tests exercise early headers, delayed bodies, late producers, retries, HTTP ranges, receiver control and frame timing. When FFmpeg is installed, they also decode both 30 fps segments individually. The Tauri demo source is independent of these checks.

Licensed under Apache-2.0. [License](LICENSE)
