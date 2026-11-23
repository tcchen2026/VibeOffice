//! A native canvas-casting sender. All network activity starts with an explicit call.
mod discovery;
mod http;
mod renderer;
mod session;
mod ts;

pub use http::MediaServer;
pub use discovery::{discover, discover_on, discover_with};
pub use renderer::Renderer;
pub use session::{FrameSink, StreamConfig, MAX_FRAME_BYTES, TARGET_SEGMENT_DURATION};
/// A standalone classic script for embedding in a native application's own webview.
pub const CANVAS_JS: &str = include_str!("../web/canvas.js");
