use crate::session::{FrameSink, Missing, StreamConfig, Timeline};
use axum::{
    body::Body,
    extract::{ConnectInfo, Path, Query, State},
    http::{header, HeaderMap, Response, StatusCode},
    routing::get,
    Router,
};
use bytes::Bytes;
use serde::Deserialize;
use std::{
    collections::HashMap,
    io,
    net::{IpAddr, SocketAddr},
    sync::{Arc, Mutex},
    time::Duration,
};
use tokio::{
    net::TcpListener,
    sync::{oneshot, Semaphore},
    time::Instant,
};

const RELOAD_WAIT: Duration = Duration::from_secs(3);
const LIVE_IDLE: Duration = Duration::from_secs(15);
type Reply = Response<Body>;

struct Hub {
    timeline: Arc<Timeline>,
    playlists: Mutex<HashMap<IpAddr, u64>>,
    clients: Arc<Semaphore>,
    hls_type: String,
    ts_type: String,
}

/// A single cast session. Dropping this owner closes the listener and wakes outstanding requests.
pub struct MediaServer {
    base: String,
    hub: Arc<Hub>,
    stop: Option<oneshot::Sender<()>>,
}

impl MediaServer {
    /// Bind the TV-facing local interface. Port zero asks the OS to choose a port.
    pub async fn bind(address: SocketAddr) -> io::Result<Self> {
        Self::bind_with_config(address, StreamConfig::default()).await
    }

    pub async fn bind_with_config(address: SocketAddr, config: StreamConfig) -> io::Result<Self> {
        Self::bind_types(
            address,
            "application/vnd.apple.mpegurl",
            "video/mp2t",
            config,
        )
        .await
    }

    pub(crate) async fn bind_types(
        address: SocketAddr,
        hls_type: &str,
        ts_type: &str,
        config: StreamConfig,
    ) -> io::Result<Self> {
        if address.ip().is_unspecified() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "choose a local interface reachable from the renderer",
            ));
        }
        let listener = TcpListener::bind(address).await?;
        let path = format!("/cast/{}", uuid::Uuid::new_v4().simple());
        let base = format!("http://{}{path}", listener.local_addr()?);
        let hub = Arc::new(Hub {
            timeline: Arc::new(Timeline::new(config)),
            playlists: Mutex::new(HashMap::new()),
            clients: Arc::new(Semaphore::new(32)),
            hls_type: hls_type.to_owned(),
            ts_type: ts_type.to_owned(),
        });
        let files = Router::new()
            .route("/{name}", get(media))
            .with_state(hub.clone());
        let app = Router::new().nest(&path, files);
        let (stop, done) = oneshot::channel();
        tokio::spawn(async move {
            let _ = axum::serve(
                listener,
                app.into_make_service_with_connect_info::<SocketAddr>(),
            )
            .with_graceful_shutdown(async {
                let _ = done.await;
            })
            .await;
        });
        Ok(Self {
            base,
            hub,
            stop: Some(stop),
        })
    }

    pub fn playlist_url(&self) -> String {
        format!("{}/live.m3u8", self.base)
    }
    pub fn session_id(&self) -> &str {
        self.base.rsplit('/').next().unwrap()
    }
    pub fn transport_stream_url(&self) -> String {
        format!("{}/live.ts", self.base)
    }
    pub fn frames(&self) -> FrameSink {
        FrameSink(self.hub.timeline.clone())
    }
    pub(crate) fn hls_type(&self) -> &str {
        &self.hub.hls_type
    }
    pub(crate) fn ts_type(&self) -> &str {
        &self.hub.ts_type
    }
}

impl Drop for MediaServer {
    fn drop(&mut self) {
        self.hub.timeline.close();
        if let Some(stop) = self.stop.take() {
            let _ = stop.send(());
        }
    }
}

#[derive(Deserialize)]
struct PlaylistQuery {
    #[serde(rename = "_HLS_msn")]
    msn: Option<u64>,
}

fn empty(status: StatusCode) -> Reply {
    Response::builder()
        .status(status)
        .body(Body::empty())
        .unwrap()
}

fn builder(mime: &str, cache: &str) -> axum::http::response::Builder {
    Response::builder()
        .header(header::CONTENT_TYPE, mime)
        .header(header::CACHE_CONTROL, cache)
        .header("transfermode.dlna.org", "Streaming")
}

fn error_status(reason: Missing) -> StatusCode {
    match reason {
        Missing::Expired => StatusCode::GONE,
        Missing::Future => StatusCode::BAD_REQUEST,
        Missing::Closed | Missing::Timeout => StatusCode::SERVICE_UNAVAILABLE,
    }
}

fn io_error(reason: Missing) -> io::Error {
    io::Error::other(format!("media unavailable: {reason:?}"))
}

async fn media(
    State(hub): State<Arc<Hub>>,
    ConnectInfo(peer): ConnectInfo<SocketAddr>,
    Path(name): Path<String>,
    Query(q): Query<PlaylistQuery>,
    headers: HeaderMap,
) -> Reply {
    match name.as_str() {
        "live.m3u8" => playlist(hub, peer.ip(), q.msn),
        "live.ts" => live(hub),
        _ => {
            let Some(seq) = name.strip_suffix(".ts").and_then(|s| s.parse::<u64>().ok()) else {
                return empty(StatusCode::NOT_FOUND);
            };
            segment(hub, seq, &headers)
        }
    }
}

fn playlist(hub: Arc<Hub>, peer: IpAddr, msn: Option<u64>) -> Reply {
    let response = builder(&hub.hls_type, "no-store");
    let latest = hub.timeline.latest();
    let target = match msn {
        Some(seq) => {
            if matches!(hub.timeline.lookup(seq), Err(Missing::Future)) {
                return empty(StatusCode::BAD_REQUEST);
            }
            seq
        }
        None => {
            let cursors = hub.playlists.lock().unwrap();
            match latest {
                Some(n) if cursors.get(&peer) == Some(&n) => n + 1,
                Some(n) => n,
                None => 0,
            }
        }
    };
    let Ok(permit) = hub.clients.clone().try_acquire_owned() else {
        return empty(StatusCode::SERVICE_UNAVAILABLE);
    };
    // Return the response *before* awaiting media. Hyper flushes headers while its body is pending.
    let stream = async_stream::try_stream! {
        let _permit = permit;
        if latest.is_none_or(|n| n < target) {
            match hub.timeline.wait(target, Instant::now() + RELOAD_WAIT).await {
                Ok(_) | Err(Missing::Expired) => (),
                Err(Missing::Timeout) if msn.is_none() && latest.is_some() => (),
                Err(e) => Err(io_error(e))?,
            }
        }
        let (last, text) = hub.timeline.playlist().ok_or_else(|| io_error(Missing::Closed))?;
        {
            let mut cursors = hub.playlists.lock().unwrap();
            if cursors.len() >= 32 && !cursors.contains_key(&peer) {
                cursors.clear();
            }
            let cursor = cursors.entry(peer).or_insert(last);
            *cursor = (*cursor).max(last);
        }
        yield Bytes::from(text);
    };
    response.body(body(stream)).unwrap()
}

fn body<S>(stream: S) -> Body
where
    S: futures_core::Stream<Item = Result<Bytes, io::Error>> + Send + 'static,
{
    Body::from_stream(stream)
}

fn live(hub: Arc<Hub>) -> Reply {
    let response = builder(&hub.ts_type, "no-store").header(header::ACCEPT_RANGES, "none");
    let Ok(permit) = hub.clients.clone().try_acquire_owned() else {
        return empty(StatusCode::SERVICE_UNAVAILABLE);
    };
    let mut seq = hub.timeline.latest().unwrap_or(0);
    let stream = async_stream::try_stream! {
        let _permit = permit;
        loop {
            match hub.timeline.wait(seq, Instant::now() + LIVE_IDLE).await {
                Ok(segment) => {
                    seq += 1;
                    yield segment.bytes;
                }
                Err(Missing::Closed) => break,
                Err(e) => Err(io_error(e))?,
            }
        }
    };
    response.body(body(stream)).unwrap()
}

// Immutable, complete segments support a single HTTP byte range, including suffix and retry reads.
fn byte_range(value: &str, len: usize) -> Option<std::ops::Range<usize>> {
    let (first, last) = value.strip_prefix("bytes=")?.split_once('-')?;
    if first.is_empty() {
        let count = last.parse::<usize>().ok()?.min(len);
        return (count > 0).then_some(len - count..len);
    }
    let start = first.parse::<usize>().ok()?;
    let end = if last.is_empty() {
        len
    } else {
        last.parse::<usize>().ok()?.saturating_add(1).min(len)
    };
    (start < end).then_some(start..end)
}

fn segment(hub: Arc<Hub>, seq: u64, headers: &HeaderMap) -> Reply {
    let known = match hub.timeline.lookup(seq) {
        Ok(s) => s,
        Err(reason) => return empty(error_status(reason)),
    };
    let mut response = builder("video/mp2t", "private, max-age=60");
    let mut range = None;
    if let Some(s) = &known {
        let len = s.bytes.len();
        response = response.header(header::ACCEPT_RANGES, "bytes");
        if let Some(h) = headers.get(header::RANGE) {
            let Some(r) = h.to_str().ok().and_then(|v| byte_range(v, len)) else {
                return Response::builder()
                    .status(StatusCode::RANGE_NOT_SATISFIABLE)
                    .header(header::CONTENT_RANGE, format!("bytes */{len}"))
                    .body(Body::empty())
                    .unwrap();
            };
            response = response
                .status(StatusCode::PARTIAL_CONTENT)
                .header(
                    header::CONTENT_RANGE,
                    format!("bytes {}-{}/{len}", r.start, r.end - 1),
                )
                .header(header::CONTENT_LENGTH, r.len());
            range = Some(r);
        } else {
            response = response.header(header::CONTENT_LENGTH, len);
        }
    }
    if let Some(s) = known {
        if s.due <= Instant::now() {
            let bytes = range.map_or_else(|| s.bytes.clone(), |r| s.bytes.slice(r));
            return response.body(Body::from(bytes)).unwrap();
        }
    }
    let Ok(permit) = hub.clients.clone().try_acquire_owned() else {
        return empty(StatusCode::SERVICE_UNAVAILABLE);
    };
    let stream = async_stream::try_stream! {
        let _permit = permit;
        let s = hub.timeline.wait(seq, Instant::now() + RELOAD_WAIT).await.map_err(io_error)?;
        yield range.map_or_else(|| s.bytes.clone(), |r| s.bytes.slice(r));
    };
    response.body(body(stream)).unwrap()
}
