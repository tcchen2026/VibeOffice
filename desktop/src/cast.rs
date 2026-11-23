//! Document windows own their selected renderer and local stream.
use crate::doc;
use dlnanything::{discover_with, MediaServer, Renderer};
use serde::Serialize;
use std::io;
use std::{
    collections::HashMap,
    sync::{
        atomic::{AtomicBool, Ordering},
        Arc, Mutex,
    },
    time::Duration,
};
use tauri::{
    ipc::{InvokeBody, Request},
    AppHandle, Manager, State, WebviewWindow,
};

#[derive(Clone, Serialize)]
#[serde(tag = "kind", content = "detail", rename_all = "camelCase")]
pub enum CastError {
    NotAnswering(String),
    Refused(String),
    Gone,
    Busy,
    Network(String),
}

impl CastError {
    fn network(detail: impl ToString) -> Self {
        Self::Network(detail.to_string())
    }
    fn control(error: io::Error) -> Self {
        match error.kind() {
            io::ErrorKind::TimedOut | io::ErrorKind::NotConnected => {
                Self::NotAnswering(error.to_string())
            }
            io::ErrorKind::Unsupported => Self::Refused(error.to_string()),
            _ => Self::network(error),
        }
    }
}

#[derive(Default)]
pub struct Casts {
    windows: Mutex<HashMap<String, Arc<Owner>>>,
    control: tauri::async_runtime::Mutex<()>,
}
#[derive(Default)]
struct Owner {
    search: Mutex<Option<tokio::sync::watch::Receiver<Search>>>,
    session: Mutex<Option<Arc<Session>>>,
}
struct Session {
    key: String,
    renderer: Renderer,
    server: MediaServer,
    playing: AtomicBool,
}

impl Casts {
    fn owner(&self, window: &WebviewWindow) -> Result<Arc<Owner>, CastError> {
        if doc::app_of(window).is_none() {
            return Err(CastError::Gone);
        }
        Ok(self
            .windows
            .lock()
            .unwrap()
            .entry(window.label().into())
            .or_default()
            .clone())
    }
    fn uses(&self, id: &str) -> bool {
        self.windows.lock().unwrap().values().any(|owner| {
            owner
                .session
                .lock()
                .unwrap()
                .as_ref()
                .is_some_and(|s| s.renderer.id == id)
        })
    }
}
impl Owner {
    fn session(&self, key: &str) -> Result<Arc<Session>, CastError> {
        self.session
            .lock()
            .unwrap()
            .as_ref()
            .filter(|s| s.key == key)
            .cloned()
            .ok_or(CastError::Gone)
    }
}
#[derive(Serialize)]
pub struct Device {
    id: String,
    name: String,
}

#[derive(Default)]
struct Search {
    devices: Vec<Renderer>,
    done: bool,
    error: Option<CastError>,
}

#[derive(Serialize)]
pub struct SearchPage {
    devices: Vec<Device>,
    done: bool,
}

fn start_search() -> tokio::sync::watch::Receiver<Search> {
    let (updates, receive) = tokio::sync::watch::channel(Search::default());
    tauri::async_runtime::spawn(async move {
        let result = discover_with(Duration::from_secs(2), |tv| {
            updates.send_modify(|s| s.devices.push(tv));
        })
        .await;
        updates.send_modify(|s| {
            s.done = true;
            s.error = result.err().map(CastError::network);
        });
    });
    receive
}

#[tauri::command]
pub async fn host_cast_scan(
    window: WebviewWindow,
    casts: State<'_, Casts>,
    since: Option<usize>,
) -> Result<SearchPage, CastError> {
    let owner = casts.owner(&window)?;
    let mut receive = {
        let mut slot = owner.search.lock().unwrap();
        if since.is_none() && slot.as_ref().is_none_or(|current| current.borrow().done) {
            *slot = Some(start_search());
        }
        slot.as_ref().ok_or(CastError::Gone)?.clone()
    };
    let cursor = since.unwrap_or(0);
    let page = receive
        .wait_for(|s| s.done || s.devices.len() > cursor)
        .await
        .map_err(CastError::network)?;
    if let Some(e) = &page.error {
        return Err(e.clone());
    }
    Ok(SearchPage {
        devices: page
            .devices
            .iter()
            .skip(cursor)
            .map(|tv| Device {
                id: tv.id.clone(),
                name: tv.name.clone(),
            })
            .collect(),
        done: page.done,
    })
}

#[tauri::command]
pub async fn host_cast_prepare(
    window: WebviewWindow,
    casts: State<'_, Casts>,
    id: String,
) -> Result<String, CastError> {
    let owner = casts.owner(&window)?;
    let _control = casts.control.lock().await;
    if owner.session.lock().unwrap().is_some() {
        return Err(CastError::Busy);
    }
    if casts.uses(&id) {
        return Err(CastError::Busy);
    }
    let renderer = owner
        .search
        .lock()
        .unwrap()
        .as_ref()
        .and_then(|search| {
            search
                .borrow()
                .devices
                .iter()
                .find(|tv| tv.id == id)
                .cloned()
        })
        .ok_or(CastError::Gone)?;
    let protocols = renderer.sink_protocols().await.unwrap_or_default();
    let server = renderer
        .media_server(&protocols)
        .await
        .map_err(CastError::network)?;
    let key = server.session_id().to_owned();
    *owner.session.lock().unwrap() = Some(Arc::new(Session {
        key: key.clone(),
        renderer,
        server,
        playing: AtomicBool::new(false),
    }));
    Ok(key)
}
#[tauri::command]
pub fn host_cast_frame(
    window: WebviewWindow,
    casts: State<Casts>,
    request: Request<'_>,
) -> Result<(), CastError> {
    let key = request
        .headers()
        .get("x-cast-session")
        .and_then(|x| x.to_str().ok())
        .ok_or(CastError::Gone)?;
    let session = casts.owner(&window)?.session(key)?;
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err(CastError::network("Expected binary H.264"));
    };
    session
        .server
        .frames()
        .push_annex_b(bytes)
        .map(|_| ())
        .map_err(CastError::network)
}
#[tauri::command]
pub async fn host_cast_play(
    window: WebviewWindow,
    casts: State<'_, Casts>,
    key: String,
) -> Result<(), CastError> {
    let _control = casts.control.lock().await;
    let session = casts.owner(&window)?.session(&key)?;
    if !session.playing.load(Ordering::Acquire) {
        session
            .renderer
            .play(&session.server, "VibeOffice")
            .await
            .map_err(CastError::control)?;
        session.playing.store(true, Ordering::Release);
    }
    Ok(())
}
#[tauri::command]
pub async fn host_cast_wait(
    window: WebviewWindow,
    casts: State<'_, Casts>,
    key: String,
) -> Result<bool, CastError> {
    let wait = casts
        .owner(&window)?
        .session(&key)?
        .server
        .wait_for_request(Duration::from_secs(10));
    Ok(wait.await)
}
#[tauri::command]
pub async fn host_cast_stop(
    window: WebviewWindow,
    casts: State<'_, Casts>,
    key: String,
) -> Result<(), CastError> {
    let _control = casts.control.lock().await;
    let owner = casts.owner(&window)?;
    let s = {
        let mut slot = owner.session.lock().unwrap();
        if slot.as_ref().is_none_or(|s| s.key != key) {
            return Ok(());
        }
        slot.take().unwrap()
    };
    let playing = s.playing.load(Ordering::Acquire);
    let receiver = s.renderer.clone();
    drop(s);
    if playing {
        receiver.stop().await.map_err(CastError::control)?;
    }
    Ok(())
}
pub fn leave(app: &AppHandle, label: &str) {
    let Some(casts) = app.try_state::<Casts>() else {
        return;
    };
    let Some(owner) = casts.windows.lock().unwrap().remove(label) else {
        return;
    };
    let session = owner.session.lock().unwrap().take();
    if let Some(s) = session {
        let app = app.clone();
        tauri::async_runtime::spawn(async move {
            let casts = app.state::<Casts>();
            let _control = casts.control.lock().await;
            let playing = s.playing.load(Ordering::Acquire);
            let receiver = s.renderer.clone();
            drop(s);
            if playing && !casts.uses(&receiver.id) {
                let _ = receiver.stop().await;
            }
        });
    }
}
