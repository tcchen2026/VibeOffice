//! Document windows own their selected renderer and local stream.
use crate::doc;
use dlnanything::{discover, MediaServer, Renderer, Transport};
use serde::Serialize;
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

#[derive(Default)]
pub struct Casts {
    windows: Mutex<HashMap<String, Arc<Owner>>>,
    control: tauri::async_runtime::Mutex<()>,
}
#[derive(Default)]
struct Owner {
    devices: Mutex<Vec<Renderer>>,
    session: Mutex<Option<Arc<Session>>>,
}
struct Session {
    key: String,
    renderer: Renderer,
    server: MediaServer,
    transport: Transport,
    playing: AtomicBool,
}

impl Casts {
    fn owner(&self, window: &WebviewWindow) -> Result<Arc<Owner>, String> {
        if doc::app_of(window).is_none() {
            return Err("Casting requires an open document window".into());
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
    fn session(&self, key: &str) -> Result<Arc<Session>, String> {
        self.session
            .lock()
            .unwrap()
            .as_ref()
            .filter(|s| s.key == key)
            .cloned()
            .ok_or_else(|| "Cast session has ended".into())
    }
}
#[derive(Serialize)]
pub struct Device {
    id: String,
    name: String,
}

#[tauri::command]
pub async fn host_cast_scan(
    window: WebviewWindow,
    casts: State<'_, Casts>,
) -> Result<Vec<Device>, String> {
    let owner = casts.owner(&window)?;
    let devices = discover(Duration::from_secs(2))
        .await
        .map_err(|e| e.to_string())?;
    let list = devices
        .iter()
        .map(|r| Device {
            id: r.id.clone(),
            name: r.name.clone(),
        })
        .collect();
    *owner.devices.lock().unwrap() = devices;
    Ok(list)
}

#[tauri::command]
pub async fn host_cast_prepare(
    window: WebviewWindow,
    casts: State<'_, Casts>,
    id: String,
    mode: String,
) -> Result<String, String> {
    let owner = casts.owner(&window)?;
    let transport = match mode.as_str() {
        "hls" => Transport::Hls,
        "ts" => Transport::MpegTs,
        _ => return Err("Unknown cast mode".into()),
    };
    let _control = casts.control.lock().await;
    if owner.session.lock().unwrap().is_some() {
        return Err("This window already has a cast".into());
    }
    if casts.uses(&id) {
        return Err("Another document is casting to this TV".into());
    }
    let renderer = owner
        .devices
        .lock()
        .unwrap()
        .iter()
        .find(|r| r.id == id)
        .cloned()
        .ok_or("Scan for TVs again")?;
    let protocols = renderer.sink_protocols().await.unwrap_or_default();
    let server = renderer
        .media_server(&protocols)
        .await
        .map_err(|e| e.to_string())?;
    let key = server.session_id().to_owned();
    *owner.session.lock().unwrap() = Some(Arc::new(Session {
        key: key.clone(),
        renderer,
        server,
        transport,
        playing: AtomicBool::new(false),
    }));
    Ok(key)
}
#[tauri::command]
pub fn host_cast_frame(
    window: WebviewWindow,
    casts: State<Casts>,
    request: Request<'_>,
) -> Result<(), String> {
    let key = request
        .headers()
        .get("x-cast-session")
        .and_then(|x| x.to_str().ok())
        .ok_or("Missing cast session")?;
    let session = casts.owner(&window)?.session(key)?;
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("Expected binary H.264".into());
    };
    session
        .server
        .frames()
        .push_annex_b(bytes)
        .map(|_| ())
        .map_err(|e| e.to_string())
}
#[tauri::command]
pub async fn host_cast_play(
    window: WebviewWindow,
    casts: State<'_, Casts>,
    key: String,
) -> Result<(), String> {
    let _control = casts.control.lock().await;
    let session = casts.owner(&window)?.session(&key)?;
    if !session.playing.load(Ordering::Acquire) {
        session
            .renderer
            .play(&session.server, session.transport, "VibeOffice")
            .await
            .map_err(|e| e.to_string())?;
        session.playing.store(true, Ordering::Release);
    }
    Ok(())
}
#[tauri::command]
pub async fn host_cast_stop(
    window: WebviewWindow,
    casts: State<'_, Casts>,
    key: String,
) -> Result<(), String> {
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
        receiver.stop().await.map_err(|e| e.to_string())?;
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
