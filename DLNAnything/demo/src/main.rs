#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]
use dlnanything::{discover, MediaServer, Renderer, Transport};
use serde::Serialize;
use std::{
    sync::{Arc, Mutex},
    time::Duration,
};
use tauri::{
    ipc::{InvokeBody, Request},
    State,
};
use tokio::sync::Mutex as AsyncMutex;

#[derive(Default)]
struct Demo {
    devices: Mutex<Vec<Renderer>>,
    active: Mutex<Option<Active>>,
    operation: AsyncMutex<()>,
}
#[derive(Clone)]
struct Active {
    key: String,
    server: Arc<MediaServer>,
    receiver: Option<Renderer>,
    transport: Transport,
}
#[derive(Serialize)]
struct DeviceInfo {
    id: String,
    name: String,
}
#[derive(Serialize)]
struct Prepared {
    key: String,
    url: String,
}

#[tauri::command]
async fn scan(state: State<'_, Demo>) -> Result<Vec<DeviceInfo>, String> {
    let devices = discover(Duration::from_secs(2))
        .await
        .map_err(|e| e.to_string())?;
    let info = devices
        .iter()
        .map(|r| DeviceInfo {
            id: r.id.clone(),
            name: r.name.clone(),
        })
        .collect();
    *state.devices.lock().unwrap() = devices;
    Ok(info)
}
fn session(state: &Demo, key: &str) -> Result<Active, String> {
    state
        .active
        .lock()
        .unwrap()
        .as_ref()
        .filter(|a| a.key == key)
        .cloned()
        .ok_or_else(|| "Cast session has ended".into())
}
#[tauri::command]
async fn prepare(
    state: State<'_, Demo>,
    receiver: Option<String>,
    mode: String,
) -> Result<Prepared, String> {
    let _guard = state.operation.lock().await;
    if state.active.lock().unwrap().is_some() {
        return Err("A cast is already open".into());
    }
    let transport = match mode.as_str() {
        "hls" => Transport::Hls,
        "ts" => Transport::MpegTs,
        _ => return Err("Unknown transport".into()),
    };
    let receiver = receiver
        .map(|id| {
            state
                .devices
                .lock()
                .unwrap()
                .iter()
                .find(|r| r.id == id)
                .cloned()
                .ok_or_else(|| "Scan for the renderer again".to_owned())
        })
        .transpose()?;
    let server = if let Some(r) = &receiver {
        let protocols = r.sink_protocols().await.unwrap_or_default();
        r.media_server(&protocols).await
    } else {
        MediaServer::bind(([127, 0, 0, 1], 0).into()).await
    }
    .map_err(|e| e.to_string())?;
    let url = match transport {
        Transport::Hls => server.playlist_url(),
        Transport::MpegTs => server.transport_stream_url(),
    };
    let key = url.rsplit('/').nth(1).unwrap().to_owned();
    *state.active.lock().unwrap() = Some(Active {
        key: key.clone(),
        server: Arc::new(server),
        receiver,
        transport,
    });
    Ok(Prepared { key, url })
}
#[tauri::command]
fn push_frame(state: State<'_, Demo>, request: Request<'_>) -> Result<(), String> {
    let key = request
        .headers()
        .get("x-cast-session")
        .and_then(|v| v.to_str().ok())
        .ok_or("Missing session")?;
    let active = session(&state, key)?;
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("Expected a binary H.264 frame".into());
    };
    active
        .server
        .frames()
        .push_annex_b(bytes)
        .map(|_| ())
        .map_err(|e| e.to_string())
}
#[tauri::command]
async fn play(state: State<'_, Demo>, key: String) -> Result<(), String> {
    let _guard = state.operation.lock().await;
    let a = session(&state, &key)?;
    if let Some(r) = a.receiver {
        r.play(&a.server, a.transport, "DLNAnything canvas")
            .await
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}
#[tauri::command]
async fn stop(state: State<'_, Demo>, key: String) -> Result<(), String> {
    let _guard = state.operation.lock().await;
    let receiver = {
        let mut slot = state.active.lock().unwrap();
        if !slot.as_ref().is_some_and(|a| a.key == key) {
            return Ok(());
        }
        slot.take().and_then(|a| a.receiver)
    };
    if let Some(r) = receiver {
        r.stop().await.map_err(|e| e.to_string())?;
    }
    Ok(())
}
fn main() {
    tauri::Builder::default()
        .manage(Demo::default())
        .invoke_handler(tauri::generate_handler![
            scan, prepare, push_frame, play, stop
        ])
        .run(tauri::generate_context!())
        .expect("could not run DLNAnything");
}
