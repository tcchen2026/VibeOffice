//! The document in each window (todos/DESKTOP_REAL_SAVE.md, Phase 1, The model): a window is one page and one
//! document, and only Rust gives a window its file. The system's files (the command line, Finder, the Open
//! dialog) and Recent Files open here, each in a window of its own or in the window that already has it; a page
//! never passes a path. Recent Files, their previews and drafts live in ~/.vibeoffice (store.rs).
//!
//! A window made for a document opens its app as `<app>/?doc`; the page asks `host_document` what it holds and
//! `host_document_bytes` for it (VO.attach, public/common/suite.js).

use std::{
    collections::{HashMap, HashSet},
    path::{Path, PathBuf},
    sync::Mutex,
};

use serde::{Deserialize, Serialize};
use tauri::{ipc::Request, ipc::Response, AppHandle, Manager, State, WebviewUrl, WebviewWindow};

use crate::{
    dialog,
    files::{self, Base, Over, SaveError, MAX_FILE},
    store::{self, Draft, Entry, Store},
    types::{self, Kind},
};

pub struct Docs {
    pub store: Store,
    /// window label -> its document
    windows: Mutex<HashMap<String, Doc>>,
    /// windows whose page was asked to close (`vo-close`) and hasn't taken it in hand yet
    asked: Mutex<HashSet<String>>,
    /// files being written now: two writes to one file never overlap
    writing: Mutex<HashSet<PathBuf>>,
}

#[derive(Default)]
struct Doc {
    /// its Recent Files entry: the file's, or a never-saved document's once it has a draft
    key: Option<String>,
    /// the file it was opened from (todos/DESKTOP_REAL_SAVE.md: Save writes only this)
    file: Option<PathBuf>,
    /// what `host_document` describes and `host_document_bytes` hands over, once
    start: Option<Start>,
    /// the page answered its close button (`host_closing`); the window may close (`host_close_window`)
    closing: bool,
    /// the version of `file` the window read or last wrote (a Save over another one asks first)
    base: Option<Base>,
    /// the window kept the version of `file` its first save replaced (once per window)
    kept: bool,
    /// the destination the user just picked in this window's Save dialog, written once by `host_write_chosen`
    choice: Option<Choice>,
    /// the folder of the window's last Save dialog
    folder: Option<PathBuf>,
    /// the file the user just picked in this window's Open dialog to take the window's place (`host_reopen`)
    next: Option<PathBuf>,
    /// the one-time word in the address of the page the window reloads to for `next` (`?doc&t=`): only that page is
    /// given the document (`host_document`), not the one it replaces, which runs on while the reload starts
    token: Option<String>,
    /// macOS: the sandbox's access to the file, from its bookmark, while the window has it (held, never read)
    #[cfg(target_os = "macos")]
    #[allow(dead_code)]
    access: Option<bookmark::Access>,
}

struct Choice {
    path: PathBuf,
    kind: &'static Kind,
    /// a Save As (a document type then becomes the window's file), not an export
    save_as: bool,
}

enum Start {
    /// the file as it is on disk
    File,
    /// the entry's unsaved version
    Draft,
    /// bytes a page gave: a file dropped on it, or a document it made (`made`: never saved, as mail merge's letters)
    Bytes {
        bytes: Vec<u8>,
        name: String,
        made: bool,
    },
}

/// What a window's starting document is, for the page (`hooks.open`'s options in suite.js).
#[derive(Serialize, Debug, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct DocInfo {
    /// the name to open the bytes as (a draft's: its own, a CSV's draft being an .xlsx)
    file_name: String,
    /// the document's name: its file's, or what a never-saved one is called
    name: String,
    /// it has been saved as a file (its unsaved version or a dropped copy aside, Save writes there)
    saved: bool,
    /// the bytes are unsaved changes
    draft: bool,
    loss_state: Option<serde_json::Value>,
    /// the window has its file (Save writes it)
    bound: bool,
    /// the file is read-only: Save goes to Save As
    read_only: bool,
}

/// A Recent Files entry as the pages list it (`VO.recent` in suite.js): no path, the folder for display only.
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RecentItem {
    key: String,
    app: String,
    name: String,
    folder: Option<String>,
    time: u64,
    size: u64,
    saved: bool,
    preview: bool,
    draft: Option<DraftItem>,
    /// open in a window now (Document Recovery leaves it out; it can't be removed)
    open: bool,
    /// open in the window that asks
    here: bool,
}

#[derive(Serialize)]
pub struct DraftItem {
    name: String,
    time: u64,
}

impl Docs {
    pub fn new(store: Store) -> Docs {
        Docs {
            store,
            windows: Mutex::default(),
            asked: Mutex::default(),
            writing: Mutex::default(),
        }
    }

    /// The window that has the document `key`, if any.
    fn window_of(&self, key: &str) -> Option<String> {
        self.windows
            .lock()
            .unwrap()
            .iter()
            .find(|(_, d)| d.key.as_deref() == Some(key))
            .map(|(label, _)| label.clone())
    }

    fn open_keys(&self) -> HashSet<String> {
        self.windows
            .lock()
            .unwrap()
            .values()
            .filter_map(|d| d.key.clone())
            .collect()
    }

    fn key_of(&self, label: &str) -> Option<String> {
        self.windows.lock().unwrap().get(label)?.key.clone()
    }

    /// A window being made gets its document, the only one its page will be given (`host_document`): a window
    /// that has one keeps it, so no document goes to a page already running, which may have parsed something
    /// hostile and holds the keys of every Recent Files entry (Threat model). One document, one window: checked
    /// under the same lock, so two opens of a file at once make one window. Else the window that has it.
    fn adopt(&self, label: &str, doc: Doc) -> Result<(), String> {
        let mut windows = self.windows.lock().unwrap();
        if let Some(holder) = holder(&windows, &doc, label) {
            return Err(holder);
        }
        if windows.contains_key(label) {
            return Err(label.into());
        }
        windows.insert(label.into(), doc);
        Ok(())
    }

    /// File ▸ Open in an app window: the window's document gives way to the file the user picked, for the page the
    /// window reloads to, which alone has the document's `token` (`host_reopen`). Else the window that has it.
    fn replace(&self, label: &str, doc: Doc) -> Result<(), String> {
        let mut windows = self.windows.lock().unwrap();
        if let Some(holder) = holder(&windows, &doc, label) {
            return Err(holder);
        }
        windows.insert(label.into(), doc);
        Ok(())
    }

    /// A window is gone: its document with it.
    pub fn forget(&self, label: &str) {
        self.windows.lock().unwrap().remove(label);
    }

    /// The close button of an app window: whether it may close now (its page said so), else the page is asked.
    pub fn may_close(&self, label: &str) -> bool {
        self.windows
            .lock()
            .unwrap()
            .get(label)
            .is_some_and(|d| d.closing)
    }

    /// `path`'s Recent Files entry, made or brought up to date as it opens: its key and its file.
    fn record(&self, app: &str, path: &Path) -> Result<(String, PathBuf), String> {
        let (key, real) = store::key_of(path).map_err(|e| describe(path, e))?;
        let meta = std::fs::metadata(&real).map_err(|e| describe(path, e))?;
        if !meta.is_file() {
            return Err(format!("“{}” is not a file", file_name(path)));
        }
        let mut entry = self
            .store
            .get(&key)
            .unwrap_or_else(|| Entry::new(app, None, ""));
        entry.app = app.into();
        entry.path = Some(real.clone());
        entry.name = file_name(&real);
        entry.size = meta.len();
        entry.mtime = meta.modified().map(store::ms).unwrap_or(0);
        entry.time = store::now_ms();
        // while the sandbox lets the app have the file (the Open dialog, Finder), for after it quits
        #[cfg(target_os = "macos")]
        if let Some(b) = bookmark::make(&real) {
            entry.bookmark = Some(b);
        }
        if let Err(e) = self.store.put(&key, &entry) {
            eprintln!("recent files: {e}");
        }
        self.store.prune(&self.open_keys());
        Ok((key, real))
    }
}

fn file_name(path: &Path) -> String {
    path.file_name()
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_default()
}

fn describe(path: &Path, e: std::io::Error) -> String {
    format!("“{}”: {e}", file_name(path))
}

/// The app of an app window (`quire-3` → quire), none for the Start Center and the licences page.
fn app_of(window: &WebviewWindow) -> Option<&'static str> {
    // the lite edition's one window has each app's page in turn: the app of the page it was sent to
    #[cfg(feature = "lite")]
    let page = window
        .state::<crate::host::Host>()
        .page_of(window.label())?
        .path()
        .trim_start_matches('/')
        .split_once('/')?
        .0
        .to_string();
    #[cfg(not(feature = "lite"))]
    let page = window.label().rsplit_once('-')?.0;
    types::APPS.iter().map(|(a, _)| *a).find(|a| *a == page)
}

fn own_app(window: &WebviewWindow) -> Result<&'static str, String> {
    app_of(window).ok_or_else(|| format!("{} has no document", window.label()))
}

/// Tells every window that Recent Files changed (`vo-recent`).
pub fn recent_changed(app: &AppHandle) {
    for w in app.webview_windows().values() {
        let _ = w.eval("window.dispatchEvent(new Event('vo-recent'));");
    }
}

/// Brings a window forward.
fn raise(app: &AppHandle, label: &str) {
    if let Some(w) = app.get_webview_window(label) {
        let _ = w.unminimize();
        let _ = w.set_focus();
    }
}

/// The window other than `label` that has `doc`'s file or entry already.
fn holder(windows: &HashMap<String, Doc>, doc: &Doc, label: &str) -> Option<String> {
    let key = doc.key.as_ref()?;
    windows
        .iter()
        .find(|(l, d)| l.as_str() != label && d.key.as_ref() == Some(key))
        .map(|(l, _)| l.clone())
}

/// Whether the page asking may be given the window's document: any page of a window made for it, only the page
/// with the word (`t`) of a window reloaded for it.
fn takes(doc: &Doc, t: Option<&str>) -> bool {
    doc.token.as_deref().is_none_or(|token| t == Some(token))
}

/// Opens `doc` in a new window of `app` (`<app>/?doc`).
fn open_in_window(app_handle: &AppHandle, app: &str, doc: Doc) -> Result<(), String> {
    // a label no window had (host::new_label counts up)
    let label = crate::host::new_label(app_handle, app)?;
    let docs = app_handle.state::<Docs>();
    if let Err(holder) = docs.adopt(&label, doc) {
        // the same file, opened at the same moment from elsewhere
        raise(app_handle, &holder);
        return Ok(());
    }
    if let Err(e) = crate::open_window(
        app_handle,
        &label,
        WebviewUrl::App(format!("{app}/?doc").into()),
    ) {
        docs.forget(&label);
        return Err(e.to_string());
    }
    Ok(())
}

/// Opens a file the system or the user named (the command line, Finder, the Open dialog): the window that already
/// has it comes forward, else a new window opens with it. Never into a page already running: a page gets only the
/// document its window was made with (`host_document`), so one that parsed something hostile can't take in another
/// file (Threat model).
pub fn open_path(app_handle: &AppHandle, path: &Path) -> Result<(), String> {
    let app = types::app_for(path)
        .ok_or_else(|| format!("VibeOffice can't open “{}”", file_name(path)))?;
    let docs = app_handle.state::<Docs>();
    let (key, real) = docs.record(app, path)?;
    recent_changed(app_handle);
    if let Some(label) = docs.window_of(&key) {
        raise(app_handle, &label);
        return Ok(());
    }
    let doc = Doc {
        key: Some(key),
        file: Some(real),
        start: Some(Start::File),
        ..Doc::default()
    };
    open_in_window(app_handle, app, doc)
}

/// The arguments that are files, relative ones from `cwd` (the switches the installer passes aren't).
pub fn files_among<I: IntoIterator<Item = impl Into<PathBuf>>>(
    args: I,
    cwd: &Path,
) -> Vec<PathBuf> {
    args.into_iter()
        .map(|a| cwd.join(a.into()))
        .filter(|p| p.is_file())
        .map(|p| std::path::absolute(&p).unwrap_or(p))
        .collect()
}

/// The labels of the Open dialog in the UI's language, from the page (English when it gives none).
#[derive(Deserialize, Default)]
#[serde(deny_unknown_fields)]
pub struct Labels {
    title: Option<String>,
    all: Option<String>,
    quire: Option<String>,
    ledger: Option<String>,
    lectern: Option<String>,
}

fn label(given: &Option<String>, english: &str) -> String {
    match given {
        Some(s) if !s.is_empty() && s.len() <= 200 && !s.chars().any(char::is_control) => s.clone(),
        _ => english.into(),
    }
}

/// The Open dialog's filters: the window's app's types, or (the Start Center) all of them and each app's.
fn open_filters(app: Option<&str>, l: &Labels) -> Vec<dialog::Filter> {
    let kind = |a: &str| match a {
        "quire" => label(&l.quire, "Documents"),
        "ledger" => label(&l.ledger, "Spreadsheets"),
        _ => label(&l.lectern, "Presentations"),
    };
    let one = |a: &str| dialog::Filter::new(kind(a), types::opens(a));
    match app {
        Some(a) => vec![one(a)],
        None => {
            let all: Vec<&str> = types::APPS
                .iter()
                .flat_map(|(_, e)| e.iter().copied())
                .collect();
            let mut list = vec![dialog::Filter::new(
                label(&l.all, "All files VibeOffice opens"),
                &all,
            )];
            list.extend(types::APPS.iter().map(|(a, _)| one(a)));
            list
        }
    }
}

/// What `host_open` did: how many files opened in windows of their own (or came forward in the one that has them),
/// and the name of the one that is to take this app window's place (`host_reopen`).
#[derive(Serialize)]
pub struct Opened {
    opened: usize,
    here: Option<String>,
}

/// File ▸ Open and the Start Center's Open File: the system's Open dialog. In an app window the first file of its
/// app takes the window's place once its page has let its document go (`host_reopen`); every other file opens in a
/// window of its own, or comes forward in the one that has it. The Start Center opens each in its own window.
#[tauri::command]
pub async fn host_open(
    app: AppHandle,
    window: WebviewWindow,
    labels: Option<Labels>,
) -> Result<Opened, String> {
    let labels = labels.unwrap_or_default();
    let own = app_of(&window);
    let files = dialog::open(
        &window,
        &label(&labels.title, "Open"),
        open_filters(own, &labels),
    )
    .await;
    let docs = app.state::<Docs>();
    let open_elsewhere = |f: &Path| {
        store::key_of(f)
            .ok()
            .and_then(|(k, _)| docs.window_of(&k))
            .is_some()
    };
    let mut here = None;
    let mut opened = 0;
    let mut errors = Vec::new();
    for f in files {
        if here.is_none() && own.is_some() && types::app_for(&f) == own && !open_elsewhere(&f) {
            here = Some(file_name(&f));
            docs.windows
                .lock()
                .unwrap()
                .entry(window.label().into())
                .or_default()
                .next = Some(f);
            continue;
        }
        match open_path(&app, &f) {
            Ok(()) => opened += 1,
            Err(e) => errors.push(e),
        }
    }
    if errors.is_empty() {
        Ok(Opened { opened, here })
    } else {
        Err(errors.join("\n"))
    }
}

/// macOS: the menu bar's File ▸ Open… with none of the suite's windows in front (main.rs, mac_menu): the system's
/// Open dialog on its own, each file in a window of its own, or the one that has it brought forward.
#[cfg(target_os = "macos")]
pub async fn open_alone(app: &AppHandle) {
    let labels = Labels::default();
    let files =
        dialog::open_alone(&label(&labels.title, "Open"), open_filters(None, &labels)).await;
    for f in files {
        if let Err(e) = open_path(app, &f) {
            eprintln!("open: {e}");
        }
    }
}

/// macOS: the suite's window in front, the Start Center's or an app's (not the licences page's), for the menu bar.
#[cfg(target_os = "macos")]
pub fn front(app: &AppHandle) -> Option<WebviewWindow> {
    app.webview_windows()
        .into_values()
        .find(|w| w.is_focused().unwrap_or(false) && (w.label() == "main" || app_of(w).is_some()))
}

/// The file the user picked in this app window's Open dialog takes the window's place, once its page has let its
/// document go (saved or given up: the "Save changes?" question); `go`: false lets the pick go instead. The window
/// reloads to its app's page for the file, whose address has a one-time word only that page knows (`token`), so the
/// page it replaces, which may have parsed something hostile, is never given the file.
#[tauri::command]
pub async fn host_reopen(app: AppHandle, window: WebviewWindow, go: bool) -> Result<(), String> {
    let docs = app.state::<Docs>();
    let next = docs
        .windows
        .lock()
        .unwrap()
        .get_mut(window.label())
        .and_then(|d| d.next.take());
    let (Some(path), true) = (next, go) else {
        return Ok(());
    };
    let own = own_app(&window)?;
    if types::app_for(&path) != Some(own) {
        return Err(format!("{own} doesn't open “{}”", file_name(&path)));
    }
    open_here(&app, &window, own, &path)
}

/// `window` goes to `app`'s page with the file `path` (File ▸ Open, `host_reopen`; the lite edition's window,
/// lite.rs). The address has a one-time word only that page knows (`token`), so the page it replaces, which may
/// have parsed something hostile, is never given the file.
pub fn open_here(
    app_handle: &AppHandle,
    window: &WebviewWindow,
    app: &str,
    path: &Path,
) -> Result<(), String> {
    match doc_address(app_handle, window.label(), app, path)? {
        Ok(page) => crate::host::navigate(app_handle, window, &page),
        Err(holder) => {
            raise(app_handle, &holder);
            Ok(())
        }
    }
}

/// The window `label` takes the file `path` as `app`'s document: the address of the page that opens it, with the
/// one-time word (`open_here`; the lite edition's first page, lite.rs), or the window that has the file already.
pub fn doc_address(
    app_handle: &AppHandle,
    label: &str,
    app: &str,
    path: &Path,
) -> Result<Result<String, String>, String> {
    let docs = app_handle.state::<Docs>();
    let (key, real) = docs.record(app, path)?;
    recent_changed(app_handle);
    let token = store::random_key();
    let doc = Doc {
        key: Some(key),
        file: Some(real),
        start: Some(Start::File),
        token: Some(token.clone()),
        ..Doc::default()
    };
    Ok(docs
        .replace(label, doc)
        .map(|()| format!("{app}/?doc&t={token}")))
}

/// An app window's Recent Files: the Start Center, the one window that lists them, brought forward (or opened again)
/// on its Recent Files.
#[tauri::command]
pub async fn host_show_recent(app: AppHandle) -> Result<(), String> {
    if let Some(main) = app.get_webview_window("main") {
        let _ = main.eval("window.dispatchEvent(new Event('vo-show-recent'));");
    }
    crate::show_start_center(&app);
    Ok(())
}

/// A Recent Files entry (`draft`: its unsaved version): the window that has it comes forward, else a new window
/// opens with it, never the page asking (as open_path), which knows the keys of every entry (`host_recent`).
#[tauri::command]
pub async fn host_open_recent(app: AppHandle, key: String, draft: bool) -> Result<(), String> {
    let docs = app.state::<Docs>();
    let mut entry = docs
        .store
        .get(&key)
        .ok_or("That file is no longer in Recent Files.")?;
    if let Some(label) = docs.window_of(&key) {
        raise(&app, &label);
        return Ok(());
    }
    if draft && entry.draft.is_none() {
        return Err(format!("“{}” has no unsaved changes", entry.name));
    }
    // macOS: the file's bookmark gives the access the sandbox forgot when the app last quit
    #[cfg(target_os = "macos")]
    let access = match entry.bookmark.as_deref().and_then(bookmark::open) {
        Some((path, access, stale)) => {
            if stale {
                entry.bookmark = bookmark::make(&path).or(entry.bookmark);
            }
            entry.path = Some(path);
            Some(access)
        }
        None => None,
    };
    if !draft {
        match &entry.path {
            Some(p) if p.is_file() => {}
            Some(_) => {
                return Err(format!(
                    "“{}” is no longer where it was: it was moved, renamed or deleted.",
                    entry.name
                ))
            }
            None => return Err(format!("“{}” was never saved", entry.name)),
        }
    }
    entry.time = store::now_ms();
    let _ = docs.store.put(&key, &entry);
    recent_changed(&app);
    let doc = Doc {
        key: Some(key),
        file: entry.path.clone(),
        start: Some(if draft { Start::Draft } else { Start::File }),
        #[cfg(target_os = "macos")]
        access,
        ..Doc::default()
    };
    let app_name = entry.app.clone();
    open_in_window(&app, &app_name, doc)
}

/// A document as bytes in a new window: a file dropped on a page (its name in `X-Name`), or one a page made
/// (`X-Made`: never saved, as mail merge's letters). An app window opens it in its own app; the Start Center
/// in the app for its name. It has no file: its first Save asks where (Phase 2).
#[tauri::command(async)]
pub fn host_new_document(
    app: AppHandle,
    window: WebviewWindow,
    request: Request<'_>,
) -> Result<(), String> {
    let tauri::ipc::InvokeBody::Raw(bytes) = request.body() else {
        return Err("no bytes".into());
    };
    if bytes.len() as u64 > MAX_FILE {
        return Err("larger than 512 MB".into());
    }
    let header = |h: &str| {
        request
            .headers()
            .get(h)
            .and_then(|v| v.to_str().ok())
            .and_then(percent_decode)
    };
    let name = header("X-Name").unwrap_or_default();
    if name.is_empty()
        || name.len() > 255
        || name.contains(['/', '\\'])
        || name.chars().any(char::is_control)
    {
        return Err(format!("not a file name: {name}"));
    }
    let made = header("X-Made").is_some();
    let target = match app_of(&window) {
        Some(a) => a,
        None => types::app_for(Path::new(&name))
            .ok_or_else(|| format!("VibeOffice can't open “{name}”"))?,
    };
    let doc = Doc {
        start: Some(Start::Bytes {
            bytes: bytes.clone(),
            name,
            made,
        }),
        ..Doc::default()
    };
    open_in_window(&app, target, doc)
}

/// `%XX` escapes (encodeURIComponent) decoded; none for bytes that aren't UTF-8.
fn percent_decode(s: &str) -> Option<String> {
    let b = s.as_bytes();
    let mut out = Vec::with_capacity(b.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'%' {
            let hex = s.get(i + 1..i + 3)?;
            out.push(u8::from_str_radix(hex, 16).ok()?);
            i += 3;
        } else {
            out.push(b[i]);
            i += 1;
        }
    }
    String::from_utf8(out).ok()
}

/// What this window's document is (null: blank, a template, the sample).
#[tauri::command]
pub fn host_document(
    window: WebviewWindow,
    docs: State<Docs>,
    t: Option<String>,
) -> Option<DocInfo> {
    let windows = docs.windows.lock().unwrap();
    let doc = windows
        .get(window.label())
        .filter(|d| takes(d, t.as_deref()))?;
    let entry = doc.key.as_deref().and_then(|k| docs.store.get(k));
    let read_only = doc.file.as_deref().is_some_and(files::read_only);
    info(doc, entry.as_ref(), read_only)
}

fn info(doc: &Doc, entry: Option<&Entry>, read_only: bool) -> Option<DocInfo> {
    let file = doc.file.as_deref().map(file_name);
    Some(match doc.start.as_ref()? {
        Start::File => DocInfo {
            file_name: file.clone()?,
            name: file?,
            saved: true,
            draft: false,
            loss_state: None,
            bound: true,
            read_only,
        },
        Start::Draft => {
            let entry = entry?;
            let draft = entry.draft.as_ref()?;
            DocInfo {
                file_name: draft.name.clone(),
                name: entry.name.clone(),
                saved: entry.path.is_some(),
                draft: true,
                loss_state: draft.loss_state.clone(),
                bound: doc.file.is_some(),
                read_only,
            }
        }
        Start::Bytes { name, made, .. } => DocInfo {
            file_name: name.clone(),
            name: name.clone(),
            saved: !made,
            draft: false,
            loss_state: None,
            bound: false,
            read_only: false,
        },
    })
}

/// This window's starting document's bytes, once. The version of its file it holds from now on is the one read
/// (an unsaved version's: the one its entry last saw).
#[tauri::command(async)]
pub fn host_document_bytes(
    window: WebviewWindow,
    docs: State<Docs>,
    t: Option<String>,
) -> Result<Response, String> {
    let (start, key, file) = {
        let mut windows = docs.windows.lock().unwrap();
        let doc = windows
            .get_mut(window.label())
            .filter(|d| takes(d, t.as_deref()))
            .ok_or("no document")?;
        (
            doc.start.take().ok_or("no document")?,
            doc.key.clone(),
            doc.file.clone(),
        )
    };
    let (bytes, base) = match start {
        Start::File => {
            let path = file.ok_or("no file")?;
            let meta = std::fs::metadata(&path).map_err(|e| describe(&path, e))?;
            let bytes = store::read_file(&path, MAX_FILE).map_err(|e| describe(&path, e))?;
            (bytes, Some(files::base_of(&meta)))
        }
        Start::Draft => {
            let key = key.as_deref().ok_or("no entry")?;
            let bytes = docs
                .store
                .draft(key)
                .map_err(|e| format!("the unsaved changes: {e}"))?;
            let base = docs
                .store
                .get(key)
                .filter(|e| e.path.is_some())
                .map(|e| Base::of_entry(e.size, e.mtime));
            (bytes, base)
        }
        Start::Bytes { bytes, .. } => (bytes, None),
    };
    if let Some(d) = docs.windows.lock().unwrap().get_mut(window.label()) {
        d.base = base;
    }
    Ok(Response::new(bytes))
}

/// The window's document, saved as its file (a Save) or as a new one (Save As): Recent Files says so, the draft
/// goes, and the windows hear of it.
fn saved(app: &AppHandle, docs: &Docs, key: &str, path: &Path, base: Base) {
    let app_name = docs
        .store
        .get(key)
        .map(|e| e.app)
        .or_else(|| types::app_for(path).map(str::to_string))
        .unwrap_or_default();
    let mut entry = docs
        .store
        .get(key)
        .unwrap_or_else(|| Entry::new(&app_name, None, ""));
    entry.path = Some(path.to_path_buf());
    entry.name = file_name(path);
    entry.size = base.size();
    entry.mtime = base.mtime();
    entry.time = store::now_ms();
    entry.draft = None;
    if let Err(e) = docs
        .store
        .put(key, &entry)
        .and_then(|()| docs.store.drop_draft(key))
    {
        eprintln!("recent files: {e}");
    }
    recent_changed(app);
}

/// Claims `path` for one write at a time; the claim ends when dropped.
struct Writing<'a>(&'a Docs, PathBuf);

impl Drop for Writing<'_> {
    fn drop(&mut self) {
        self.0.writing.lock().unwrap().remove(&self.1);
    }
}

fn claim<'a>(docs: &'a Docs, path: &Path) -> Result<Writing<'a>, SaveError> {
    if !docs.writing.lock().unwrap().insert(path.to_path_buf()) {
        return Err(files::fail(
            "busy",
            format!("“{}” is being saved already.", file_name(path)),
        ));
    }
    Ok(Writing(docs, path.to_path_buf()))
}

fn raw_body<'a>(request: &'a Request<'_>) -> Result<&'a [u8], SaveError> {
    match request.body() {
        tauri::ipc::InvokeBody::Raw(b) => Ok(b),
        _ => Err(files::fail("io", "no bytes")),
    }
}

/// Save (Ctrl+S): the bytes into the window's own file, in place, without asking (`X-Replace`: even though another
/// program changed it, once the user said Replace It). The version it replaces is kept, once per window.
#[tauri::command(async)]
pub fn host_save(
    app: AppHandle,
    window: WebviewWindow,
    docs: State<Docs>,
    request: Request<'_>,
) -> Result<(), SaveError> {
    let bytes = raw_body(&request)?;
    let own = own_app(&window).map_err(|e| files::fail("io", e))?;
    let (path, key, base, kept) = {
        let windows = docs.windows.lock().unwrap();
        let d = windows.get(window.label());
        match d.and_then(|d| Some((d.file.clone()?, d.key.clone()?, d.base, d.kept))) {
            Some(t) => t,
            None => {
                return Err(files::fail(
                    "missing",
                    "This document has no file of its own yet.",
                ))
            }
        }
    };
    let kind = types::kind_of(own, &path).ok_or_else(|| {
        files::fail(
            "type",
            format!("{own} doesn't save “{}”.", file_name(&path)),
        )
    })?;
    let over = match (request.headers().get("X-Replace"), base) {
        (Some(_), _) | (None, None) => Over::Anything,
        (None, Some(b)) => Over::Unchanged(b),
    };
    let _writing = claim(&docs, &path)?;
    let new = files::write(&path, kind, bytes, over, |p| {
        if !kept {
            docs.store.keep_previous(&key, p);
        }
    })?;
    if let Some(d) = docs.windows.lock().unwrap().get_mut(window.label()) {
        d.base = Some(new);
        d.kept = true;
    }
    saved(&app, &docs, &key, &path, new);
    Ok(())
}

/// The Mac App Sandbox refused a save (`denied`): the system's folder chooser asks the user for the folder of the
/// file (the window's own for `save`, the Save dialog's pick for `chosen`), opening there, `text` its message in the
/// UI's language. A folder that holds the file lets the app in until it quits (nothing is kept for after); the page
/// then saves again. False when the user cancelled; an error for a folder that doesn't hold the file. Only the
/// sandbox asks (`__VO_HOST__.sandboxed`).
#[tauri::command]
pub async fn host_allow_access(
    app: AppHandle,
    window: WebviewWindow,
    how: String,
    text: Option<String>,
) -> Result<bool, SaveError> {
    let docs = app.state::<Docs>();
    let path = {
        let windows = docs.windows.lock().unwrap();
        let d = windows.get(window.label());
        match how.as_str() {
            "save" => d.and_then(|d| d.file.clone()),
            "chosen" => d.and_then(|d| Some(d.choice.as_ref()?.path.clone())),
            other => return Err(files::fail("io", format!("not a save: {other}"))),
        }
    };
    let path = path.ok_or_else(|| files::fail("missing", "There is no file to save to."))?;
    let folder = path
        .parent()
        .map(Path::to_path_buf)
        .ok_or_else(|| files::fail("io", "The file has no folder."))?;
    #[cfg(target_os = "macos")]
    {
        let text = label(&text, "Allow access to the folder to save the file there");
        let Some(picked) = dialog::folder(&window, &text, &folder).await else {
            return Ok(false);
        };
        let real = |p: &Path| std::fs::canonicalize(p).unwrap_or_else(|_| p.to_path_buf());
        if !real(&path).starts_with(real(&picked)) {
            return Err(files::fail(
                "elsewhere",
                format!(
                    "“{}” isn't in that folder. Choose the folder it is in: {}",
                    file_name(&path),
                    folder.display()
                ),
            ));
        }
        // the chooser's grant lasts until the app quits
        Ok(true)
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = (text, folder);
        Err(files::fail(
            "io",
            "Only the Mac's sandbox asks for folders.",
        ))
    }
}

/// One of the window's app's types for the Save dialog, its label in the UI's language.
#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Offer {
    id: String,
    label: String,
}

/// What `host_choose` returns: the name and type the user picked.
#[derive(Serialize)]
pub struct Chosen {
    name: String,
    #[serde(rename = "type")]
    kind: String,
}

/// Save As (`saveAs`) and every export (`export`): the system's Save dialog, in the folder of the window's file
/// (else the last one this window used, else Documents), with `name` (a bare name, no extension) and the types
/// offered, `type` first; Rust makes the filters from its own table. Remembers the choice for `host_write_chosen`;
/// none when cancelled. The lite edition has no dialog: the page asked for the name, and the file is `name` with
/// `type`'s extension in its Documents (lite.rs); one that is there is replaced only when the page asks again with
/// `replace`, after asking the user (error `exists`).
#[tauri::command]
pub async fn host_choose(
    app: AppHandle,
    window: WebviewWindow,
    purpose: String,
    types: Vec<Offer>,
    r#type: String,
    name: String,
    title: Option<String>,
    replace: Option<bool>,
) -> Result<Option<Chosen>, SaveError> {
    let bad = |m: String| files::fail("type", m);
    let own = own_app(&window).map_err(|e| files::fail("io", e))?;
    let save_as = match purpose.as_str() {
        "saveAs" => true,
        "export" => false,
        other => return Err(bad(format!("not a purpose: {other}"))),
    };
    let mut offered: Vec<(&'static Kind, String)> = Vec::new();
    for o in &types {
        let kind =
            types::kind(own, &o.id).ok_or_else(|| bad(format!("{own} doesn't save {}", o.id)))?;
        offered.push((kind, label(&Some(o.label.clone()), kind.exts[0])));
    }
    let first = offered
        .iter()
        .position(|(k, _)| k.id == r#type)
        .ok_or_else(|| bad(format!("{} isn't offered", r#type)))?;
    offered.swap(0, first);
    if !bare_name(&name) {
        return Err(bad(format!("not a file name: {name}")));
    }
    let docs = app.state::<Docs>();
    #[cfg(feature = "lite")]
    let (kind, path) = {
        let _ = title;
        let kind = offered[0].0;
        let path = crate::lite::document(&format!("{name}.{}", kind.exts[0]))
            .map_err(|e| files::fail("io", e))?;
        if path.exists() && replace != Some(true) {
            return Err(SaveError {
                file: Some(file_name(&path)),
                ..files::fail("exists", format!("“{}” already exists.", file_name(&path)))
            });
        }
        (kind, path)
    };
    #[cfg(not(feature = "lite"))]
    let (kind, path) = {
        let _ = replace;
        let (file, last) = {
            let windows = docs.windows.lock().unwrap();
            let d = windows.get(window.label());
            (
                d.and_then(|d| d.file.clone()),
                d.and_then(|d| d.folder.clone()),
            )
        };
        let folder = file
            .as_deref()
            .and_then(Path::parent)
            .map(Path::to_path_buf)
            .or(last)
            .or_else(|| app.path().document_dir().ok());
        let filters = offered
            .iter()
            .map(|(k, l)| dialog::Filter::new(l.clone(), k.exts))
            .collect();
        let title = label(&title, if save_as { "Save As" } else { "Export" });
        let suggested = format!("{name}.{}", offered[0].0.exts[0]);
        let Some((picked, filter)) =
            dialog::save(&window, &title, filters, folder, &suggested).await
        else {
            return Ok(None);
        };
        // the type: the extension typed, if it is one offered; else the filter's (GTK says which), the extension
        // added
        let typed = types::ext(&picked)
            .and_then(|e| offered.iter().find(|(k, _)| k.exts.contains(&e.as_str())));
        let (kind, path) = match typed {
            Some((k, _)) => (*k, picked.clone()),
            None => {
                let k = offered[filter.unwrap_or(0).min(offered.len() - 1)].0;
                let mut p = picked.clone().into_os_string();
                p.push(format!(".{}", k.exts[0]));
                (k, PathBuf::from(p))
            }
        };
        // the dialog asked about replacing the name it showed, not one with an extension added
        if path != picked
            && path.exists()
            && !dialog::confirm(
                &window,
                &title,
                &format!("“{}” already exists. Replace it?", file_name(&path)),
            )
            .await
        {
            return Ok(None);
        }
        (kind, path)
    };
    if let Ok((key, _)) = store::key_of(&path) {
        if docs.window_of(&key).is_some_and(|w| w != window.label()) {
            return Err(SaveError {
                file: Some(file_name(&path)),
                ..files::fail(
                    "open-elsewhere",
                    format!(
                        "“{}” is open in another VibeOffice window.",
                        file_name(&path)
                    ),
                )
            });
        }
    }
    let chosen = Chosen {
        name: file_name(&path),
        kind: kind.id.into(),
    };
    let mut windows = docs.windows.lock().unwrap();
    let d = windows.entry(window.label().into()).or_default();
    d.folder = path.parent().map(Path::to_path_buf);
    d.choice = Some(Choice {
        path,
        kind,
        save_as,
    });
    Ok(Some(chosen))
}

/// A bare file name a page suggests: no folder, drive, stream or device in it.
fn bare_name(name: &str) -> bool {
    !name.trim().is_empty()
        && name.len() <= 200
        && !name.contains(['/', '\\', ':', '*', '?', '"', '<', '>', '|'])
        && !name.chars().any(char::is_control)
        && name != ".."
        && !files::reserved(name)
}

/// What `host_write_chosen` returns: the file's name and type, and whether it is now the window's file.
#[derive(Serialize)]
pub struct Written {
    name: String,
    #[serde(rename = "type")]
    kind: String,
    bound: bool,
}

/// Writes the window's choice (`host_choose`), once. A Save As to a document type makes it the window's file: its
/// Recent Files entry is the new file's (a never-saved document's own entry gives it its preview, then goes). A copy
/// of a file from the internet keeps its Mark of the Web (Windows).
#[tauri::command(async)]
pub fn host_write_chosen(
    app: AppHandle,
    window: WebviewWindow,
    docs: State<Docs>,
    request: Request<'_>,
) -> Result<Written, SaveError> {
    let bytes = raw_body(&request)?;
    let (choice, own_file, old_key) = {
        let mut windows = docs.windows.lock().unwrap();
        let d = windows
            .get_mut(window.label())
            .ok_or_else(|| files::fail("io", "nothing was chosen"))?;
        (
            d.choice
                .take()
                .ok_or_else(|| files::fail("io", "nothing was chosen"))?,
            d.file.clone(),
            d.key.clone(),
        )
    };
    let key = store::key_of(&choice.path).ok().map(|(k, _)| k);
    let _writing = claim(&docs, &choice.path)?;
    let result = files::write(&choice.path, choice.kind, bytes, Over::Chosen, |p| {
        if let Some(k) = &key {
            docs.store.keep_previous(k, p);
        }
    });
    let new = match result {
        Ok(b) => b,
        Err(e) => {
            // kept for Try Again
            if let Some(d) = docs.windows.lock().unwrap().get_mut(window.label()) {
                d.choice.get_or_insert(choice);
            }
            return Err(e);
        }
    };
    if let Some(from) = &own_file {
        files::carry_mark(from, &choice.path);
    }
    let bound = choice.save_as && choice.kind.binds;
    if bound {
        let app_name = choice.kind.app;
        let (key, real) = docs
            .record(app_name, &choice.path)
            .map_err(|e| files::fail("io", e))?;
        if let Some(old) = old_key.filter(|o| *o != key) {
            match docs.store.get(&old) {
                // a document never saved: its entry was for its draft alone
                Some(e) if e.path.is_none() => {
                    if let Some(jpeg) = docs.store.preview(&old) {
                        let _ = docs.store.set_preview(&key, &jpeg);
                    }
                    let _ = docs.store.remove(&old);
                }
                // the file it was opened from stays in the list, without the changes now saved elsewhere
                Some(_) => {
                    let _ = docs.store.drop_draft(&old);
                }
                None => {}
            }
        }
        if let Some(d) = docs.windows.lock().unwrap().get_mut(window.label()) {
            d.key = Some(key.clone());
            d.file = Some(real.clone());
            d.base = Some(new);
            d.kept = true;
            #[cfg(target_os = "macos")]
            {
                d.access = None;
            }
        }
        saved(&app, &docs, &key, &real, new);
    }
    Ok(Written {
        name: file_name(&choice.path),
        kind: choice.kind.id.into(),
        bound,
    })
}

/// The page gave up after choosing (the compatibility check's Cancel): the choice goes.
#[tauri::command]
pub fn host_cancel_choice(window: WebviewWindow, docs: State<Docs>) {
    if let Some(d) = docs.windows.lock().unwrap().get_mut(window.label()) {
        d.choice = None;
    }
}

/// File ▸ Versions…: the versions of this window's file its saves replaced, newest first.
#[derive(Serialize)]
pub struct VersionItem {
    id: String,
    time: u64,
    size: u64,
}

#[tauri::command(async)]
pub fn host_versions(window: WebviewWindow, docs: State<Docs>) -> Result<Vec<VersionItem>, String> {
    let Some(key) = docs.key_of(window.label()) else {
        return Ok(Vec::new());
    };
    Ok(docs
        .store
        .versions(&key)
        .into_iter()
        .map(|(id, time, size)| VersionItem { id, time, size })
        .collect())
}

/// One of them in a new window, as a document never saved: the user saves it where they like.
#[tauri::command(async)]
pub fn host_open_version(
    app: AppHandle,
    window: WebviewWindow,
    docs: State<Docs>,
    id: String,
) -> Result<(), String> {
    let own = own_app(&window)?;
    let key = docs.key_of(window.label()).ok_or("no versions")?;
    if id.split('.').next() != Some(key.as_str()) {
        return Err("not a version of this document".into());
    }
    let bytes = docs.store.version(&id).map_err(|e| e.to_string())?;
    let name = docs
        .store
        .get(&key)
        .map(|e| e.name)
        .unwrap_or_else(|| "Document".into());
    let doc = Doc {
        start: Some(Start::Bytes {
            bytes,
            name,
            made: true,
        }),
        ..Doc::default()
    };
    open_in_window(&app, own, doc)
}

/// Recent Files, newest first.
#[tauri::command(async)]
pub fn host_recent(window: WebviewWindow, docs: State<Docs>) -> Result<Vec<RecentItem>, String> {
    let open = docs.open_keys();
    let mine = docs.key_of(window.label());
    Ok(docs
        .store
        .list()
        .into_iter()
        .map(|(key, e, preview)| RecentItem {
            open: open.contains(&key),
            here: mine.as_deref() == Some(key.as_str()),
            key,
            app: e.app,
            name: e.name,
            folder: e
                .path
                .as_deref()
                .and_then(Path::parent)
                .map(|p| p.to_string_lossy().into_owned()),
            time: e.time,
            size: e.size,
            saved: e.path.is_some(),
            preview,
            draft: e.draft.map(|d| DraftItem {
                name: d.name,
                time: d.time,
            }),
        })
        .collect())
}

/// A Recent Files entry's preview (a JPEG).
#[tauri::command(async)]
pub fn host_recent_preview(docs: State<Docs>, key: String) -> Result<Response, String> {
    docs.store
        .preview(&key)
        .map(Response::new)
        .ok_or_else(|| "no preview".into())
}

fn not_open(docs: &Docs, key: &str) -> Result<(), String> {
    match docs.window_of(key) {
        Some(_) => Err("That document is open in a window.".into()),
        None => Ok(()),
    }
}

/// Removes an entry (the Start Center's ×), not of a document open in a window.
#[tauri::command(async)]
pub fn host_recent_remove(app: AppHandle, docs: State<Docs>, key: String) -> Result<(), String> {
    not_open(&docs, &key)?;
    docs.store.remove(&key).map_err(|e| e.to_string())?;
    recent_changed(&app);
    Ok(())
}

/// Clear Recent Documents: every entry without unsaved changes, but those open in a window.
#[tauri::command(async)]
pub fn host_recent_clear(app: AppHandle, docs: State<Docs>) -> Result<(), String> {
    docs.store.clear(&docs.open_keys());
    recent_changed(&app);
    Ok(())
}

/// Document Recovery's Discard: an entry's unsaved changes, not of a document open in a window.
#[tauri::command(async)]
pub fn host_discard_draft(app: AppHandle, docs: State<Docs>, key: String) -> Result<(), String> {
    not_open(&docs, &key)?;
    docs.store.drop_draft(&key).map_err(|e| e.to_string())?;
    recent_changed(&app);
    Ok(())
}

/// The preview of this window's document (a JPEG), for its Recent Files entry; nothing while it has none.
#[tauri::command(async)]
pub fn host_set_preview(
    app: AppHandle,
    window: WebviewWindow,
    docs: State<Docs>,
    request: Request<'_>,
) -> Result<(), String> {
    let tauri::ipc::InvokeBody::Raw(jpeg) = request.body() else {
        return Err("no bytes".into());
    };
    let Some(key) = docs.key_of(window.label()) else {
        return Ok(());
    };
    docs.store
        .set_preview(&key, jpeg)
        .map_err(|e| e.to_string())?;
    recent_changed(&app);
    Ok(())
}

/// What a draft says about itself, before its bytes in the body (`VO.host.setDraft`).
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct DraftInfo {
    /// the document's name
    name: String,
    /// the name its unsaved version opens as, when another
    draft_name: Option<String>,
    loss_state: Option<serde_json::Value>,
}

/// This window's unsaved changes, in the app's own format: kept in its Recent Files entry (a never-saved
/// document gets an entry of its own). The body: the length of a JSON `DraftInfo` (4 bytes, little-endian), it,
/// then the bytes.
#[tauri::command(async)]
pub fn host_set_draft(
    app: AppHandle,
    window: WebviewWindow,
    docs: State<Docs>,
    request: Request<'_>,
) -> Result<(), String> {
    let own = own_app(&window)?;
    let tauri::ipc::InvokeBody::Raw(body) = request.body() else {
        return Err("no bytes".into());
    };
    let (info, bytes) = split_draft(body)?;
    let key = {
        let mut windows = docs.windows.lock().unwrap();
        let doc = windows.entry(window.label().into()).or_default();
        doc.key.get_or_insert_with(store::random_key).clone()
    };
    let mut entry = docs
        .store
        .get(&key)
        .unwrap_or_else(|| Entry::new(own, None, &info.name));
    if entry.path.is_none() {
        entry.name = info.name.clone();
    }
    entry.time = store::now_ms();
    entry.draft = Some(Draft {
        name: info.draft_name.unwrap_or(info.name),
        time: entry.time,
        loss_state: info.loss_state,
    });
    docs.store
        .set_draft(&key, &entry, bytes)
        .map_err(|e| e.to_string())?;
    recent_changed(&app);
    Ok(())
}

fn split_draft(body: &[u8]) -> Result<(DraftInfo, &[u8]), String> {
    let n = body
        .get(..4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]) as usize)
        .ok_or("no draft")?;
    let json = body.get(4..4 + n).ok_or("no draft")?;
    let info: DraftInfo = serde_json::from_slice(json).map_err(|e| format!("draft: {e}"))?;
    let name_ok = |s: &str| {
        !s.is_empty()
            && s.len() <= 255
            && !s.contains(['/', '\\'])
            && !s.chars().any(char::is_control)
    };
    if !name_ok(&info.name) || !info.draft_name.as_deref().is_none_or(name_ok) {
        return Err("draft: not a document name".into());
    }
    Ok((info, &body[4 + n..]))
}

/// This window's unsaved changes are gone (saved, or the user chose not to keep them).
#[tauri::command(async)]
pub fn host_clear_draft(
    app: AppHandle,
    window: WebviewWindow,
    docs: State<Docs>,
) -> Result<(), String> {
    let Some(key) = docs.key_of(window.label()) else {
        return Ok(());
    };
    docs.store.drop_draft(&key).map_err(|e| e.to_string())?;
    if docs.store.get(&key).is_none() {
        if let Some(d) = docs.windows.lock().unwrap().get_mut(window.label()) {
            d.key = None;
        }
    }
    recent_changed(&app);
    Ok(())
}

/// The Window menu: this app's windows, by their titles.
#[derive(Serialize)]
pub struct WindowItem {
    id: String,
    title: String,
    current: bool,
}

#[tauri::command]
pub fn host_windows(app: AppHandle, window: WebviewWindow) -> Result<Vec<WindowItem>, String> {
    let own = own_app(&window)?;
    let mut list: Vec<(usize, WindowItem)> = app
        .webview_windows()
        .into_iter()
        .filter_map(|(label, w)| {
            let (page, n) = label.rsplit_once('-')?;
            let n: usize = n.parse().unwrap_or(0);
            (page == own).then(|| {
                (
                    n,
                    WindowItem {
                        current: label == window.label(),
                        title: w.title().unwrap_or_default(),
                        id: label,
                    },
                )
            })
        })
        .collect();
    list.sort_by_key(|(n, _)| *n);
    Ok(list.into_iter().map(|(_, w)| w).collect())
}

/// The Window menu: brings another window of this app forward.
#[tauri::command]
pub fn host_focus(app: AppHandle, window: WebviewWindow, id: String) -> Result<(), String> {
    let own = own_app(&window)?;
    if id.rsplit_once('-').map(|(p, _)| p) != Some(own) {
        return Err(format!("not a window of {own}: {id}"));
    }
    raise(&app, &id);
    Ok(())
}

/// The page took its close button (`vo-close`) in hand: no "not answering" question then.
#[tauri::command]
pub fn host_closing(window: WebviewWindow, docs: State<Docs>) {
    docs.asked.lock().unwrap().remove(window.label());
}

/// The close button of an app window: the page is asked (`vo-close`) and runs its "Save changes?" question,
/// then closes the window itself (`host_close_window`). A page that doesn't take it in hand within 10 s (it hangs)
/// gets a question from here: close anyway? Its unsaved changes stay in Recent Files (the draft).
pub fn close_requested(window: &tauri::Window) {
    let app = window.app_handle();
    let Some(w) = app.get_webview_window(window.label()) else {
        return;
    };
    app.state::<Docs>()
        .asked
        .lock()
        .unwrap()
        .insert(w.label().into());
    let _ = w.eval("window.dispatchEvent(new Event('vo-close'));");
    std::thread::spawn(move || {
        std::thread::sleep(std::time::Duration::from_secs(10));
        if !w.state::<Docs>().asked.lock().unwrap().contains(w.label()) {
            return;
        }
        let close = tauri::async_runtime::block_on(dialog::confirm(
            &w,
            "VibeOffice",
            "This window isn't answering. Close it anyway? Its unsaved changes are kept in Recent Files.",
        ));
        if close {
            allow_close(&w);
        }
    });
}

/// Lets an app window close (File ▸ Exit, the page's answer to `vo-close`), and closes it.
pub fn allow_close(window: &WebviewWindow) {
    let docs = window.state::<Docs>();
    docs.windows
        .lock()
        .unwrap()
        .entry(window.label().into())
        .or_default()
        .closing = true;
    docs.asked.lock().unwrap().remove(window.label());
    let _ = window.close();
}

/// macOS: the sandbox forgets a file's grant when the app quits, so a Recent Files entry keeps a security-scoped
/// bookmark of its file, made while the grant holds (the Open dialog, Finder); opening the entry later resolves it and
/// holds the access while a window has the file.
#[cfg(target_os = "macos")]
mod bookmark {
    use objc2::{rc::Retained, runtime::Bool};
    use objc2_foundation::{
        NSData, NSString, NSURLBookmarkCreationOptions, NSURLBookmarkResolutionOptions, NSURL,
    };
    use std::path::{Path, PathBuf};

    pub fn make(path: &Path) -> Option<String> {
        let url = NSURL::fileURLWithPath(&NSString::from_str(&path.to_string_lossy()));
        let data = url
            .bookmarkDataWithOptions_includingResourceValuesForKeys_relativeToURL_error(
                NSURLBookmarkCreationOptions::WithSecurityScope,
                None,
                None,
            )
            .ok()?;
        Some(crate::store::hex(&data.to_vec()))
    }

    /// The access to a file a bookmark gave, held until dropped.
    pub struct Access(Retained<NSURL>);
    // an NSURL doesn't change: any thread may hold it
    unsafe impl Send for Access {}
    unsafe impl Sync for Access {}
    impl Drop for Access {
        fn drop(&mut self) {
            unsafe { self.0.stopAccessingSecurityScopedResource() }
        }
    }

    /// Where a bookmark's file is now, with access to it, and whether the bookmark should be made again.
    pub fn open(hex: &str) -> Option<(PathBuf, Access, bool)> {
        let data = NSData::with_bytes(&crate::store::unhex(hex)?);
        let mut stale = Bool::NO;
        let url = unsafe {
            NSURL::URLByResolvingBookmarkData_options_relativeToURL_bookmarkDataIsStale_error(
                &data,
                NSURLBookmarkResolutionOptions::WithSecurityScope,
                None,
                &mut stale,
            )
        }
        .ok()?;
        if !unsafe { url.startAccessingSecurityScopedResource() } {
            return None;
        }
        let path = PathBuf::from(url.path()?.to_string());
        Some((path, Access(url), stale.as_bool()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_window_keeps_the_document_it_was_made_with() {
        let root = std::env::temp_dir().join(format!("vo-adopt-{}", std::process::id()));
        let docs = Docs::new(Store::new(root));
        let bytes = |name: &str| Doc {
            start: Some(Start::Bytes {
                bytes: b"dropped".to_vec(),
                name: name.into(),
                made: false,
            }),
            ..Doc::default()
        };
        docs.adopt("quire-1", bytes("Dropped.docx")).unwrap();
        let file = |key: &str| Doc {
            key: Some(key.into()),
            file: Some(PathBuf::from(format!("/d/{key}.docx"))),
            start: Some(Start::File),
            ..Doc::default()
        };
        assert!(
            docs.adopt("quire-1", file("Secret")).is_err(),
            "a window with dropped bytes and no file takes no other file"
        );
        {
            let w = docs.windows.lock().unwrap();
            assert!(
                w["quire-1"].file.is_none()
                    && matches!(w["quire-1"].start, Some(Start::Bytes { .. }))
            );
        }
        docs.adopt("quire-2", file("Report")).unwrap();
        assert_eq!(
            docs.adopt("quire-3", file("Report")),
            Err("quire-2".into()),
            "two opens at once: one window, the other brought forward"
        );
        assert!(!docs.windows.lock().unwrap().contains_key("quire-3"));
    }

    #[test]
    fn a_window_reloaded_for_a_file_gives_it_only_to_the_page_with_its_word() {
        let reloaded = Doc {
            token: Some("w1".into()),
            start: Some(Start::File),
            ..Doc::default()
        };
        assert!(
            !takes(&reloaded, None),
            "the page it replaces doesn't know the word"
        );
        assert!(!takes(&reloaded, Some("w2")));
        assert!(takes(&reloaded, Some("w1")));
        assert!(
            takes(&Doc::default(), None),
            "a window made for its document: its one page"
        );
        let docs = Docs::new(Store::new(std::env::temp_dir().join("vo-replace")));
        let file = |key: &str| Doc {
            key: Some(key.into()),
            file: Some(PathBuf::from(format!("/d/{key}.docx"))),
            ..Doc::default()
        };
        docs.adopt("quire-1", file("Report")).unwrap();
        assert_eq!(
            docs.replace("quire-2", file("Report")),
            Err("quire-1".into()),
            "open in another window"
        );
        docs.replace("quire-2", file("Notes")).unwrap();
        assert_eq!(docs.key_of("quire-2").as_deref(), Some("Notes"));
    }

    #[test]
    fn a_window_tells_its_page_what_it_holds() {
        let file = |start| Doc {
            key: Some("k".into()),
            file: Some(PathBuf::from("/d/Report.docx")),
            start: Some(start),
            ..Doc::default()
        };
        let i = info(&file(Start::File), None, false).unwrap();
        assert_eq!(
            (i.name.as_str(), i.saved, i.draft, i.bound),
            ("Report.docx", true, false, true)
        );
        let mut entry = Entry::new("ledger", Some(PathBuf::from("/d/Budget.csv")), "Budget.csv");
        entry.draft = Some(Draft {
            name: "Budget.xlsx".into(),
            time: 1,
            loss_state: Some(serde_json::json!({ "entries": [] })),
        });
        let i = info(&file(Start::Draft), Some(&entry), false).unwrap();
        assert_eq!(
            (i.file_name.as_str(), i.name.as_str(), i.saved, i.draft),
            ("Budget.xlsx", "Budget.csv", true, true)
        );
        assert!(i.loss_state.is_some());
        let made = Doc {
            start: Some(Start::Bytes {
                bytes: vec![1],
                name: "Letters1.docx".into(),
                made: true,
            }),
            ..Doc::default()
        };
        let i = info(&made, None, false).unwrap();
        assert_eq!((i.saved, i.bound), (false, false));
        assert!(info(&Doc::default(), None, false).is_none(), "blank");
    }

    #[test]
    fn a_draft_carries_its_names_and_refuses_paths() {
        let body = |json: &str, bytes: &[u8]| {
            let mut b = (json.len() as u32).to_le_bytes().to_vec();
            b.extend_from_slice(json.as_bytes());
            b.extend_from_slice(bytes);
            b
        };
        let b = body(
            r#"{"name":"Budget.csv","draftName":"Budget.xlsx","lossState":{"entries":[]}}"#,
            b"PK",
        );
        let (info, bytes) = split_draft(&b).unwrap();
        assert_eq!(
            (info.name.as_str(), info.draft_name.as_deref(), bytes),
            ("Budget.csv", Some("Budget.xlsx"), &b"PK"[..])
        );
        for bad in [
            r#"{"name":"../x.docx"}"#,
            r#"{"name":""}"#,
            r#"{"name":"a","draftName":"C:\\x"}"#,
            r#"{"name":"a","path":"/x"}"#,
        ] {
            assert!(split_draft(&body(bad, b"")).is_err(), "{bad}");
        }
        assert!(split_draft(&[200, 0, 0, 0, b'{']).is_err());
        assert_eq!(
            percent_decode("R%C3%A9sum%C3%A9%20(1).docx").as_deref(),
            Some("Résumé (1).docx")
        );
        assert_eq!(percent_decode("%FF"), None);
    }

    #[test]
    fn the_open_dialog_offers_the_apps_types() {
        let l = Labels::default();
        let start = open_filters(None, &l);
        assert_eq!(start.len(), 4);
        assert!(
            start[0].exts.contains(&"docx".to_string())
                && start[0].exts.contains(&"pptx".to_string())
        );
        let quire = open_filters(
            Some("quire"),
            &Labels {
                quire: Some("Word documents".into()),
                ..Labels::default()
            },
        );
        assert_eq!(quire.len(), 1);
        assert_eq!(quire[0].name, "Word documents");
        assert!(!quire[0].exts.contains(&"xlsx".to_string()));
        assert_eq!(label(&Some("a\nb".into()), "Open"), "Open");
    }

    #[test]
    fn the_files_among_the_arguments_relative_to_the_starts_folder() {
        let dir = std::env::temp_dir().join(format!("vo-args-{}", std::process::id()));
        std::fs::create_dir_all(dir.join("sub")).unwrap();
        std::fs::write(dir.join("sub").join("Budget.csv"), "a,b").unwrap();
        let absolute = dir.join("sub").join("Budget.csv");
        assert_eq!(
            files_among(
                ["-choose-defaults", "sub/Budget.csv", "missing.docx", "sub"],
                &dir
            ),
            std::slice::from_ref(&absolute)
        );
        assert!(files_among(["-pin-taskbar"], &std::env::temp_dir()).is_empty());
        assert_eq!(
            files_among([absolute.clone()], &std::env::temp_dir()),
            [absolute]
        );
        std::fs::remove_dir_all(&dir).unwrap();
    }
}
