// No console window in release builds.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

#[cfg(all(feature = "lite", target_os = "linux"))]
mod cardputer;
mod dialog;
mod doc;
mod files;
mod host;
#[cfg(any(windows, test))]
mod installation;
#[cfg(windows)]
mod installed;
#[cfg(target_os = "linux")]
mod instance;
#[cfg(feature = "lite")]
mod lite;
mod net;
mod settings;
mod store;
mod types;

use tauri::{
    webview::{NewWindowResponse, PageLoadEvent, PermissionKind, PermissionResponse},
    AppHandle, Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder,
};

/// A window of the suite: the Start Center ("main") or an app (host_new_window). Every window gets the
/// host script, title, link, drop and UI size handling, and the network rules (net.rs): on Windows and macOS
/// it starts on about:blank and goes to its page once they are on (`net::guard`).
pub fn open_window(app: &AppHandle, label: &str, url: WebviewUrl) -> tauri::Result<WebviewWindow> {
    let page = net::page_url(app, url.clone());
    // the one page this window loads (host::Page); the Start Center's first, when probing, the probe's
    #[cfg(any(windows, target_os = "macos"))]
    let probe = label == "main" && net::probe::wanted();
    #[cfg(not(any(windows, target_os = "macos")))]
    let probe = false;
    if let Some(page) = &page {
        app.state::<host::Host>()
            .expect(label, host::Page::new(page.clone(), probe));
    }
    #[cfg(any(windows, target_os = "macos"))]
    let url = WebviewUrl::External(net::BLANK.parse().expect("about:blank"));
    let builder = WebviewWindowBuilder::new(app, label, url);
    // the pages from the dev server: the settings as they are now (otherwise served with common/look.js)
    #[cfg(not(feature = "custom-protocol"))]
    let builder =
        builder.initialization_script(settings::script(&app.state::<settings::Settings>()));
    // the lite edition's one window: the device's whole screen, as the compositor gives it (cardputer/), with no
    // title bar; elsewhere a window of its size
    #[cfg(feature = "lite")]
    let builder = builder
        .inner_size(320.0, 170.0)
        .decorations(false)
        .data_directory(lite::web_data());
    #[cfg(not(feature = "lite"))]
    let builder = builder
        .inner_size(1280.0, 820.0)
        .min_inner_size(640.0, 480.0);
    let window = builder
        .title("VibeOffice")
        .initialization_script(host::init_script())
        .initialization_script_for_all_frames(net::NO_WEBRTC)
        .initialization_script_for_all_frames(host::NO_FILE_PICKERS)
        .initialization_script_for_all_frames(host::NO_STORAGE)
        // nothing the web view would keep outlives the app (WebView2 InPrivate, WKWebView's non-persistent data
        // store). Not on Linux: Tauri 2.12 registers tauri:// once per shared web context, while an incognito
        // WebKitGTK view gets a context of its own, so from the second window on no page would load there.
        // The lite edition has one window, so it can be (its whole profile would be on disk otherwise).
        .incognito(cfg!(any(windows, target_os = "macos", feature = "lite")))
        // Windows: https://tauri.localhost, a secure context, where WebView2 doesn't look up the names of links
        // on the page ahead (DNS prefetch); macOS and Linux serve tauri://localhost either way
        .use_https_scheme(true)
        // No camera, microphone, location, notifications, …: only Local Font Access (Settings ▸ Fonts on
        // this computer) and reading the clipboard (Paste) go on as the engine handles them.
        .on_permission_request(|_, kind| match kind {
            PermissionKind::LocalFonts | PermissionKind::ClipboardRead => {
                PermissionResponse::Default
            }
            _ => PermissionResponse::Deny,
        })
        // The pages take dropped files themselves (Start Center, each app).
        .disable_drag_drop_handler()
        // The window loads its one page (host::Page). Links out of the app open in the default browser if
        // they may (host::leave: the Start Center's Update and GitHub), never in this window or a bare popup.
        // The suite's own windows come from VO.host.newWindow (host_new_window), with the network rules.
        .on_navigation({
            let (app, label) = (app.clone(), label.to_string());
            move |url| {
                if !host::is_external(url) {
                    return app.state::<host::Host>().may_load(&label, url);
                }
                host::leave(&label, url);
                false
            }
        })
        .on_new_window({
            let label = label.to_string();
            move |url, _| {
                if host::is_external(&url) {
                    host::leave(&label, &url);
                } else {
                    eprintln!("not opened: a bare window for {url}");
                }
                NewWindowResponse::Deny
            }
        })
        // Pages save through the host (doc.rs): a web view download would write where no one asked
        .on_download(|_, _| false)
        // The system title bar shows the page's title, as in the installed PWA (not about:blank's, nor
        // the network probe's).
        .on_document_title_changed(|window, title| {
            #[cfg(any(windows, target_os = "macos"))]
            if title == net::BLANK || net::probe::reported(&window, &title) {
                return;
            }
            let _ = window.set_title(&title);
        })
        // The UI size (Settings), as each page starts: WebKitGTK resets the zoom on the first load. The lite
        // edition's start page is made for its screen: always 100 %.
        .on_page_load(|window, payload| {
            if payload.event() == PageLoadEvent::Started {
                let scale = if cfg!(feature = "lite") && payload.url().path() == "/lite.html" {
                    1.0
                } else {
                    window.state::<settings::Settings>().ui_scale()
                };
                let _ = window.set_zoom(scale);
            }
            // The first start after installing: what the user chose in the installer
            #[cfg(windows)]
            if payload.event() == PageLoadEvent::Finished
                && window.label() == "main"
                && matches!(payload.url().scheme(), "http" | "https")
            {
                installed::after_install(&window);
            }
        })
        .build()?;
    #[cfg(feature = "lite")]
    lite::web_view(&window);
    #[cfg(any(windows, target_os = "macos"))]
    net::guard(&window, page)?;
    Ok(window)
}

/// The Start Center's page: index.html, the lite edition's own lite.html.
fn start_page() -> WebviewUrl {
    #[cfg(feature = "lite")]
    return WebviewUrl::App(lite::START.into());
    #[cfg(not(feature = "lite"))]
    WebviewUrl::default()
}

/// Windows can be made (setup has run; on macOS WebKit has compiled the network rules). Until then, files that
/// arrive (Finder, at start) wait in PENDING.
static STARTED: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);
#[cfg(target_os = "macos")]
static PENDING: std::sync::Mutex<Vec<std::path::PathBuf>> = std::sync::Mutex::new(Vec::new());

/// Opens files the system handed over, each in a window of its own (doc::open_path); how many opened.
fn open_files(app: &AppHandle, files: &[std::path::PathBuf]) -> usize {
    files
        .iter()
        .filter(|f| match doc::open_path(app, f) {
            Ok(_) => true,
            Err(e) => {
                eprintln!("{e}");
                false
            }
        })
        .count()
}

/// The Start Center, brought forward, or opened again if it was closed.
fn show_start_center(app: &AppHandle) {
    match app.get_webview_window("main") {
        Some(main) => {
            let _ = main.unminimize();
            let _ = main.set_focus();
            #[cfg(windows)]
            installed::after_install(&main);
        }
        // (on Windows it does what the installer asked as it loads)
        None => {
            if let Err(e) = open_window(app, "main", start_page()) {
                eprintln!("the Start Center: {e}");
            }
        }
    }
}

/// Another start of VibeOffice while this one runs (Windows: `second_start`; Linux: instance.rs): its files
/// open here, each in a window of its own; without any, the Start Center comes forward.
#[cfg(any(windows, target_os = "linux"))]
fn another_start(app: &AppHandle, args: Vec<std::ffi::OsString>, cwd: std::path::PathBuf) {
    // the other process waits for its handler: windows are made in a task, after that has returned
    let app = app.clone();
    tauri::async_runtime::spawn(async move {
        let files = if cfg!(feature = "lite") {
            Vec::new()
        } else {
            doc::files_among(args, &cwd)
        };
        if open_files(&app, &files) == 0 {
            show_start_center(&app);
        }
    });
}

/// Another start on Windows (tauri-plugin-single-instance). What the installer asked for (pin, Default apps) is
/// done here.
#[cfg(windows)]
fn second_start(app: &AppHandle, args: Vec<String>, cwd: String) {
    let args = args.get(1..).unwrap_or_default().to_vec();
    installed::take_args(&args);
    another_start(app, args.into_iter().map(Into::into).collect(), cwd.into());
}

/// An app window's close button: its page answers (doc::close_requested), not the window. The Start Center and
/// the licences page close at once. A window gone takes its document with it.
fn window_event(window: &tauri::Window, event: &tauri::WindowEvent) {
    let page = window.label().rsplit_once('-').map(|(p, _)| p);
    match event {
        tauri::WindowEvent::CloseRequested { api, .. }
            if page.is_some_and(types::is_app)
                && !window.state::<doc::Docs>().may_close(window.label()) =>
        {
            api.prevent_close();
            doc::close_requested(window);
        }
        tauri::WindowEvent::Destroyed => window.state::<doc::Docs>().forget(window.label()),
        _ => {}
    }
}

/// The menu bar on macOS: Tauri's default one, without Edit ▸ Undo and Redo, which would run WebKit's own
/// undo rather than the document's (⌘Z and ⌘Y reach the page, whose Edit menu has them). Cut, Copy, Paste
/// and Select All stay: WKWebView gets ⌘X, ⌘C, ⌘V and ⌘A through them.
#[cfg(target_os = "macos")]
fn mac_menu(app: &AppHandle) -> tauri::Result<tauri::menu::Menu<tauri::Wry>> {
    use tauri::menu::{
        AboutMetadata, Menu, MenuItem, PredefinedMenuItem as P, Submenu, HELP_SUBMENU_ID,
        WINDOW_SUBMENU_ID,
    };
    // no shortcuts: ⌘N and ⌘O stay the page's, where they make a document of its own app and open into its window
    let item = |id: &str, text: &str| MenuItem::with_id(app, id, text, true, None::<&str>);
    let info = app.package_info();
    let about = AboutMetadata {
        name: Some(info.name.clone()),
        version: Some(info.version.to_string()),
        ..Default::default()
    };
    Menu::with_items(
        app,
        &[
            &Submenu::with_items(
                app,
                info.name.clone(),
                true,
                &[
                    &P::about(app, None, Some(about))?,
                    &P::separator(app)?,
                    &P::services(app, None)?,
                    &P::separator(app)?,
                    &P::hide(app, None)?,
                    &P::hide_others(app, None)?,
                    &P::show_all(app, None)?,
                    &P::separator(app)?,
                    &P::quit(app, None)?,
                ],
            )?,
            &Submenu::with_items(
                app,
                "File",
                true,
                &[
                    &item("vo-new-quire", "New Document")?,
                    &item("vo-new-ledger", "New Spreadsheet")?,
                    &item("vo-new-lectern", "New Presentation")?,
                    &P::separator(app)?,
                    &item("vo-open", "Open…")?,
                    &P::separator(app)?,
                    &P::close_window(app, None)?,
                ],
            )?,
            &Submenu::with_items(
                app,
                "Edit",
                true,
                &[
                    &P::cut(app, None)?,
                    &P::copy(app, None)?,
                    &P::paste(app, None)?,
                    &P::select_all(app, None)?,
                ],
            )?,
            &Submenu::with_items(app, "View", true, &[&P::fullscreen(app, None)?])?,
            &Submenu::with_id_and_items(
                app,
                WINDOW_SUBMENU_ID,
                "Window",
                true,
                &[
                    &P::minimize(app, None)?,
                    &P::maximize(app, None)?,
                    &P::separator(app)?,
                    &item("vo-start", "Start Center")?,
                    &P::separator(app)?,
                    &P::close_window(app, None)?,
                ],
            )?,
            &Submenu::with_id_and_items(
                app,
                HELP_SUBMENU_ID,
                "Help",
                true,
                &[&item("vo-help", "VibeOffice Help")?],
            )?,
        ],
    )
}

/// The Mac's own items in the menu bar (mac_menu), which App Review tries with every window closed: a new document of
/// each app, Open… (the window in front's own, else the system's dialog on its own), the Start Center again, and
/// Help (the app in front's Help pane, else the project's page in the browser). The pages hear `vo-menu`.
#[cfg(target_os = "macos")]
fn mac_menu_event(app: &AppHandle, id: &str) {
    const HELP: &str = "https://github.com/tcchen2026/VibeOffice#readme";
    let (app, id) = (app.clone(), id.to_string());
    let tell = |w: &tauri::WebviewWindow, what: &str| {
        let _ = w.eval(format!(
            "window.dispatchEvent(new CustomEvent('vo-menu', {{ detail: '{what}' }}));"
        ));
    };
    // windows are made in a task, after the menu's handler has returned (as another_start)
    tauri::async_runtime::spawn(async move {
        match id.as_str() {
            "vo-new-quire" | "vo-new-ledger" | "vo-new-lectern" => {
                if let Err(e) = host::open_blank(&app, &id["vo-new-".len()..]) {
                    eprintln!("new window: {e}");
                }
            }
            "vo-open" => match doc::front(&app) {
                Some(w) => tell(&w, "open"),
                None => doc::open_alone(&app).await,
            },
            "vo-start" => show_start_center(&app),
            "vo-help" => match doc::front(&app).filter(|w| w.label() != "main") {
                Some(w) => tell(&w, "help"),
                None => {
                    let _ = host::open_external(&HELP.parse().expect("a URL"));
                }
            },
            _ => {}
        }
    });
}

fn main() {
    // The lite edition (docs/desktop.md, Lite edition), for a device that swaps: JavaScriptCore without its JIT,
    // about 30 MiB less for the pages at the same speed; GTK without OpenGL, so WebKit draws on the CPU (its web
    // process half the size), and hands its frames over in shared memory, not GPU buffers (the app's own process
    // 11 MiB smaller). Set beforehand, each stays as set (measuring).
    #[cfg(feature = "lite")]
    for (var, value) in [
        ("JSC_useJIT", "false"),
        ("GDK_GL", "disable"),
        ("WEBKIT_DISABLE_DMABUF_RENDERER", "1"),
    ] {
        if std::env::var_os(var).is_none() {
            std::env::set_var(var, value);
        }
    }
    // `-uninstall` ends here
    #[cfg(windows)]
    if !installed::start() {
        return;
    }
    let mut context = tauri::generate_context!();
    settings::serve(&mut context);
    let builder = tauri::Builder::default();
    // One VibeOffice per user session on Windows: a second start (the Start menu, a double-click on a
    // file) hands its arguments to this one and exits. The plugin goes first, before anything is made.
    #[cfg(windows)]
    let builder = builder.plugin(tauri_plugin_single_instance::init(second_start));
    // and on Linux (instance.rs): a later start hands its files over here and exits
    #[cfg(target_os = "linux")]
    let listener = {
        let cwd = std::env::current_dir().unwrap_or_default();
        let args: Vec<_> = std::env::args_os().skip(1).collect();
        match instance::folder(std::env::home_dir()) {
            Some(dir) => match instance::claim(&dir, &cwd, &args) {
                instance::Claim::Sent => return,
                instance::Claim::First(listener) => listener,
            },
            None => None,
        }
    };
    #[cfg(target_os = "macos")]
    let builder = builder
        .menu(mac_menu)
        .on_menu_event(|app, event| mac_menu_event(app, event.id().as_ref()));
    let app = builder
        .manage(host::Host::default())
        .invoke_handler(tauri::generate_handler![
            host::host_open_url,
            settings::host_set_setting,
            host::host_new_window,
            host::host_close_window,
            doc::host_document,
            doc::host_document_bytes,
            doc::host_open,
            doc::host_open_recent,
            doc::host_reopen,
            doc::host_show_recent,
            doc::host_new_document,
            doc::host_recent,
            doc::host_recent_preview,
            doc::host_recent_remove,
            doc::host_recent_clear,
            doc::host_discard_draft,
            doc::host_set_preview,
            doc::host_set_draft,
            doc::host_clear_draft,
            doc::host_windows,
            doc::host_focus,
            doc::host_closing,
            doc::host_save,
            doc::host_allow_access,
            doc::host_choose,
            doc::host_write_chosen,
            doc::host_cancel_choice,
            doc::host_versions,
            doc::host_open_version,
            #[cfg(feature = "lite")]
            lite::host_lite_list,
            #[cfg(feature = "lite")]
            lite::host_lite_open,
            #[cfg(feature = "lite")]
            lite::host_lite_new,
            #[cfg(feature = "lite")]
            lite::host_lite_home,
            #[cfg(feature = "lite")]
            lite::host_lite_licenses,
            #[cfg(feature = "lite")]
            lite::host_lite_quit,
            #[cfg(feature = "lite")]
            lite::host_lite_show
        ])
        .on_window_event(window_event)
        .setup(move |app| {
            #[cfg(target_os = "linux")]
            if let Some(listener) = listener {
                let app = app.handle().clone();
                instance::listen(listener, move |cwd, args| another_start(&app, args, cwd));
            }
            let home = app.path().home_dir().ok();
            // an earlier version kept the UI size in settings.json, in the app's config folder
            let old = app
                .path()
                .app_config_dir()
                .ok()
                .map(|d| d.join("settings.json"));
            app.manage(settings::Settings::load(
                store::Store::at(home.clone()),
                old.as_deref(),
            ));
            // expired versions and what a crash left, away from the windows' start
            let sweep = store::Store::at(home.clone());
            std::thread::spawn(move || sweep.sweep());
            let docs = doc::Docs::new(store::Store::at(home));
            #[cfg(feature = "lite")]
            lite::init(docs.store.root());
            // the CardputerZero's keyboard, read here and handed to the compositor
            #[cfg(all(feature = "lite", target_os = "linux"))]
            cardputer::start();
            app.manage(docs);
            // the files on the command line (Open with, a double-click on a file: "%1") open in their own windows;
            // without any, the Start Center
            let cwd = std::env::current_dir().unwrap_or_default();
            let files = if cfg!(feature = "lite") {
                Vec::new()
            } else {
                doc::files_among(std::env::args_os().skip(1), &cwd)
            };
            let start = {
                let app = app.handle().clone();
                move || {
                    STARTED.store(true, std::sync::atomic::Ordering::SeqCst);
                    #[cfg(target_os = "macos")]
                    let files = [files, std::mem::take(&mut *PENDING.lock().unwrap())].concat();
                    if open_files(&app, &files) > 0 && !cfg!(target_os = "macos") {
                        return;
                    }
                    #[cfg(feature = "lite")]
                    let first = lite::first_page(&app, std::env::args_os().skip(1));
                    #[cfg(not(feature = "lite"))]
                    let first = start_page();
                    if let Err(e) = open_window(&app, "main", first) {
                        eprintln!("the Start Center: {e}");
                        app.exit(1);
                    }
                }
            };
            // macOS: once WebKit has compiled the network rules (without them, no window opens)
            #[cfg(target_os = "macos")]
            net::prepare(app.handle(), start);
            #[cfg(not(target_os = "macos"))]
            start();
            Ok(())
        })
        .build(context)
        .expect("error while starting VibeOffice");
    #[cfg(not(target_os = "macos"))]
    app.run(|app, event| {
        if let tauri::RunEvent::Exit = event {
            app.state::<settings::Settings>().flush();
        }
    });
    #[cfg(target_os = "macos")]
    app.run(|app, event| match event {
        // Files from Finder (a double-click, Open With, a drop on the Dock icon), at start or later. The
        // sandbox lets the app read what the user opened this way.
        tauri::RunEvent::Opened { urls } => {
            let files: Vec<_> = urls.iter().filter_map(|u| u.to_file_path().ok()).collect();
            if STARTED.load(std::sync::atomic::Ordering::SeqCst) {
                let app = app.clone();
                tauri::async_runtime::spawn(async move { open_files(&app, &files) });
            } else {
                PENDING.lock().unwrap().extend(files);
            }
        }
        // A click on the Dock icon with no window open brings the Start Center back
        tauri::RunEvent::Reopen {
            has_visible_windows: false,
            ..
        } if app.webview_windows().is_empty() => {
            let app = app.clone();
            tauri::async_runtime::spawn(async move { show_start_center(&app) });
        }
        // The app keeps running, with its menu bar, after its last window closes; Quit (⌘Q) ends it
        tauri::RunEvent::ExitRequested {
            code: None, api, ..
        } => api.prevent_exit(),
        tauri::RunEvent::Exit => app.state::<settings::Settings>().flush(),
        _ => {}
    });
}
