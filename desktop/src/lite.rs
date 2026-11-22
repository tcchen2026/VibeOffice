//! The lite edition (docs/desktop.md, Lite edition): a keyboard-first VibeOffice for small Linux devices with little
//! memory (the CardputerZero: cardputer/). One window, `main`, has the start page (public/lite.html) or one app's
//! page with its document, in turn: one web process at a time. The documents are the files in
//! ~/VibeOffice/Documents, listed and opened by name; a Save As names a file there (doc.rs `host_choose`). Nothing
//! else is kept: no Recent Files, previews, drafts or previous versions (store.rs), no web view profile (main.rs).

use std::{
    path::{Path, PathBuf},
    sync::OnceLock,
};

use serde::Serialize;
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindow};

use crate::{doc::Docs, host, store, types};

/// The start page, the window's first and the one each document goes back to.
pub const START: &str = "lite.html";

/// ~/VibeOffice/Documents (`init`).
static DOCUMENTS: OnceLock<PathBuf> = OnceLock::new();

/// What a new Documents folder starts with: each app's sample, saved by the app (tools/lite-examples.mjs).
const EXAMPLES: [(&str, &[u8]); 3] = [
    (
        "Welcome to Quire.docx",
        include_bytes!("../lite-examples/Welcome to Quire.docx"),
    ),
    (
        "Quarterly Sales.xlsx",
        include_bytes!("../lite-examples/Quarterly Sales.xlsx"),
    ),
    (
        "Harbor & Pine.pptx",
        include_bytes!("../lite-examples/Harbor & Pine.pptx"),
    ),
];

/// The Documents folder in the edition's own (`store::Store::root`).
pub fn init(root: &Path) {
    let _ = DOCUMENTS.set(documents_in(root));
}

/// `root`'s Documents folder; made if missing, with the examples in it (one that is there, even empty, stays as it
/// is).
fn documents_in(root: &Path) -> PathBuf {
    let dir = root.join("Documents");
    if dir.symlink_metadata().is_err() {
        let made = std::fs::create_dir_all(&dir).and_then(|()| {
            EXAMPLES
                .iter()
                .try_for_each(|(name, bytes)| std::fs::write(dir.join(name), bytes))
        });
        if let Err(e) = made {
            eprintln!("{}: {e}", dir.display());
        }
    }
    dir
}

fn documents() -> Result<&'static Path, String> {
    DOCUMENTS
        .get()
        .map(PathBuf::as_path)
        .ok_or_else(|| "no Documents folder".into())
}

/// A name the pages give for a file in Documents: one name, not a path (no separators), not hidden.
fn name_ok(name: &str) -> bool {
    !name.trim().is_empty()
        && name.len() <= 255
        && !name.starts_with('.')
        && !name.contains(['/', '\\'])
        && !name.chars().any(char::is_control)
}

/// The file `name` in Documents, there or not. A link there is refused, so every file opened or written is in the
/// folder.
pub fn document(name: &str) -> Result<PathBuf, String> {
    if !name_ok(name) {
        return Err(format!("not a document's name: {name}"));
    }
    let path = documents()?.join(name);
    if path.symlink_metadata().is_ok_and(|m| !m.is_file()) {
        return Err(format!("“{name}” is not a file in Documents"));
    }
    Ok(path)
}

/// A document as the start page lists it.
#[derive(Serialize, Debug, PartialEq)]
pub struct Item {
    name: String,
    app: &'static str,
    size: u64,
    /// last changed, ms since 1970
    modified: u64,
}

/// The files in `dir` an app opens, newest first: plain files only (no links or folders), no hidden ones.
fn list(dir: &Path) -> Vec<Item> {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return Vec::new();
    };
    let mut items: Vec<Item> = entries
        .flatten()
        .filter(|e| e.file_type().is_ok_and(|t| t.is_file()))
        .filter_map(|e| {
            let name = e.file_name().into_string().ok()?;
            let app = types::app_for(Path::new(&name))?;
            if !name_ok(&name) {
                return None;
            }
            let meta = e.metadata().ok()?;
            Some(Item {
                name,
                app,
                size: meta.len(),
                modified: meta.modified().map(store::ms).unwrap_or(0),
            })
        })
        .collect();
    items.sort_by(|a, b| b.modified.cmp(&a.modified).then(a.name.cmp(&b.name)));
    items
}

/// The window's first page: the start page, or the document named first on the command line, a file in Documents
/// an app opens (`vibeoffice Agreement.docx`: for measuring and testing; M5's launcher names none).
pub fn first_page(
    app: &AppHandle,
    args: impl IntoIterator<Item = std::ffi::OsString>,
) -> WebviewUrl {
    let named = args.into_iter().next().map(|arg| {
        let name = arg.to_string_lossy().into_owned();
        let path = document(&name)?;
        let own = types::app_for(&path)
            .filter(|_| path.is_file())
            .ok_or_else(|| format!("“{name}” is not a document in Documents"))?;
        crate::doc::doc_address(app, "main", own, &path)?
    });
    let page = match named {
        Some(Ok(page)) => page,
        Some(Err(e)) => {
            eprintln!("{e}");
            START.into()
        }
        None => START.into(),
    };
    WebviewUrl::App(page.into())
}

/// The start page's list: the documents in Documents, and where the list leaves room, how files get there.
#[derive(Serialize)]
pub struct Listing {
    documents: Vec<Item>,
    /// the folder as its user would type it: ~/VibeOffice/Documents
    folder: String,
    /// for an SFTP app: the user at each address the device has on a network, Wi-Fi's first (none off any); not
    /// without an SSH server listening
    sftp: Option<Vec<String>>,
}

#[tauri::command]
pub fn host_lite_list() -> Result<Listing, String> {
    let dir = documents()?;
    let home = std::env::var_os("HOME").map(PathBuf::from);
    let folder = match home.as_deref().and_then(|h| dir.strip_prefix(h).ok()) {
        Some(inside) => Path::new("~").join(inside).display().to_string(),
        None => dir.display().to_string(),
    };
    let sftp = ssh_listens().then(|| {
        let user = std::env::var("USER")
            .or_else(|_| std::env::var("LOGNAME"))
            .unwrap_or_default();
        addresses()
            .iter()
            .map(|ip| format!("{user}@{ip}"))
            .collect()
    });
    Ok(Listing {
        documents: list(dir),
        folder,
        sftp,
    })
}

/// Whether an SSH server takes connections here (TCP port 22, listening: /proc/net/tcp and tcp6).
fn ssh_listens() -> bool {
    ["/proc/net/tcp", "/proc/net/tcp6"]
        .iter()
        .any(|f| std::fs::read_to_string(f).is_ok_and(|table| listens(&table, 22)))
}

/// Whether a /proc/net/tcp table has a socket listening on `port` (local address `IP:PORT` in hex, state 0A).
fn listens(table: &str, port: u16) -> bool {
    let local = format!(":{port:04X}");
    table.lines().skip(1).any(|line| {
        let mut cols = line.split_whitespace();
        let at = cols.nth(1).unwrap_or_default();
        let state = cols.nth(1).unwrap_or_default();
        at.ends_with(&local) && state == "0A"
    })
}

/// The device's IPv4 addresses on a network (not loopback or link-local), Wi-Fi's (wl…) first.
fn addresses() -> Vec<std::net::Ipv4Addr> {
    let mut found: Vec<(bool, std::net::Ipv4Addr)> = Vec::new();
    #[cfg(target_os = "linux")]
    // SAFETY: getifaddrs fills `list` with a linked list valid until freeifaddrs; each entry's name is a C string,
    // its address null or a sockaddr of the family it gives (sockaddr_in for AF_INET)
    unsafe {
        let mut list: *mut libc::ifaddrs = std::ptr::null_mut();
        if libc::getifaddrs(&mut list) != 0 {
            return Vec::new();
        }
        let mut at = list;
        while let Some(ifa) = at.as_ref() {
            if !ifa.ifa_addr.is_null() && i32::from((*ifa.ifa_addr).sa_family) == libc::AF_INET {
                let sin = &*(ifa.ifa_addr as *const libc::sockaddr_in);
                let ip = std::net::Ipv4Addr::from(u32::from_be(sin.sin_addr.s_addr));
                let wifi = std::ffi::CStr::from_ptr(ifa.ifa_name)
                    .to_bytes()
                    .starts_with(b"wl");
                if !ip.is_loopback() && !ip.is_link_local() && !ip.is_unspecified() {
                    found.push((!wifi, ip));
                }
            }
            at = ifa.ifa_next;
        }
        libc::freeifaddrs(list);
    }
    found.sort();
    found.dedup();
    found.into_iter().map(|(_, ip)| ip).collect()
}

/// The start page opens the document `name`: the window goes to its app's page with it (`doc::open_here`).
#[tauri::command]
pub async fn host_lite_open(
    app: AppHandle,
    window: WebviewWindow,
    name: String,
) -> Result<(), String> {
    let path = document(&name)?;
    if !path.is_file() {
        return Err(format!("“{name}” is no longer in Documents"));
    }
    let own = types::app_for(&path).ok_or_else(|| format!("VibeOffice can't open “{name}”"))?;
    crate::doc::open_here(&app, &window, own, &path)
}

/// A new blank document of `page` (quire, ledger, lectern), in the window, once its page has let its own go.
#[tauri::command]
pub async fn host_lite_new(
    app: AppHandle,
    window: WebviewWindow,
    page: String,
) -> Result<(), String> {
    if !types::is_app(&page) {
        return Err(format!("not one of the apps: {page}"));
    }
    let start = host::Start {
        new: true,
        template: None,
    };
    let path = host::start_path(&app, &page, &start)?;
    app.state::<Docs>().forget(window.label());
    host::navigate(&app, &window, &path)
}

/// Back to the start page, once the app's page has let its document go (File ▸ Close, Exit, Open).
#[tauri::command]
pub async fn host_lite_home(app: AppHandle, window: WebviewWindow) -> Result<(), String> {
    app.state::<Docs>().forget(window.label());
    host::navigate(&app, &window, START)
}

/// The licences page (Help on the start page), which comes back with Esc (oss-licenses.js).
#[tauri::command]
pub async fn host_lite_licenses(app: AppHandle, window: WebviewWindow) -> Result<(), String> {
    host::navigate(&app, &window, "oss-licenses.html")
}

/// The start page's Quit.
#[tauri::command]
pub fn host_lite_quit(app: AppHandle) {
    app.exit(0);
}

/// Lectern's slide show (`on`) and its end: with an HDMI screen connected, the window goes to the desktop there (M5's
/// labwc, full screen) for the show and comes back after it, while the device's own screen says so in a few words of
/// its own (`hdmi`). Whether it went.
#[tauri::command]
pub async fn host_lite_show(
    app: AppHandle,
    window: WebviewWindow,
    on: bool,
) -> Result<bool, String> {
    let (tx, rx) = std::sync::mpsc::channel();
    app.run_on_main_thread(move || {
        let _ = tx.send(hdmi::show(&window, on));
    })
    .map_err(|e| e.to_string())?;
    rx.recv().map_err(|e| e.to_string())?
}

#[cfg(target_os = "linux")]
mod hdmi {
    use std::{cell::RefCell, path::PathBuf};

    use gtk::prelude::*;
    use tauri::WebviewWindow;

    /// What the device's screen shows while the window is on HDMI.
    const NOTE: &str = "<span size='large' weight='bold'>Showing on HDMI</span>\n\n\
        Right, Space: next slide\nLeft: previous slide\nEsc: end the show";

    /// The show on HDMI: the window's screen to go back to, the note on the device's screen.
    struct Away {
        home: gtk::gdk::Screen,
        note: gtk::Window,
    }

    thread_local! {
        static AWAY: RefCell<Option<Away>> = const { RefCell::new(None) };
        /// The HDMI desktop's display, by its name, open from the first show on: closed after one, it took visuals
        /// the window still used with it (GTK crashed).
        static DISPLAY: RefCell<Option<(String, gtk::gdk::Display)>> = const { RefCell::new(None) };
    }

    /// The Wayland display of the desktop on a connected HDMI screen: the live socket in $XDG_RUNTIME_DIR other than
    /// ours (cage's). VIBEOFFICE_HDMI=wayland-N names one, to try this without the device.
    fn display() -> Option<String> {
        if let Some(name) = std::env::var_os("VIBEOFFICE_HDMI") {
            return name.into_string().ok();
        }
        let connected = std::fs::read_dir("/sys/class/drm")
            .ok()?
            .flatten()
            .any(|e| {
                e.file_name().to_string_lossy().contains("-HDMI-")
                    && std::fs::read_to_string(e.path().join("status"))
                        .is_ok_and(|s| s.trim() == "connected")
            });
        if !connected {
            return None;
        }
        let ours = std::env::var("WAYLAND_DISPLAY").unwrap_or_default();
        let dir = PathBuf::from(std::env::var_os("XDG_RUNTIME_DIR")?);
        let mut names: Vec<String> = std::fs::read_dir(&dir)
            .ok()?
            .flatten()
            .filter_map(|e| e.file_name().into_string().ok())
            .filter(|n| n.starts_with("wayland-") && !n.ends_with(".lock") && *n != ours)
            .collect();
        names.sort();
        names
            .into_iter()
            .find(|n| std::os::unix::net::UnixStream::connect(dir.join(n)).is_ok())
    }

    pub fn show(window: &WebviewWindow, on: bool) -> Result<bool, String> {
        let gtk_window = window.gtk_window().map_err(|e| e.to_string())?;
        if !on {
            if let Some(away) = AWAY.with(|a| a.borrow_mut().take()) {
                gtk_window.unfullscreen();
                gtk_window.set_screen(&away.home);
                gtk_window.present();
                away.note.close();
                crate::cardputer::follow(None);
            }
            return Ok(false);
        }
        if AWAY.with(|a| a.borrow().is_some()) {
            return Ok(true);
        }
        let Some(name) = display() else {
            return Ok(false);
        };
        let display = DISPLAY.with(|d| {
            let mut d = d.borrow_mut();
            match &*d {
                Some((open, display)) if *open == name => Ok(display.clone()),
                _ => {
                    let display = gtk::gdk::Display::open(&name)
                        .ok_or_else(|| format!("no display {name}"))?;
                    *d = Some((name.clone(), display.clone()));
                    Ok::<_, String>(display)
                }
            }
        })?;
        let note = gtk::Window::new(gtk::WindowType::Toplevel);
        note.set_decorated(false);
        let label = gtk::Label::new(None);
        label.set_markup(NOTE);
        label.set_justify(gtk::Justification::Center);
        note.add(&label);
        note.show_all();
        note.fullscreen();
        let home = gtk::prelude::GtkWindowExt::screen(&gtk_window).ok_or("no screen")?;
        gtk_window.set_screen(&display.default_screen());
        gtk_window.fullscreen();
        gtk_window.present();
        crate::cardputer::follow(Some(name));
        AWAY.with(|a| *a.borrow_mut() = Some(Away { home, note }));
        Ok(true)
    }
}

#[cfg(not(target_os = "linux"))]
mod hdmi {
    pub fn show(_: &tauri::WebviewWindow, _: bool) -> Result<bool, String> {
        Ok(false)
    }
}

/// WebKit's own folder (an empty cache, salts), even for an incognito web view: in $XDG_RUNTIME_DIR, which is gone
/// when the user logs out, not in ~/.local/share.
pub fn web_data() -> PathBuf {
    std::env::var_os("XDG_RUNTIME_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("vibeoffice-lite")
}

/// When the process started (`clock`, first thing in main), for the log's timings: on the device the window comes
/// and pages load seconds apart, and the log ($XDG_RUNTIME_DIR/vibeoffice.log, cardputer/launch.sh) says when.
static T0: OnceLock<std::time::Instant> = OnceLock::new();

pub fn clock() {
    T0.get_or_init(std::time::Instant::now);
}

fn seconds() -> f64 {
    T0.get().map_or(0.0, |t| t.elapsed().as_secs_f64())
}

/// The window has just been made: VibeOffice's own splash covers it until its first page has loaded (`loaded`), as
/// the window shows long before WebKit has drawn anything (on the device, seconds of an empty screen).
pub fn starting(window: &WebviewWindow) {
    eprintln!("VibeOffice Lite: the window after {:.1} s", seconds());
    #[cfg(target_os = "linux")]
    {
        // the app's icon as the window has it (from icons/icon.png)
        let icon = window
            .app_handle()
            .default_window_icon()
            .map(|i| (i.rgba().to_vec(), i.width() as i32, i.height() as i32));
        let w = window.clone();
        let _ = window.run_on_main_thread(move || {
            if let Ok(parent) = w.gtk_window() {
                toast::splash(&parent, icon);
            }
        });
    }
    #[cfg(not(target_os = "linux"))]
    let _ = window;
}

/// The page the window goes to (host::navigate, at once) or starts to load (main.rs): an app's (a document opened or
/// made from the list) shows a note of its own until it has loaded (`loaded`), as the page itself can show nothing
/// before its scripts have run (on the device, seconds), and a hint about the pointer keys.
pub fn loading(window: &WebviewWindow, url: &tauri::Url) {
    #[cfg(target_os = "linux")]
    if types::is_app(url.path().trim_matches('/')) {
        let w = window.clone();
        let _ = window.run_on_main_thread(move || {
            if let Ok(parent) = w.gtk_window() {
                toast::show(&parent);
            }
        });
    }
    #[cfg(not(target_os = "linux"))]
    let _ = (window, url);
}

/// The page has loaded: the splash or the note goes.
pub fn loaded(window: &WebviewWindow, url: &tauri::Url) {
    eprintln!(
        "VibeOffice Lite: {} loaded after {:.1} s",
        url.path(),
        seconds()
    );
    #[cfg(target_os = "linux")]
    let _ = window.run_on_main_thread(toast::hide);
    #[cfg(not(target_os = "linux"))]
    let _ = window;
}

/// What covers the window for a while, in a popup of the kind a tooltip is (under Wayland part of the window's
/// surface, not a window cage would show full screen; no focus of its own, so keys stay with the page), native, no
/// web view: the splash at the start (`splash`: the whole window, the start page's colours, the Classic look's
/// --surface, --ink and --ink-7) and the note while an app's page loads (`show`: two lines in the middle, as the apps'
/// toasts look, public/common/ui.css `.toast`, the Classic look's tooltip colours). One at a time.
#[cfg(target_os = "linux")]
mod toast {
    use std::cell::RefCell;

    use gtk::gdk_pixbuf::{Colorspace, InterpType, Pixbuf};
    use gtk::prelude::*;

    const TEXT: &str =
        "<b>Loading…</b>\n<small>Hint: Sym + F X Z C, Enter, Space for the mouse</small>";
    const STYLE: &str = "window { background: #ffffe1; color: #000; border: 1px solid #000; }";
    const SPLASH: &str = "<span size='large' weight='bold'>VibeOffice Lite</span>\n\
        <small><span foreground='#777777'>Starting…</span></small>";
    const SPLASH_STYLE: &str = "window { background: #fff; color: #000; } label { color: #000; }";

    thread_local! {
        static TOAST: RefCell<Option<gtk::Window>> = const { RefCell::new(None) };
    }

    /// A popup over `parent`, in `style`, unless one is up
    fn popup(parent: &gtk::ApplicationWindow, style: &str) -> Option<gtk::Window> {
        if TOAST.with(|t| t.borrow().is_some()) {
            return None;
        }
        let note = gtk::Window::new(gtk::WindowType::Popup);
        note.set_type_hint(gtk::gdk::WindowTypeHint::Tooltip);
        note.set_transient_for(Some(parent));
        let css = gtk::CssProvider::new();
        if css.load_from_data(style.as_bytes()).is_ok() {
            note.style_context()
                .add_provider(&css, gtk::STYLE_PROVIDER_PRIORITY_APPLICATION);
        }
        Some(note)
    }

    fn up(note: gtk::Window) {
        note.show();
        TOAST.with(|t| *t.borrow_mut() = Some(note));
    }

    pub fn show(parent: &gtk::ApplicationWindow) {
        let Some(note) = popup(parent, STYLE) else {
            return;
        };
        let label = gtk::Label::new(None);
        label.set_markup(TEXT);
        label.set_justify(gtk::Justification::Center);
        label.set_line_wrap(true);
        label.set_margin_top(4);
        label.set_margin_bottom(4);
        label.set_margin_start(8);
        label.set_margin_end(8);
        note.add(&label);
        label.show();
        // as wide as its text, at most the window's width less a little
        let (width, height) = parent.size();
        let w = note.preferred_width().1.min(width - 16).max(1);
        let h = note.preferred_height_for_width(w).1;
        note.set_size_request(w, h);
        let (x, y) = parent.position();
        note.move_(x + (width - w).max(0) / 2, y + (height - h).max(0) / 2);
        up(note);
    }

    /// The splash: the app's icon (`icon`: RGBA, width, height) and name over the whole window
    pub fn splash(parent: &gtk::ApplicationWindow, icon: Option<(Vec<u8>, i32, i32)>) {
        let Some(note) = popup(parent, SPLASH_STYLE) else {
            return;
        };
        let column = gtk::Box::new(gtk::Orientation::Vertical, 6);
        column.set_valign(gtk::Align::Center);
        if let Some((rgba, w, h)) = icon {
            let full = Pixbuf::from_bytes(
                &gtk::glib::Bytes::from_owned(rgba),
                Colorspace::Rgb,
                true,
                8,
                w,
                h,
                w * 4,
            );
            if let Some(small) = full.scale_simple(48, 48, InterpType::Bilinear) {
                column.add(&gtk::Image::from_pixbuf(Some(&small)));
            }
        }
        let label = gtk::Label::new(None);
        label.set_markup(SPLASH);
        label.set_justify(gtk::Justification::Center);
        column.add(&label);
        note.add(&column);
        column.show_all();
        let (width, height) = parent.size();
        note.set_size_request(width, height);
        let (x, y) = parent.position();
        note.move_(x, y);
        up(note);
    }

    pub fn hide() {
        if let Some(note) = TOAST.with(|t| t.borrow_mut().take()) {
            note.close();
        }
    }
}

/// The window's web view keeps no page it left (no back-forward cache) and caches little: WebKit's document viewer
/// model. The memory goes back as the window moves between the start page and a document.
pub fn web_view(window: &WebviewWindow) {
    #[cfg(target_os = "linux")]
    let _ = window.with_webview(|webview| {
        use webkit2gtk::{CacheModel, WebContextExt, WebViewExt};
        if let Some(context) = webview.inner().context() {
            context.set_cache_model(CacheModel::DocumentViewer);
        }
    });
    #[cfg(not(target_os = "linux"))]
    let _ = window;
}

#[cfg(test)]
mod tests {
    use super::*;

    fn temp(name: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!("vo-lite-{name}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn an_ssh_server_is_seen_listening() {
        let table = "  sl  local_address rem_address   st tx_queue rx_queue tr tm->when retrnsmt   uid  timeout inode
   0: 00000000:0016 00000000:0000 0A 00000000:00000000 00:00000000 00000000     0        0 1234 1 0 100 0 0 10 0
   1: 0100007F:1F90 0100007F:0016 01 00000000:00000000 00:00000000 00000000  1000        0 5678 1 0 20 4 30 10 -1";
        assert!(listens(table, 22));
        // port 22 only as the far end of a connection, or another port listening: no
        assert!(!listens(table, 8080));
        let connected = table.replace(" 0A ", " 01 ");
        assert!(!listens(&connected, 22));
        assert!(!listens("", 22));
    }

    #[test]
    fn a_new_documents_folder_starts_with_the_examples() {
        let root = temp("examples");
        let dir = documents_in(&root);
        let mut listed: Vec<_> = list(&dir).into_iter().map(|i| (i.name, i.app)).collect();
        listed.sort();
        assert_eq!(
            listed,
            [
                ("Harbor & Pine.pptx".to_string(), "lectern"),
                ("Quarterly Sales.xlsx".to_string(), "ledger"),
                ("Welcome to Quire.docx".to_string(), "quire"),
            ]
        );
        for (name, bytes) in EXAMPLES {
            assert!(bytes.starts_with(b"PK\x03\x04"), "{name} is a package");
        }
        // a folder that is there stays as it is, emptied or not
        for (name, _) in EXAMPLES {
            std::fs::remove_file(dir.join(name)).unwrap();
        }
        documents_in(&root);
        assert!(list(&dir).is_empty());
        let _ = std::fs::remove_dir_all(&root);
    }

    #[test]
    fn names() {
        for ok in ["Report.docx", "a b.xlsx", "über.md", "x"] {
            assert!(name_ok(ok), "{ok}");
        }
        for bad in [
            "",
            " ",
            ".hidden.docx",
            "..",
            "../x.docx",
            "a/b.docx",
            "a\\b.docx",
            "a\nb",
            &"x".repeat(256),
        ] {
            assert!(!name_ok(bad), "{bad:?}");
        }
    }

    #[test]
    fn lists_documents_newest_first() {
        let dir = temp("list");
        let write = |name: &str, age: u64| {
            let p = dir.join(name);
            std::fs::write(&p, b"x").unwrap();
            let t = std::time::SystemTime::now() - std::time::Duration::from_secs(age);
            std::fs::File::options()
                .write(true)
                .open(&p)
                .unwrap()
                .set_modified(t)
                .unwrap();
        };
        write("old.docx", 300);
        write("new.xlsx", 10);
        write("mid.pptx", 100);
        write("notes.exe", 0); // no app opens it
        write(".hidden.docx", 0);
        std::fs::create_dir(dir.join("folder.docx")).unwrap();
        #[cfg(unix)]
        std::os::unix::fs::symlink(dir.join("old.docx"), dir.join("link.docx")).unwrap();
        let names: Vec<_> = list(&dir).into_iter().map(|i| (i.name, i.app)).collect();
        assert_eq!(
            names,
            [
                ("new.xlsx".to_string(), "ledger"),
                ("mid.pptx".to_string(), "lectern"),
                ("old.docx".to_string(), "quire"),
            ]
        );
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn documents_stay_in_the_folder() {
        let dir = temp("doc");
        let _ = DOCUMENTS.set(dir.clone());
        let dir = documents().unwrap().to_path_buf();
        assert_eq!(document("a.docx").unwrap(), dir.join("a.docx"));
        assert!(document("../a.docx").is_err());
        assert!(document("/etc/passwd").is_err());
        #[cfg(unix)]
        {
            std::os::unix::fs::symlink("/etc/passwd", dir.join("out.docx")).unwrap();
            assert!(document("out.docx").is_err());
        }
    }
}
