//! The network the pages may use: https://vibeoffice.work/ alone, for the update check (updater/current.json
//! and whatsnew.json). Native TV casting uses separate sockets (cast.rs). Each window's web view refuses every other
//! http(s) and ws(s) request, whatever makes it (a page, a picture or frame in a document, a worker):
//! WKWebView through a content rule list (macOS), WebView2 by answering the request itself (Windows).
//! Links still open in the default browser (host::open_external), which is not the app.
//!
//! It fails closed: a window starts on about:blank and goes to its page only once the rules are on its web
//! view (`guard`); if they can't be, the window closes and says why, and the app quits when it was the only
//! one. `VO_NET_PROBE=1` checks that they work: before the Start Center's page, its window tries to reach
//! another site and prints `network probe: blocked` (or what it reached) on stderr (the macOS smoke run).

use tauri::{AppHandle, Url, WebviewUrl};
#[cfg(any(windows, target_os = "macos"))]
use tauri::{Manager, WebviewWindow};

/// The one site the pages may reach.
const SITE_HOST: &str = "vibeoffice.work";

/// Whether a web view may load `url`: the site over https, the app's own pages and IPC (Tauri's
/// https://tauri.localhost and https://ipc.localhost on Windows; the dev server in a dev build), and
/// anything that isn't the network (data:, blob:, about:, tauri:, ipc:). Windows asks this for every
/// request; macOS says the same in `rules`.
#[cfg_attr(not(windows), allow(dead_code))]
pub fn allowed(url: &str) -> bool {
    let Ok(url) = tauri::Url::parse(url) else {
        return false;
    };
    match url.scheme() {
        "https" if url.host_str() == Some(SITE_HOST) && url.port().is_none() => true,
        "https" if matches!(url.host_str(), Some("tauri.localhost" | "ipc.localhost")) => true,
        "http" | "https" => {
            !cfg!(feature = "custom-protocol") && url.host_str() == Some("127.0.0.1")
        }
        "ws" | "wss" => false,
        _ => true,
    }
}

/// WebKit's content rules (the format of Safari's content blockers): block every http(s) and ws(s)
/// request, then let the site's through (and the dev server's in a dev build). WebKit's url-filter is a
/// small subset of regular expressions: no alternation, so one rule per scheme.
#[cfg_attr(not(target_os = "macos"), allow(dead_code))]
fn rules() -> String {
    let block = |filter: &str| {
        format!(r#"{{"trigger":{{"url-filter":"{filter}"}},"action":{{"type":"block"}}}}"#)
    };
    let allow = |filter: &str| {
        format!(
            r#"{{"trigger":{{"url-filter":"{filter}"}},"action":{{"type":"ignore-previous-rules"}}}}"#
        )
    };
    let mut rules = vec![
        block("^https?://"),
        block("^wss?://"),
        allow(&format!("^https://{}/", SITE_HOST.replace('.', "\\\\."))),
    ];
    if !cfg!(feature = "custom-protocol") {
        rules.push(allow("^http://127\\\\.0\\\\.0\\\\.1[:/]"));
    }
    format!("[{}]", rules.join(","))
}

/// Where a window starts, until its rules are on (`guard`).
#[cfg(any(windows, target_os = "macos"))]
pub const BLANK: &str = "about:blank";

/// The address Tauri gives `url` (as its manager/webview.rs does): the app's own pages at tauri://localhost
/// (macOS, Linux) or https://tauri.localhost (Windows, `use_https_scheme`), the dev server in a dev build.
pub fn page_url(app: &AppHandle, url: WebviewUrl) -> Option<Url> {
    let base = if cfg!(feature = "custom-protocol") {
        Url::parse(if cfg!(windows) {
            "https://tauri.localhost/"
        } else {
            "tauri://localhost/"
        })
        .ok()?
    } else {
        app.config().build.dev_url.clone()?
    };
    match url {
        WebviewUrl::App(path) if path.to_str() == Some("index.html") => Some(base),
        WebviewUrl::App(path) => base.join(&path.to_string_lossy()).ok(),
        WebviewUrl::External(url) | WebviewUrl::CustomProtocol(url) => Some(url),
        _ => None,
    }
}

/// A window open_window has just made, on about:blank: the rules go on its web view, then it goes to `page`.
/// If they can't, it never gets there (`refuse`).
///
/// It goes on from here, not from inside with_webview: Tauri holds the web view's lock while that runs the
/// closure, which it does at once when called on the main thread, and navigate takes the same lock (the
/// main thread hung there, 2026-10-17). From another thread the closure runs later, on the main thread, and
/// a thread of its own waits for it.
#[cfg(any(windows, target_os = "macos"))]
pub fn guard(window: &WebviewWindow, page: Option<Url>) -> tauri::Result<()> {
    use std::sync::mpsc::{channel, TryRecvError};
    let (applied, result) = channel();
    window.with_webview(move |webview| {
        let _ = applied.send(apply(webview));
    })?;
    let w = window.clone();
    let go = move |result: Result<(), String>| {
        match result.and(page.ok_or_else(|| "the window has no page".to_string())) {
            Ok(page) if w.label() == "main" && probe::wanted() => probe::start(&w, page),
            Ok(page) => {
                let _ = w.navigate(page);
            }
            // on the main thread, which the alert needs (AppKit)
            Err(why) => {
                let window = w.clone();
                let _ = w.run_on_main_thread(move || refuse(&window, &why));
            }
        }
    };
    let gone = || Err("the window's web view is gone".to_string());
    match result.try_recv() {
        Ok(result) => go(result),
        Err(TryRecvError::Disconnected) => go(gone()),
        Err(TryRecvError::Empty) => {
            std::thread::spawn(move || go(result.recv().unwrap_or_else(|_| gone())));
        }
    }
    Ok(())
}

/// The rules on a window's web view: WebKit's compiled list (macOS), WebView2's request filter (Windows). On
/// Windows WebView2's browser keys go off here too, before the page (`win::browser_keys_off`).
#[cfg(any(windows, target_os = "macos"))]
fn apply(webview: tauri::webview::PlatformWebview) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    return mac::apply(webview.controller());
    #[cfg(windows)]
    {
        let controller = webview.controller();
        if let Err(e) = win::browser_keys_off(&controller) {
            eprintln!("browser keys: {e}");
        }
        win::apply(&controller, webview.environment()).map_err(|e| e.to_string())
    }
}

/// WebRTC, which neither the pages' policy nor the rules above cover (its connections aren't requests the web
/// view shows us): gone from every frame before its scripts (open_window). Linux's WebKitGTK has it off already.
pub const NO_WEBRTC: &str = "for (const k of ['RTCPeerConnection', 'webkitRTCPeerConnection', 'RTCDataChannel']) delete window[k];";

/// A window whose rules couldn't be set: it closes without loading its page, and says why; with no other
/// window open, the app quits.
#[cfg(any(windows, target_os = "macos"))]
fn refuse(window: &WebviewWindow, why: &str) {
    eprintln!("network rules: {why}");
    alert(&format!(
        "VibeOffice could not limit this window to https://{SITE_HOST}/, so it did not open it.\n\n{why}"
    ));
    let app = window.app_handle().clone();
    if app
        .webview_windows()
        .keys()
        .all(|label| label == window.label())
    {
        app.exit(1);
    } else {
        let _ = window.destroy();
    }
}

/// A message the user must read, in the system's own dialog.
#[cfg(any(windows, target_os = "macos"))]
fn alert(text: &str) {
    #[cfg(target_os = "macos")]
    mac::alert(text);
    #[cfg(windows)]
    crate::installed::message(
        text,
        windows::Win32::UI::WindowsAndMessaging::MB_OK
            | windows::Win32::UI::WindowsAndMessaging::MB_ICONERROR,
    );
}

/// `VO_NET_PROBE=1`: whether the rules refuse what they should, seen from a page. Before going to its page,
/// the Start Center's window loads a page of its own (a data: URL, with no policy of its own to block
/// anything first), which tries a picture, a fetch and a WebSocket on other sites, with a picture from
/// https://vibeoffice.work/ as the control, and puts the result in its title; `reported` prints it and sends
/// the window on to its page.
#[cfg(any(windows, target_os = "macos"))]
pub mod probe {
    use std::{collections::HashMap, sync::Mutex};
    use tauri::{Url, WebviewWindow};

    /// The page each probing window goes to afterwards
    static NEXT: Mutex<Option<HashMap<String, Url>>> = Mutex::new(None);

    pub fn wanted() -> bool {
        std::env::var_os("VO_NET_PROBE").is_some_and(|v| v == "1")
    }

    const PAGE: &str = r#"<!doctype html><title>network probe</title><script>
const image = (src) => new Promise((done) => { const i = new Image(); i.onload = () => done(true); i.onerror = () => done(false);
  setTimeout(() => done(false), 10000); i.src = src + '?' + Date.now(); });
const socket = (src) => new Promise((done) => { try { const s = new WebSocket(src); s.onopen = () => { s.close(); done(true); };
  s.onerror = () => done(false); setTimeout(() => done(false), 10000); } catch (e) { done(false); } });
(async () => {
  const control = await image('https://vibeoffice.work/icons/icon-192.png'), reached = [];
  if (await image('https://github.com/favicon.ico')) reached.push('a picture');
  if (await fetch('https://api.github.com/zen', { cache: 'no-store' }).then(() => true, () => false)) reached.push('a fetch');
  if (await socket('wss://echo.websocket.org/')) reached.push('a WebSocket');
  document.title = 'network probe: ' + (reached.length ? 'REACHED ' + reached.join(', ')
    : control ? 'blocked' : 'no control (https://vibeoffice.work/ did not load either)');
})();
</script>"#;

    pub fn start(window: &WebviewWindow, page: Url) {
        NEXT.lock()
            .unwrap()
            .get_or_insert_with(HashMap::new)
            .insert(window.label().into(), page);
        let encoded: String = PAGE
            .bytes()
            .map(|b| match b {
                b'a'..=b'z' | b'A'..=b'Z' | b'0'..=b'9' | b'-' | b'.' | b'_' => {
                    (b as char).to_string()
                }
                _ => format!("%{b:02X}"),
            })
            .collect();
        let url = format!("data:text/html;charset=utf-8,{encoded}");
        let _ = window.navigate(Url::parse(&url).expect("a data: URL"));
    }

    /// A title from the probe page (not one to show): the result is printed and the window goes on to its page.
    pub fn reported(window: &WebviewWindow, title: &str) -> bool {
        if !title.starts_with("network probe") {
            return false;
        }
        if let Some(result) = title.strip_prefix("network probe: ") {
            eprintln!("network probe: {result}");
            let next = NEXT
                .lock()
                .unwrap()
                .as_mut()
                .and_then(|n| n.remove(window.label()));
            if let Some(page) = next {
                let _ = window.navigate(page);
            }
        }
        true
    }
}

#[cfg(target_os = "macos")]
pub use mac::prepare;

#[cfg(target_os = "macos")]
mod mac {
    use block2::RcBlock;
    use objc2::{rc::Retained, MainThreadMarker};
    use objc2_app_kit::NSAlert;
    use objc2_foundation::{NSError, NSString};
    use objc2_web_kit::{WKContentRuleList, WKContentRuleListStore, WKUserContentController};
    use std::cell::RefCell;

    thread_local! {
        /// The compiled rules, once WebKit has them (main thread only, as WebKit is)
        static RULES: RefCell<Option<Retained<WKContentRuleList>>> = const { RefCell::new(None) };
    }

    /// Compiles the rules (WebKit keeps them, compiled, in the app's container), then calls `ready` on the
    /// main thread: setup opens the Start Center only then. If WebKit refuses them, no window opens: the
    /// app says why and quits.
    pub fn prepare(app: &super::AppHandle, ready: impl FnOnce() + 'static) {
        let mtm = MainThreadMarker::new().expect("network rules: not on the main thread");
        let refused = {
            let app = app.clone();
            move |why: String| {
                eprintln!("network rules: {why}");
                alert(&format!(
                    "VibeOffice could not limit its windows to https://{}/, so it did not start.\n\n{why}",
                    super::SITE_HOST
                ));
                app.exit(1);
            }
        };
        let Some(store) = (unsafe { WKContentRuleListStore::defaultStore(mtm) }) else {
            return refused("no content rule list store".into());
        };
        let then = RefCell::new(Some((ready, refused)));
        let done = RcBlock::new(move |list: *mut WKContentRuleList, error: *mut NSError| {
            let Some((ready, refused)) = then.borrow_mut().take() else {
                return;
            };
            match unsafe { Retained::retain(list) } {
                Some(list) => {
                    eprintln!("network rules: only https://{}/", super::SITE_HOST);
                    RULES.set(Some(list));
                    ready();
                }
                None => {
                    let why =
                        unsafe { error.as_ref() }.map(|e| e.localizedDescription().to_string());
                    refused(format!("WebKit refused them: {}", why.unwrap_or_default()));
                }
            }
        });
        unsafe {
            store.compileContentRuleListForIdentifier_encodedContentRuleList_completionHandler(
                Some(&NSString::from_str("vibeoffice-network")),
                Some(&NSString::from_str(&super::rules())),
                Some(&done),
            );
        }
    }

    /// The rules on a web view's user content controller (wry's WKUserContentController).
    pub fn apply(controller: *mut std::ffi::c_void) -> Result<(), String> {
        let controller = unsafe { Retained::retain(controller.cast::<WKUserContentController>()) }
            .ok_or("the web view has no user content controller")?;
        RULES.with_borrow(|list| match list {
            Some(list) => {
                unsafe { controller.addContentRuleList(list) };
                Ok(())
            }
            None => Err("the rules are not compiled".into()),
        })
    }

    /// A modal alert (NSAlert), on the main thread.
    pub fn alert(text: &str) {
        let Some(mtm) = MainThreadMarker::new() else {
            return;
        };
        let alert = NSAlert::new(mtm);
        alert.setMessageText(&NSString::from_str("VibeOffice"));
        alert.setInformativeText(&NSString::from_str(text));
        alert.runModal();
    }
}

#[cfg(windows)]
mod win {
    use webview2_com::{
        take_pwstr,
        Microsoft::Web::WebView2::Win32::{
            ICoreWebView2Controller, ICoreWebView2Environment, ICoreWebView2Settings3,
            ICoreWebView2_22, COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL,
            COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_ALL,
        },
        WebResourceRequestedEventHandler,
    };
    use windows::core::{Interface, HSTRING, PWSTR};

    /// WebView2's own keys off: F5 and Ctrl+R (reload), Ctrl+P (its print preview), Ctrl+F (its find bar),
    /// Ctrl+S (save the page), Ctrl + and − (zoom; Settings has the UI size). Editing keys stay, and the apps'
    /// own shortcuts reach the page.
    pub fn browser_keys_off(controller: &ICoreWebView2Controller) -> windows::core::Result<()> {
        unsafe {
            controller
                .CoreWebView2()?
                .Settings()?
                .cast::<ICoreWebView2Settings3>()?
                .SetAreBrowserAcceleratorKeysEnabled(false)
        }
    }

    /// Every request of the web view comes here first (also from workers, where WebView2 is new enough:
    /// ICoreWebView2_22); one `allowed` refuses gets a 403 without reaching the network. wry's own handler
    /// on the same event answers the app's pages (tauri.localhost), which this one lets through.
    pub fn apply(
        controller: &ICoreWebView2Controller,
        env: ICoreWebView2Environment,
    ) -> windows::core::Result<()> {
        unsafe {
            let webview = controller.CoreWebView2()?;
            let every = HSTRING::from("*");
            match webview.cast::<ICoreWebView2_22>() {
                Ok(webview) => webview.AddWebResourceRequestedFilterWithRequestSourceKinds(
                    &every,
                    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL,
                    COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_ALL,
                )?,
                Err(_) => webview
                    .AddWebResourceRequestedFilter(&every, COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL)?,
            }
            let handler = WebResourceRequestedEventHandler::create(Box::new(move |_, args| {
                let Some(args) = args else {
                    return Ok(());
                };
                let mut uri = PWSTR::null();
                args.Request()?.Uri(&mut uri)?;
                if !super::allowed(&take_pwstr(uri)) {
                    let refused = env.CreateWebResourceResponse(
                        None,
                        403,
                        &HSTRING::from("Forbidden"),
                        &HSTRING::from("Content-Type: text/plain"),
                    )?;
                    args.SetResponse(&refused)?;
                }
                Ok(())
            }));
            let mut token = 0;
            webview.add_WebResourceRequested(&handler, &mut token)?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn only_the_site_and_the_app_itself() {
        for url in [
            "https://vibeoffice.work/updater/current.json",
            "https://vibeoffice.work/updater/whatsnew.json",
            "https://tauri.localhost/quire/index.html",
            "https://ipc.localhost/host_read",
            "tauri://localhost/index.html",
            "ipc://localhost/host_read",
            "blob:https://tauri.localhost/0b6c",
            "data:image/png;base64,AAAA",
            "about:blank",
        ] {
            assert!(allowed(url), "{url}");
        }
        for url in [
            "http://tauri.localhost/quire/index.html",
            "http://ipc.localhost/host_read",
            "http://vibeoffice.work/updater/current.json",
            "https://vibeoffice.work:8443/",
            "https://www.vibeoffice.work/",
            "https://vibeoffice.work.example.com/",
            "https://vibeoffice.work@example.com/",
            "https://example.com/picture.png",
            "http://localhost:8080/",
            "http://192.168.1.1/",
            "wss://vibeoffice.work/",
            "not a url",
        ] {
            assert!(!allowed(url), "{url}");
        }
    }

    #[test]
    fn rules_are_json_webkit_reads() {
        let rules: serde_json::Value = serde_json::from_str(&rules()).expect("valid JSON");
        let filters: Vec<(&str, &str)> = rules
            .as_array()
            .unwrap()
            .iter()
            .map(|r| {
                (
                    r["trigger"]["url-filter"].as_str().unwrap(),
                    r["action"]["type"].as_str().unwrap(),
                )
            })
            .collect();
        assert_eq!(
            filters[..3],
            [
                ("^https?://", "block"),
                ("^wss?://", "block"),
                ("^https://vibeoffice\\.work/", "ignore-previous-rules")
            ]
        );
        assert!(
            filters.iter().all(|(f, _)| !f.contains('|')),
            "WebKit's url-filter has no alternation"
        );
    }
}
