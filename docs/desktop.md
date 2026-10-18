# VibeOffice desktop app (Tauri): Windows and macOS, Linux too

`desktop/` wraps the suite in `public/` in a native window with [Tauri 2](https://tauri.app). The pages are the same files as the web version, embedded in the binary (brotli) and served from `https://tauri.localhost/` (Windows) or `tauri://localhost/` (macOS, Linux). They draw in the system's web view: WebView2 (Edge) on Windows, which every supported Windows 10 and 11 already has; WKWebView (WebKit, Safari's engine) on macOS; WebKitGTK 4.1 on Linux. The pages reach the system only through `VO.host` (Host interface below).

| File | What |
|---|---|
| `desktop/Cargo.toml` | the crate (on macOS also `objc2-foundation` and `objc2-app-kit`, which Tauri builds anyway, with only the classes used); size-tuned release profile (`opt-level = "z"`, LTO, one codegen unit, `panic = "abort"`, stripped); Tauri without `dynamic-acl` |
| `desktop/.cargo/config.toml` | builds for `x86_64-pc-windows-gnu` with mingw-w64 by default |
| `desktop/tauri.conf.json` | app name and identifier, `frontendDist: ../public`, `devUrl` (tools/serve.py), `removeUnusedCommands`; Tauri's bundler is off (the installer is `setup/`) |
| `desktop/src/main.rs` | the files the app is started with, each in a window of its own, else the Start Center; one instance on Windows (`second_start`) and Linux (`instance.rs`); every window (`open_window`), incognito but on Linux (Settings and storage, below): system title bar follows `document.title`, file drops go to the page, links out of the app go to the default browser, the UI size on each page load, `__VO_HOST__` before each page; an app window's close button goes to its page (`window_event`); on macOS the menu bar (`mac_menu`), Finder's files (`RunEvent::Opened`), the Dock (`Reopen`) and staying open with no window (macOS, below) |
| `desktop/src/host.rs` | the native side of `VO.host`: the window and link commands, the navigation and link rules, the script that takes the browser's storage away (`NO_STORAGE`) |
| `desktop/src/settings.rs` | the settings (`VO.settings`) in `~/.vibeoffice/config.json`: `host_set_setting`, the UI size, served at the top of `common/look.js` (Settings and storage, below) |
| `desktop/src/instance.rs` | Linux: one VibeOffice per user and display, a lock and a Unix socket |
| `desktop/src/files.rs` | writing a user's file (Saving, below): the checks, the replace on each system, the error codes |
| `desktop/src/doc.rs` | the document in each window (Documents and windows, below): opening files and Recent Files into windows, the window's starting document, drafts and previews, the Window menu, the close question; on macOS Recent Files' bookmarks |
| `desktop/src/store.rs` | `~/.vibeoffice`: `config.json` (the settings), Recent Files, their previews and drafts, each file written whole (`write_atomic`) |
| `desktop/src/types.rs` | which app opens which file: the table `VO.APPS` is (a test compares them) |
| `desktop/src/dialog.rs` | the system's Open dialog and a yes-or-no question: rfd on Windows and macOS, GTK's native chooser on Linux |
| `desktop/src/net.rs` | the network rules: each window's web view reaches `https://vibeoffice.work/` alone, or the window never loads its page; `VO_NET_PROBE` (Network, below) |
| `desktop/src/installation.rs` | what the installer writes and `-uninstall` removes, in one list: files, shortcuts, HKCU registrations (shared with `setup/`) |
| `desktop/src/installed.rs` | Windows: `vibeoffice.exe -uninstall`, and the first start after installing (taskbar pin request, Default apps) |
| `desktop/app.manifest` | common controls v6 and `asInvoker` (never asks for administrator rights) |
| `desktop/macos/` | `Info.plist` (the bundle's, filled in by `build.sh`), `direct.entitlements` and `appstore.entitlements`, `build.sh` (the macOS release, on a Mac: macOS, below) |
| `desktop/linux/` | `build.sh` (the AppImage: Linux, below), `AppRun` (its entry point: the check for WebKitGTK 4.1), `vibeoffice.desktop`, `smoke.sh` and `on-system.sh` (the CI's runs) |
| `desktop/release-notes.md` | the GitHub release's notes, every download (any of the workflows may create the release) |
| `desktop/setup/` | the installer, `VibeOfficeSetup.exe`: carries the app; its dialog, manifest and version in `setup.rc` |
| `desktop/permissions/host.toml`, `desktop/capabilities/` | which `host_*` commands each kind of window may call (Host interface, below); no Tauri core commands |
| `desktop/icons/` | `icon.icns` (macOS: the app icon on Apple's grid, 16–1024 px, from `public/icons/icon.svg` by `tools/desktop-icons.py`), `icon.ico` (16–256 px; the exe and the installer), `quire.ico`, `ledger.ico`, `lectern.ico` (document icons, resources 40101–40103 in the exe, numbered after the app icon's 32512 so that stays the exe's first, generated from each app's SVG by `tools/desktop-icons.py`), and `icon.png` (128 px, Linux window icon; Tauri embeds it as raw RGBA, so keep it small) |
| `desktop/licenses/` | license texts of what ships with the app but is not a crate (`WebView2Loader.txt`, from the WebView2 SDK package; `appimage-runtime.txt`, the AppImage runtime's and its parts'), and in `crates/` those of the crates whose package has none (taken from their repositories), for `public/oss-licenses.html` |
| `desktop/patches/option-ext/` | our own stand-in for the `option-ext` crate (MPL-2.0), which `dirs-sys` uses through Tauri: the one method it needs, Apache-2.0, put in its place by `[patch.crates-io]` in `Cargo.toml` (AGENTS.md, Code from other projects) |
| `desktop/build.sh` | release build: the exe, the installer and the zip in `desktop/target/dist/` |
| `.github/workflows/release.yml` | CI, run by hand: every download in one run: the version's release as a draft, then Windows, the macOS DMG and Linux one after another (the workflows below), published once all three files are in it |
| `.github/workflows/windows.yml` | CI, run by hand or by `release.yml`: `build.sh` on Ubuntu; the zip as the run's artifact and a GitHub release |
| `.github/workflows/linux.yml` | CI, run by hand or by `release.yml`: `linux/build.sh` on Ubuntu 22.04, the AppImage in the version's GitHub release; with *smoke*, first a smoke run there and each system in a container (without WebKitGTK, then with it), screenshots |
| `.github/workflows/macos.yml` | CI, run by hand or by `release.yml`: `macos/build.sh dmg` on a macOS runner, with *smoke* a smoke run with screenshots; the DMG in the version's GitHub release (macOS, below) |
| `.github/workflows/appstore.yml` | CI, run by hand: `macos/build.sh appstore`, the App Store package to App Store Connect and as the run's artifact (macOS, below) |
| `tests/desktop/` | `host.test.mjs` and `host-stub.js` (Host interface, below); `setup-wine.py`, the installer end to end in Wine (The installer, below) |

## Host interface

One interface between the pages and what runs them, so a page never asks which: `VO.host` in `public/common/suite.js`. In a browser (tab or installed PWA) it is built on the browser's APIs; in the desktop app on the commands in `desktop/src/host.rs`, called over Tauri's IPC (`window.__TAURI_INTERNALS__.invoke`). The desktop app announces itself before each page's scripts with `window.__VO_HOST__ = { platform, version, appVersion, edition }` (`edition`: `'appstore'` for the Mac App Store's, else `'direct'`; written as JSON, frozen, on a property the page can't replace). The settings come with `common/look.js` (Settings and storage, below).

| `VO.host` | Web | Desktop |
|---|---|---|
| `kind` | `'web'` | `'desktop'` |
| `platform` | `'web'` | `'windows'`, `'macos'`, `'linux'` (Rust's `std::env::consts::OS`) |
| `updates` | `true` (the page never checks) | `true`; `false` in the Mac App Store edition, which the store updates (the Start Center never reads `current.json`; old, it still shows Check for updates, for the App Store) |
| `version` | `0` | `host::VERSION`, raised when a command changes shape (5 since the settings moved to `~/.vibeoffice`) |
| `onFiles(fn)` | `launchQueue` (the PWA's file handlers) | nothing: the native side opens the files the app is started with, or that another start (Windows) or Finder (macOS) hands over, each in a window of its own |
| `openUrl(url)` | a new tab | `host_open_url`, from the Start Center: the update page and our GitHub only (Links, below), to the system (`ShellExecuteW`, `NSWorkspace openURL:`, `xdg-open`) |
| `newWindow(url)` | `window.open` (false if blocked) | `host_new_window`: told the page (`quire`, `ledger`, `lectern`, `licenses`) and how it starts (`VO.launch`: `new`, `template`), never an address; Rust makes the address (`host::page_path`) and a full window with the host script and handlers, labelled `<page>-<n>`. An app window opens a blank document or a template of its own app, and the licences page; at most 64 windows. A Recent Files card's address (`recent`, `draft`) goes to `openRecent` |
| `open()`, `reopen(go)` | — | `host_open`: the system's Open dialog (the window's app's types; the Start Center's, all): `{ opened, here }`. In an app window the first file of its app is held (`here`, its name) until the page has let its document go, then `host_reopen` reloads the window with it; every other file opens in a window of its own (Documents and windows, below) |
| `openRecent(key, draft)`, `showRecent()` | — | the Start Center's `host_open_recent`: a Recent Files entry, by its key, in a window of its own; an app window's `host_show_recent`: the Start Center on its Recent Files |
| `newDocument(blob, name, made)` | — | `host_new_document`: bytes in a new window (raw, the name in a header): a file dropped on a window that holds a document, or one a page made (`made`: never saved) |
| `document()` | — | `host_document`, `host_document_bytes`: this window's starting document, once (`{ info, file }`, or null) |
| `windows()`, `focus(id)` | `[]` | `host_windows`, `host_focus`: the app's other windows, for its Window menu |
| `closeWindow()` | `false` | `host_close_window`: closes an app window (true), once its page has its document saved or given up; the Start Center's (`main`) stays (false) |
| `file`, `save(bytes, replace)`, `choose(args)`, `writeChosen(bytes)`, `cancelChoice()`, `versions()`, `openVersion(id)` | — | saving (Saving, below; used through `L.file`): the window's file (`{ bound, readOnly, name }`), `host_save`, `host_choose`, `host_write_chosen`, `host_cancel_choice`, `host_versions`, `host_open_version` |
| `sandboxed`, `allowAccess(how, text)` | `false` | the Mac App Sandbox: `host_allow_access`, the folder chooser for a refused save's folder (Saving, below) |
| `uiScale()`, `setUiScale(n)` | `null` (the browser's own zoom does it) | the setting `vibeoffice.uiScale` (`host_set_setting`): the web view's zoom, 0.5–4, every window |

`VO.VERSION` is `appVersion` in the desktop app (the build's version, below).

`fn(files)` gets `File` objects; the Start Center calls `onFiles` and opens the first file in its app (`VO.handoff`). The service worker is registered only on the web.

Each kind of window calls only its own commands (`desktop/permissions/host.toml`, granted in `desktop/capabilities/`; Tauri refuses the rest, and every Tauri core command, which nothing grants): the Start Center (`main`) Open File, Recent Files, the one window that lists them (`host_recent`, `host_recent_preview`, `host_recent_remove`, `host_recent_clear`, `host_open_recent`, `host_discard_draft`), a dropped file (`host_new_document`), app windows, the update page and the settings (`host_set_setting`); the apps (`quire-*`, `ledger-*`, `lectern-*`) their document (`host_document`, `host_document_bytes`, `host_set_draft`, `host_clear_draft`, `host_set_preview`), saving it (`host_save`, `host_allow_access`, `host_choose`, `host_write_chosen`, `host_cancel_choice`, `host_versions`, `host_open_version`), File ▸ Open into the window (`host_open`, `host_reopen`), the Start Center for Recent Files (`host_show_recent`; no entry's key), the Window menu, closing, the licences page and the settings; the licences page (`licenses-*`) none (it reads the settings it is served). A new command goes into a permission there, and into `tests/desktop/host-stub.js`, which refuses the same (a test compares them).

Deadlocks: a command without `async` runs on the main thread, which runs every window, so one that makes a window is `async` (`host_new_window`: from a sync command that deadlocks on Windows). Tauri holds a web view's lock while it runs a `with_webview` closure, at once when called on the main thread, so the closure calls no method of that window or web view (`navigate`, `eval`, `set_zoom`): `net::guard` did, and the app hung at start (2026-10-17).

Links: the app sends an address to the default browser only from the Start Center and only from its list, `host::is_ours`: the update page (`https://vibeoffice.work/updater/`) and our GitHub (`https://github.com/tcchen2026/VibeOffice`); https, no port, no user name (`host::may_open`, for `host_open_url` and for a navigation or new window that leaves the app: `on_navigation`, `on_new_window` → `host::leave`). Anything else opens nothing (a line on stderr); the pages show such an address with Copy instead, for the user to paste into their browser: a document's hyperlinks (`ui.linkDialog`; docs/suite.md, Hyperlinks) and the licences page's links (`oss-licenses.js`, in the desktop app only).

A window loads one page, the one Rust opened it with (`host::Page`, from `on_navigation`): before it about:blank (where Windows and macOS windows start, `net::guard`) and, in the Start Center's window with `VO_NET_PROBE`, the probe's `data:` page; then the page itself, again when it reloads (the Start Center does when its language changes), with any fragment (its `#create`), and its own `blob:` addresses (on Linux a save starts as one). Other pages, the page with another query and every other scheme are refused, with a line on stderr (`not loaded: …`), and a page can't open a bare window of its own: the suite's windows come from `VO.host.newWindow`, with the network rules, and `VO.go` reports a window it couldn't open rather than changing the page.

UI size (Settings on the Start Center): the web view's zoom, as Ctrl + in a browser, so canvases stay sharp and clicks land where they should. CSS `zoom` on the page was measured and does not work: at 200 % every click on Ledger's grid landed on B4 and menus opened at twice their place. The setting `vibeoffice.uiScale` keeps it (`settings.json` in the app's config folder, `%APPDATA%\work.vibeoffice.desktop\` or `~/.config/work.vibeoffice.desktop/`, where earlier versions kept it, is read while there is no `config.json`), and it is applied as each page starts loading (WebKitGTK resets the zoom on the first load).

Rules for growing it:

- Pages call `VO.host`, never Tauri, `launchQueue` or the File System Access API directly; both sides get each new method, the web side doing what the browser can (or reporting that it cannot).
- A page never names a file: the native side gives each window its document (the command line, Finder, the Open dialog, Recent Files by key), and the window's commands act on that document only; paths come out only for display (a Recent Files entry's folder).
- Commands are named `host_<verb>`, take named arguments and return JSON or raw bytes (`tauri::ipc::Response`, an `ArrayBuffer` in the page) for file contents; bytes go in as a raw body (`Uint8Array`), what they are in headers or a JSON head (`host_set_draft`), never as JSON arrays.
- App commands are allowed only for the app's own pages: Tauri refuses them from any other origin while the app has no ACL manifest.

Windows: as when the PWA is installed, the Start Center stays in its window and each app opens in one of its own (`VO.go` → `newWindow`). Started with files (a double-click, Open with), the app opens each in its app's window, without the Start Center. One VibeOffice runs per user session (`tauri-plugin-single-instance`, Tauri's own, registered first in `main.rs`): a second start (the Start menu, a double-click on a file, Open with) hands its arguments to the running one and exits. Its files open there, each in a window of its own (or the window that has it comes forward); with none, the Start Center comes to the front, or opens again if it was closed. What the installer asked for (`-pin-taskbar`, `-choose-defaults`) is done by the running one too. A page cannot simply `window.open`: WebView2 would make a bare popup without the host script or IPC, so the app's own new windows are made natively (`open_window` in `main.rs`, from an async command: a sync one deadlocks on Windows).

Linux has one VibeOffice per user and display too (`instance.rs`; the plugin would bring D-Bus): the first start holds a lock (`vibeoffice.<display>.lock`) and listens on a Unix socket beside it (`.sock`, mode 0600) in `$XDG_RUNTIME_DIR`, else `~/.vibeoffice` (`VIBEOFFICE_HOME` for tests); a later start sends its folder and arguments, which the first takes as a command line (existing files only), and exits. If the first doesn't answer within 5 s, the later one runs on its own. macOS has one instance by LaunchServices (`RunEvent::Opened`).

### Settings and storage

The pages keep nothing in the web view. Every window's script (`host::NO_STORAGE`, all frames) replaces `localStorage`, `sessionStorage`, `indexedDB`, `caches`, `BroadcastChannel`, `navigator.storage` and `navigator.serviceWorker` with getters that throw, so a use fails at once instead of quietly working until the app quits; the web views are incognito (WebView2 InPrivate, WKWebView's non-persistent data store), except on Linux, where Tauri 2.12 registers `tauri://` once per shared web context while an incognito WebKitGTK view gets a context of its own, so from the second window on no page would load. Recent Files, previews and drafts are `doc.rs`'s; the settings are `settings.rs`'s.

A setting is a string by key, as `localStorage` keeps them on the web: `VO.settings` (`get(key, default)`, `set(key, value)`, `remove(key)`, `on(fn)` for other windows' changes, `kept()`), in `common/look.js`, the first script of every page; `L.store`, the looks, the language, local fonts, the update question's day and Scout go through it. In the desktop app they live in `~/.vibeoffice/config.json`: keys `<app>.<name>` (`[a-z][a-z0-9]*`, a dot, 1–64 of `A-Za-z0-9_.-`), a value at most 1 MB, all of them 8 MB; changes are written together, 300 ms after the first, atomically, and as the app ends. Rust acts on two: `vibeoffice.uiScale` (0.5–4, every window's zoom at once) and `vibeoffice.look`, Paper while the user hasn't picked one. Each page gets them at the top of `common/look.js` (`window.__VO_SETTINGS__`; `settings::Served` wraps the built-in pages), so a page has them as they are when it loads, a reload too (the Start Center reloads when its language changes, after `kept()`); a build that loads the pages from the dev server (`--no-default-features`) puts them in each window's init script instead, as they were when the window was made. A page's `set` changes its own copy at once and sends it (`host_set_setting`, one at a time); the other windows get `vo-settings` `{ key, value }`, as the web's `storage` event. Nothing is migrated from the web view storage of earlier builds.

`node --test tests/desktop/host.test.mjs` checks the page side in Chromium: `tests/desktop/host-stub.js` plays the desktop host (run before each page with `tools/shot.mjs --pre`; this tab plays each new window; `?file=`, `?pick=`, `?recent=` and `?update=button|dialog` set the scene): a window made for a file opens it, with no browser storage or service worker and the build's version; the settings (served as a page starts, after a reload too, kept by the host, changed by another window), with the stub taking the storage APIs away as the app does and a static check that only the web's own code names them; Open File and File ▸ Open (here, or a new window); a dropped file; Recent Files from the native side; the close button's question, and Esc in it cancelling in every app; Ledger saving the CSV it opened and its recovered draft as CSV; a save the Mac sandbox refuses asking for the folder; `openUrl` refuses anything but VibeOffice's own addresses; Settings sets the UI size (the web points to the browser's zoom); the update policies; versions compare by date.

## Documents and windows

A window holds one document, and only the native side gives a window its file (`desktop/src/doc.rs`; the design in
todos/DESKTOP_REAL_SAVE.md, Phase 1):

- **Opening.** The files the app is started with, those another start hands over (Windows) and Finder's (macOS), the
  Open dialog's (`VO.host.open`: the Start Center's Open File, the apps' File ▸ Open) and Recent Files
  (`VO.host.openRecent`) each open in their app's window, made for them (`<app>/?doc`; the page asks
  `host_document`), or bring forward the window that already has the file (one Recent Files key per file:
  `store::key_of`, the sha256 of its real path, lower case on Windows). No page already running is ever given
  another document: it may have parsed something hostile. A window gets its document as it is made (`Docs::adopt`),
  or by reloading: File ▸ Open in an app window holds the file the user picked (`next`), the page asks "Save changes?"
  as when it closes (Cancel keeps everything), and `host_reopen` reloads the window to `<app>/?doc&t=<word>`; only the
  page with that one-time word is given the document (`Doc::token`), not the one it replaces, which runs on while the
  reload starts. `adopt` and `replace` refuse a file another window has, under the windows lock, so two opens of one
  file at once make one window and bring it forward.
- **Recent Files are the Start Center's.** It alone lists them (`host_recent`, previews), opens them by key
  (`host_open_recent`), removes them and discards their unsaved changes. In the apps, File ▸ Recent Files..., the
  task pane's link and the Document Recovery pane bring the Start Center forward on its Recent Files
  (`host_show_recent`, then `vo-show-recent` there); an app page holds no entry's key.
- **One document per window.** On the desktop (`VO.oneDocument`) the apps' New, templates and Open make windows
  (`VO.go`, `VO.openDialog`), a file dropped on a window with a document goes to a new window as bytes
  (`VO.elsewhere`), as does a document a page makes (Quire's New with content and mail merge's letters:
  `A.toWindow`, `VO.made`); File ▸ Close and Exit close the window; the Window menu lists the app's windows
  (`VO.windowItems`). Ledger logs a console error if a second workbook reaches a window anyway.
- **Recent Files** live in `~/.vibeoffice/recent_files/<yyyy-mm>/<key>.recent` (JSON: app, path, name, size, when,
  the draft's name and loss state; on macOS a security-scoped bookmark), with `.recent.preview` (JPEG, 512 KB at most)
  and `.recent.draft` (the unsaved version in the app's own format, 50 MB at most) beside it. No user's file name is
  in a path the app makes; only names of that form are read; each file is written to a temp file, synced and renamed
  over the old one. `VIBEOFFICE_HOME` puts the folder elsewhere (tests). The 30 newest entries without a draft are
  kept, every entry with one, and those open in a window. A window changes only its own entry: its draft
  (`host_set_draft`, a never-saved document getting an entry of its own), its preview, clearing its draft; Document
  Recovery's Discard and the Start Center's × refuse a document open in a window. Every window hears `vo-recent` when
  the list changes. On macOS the sandbox forgets a file's grant when the app quits: the bookmark, made while the grant
  holds, opens it again (entitlement `files.bookmarks.app-scope`).
- **Closing.** An app window's close button asks its page (`vo-close`); the page answers at once (`host_closing`),
  asks "Save changes?" for its document (`hooks.close`) and closes the window (`host_close_window`), or stays on
  Cancel. A page that doesn't answer in 10 s gets the native side's own question (close anyway? its unsaved changes
  stay in Recent Files). The Start Center and the licences page close at once.

Saving writes the window's file (Saving, below).

## Saving

The desktop app saves to real files (todos/DESKTOP_REAL_SAVE.md, Phase 2; `desktop/src/files.rs`, `doc.rs`; the pages'
side is `L.file` in `public/common/core.js`):

- **Save** (Ctrl+S) writes the window's own file in place, without asking, as Office does (`host_save`). A window
  without one (a new document, a dropped file, a page-made one), a read-only file (its title says "(Read-Only)"), or
  a type the app can't write (.doc, .xls) goes to Save As. Changes made while a save runs stay unsaved
  (`VO.changeCount`). A workbook's password to open is set in Ledger's Tools ▸ Options ▸ Security (the web's Save As
  ▸ Tools has it too).
- **Save As and every export** (PDF, web page, PNG, CSV, text, Send To) ask with the system's Save dialog
  (`host_choose`: in the folder of the window's file, else the last one this window used, else Documents; the app's
  types as filters, from Rust's own table, `types::KINDS`; the system asks before replacing a file), then write only
  what was picked there (`host_write_chosen`). A typed extension of another offered type switches to it; a name
  without one gets the chosen type's (on Linux, where the dialog doesn't add it, a replace is asked again for the
  final name). A Save As to a document type (Office formats, Quire's Markdown, Ledger's XML Spreadsheet) makes that
  file the window's own and its Recent Files entry; an export changes nothing about the document.
- **The write**: the type must be the window's app's and match the extension and the bytes' start (a ZIP or an
  encrypted Office file, `%PDF-`, `{\rtf`, the PNG signature, text); at most 512 MB; a link, folder or device in the
  file's place is refused; for Save, the file must be the version the window read or last wrote (`changed`
  otherwise). The version on disk is kept first (below); the bytes go to a temp file beside it
  (`.~vibeoffice-<hex>.tmp`), synced, then replace it: `ReplaceFileW` on Windows (keeping the file's permissions,
  attributes and streams, its Mark of the Web among them; a sharing violation is tried again a few times), a rename on
  Linux (the file's mode kept, the folder synced; in a folder the user can't write, or a file with another owner or
  other hard links, written in place, the old bytes put back if that fails), and on macOS, whose sandbox grants the
  file but not its folder, `replaceItemAtURL:` from the item replacement folder (NSDocument's safe save). A failed save
  leaves the file as it was. A copy saved from a file that came from the internet keeps its Mark of the Web (Windows,
  `files::carry_mark`). After the write the Recent Files entry is brought up to date and its draft goes.
- **When it fails** the page says why, with the next step and that the file hasn't changed (`L.file.write`):
  `changed` (Replace It, Save As…), `missing` (moved, renamed or deleted: Save As…, never made again silently),
  `denied` (read-only or no permission: Save As…; in the Mac's sandbox Allow Access… first, below), `busy` (another
  program holds it: Try Again, Save As…), `full`,
  `open-elsewhere` (the Save dialog picked a file another window has open: back to the dialog), and the rest with the
  system's words.
- **Allow Access… (macOS, both editions in the App Sandbox; `__VO_HOST__.sandboxed`)**: a save the sandbox refuses
  (`denied`) offers Allow Access…, Save As… and Cancel. Allow Access… (`host_allow_access`) shows the system's folder
  chooser at the file's folder, its message saying why; a folder that holds the file lets the app in until it quits
  (nothing is kept for after: the next run asks again), and the save runs again. Another folder is refused, naming
  the right one; Cancel in the chooser asks again. Type-checked, not yet run on a Mac.
- **Previous versions**: the version a window's first save replaces, and a file an export replaces, are kept in
  `~/.vibeoffice/previous_versions/<yyyy-mm>/<file's key>.<ms>.previous` (no user file name), 30 days, five a file,
  1 GB in all and no file's over half of it, none over 200 MB. Over 1 GB, space comes first from the file whose
  versions take the most, its oldest first, and a file's last version goes last (then the largest), so one large
  file saved often trims its own history, not every other file's (`store::prune_versions_within`). Pruned after each
  version kept and at start, when the temp files a crash left (an hour old) go too (`Store::sweep`). A file's key is
  its path: a renamed or moved file starts a new history. File ▸ Versions… (desktop only; the ribbon's File tab too)
  lists them; Open shows one in a new window as a document never saved.
- Web view downloads are refused (`on_download`), and the File System Access pickers are removed from every frame
  (`host::NO_FILE_PICKERS`), so nothing else writes a file.
- One write at a time per window (`L.file.write` queues them), and two writes to one file never overlap (`Docs::writing`).

## Versions and updates

A release is made with Actions ▸ Release (`release.yml`): the version's GitHub release as a draft, then Windows, the
macOS DMG and Linux one after another, each adding its download, and the release published once all three are in it.
A system that fails stops the ones after it and leaves the draft: its own workflow, run again on the same commit, adds
the missing file, then the draft is published by hand. The Mac App Store edition goes to App Store Connect from its own
workflow (`appstore.yml`).

A version is `v<yyyy>.<mm>.<dd>-<6 hex>`, the date (UTC) and commit of the tree it was built from, e.g. `v2026.10.11-d2134d`: `tools/current-tree-version.sh [commit]` prints it. `build.sh` stamps it into the app and the installer (`VO_VERSION`, shown as `VO.VERSION` on the Start Center and in About); the release is tagged with it, and the zip keeps one name. Windows wants numbers: the exes' file version is `2026.10.11`. A plain `cargo build` says `dev`.

`public/updater/current.json`, edited by hand (nothing updates it): per policy, per platform (`web`, `macos`, `windows`, `linux`), the version an edition must be at least. The Mac App Store edition never reads it (`VO.host.updates`): the store updates it, and it has no network; its Check for updates opens the update page as `macos_store`.

    { "show-update-button": { "windows": "v2026.10.11-000000", … },
      "show-update-dialog": { "windows": "v2026.10.11-000000", … } }

`public/updater/whatsnew.json`, also by hand: one list for every platform, version → plain text (line breaks kept).

    { "v2026.10.11-000000": "VibeOffice for Windows and Linux: …" }

When the Start Center opens, the desktop app reads `https://vibeoffice.work/updater/current.json` (the site sends `Access-Control-Allow-Origin: *`; nothing about the user or their files is sent). Older than `show-update-dialog`: the Update dialog, at most once a day (`vibeoffice.updateAsked`), listing the `whatsnew.json` entries newer than this version, newest first (by date; the dialog shows without them if that file cannot be read); older than `show-update-button`: the Update VibeOffice button in the Start Center. `whatsnew.json` is read only when an update is due. Update opens `https://vibeoffice.work/updater/?from_version=<version>&current_platform=<platform>` in the default browser: the update page (`public/updater/index.html`, in English), which sends people to the latest GitHub release (`releases/latest`: not a pre-release or a draft) and shows the steps for the `current_platform` it was opened with (`public/updater/updater.js`): Windows, macOS, Linux (no download yet: the source), the browser (reload), or `macos_store` (update from the App Store); without one, every system. So a release is published (both workflows) before `current.json` asks for it. Offline, or a file that is not this JSON: nothing is shown from it. Apart from `current.json`, a version two months old or more (`VO.monthsOld`: whole calendar months from its day; `VO.isOutdated`) shows Check for updates with its age ("This version is 5 months old."), which opens the same page, in the Mac App Store edition too; it waits up to 5 s for the check and gives way to the Update button when `current.json` asks, so the two never show together. The site and the installed PWA show it too (docs/suite.md). A `dev` build is newer than any dated version and never asks. To try both against the published file without asking anyone else: keep `show-update-dialog` older than `show-update-button` (`v2026.10.05` and `v2026.10.11`) and build exes stamped in between and before (`VO_VERSION=v2026.10.08-000000` → the button only; `v2026.10.01-000000` → the dialog), e.g. `VO_VERSION=… TAURI_CONFIG='{"version":"2026.10.8"}' cargo build --release`; the dialog shows once a day per profile (`vibeoffice.updateAsked`). `tests/desktop/host.test.mjs` checks the repo's `current.json` the same way.

## Network

The pages may reach `https://vibeoffice.work/` alone, for the update check (`updater/current.json` and `whatsnew.json`);
the app itself makes no connections (`desktop/src/net.rs`). Every other http, https, ws and wss request a window makes,
whatever makes it (a page, a picture or frame in a document, a worker), is refused before it leaves the machine:

- macOS: a WebKit content rule list (the format of Safari's content blockers): block `^https?://` and `^wss?://`, then
  let `^https://vibeoffice\.work/` through. WebKit compiles it at start (and keeps it compiled in the app's container);
  the Start Center opens only then.
- Windows: WebView2 hands every request of a window to `net::allowed` first (`AddWebResourceRequestedFilter` for `*`,
  workers included where WebView2 has `ICoreWebView2_22`); a refused one gets a 403 instead. WebSocket connections
  may not pass through that event (WebView2 documents it for HTTP requests), but no page opens one.
- Linux: no rules (no release).

It fails closed (`net::guard`). Each window starts on `about:blank` and goes to its page only once the rules are on
its web view, so no page of the app runs without them. If they can't be set, the window closes without loading its
page and a system alert says why; when it was the only window, the app quits (exit code 1). On macOS, if WebKit
refuses to compile the rules, no window opens at all.

The rules are checked from a page: with `VO_NET_PROBE=1`, the Start Center's window first loads a probe page of its
own (a `data:` URL with no policy of its own, so only the rules can refuse anything). It tries a picture from
github.com, a fetch of api.github.com and a WebSocket to echo.websocket.org, with a picture from vibeoffice.work as the
control, and prints `network probe: blocked` on stderr, or `REACHED` and what got through. Then the window goes on to
the Start Center. The macOS smoke run (*smoke* ticked) sets it and fails without `blocked`. On Windows a WebSocket may get through
WebView2's filter; the pages' own Content-Security-Policy refuses it there (todos/HARDENING.md, phase 1).

The app's own pages and IPC (`tauri://`, `ipc://`; `https://tauri.localhost` and `https://ipc.localhost` on Windows, a secure context, where WebView2 doesn't look up links' host names ahead) and
`data:` and `blob:` are not the network; a dev build also reaches its dev server (127.0.0.1). Our own links still open
in the default browser, which is not the app. The pages' Content-Security-Policy (docs/suite.md, Security) is the
same limit from the other side: `connect-src` names the site and the IPC sources (`ipc:`, `https://ipc.localhost`). The sandbox can't name a host: the DMG has the network entitlement
(`network.client`) for the update check, which is why the rules live in the web view; the App Store edition has none.
What the system and WebView2 do on their own (Edge's updates, certificate checks) is theirs.

WebRTC isn't a request the web view shows us, so neither the rules nor the policy cover it: `net::NO_WEBRTC`, run in
every frame before its scripts, removes `RTCPeerConnection`, `webkitRTCPeerConnection` and `RTCDataChannel`
(WebKitGTK has WebRTC off already). Permission requests (camera, microphone, location, notifications, …) are refused
natively (`on_permission_request` in `main.rs`), except Local Font Access (Settings ▸ Fonts on this computer) and
reading the clipboard (Paste), which go on as the engine handles them; on macOS the sandbox grants no device anyway.
On Windows, WebView2's own browser keys are off (`net::win::browser_keys_off`, before the page: F5 and Ctrl+R reload,
Ctrl+P, Ctrl+F, Ctrl+S, the zoom keys); editing keys and the apps' own shortcuts stay.

## Building

Windows (below); macOS: `desktop/macos/build.sh`, on a Mac (macOS, below).

Needs Rust with the `x86_64-pc-windows-gnu` target, mingw-w64 (`x86_64-w64-mingw32-gcc`, `windres`) and `zip`; builds on Linux, no Windows SDK.

    desktop/build.sh        # desktop/target/dist/: VibeOffice-Windows-x64.zip, holding VibeOfficeSetup.exe

`cd desktop && cargo build --release` builds the exe alone (`target/x86_64-pc-windows-gnu/release/`); `cargo build --release -p vibeoffice-setup` then builds the installer around it (`VO_PAYLOAD`: another folder holding `vibeoffice.exe` and `WebView2Loader.dll`). `cargo build --no-default-features` leaves out the embedded pages and loads `devUrl` (http://127.0.0.1:8760/, `tools/serve.py`), so page edits show on reload without rebuilding.

CI runs only by hand: Actions ▸ Windows ▸ Run workflow (pick the branch; tick *pre-release* if it is one). It runs `build.sh` on Ubuntu, keeps the installer as the run's artifact (`VibeOffice-Windows-x64`: GitHub zips it, so it downloads as the same zip) and publishes `VibeOffice-Windows-x64.zip` as a GitHub release tagged with the version (`v2026.10.11-d2134d`, at the commit built); running it again on the same commit replaces the release's file. The zip's name is the same in every release, so `releases/latest/download/VibeOffice-Windows-x64.zip` is always the newest. `public/updater/current.json` is not touched: raising the versions there is what makes installed copies offer the update.

## The installer

`VibeOfficeSetup.exe` (`desktop/setup/`), alone in the release zip, installs VibeOffice for the user who runs it, without administrator rights, as Chrome does. It carries the app (`vibeoffice.exe`, `WebView2Loader.dll`, `LICENSE.txt`) and writes it into `%LOCALAPPDATA%\VibeOffice`. What it writes is listed once, in `desktop/src/installation.rs`, which `vibeoffice.exe -uninstall` removes again:

- desktop and Start menu shortcuts (`VibeOffice.lnk`, with the AppUserModelID `work.vibeoffice.desktop`, which the app sets on its process too);
- under HKCU: Apps & features (`Uninstall\VibeOffice`: version, size, icon, `UninstallString` = `"…\vibeoffice.exe" -uninstall`), `App Paths`, `Applications\vibeoffice.exe`, and the file types `.docx` (Quire Document), `.xlsx` (Ledger Worksheet) and `.pptx` (Lectern Presentation): a ProgID each (`VibeOffice.docx`, … with `"…\vibeoffice.exe" "%1"` and the app's document icon), listed under the extension's `OpenWithProgids` and in the app's Default apps capabilities (`RegisteredApplications`). The extension's own default and `UserChoice` are Windows' and the user's: never written (UCPD protects UserChoice anyway);
- Explorer is told the file types changed (`SHChangeNotify`).

It writes the files first, then the registry (so Apps & features can always uninstall), then the shortcuts (one that can't be made is a warning, not a failed install), and starts the app outside any job object Windows may run installers in. Each step, and the app's first start (the pin request, Default apps), goes to `%TEMP%\VibeOfficeSetup.log`.

The dialog (`setup.rc`) says where it installs and offers two choices:

- **Pin VibeOffice to the taskbar** (ticked for a new install): the app, at its first start, asks Windows (`TaskbarManager.RequestPinCurrentAppAsync`), and Windows asks the user. Windows lets only the app itself ask, in front and with a Start menu entry, never an installer. The box shows only where Windows lets desktop apps ask: `TaskbarManager` with desktop app support, and no Limited Access Feature token needed (from KB5074105, builds 26100/26200.7705); elsewhere (Windows 10, older Windows 11) the dialog says how to pin it from the Start menu;
- **Open default app settings, in which you can set VibeOffice as default app by yourself.** Ticked unless the user has already chosen VibeOffice for `.docx`, `.xlsx` and `.pptx` (their `UserChoice`, or `UserChoiceLatest` on newer Windows 11; only read). Ticked, the app opens `ms-settings:defaultapps?registeredAppUser=VibeOffice`, VibeOffice's page in Default apps, after the pin question; if the app can't be started, the installer opens it. The file types are registered either way; which app opens them is the user's choice, made there.

The installer starts the app with the choices (`-pin-taskbar`, `-choose-defaults`); the app acts on them once its Start Center has loaded (`installed.rs`). Run again with this very app installed (its `vibeoffice.exe` byte for byte, whatever the version says), it offers **Reinstall**, which also repairs what is missing; another build gets **Update**. Files dropped on the installer open in the app. A running VibeOffice keeps running during an update: its files, which Windows won't let anyone replace, are renamed `*.old` and the app removes them at its next start. Until it has been closed entirely, a start goes to the running (old) one, which opens Default apps as asked. A DLL in the folder that is not the app's is deleted.

Uninstall (Apps & features, or `vibeoffice.exe -uninstall`) asks first; **Also delete VibeOffice's settings, Recent Files list and drafts** (unticked) also removes `%LOCALAPPDATA%\work.vibeoffice.desktop` (the WebView2 profile) and `%APPDATA%\work.vibeoffice.desktop`. With VibeOffice open, it asks the user to close it (Retry). It removes the shortcuts (and the taskbar pin's) that point to the app, VibeOffice's own keys, and only its own values in shared ones (`OpenWithProgids`, `RegisteredApplications`); the folder goes once the uninstaller has ended (a hidden `cmd.exe` from System32 waits for it).

The installer runs from Downloads, so it loads nothing from its own folder: it imports only Known DLLs and API sets (`build.sh` fails otherwise), calls `SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32)` before anything else, and gets comctl32 v6 from side-by-side (its manifest). Per user, the folder is writable by the user, as Chrome's per-user install is; integrity levels keep low-integrity and sandboxed processes out.

`python3 tests/desktop/setup-wine.py desktop/target/dist/VibeOfficeSetup.exe` runs it end to end in Wine (needs wine, Xvfb, xdotool, ImageMagick), in its own prefix in `./tmp/setup-wine` with screenshots of both dialogs: install from a folder holding only the installer (files, shortcuts, registry, no extension default), again with that version (no dialog), uninstall (all gone, another program's `OpenWithProgids` entry kept). Regenerate the document icons after changing an app's SVG with `python3 tools/desktop-icons.py` (`rsvg-convert` from librsvg2-bin, and ImageMagick).

## macOS: the DMG and the Mac App Store

One universal binary (Apple silicon and Intel, macOS 13 or later), bundled twice, the two editions told apart by
`VODistribution` in the bundle's `Info.plist`, which the signature seals (`host::app_store`):

- **direct**: `VibeOffice-macOS.dmg` in the GitHub release, signed with Developer ID, notarized and stapled. It checks
  for updates as Windows does (`updater/current.json`, `macos`).
- **appstore**: `VibeOffice-macOS-AppStore.pkg` for App Store Connect (TestFlight, then review), signed for the store
  with its provisioning profile. `edition: 'appstore'`: the Start Center never reads `current.json`, the store updates it
  (App Review guideline 2.4.5), so it has no network entitlement either. Two months old, it shows Check for updates,
  which opens the update page in the browser with `current_platform=macos_store`: update from the App Store.

Both run in the App Sandbox, so trying the DMG tries the store's edition too (`desktop/macos/*.entitlements`): the
files the user picks in the Open and Save dialogs or opens from Finder, their bookmarks (Recent Files after a restart),
printing, and in the DMG the network for the update check. Links go to the browser through
NSWorkspace, which the sandbox allows; no program is started.

The Mac's ways (`main.rs`):

- Files from Finder (a double-click, Open With, a drop on the Dock icon) arrive as `RunEvent::Opened`, at start or
  while the app runs, and each opens in a window of its own (Documents and windows, above); the sandbox lets the app
  read them, and their Recent Files entries keep a bookmark for after it quits. `Info.plist` lists .docx, .xlsx
  and .pptx (Quire Document, Ledger Worksheet, Lectern Presentation), as the Windows installer does: in Open With, never
  taking the default (`LSHandlerRank` Alternate); the system draws their icons from the app's.
- The app stays open, with its menu bar, after its last window closes; a click on the Dock icon or Window ▸ Start
  Center brings the Start Center back; Quit (⌘Q) ends it.
- The menu bar (`mac_menu`): the app menu (About, Services, Hide, Quit); File ▸ New Document, New Spreadsheet, New
  Presentation, Open…, Close Window; Edit with Cut, Copy, Paste and Select All (WKWebView gets ⌘X, ⌘C, ⌘V and ⌘A
  through them); View ▸ Full Screen; Window (Minimize, Zoom, Start Center, the open windows); Help ▸ VibeOffice Help.
  The File, Window and Help items work with no window open, as App Review checks (`mac_menu_event`): Open… and Help
  run the Open and Help of the window in front (the page hears `vo-menu`; not while a dialog of its own is open),
  with none of ours in front Open… is the system's dialog on its own and Help the project's page in the browser. They
  have no shortcuts: ⌘N and ⌘O stay the page's. No Undo or Redo there: they would run WebKit's own undo, not the
  document's; ⌘Z reaches the page, as in Safari. The app's commands are in the window's own menus; the menu bar
  doesn't mirror them yet (to do).
- ⌘ does what Ctrl does in every app, and on a Mac shortcuts show as ⌘, ⌥ and ⇧ (`ui.keyText`; Safari too).

**Building**, on a Mac with Xcode's command line tools and Rust's `aarch64-apple-darwin` and `x86_64-apple-darwin`:

    desktop/macos/build.sh        # desktop/target/dist/: VibeOffice.app, and the DMG and package it can sign

It builds both binaries, joins them (`lipo`) and bundles `VibeOffice.app` for one edition: `build.sh` (or
`build.sh dmg`) the direct one, `build.sh appstore` the Mac App Store's. By the identities in the keychain (found by
name; `MACOS_DEVELOPER_ID`, `MACOS_APPSTORE_ID`, `MACOS_INSTALLER_ID` choose others):

- `dmg`, none: the app signed ad hoc, to try on that Mac (and the CI smoke run);
- `dmg`, Developer ID Application: the DMG (the app and an Applications link), notarized with an App Store Connect API key
  (`ASC_KEY_ID`, `ASC_ISSUER_ID`, `ASC_KEY_PATH`, the .p8 file) and stapled. Its window shows the two side by side
  on a light sky with an arrow from the app to Applications and Scout under it, pointing the same way; the volume has
  the app's icon. dmgbuild (MIT; Python 3.9 or later, fetched into `target/dmgbuild` from
  `macos/dmg-requirements.txt`, pinned with hashes) writes the window's layout (`macos/dmg-settings.py`) without
  Finder, which can't be scripted reliably on CI machines. The background is `macos/dmg/background.png` and its `@2x`,
  joined into one HiDPI TIFF; `node tools/dmg-background.mjs` draws them in Chromium (Scout by his own rig,
  GestureRight) from the settings' window size and icon places, so run it after changing those (`--preview` adds
  stand-ins for the icons, in `tmp/dmg-preview.png`). Finder draws the icons' names in black on a background picture,
  in Dark Mode too, so the picture stays light under them;
- `appstore`, Apple Distribution and Mac Installer Distribution, with `MACOS_PROVISION_PROFILE` (the profile's path):
  the App Store package; without them it fails.

The bundle's version is the date (`2026.10.15`), its build number the commit's UTC time (`2026.1015.641`), which rises
with every commit, as App Store Connect requires.

From Linux the crate only type-checks for macOS: linking needs Apple's SDK, which its licence allows on Apple hardware
only, so nothing here links, bundles or runs. After `rustup target add aarch64-apple-darwin x86_64-apple-darwin`
(clang compiles the one Objective-C file of `objc2` for the check):

    cd desktop && CC_aarch64_apple_darwin=clang CFLAGS_aarch64_apple_darwin=--target=arm64-apple-macos11 \
      cargo check --release --target aarch64-apple-darwin

**CI**, run by hand: Actions ▸ macOS ▸ Run workflow (`.github/workflows/macos.yml`, on `macos-15`; also run by
`release.yml` with Windows and Linux) builds the DMG; with
*smoke* ticked (off by default) it then opens the app alone and with a .docx, a .csv and a .md (`open -a`, as Finder does) and keeps a screenshot after each
(artifact `VibeOffice-macOS-smoke`, with the app's log), and one of the DMG's window opened in Finder when there is a
DMG (`0-dmg.png`). With the signing secrets it adds the DMG to the version's GitHub release (the tag Windows uses, so
one release holds both; `desktop/release-notes.md`). Without them the app is signed ad hoc and nothing is released:
artifact `VibeOffice-macOS-unsigned`, which starts after `xattr -dr com.apple.quarantine VibeOffice.app`. The App
Store edition has its own: Actions ▸ Mac App Store ▸ Run workflow (`.github/workflows/appstore.yml`) builds the
package, keeps it as the run's artifact (`VibeOffice-macOS-AppStore`, for Transporter) and, with *upload* ticked (the
default), uploads it to App Store Connect (`altool`); it fails without the App Store secrets. App Store Connect takes
each build number once, and the build number is the commit's time: one upload a commit.

The signing material is in the secrets of the `release` environment (Settings ▸ Environments; limit it to `main`, so
runs from other branches and from forks never see them):

| Secret | What |
|---|---|
| `MACOS_DEVELOPER_ID_P12`, `MACOS_DEVELOPER_ID_PASSWORD` | the Developer ID Application certificate and key, a .p12 in base64 (only the team's Account Holder can create it) |
| `MACOS_APPSTORE_P12`, `MACOS_APPSTORE_PASSWORD` | the Apple Distribution and Mac Installer Distribution certificates and keys, one .p12 in base64 |
| `MACOS_PROVISION_PROFILE` | the Mac App Store Connect provisioning profile for `work.vibeoffice.desktop`, in base64 |
| `APPLE_TEAM_ID` | the team ID (otherwise read from the identity's name) |
| `ASC_KEY_ID`, `ASC_ISSUER_ID`, `ASC_KEY_P8` | an App Store Connect API team key (the .p8 file's text) with the lowest role that works (Developer; App Manager if the upload is refused): notarizing and uploading |

Before the first upload: the app's record in App Store Connect (bundle ID `work.vibeoffice.desktop`) and the answers to
its export compliance questions (`crypto.js` encrypts and opens password-protected Office files with standard AES and
SHA), after which `ITSAppUsesNonExemptEncryption` goes into `Info.plist`. Submitting for review waits for real saving
to be tried on a Mac (todos/DESKTOP_REAL_SAVE.md, Phase 2: built, run on Linux only).

## Linux

One download, `VibeOffice-Linux-x86_64.AppImage` (`desktop/linux/build.sh`), for x86_64 systems with WebKitGTK 4.1:
Ubuntu 22.04 to 26.04, Debian 12 and 13, Fedora. Not sandboxed: no Flatpak, Snap or `.deb` (on macOS both editions
stay in the App Sandbox). Debian 10 and 11 are left out: they have no WebKitGTK 4.1, which needs libsoup 3 (Debian has
it from 12 on).

- **The system's WebKitGTK.** The app links the system's WebKitGTK 4.1, GTK 3 (LGPL-2.1-or-later) and libsoup 3
  dynamically, which the LGPL allows (AGENTS.md, Code from other projects); the AppImage carries none of them. Built on
  Ubuntu 22.04 (glibc 2.35, the oldest of these systems) it starts on every newer one; built on a newer system it needs
  that system's glibc (on Debian 13, 2.39).
- **`AppRun`**, the AppImage's entry point, asks the system's loader which of the app's libraries are missing
  (`LD_TRACE_LOADED_OBJECTS`, what `ldd` does; plain `sh`, as a bare system may lack awk and sed). When one is, it says
  how to install WebKitGTK 4.1 with the system's own command, from `/etc/os-release` (apt, dnf, zypper, pacman), in a
  dialog (zenity, kdialog or xmessage) and on stderr, in English; otherwise the app takes its place.
- **The AppImage** is the AppImage project's runtime (type2-runtime 20251108, a static program, pinned by its
  checksum) followed by a squashfs (zstd) of the folder: `AppRun`, `usr/bin/vibeoffice`, `vibeoffice.desktop` (.docx,
  .xlsx and .pptx for an AppImage tool's Open With) and the icon (`public/icons/icon.svg`). The runtime mounts it with
  FUSE (`fusermount3`); without FUSE, `--appimage-extract-and-run` or `APPIMAGE_EXTRACT_AND_RUN=1`. Its licences,
  libfuse's (LGPL-2.1, inside the runtime) among them, are in the open source licenses
  (`desktop/licenses/appimage-runtime.txt`).
- **CI** (`.github/workflows/linux.yml`, run by hand): builds on `ubuntu-22.04` and puts the AppImage into the
  version's release. With *smoke* ticked (off by default) it first runs the smoke run there (`desktop/linux/smoke.sh`:
  the Start Center, then a .docx, a CSV and a .md handed to the running app, each found by its window's title, a
  screenshot after each), then Ubuntu 24.04 and 26.04, Debian 12 and 13 and Fedora, each in a container
  (`desktop/linux/on-system.sh`): without WebKitGTK the AppImage must not start and must name that system's command,
  with it the smoke run; the release waits for them all.

Building here needs the development files of WebKitGTK 4.1, GTK 3 and libsoup 3 (`libwebkit2gtk-4.1-dev`,
`libgtk-3-dev`, `libsoup-3.0-dev`), and squashfs-tools for the AppImage:

    cd desktop && cargo build --release --target x86_64-unknown-linux-gnu
    desktop/linux/build.sh

Without root to install them, the development packages can be unpacked into a folder instead (`apt-get download` those
three and what they depend on, `dpkg -x` each into `tmp/linuxdev/root`), with pkg-config pointed there; the runtime
libraries are the system's:

    R=$PWD/tmp/linuxdev/root
    export PKG_CONFIG_SYSROOT_DIR=$R PKG_CONFIG_PATH=$R/usr/lib/x86_64-linux-gnu/pkgconfig:$R/usr/share/pkgconfig

Run it with a file to open: `target/x86_64-unknown-linux-gnu/release/vibeoffice ~/Budget.csv`, or the AppImage the
same way; `desktop/linux/smoke.sh <AppImage> <folder>` on a display (or with Xvfb) does what the CI does.

## Size

| Build (2026-10-11, Tauri 2.12.3, rustc 1.95) | Size |
|---|---|
| `vibeoffice.exe`, release | 7.4 MB (7.39) |
| of which `public/` (11 MB, brotli quality 9) | about 4.1 MB |
| of which the three document icons | 315 KB |
| `WebView2Loader.dll`, installed beside the exe (GNU target) | 160 KB |
| `VibeOfficeSetup.exe`: the above, `LICENSE.txt` and its own 425 KB (code 283 KB, icon and dialog 57 KB) | 8.0 MB |
| `VibeOffice-Windows-x64.zip` | 5.8 MB |
| `vibeoffice`, Linux release | 7.2 MB |

Licenses: `tools/oss-licenses.py` lists every crate compiled into the app and its installer, on Windows, macOS and Linux (Linux's native GTK and WebKitGTK libraries are not crates and are not covered: Linux above; 168 today: the workspace's members and their normal dependencies all the way down, with features as the build resolves them, from `cargo tree`; not what only build scripts or Tauri's procedural macros use, such as `cssparser`, `selectors` and `dtoa-short`, MPL-2.0, which never reach the binary), with the license files from each package, and the DLL installed beside the exe (`WebView2Loader.dll`), in `public/oss-licenses.html`; run it after changing a dependency. It stops if a shipped crate's licence needs copyleft; `--check` (also in `node --test tests/licenses.test.mjs`) checks the page against the crates that ship.

No exe packers (UPX): antivirus flags packed exes. The largest pages are the thesaurus (`common/dict/en.thes`, 2 MB) and the two word lists (650 KB each).

## Results

- Linux (Debian 13, WebKitGTK 2.52.6, Xfce): Start Center, all three apps; `vibeoffice Budget.csv` opens it in Ledger and `Notes.md` in Quire, the title bar follows the document; Recent Files and the look are kept between runs. With no network at all (`unshare -rn`), a fresh profile: everything loads, in Paper; the GitHub button goes to `xdg-open`; UI size 150 % and 200 % scale the whole UI, sharp, and are back after a restart; the Start Center keeps Settings reachable in its narrow layout. Started with `Budget.csv`, the Start Center stays and Ledger opens in a window of its own; Quire from the Start Center opens a third window, and its File ▸ Exit closes only that one.
- Linux (2026-10-17), settings and one instance: a fresh `~/.vibeoffice` starts in Paper at the 125 % an old `settings.json` held; Français in Settings writes `config.json` (folder 0700) and the Start Center reloads in French; Quire opens in a second window in French; Paper 2016 picked in the Start Center turns the open Quire window to the ribbon. `vibeoffice Budget.csv` while it runs hands the file over and exits (0.45 s), Ledger opens it in the running app; with the Start Center closed, a start without files opens it again; once every window is closed the process ends, and the next start takes over the old socket.
- Linux AppImage (2026-10-18, built here on Debian 13, 7.2 MB): `smoke.sh` on Xvfb, mounted with FUSE: the Start
  Center, then `linecases.docx`, `Budget.csv` and `Notes.md` handed over by second starts, each in its window, and
  nothing left running or mounted after. `AppRun` with a library missing (a stand-in app): the zenity dialog and the
  message on stderr, exit 1; the command for Fedora, Ubuntu, Linux Mint, openSUSE and the general advice for Gentoo.
  Not yet run in CI (`linux.yml`: Ubuntu 22.04, the other systems in containers).
- macOS (2026-10-15, the CI's first run, macos-15): builds universal (arm64 and x86_64, 15.3 MB), signed ad hoc; the smoke run opens the Start Center, then `linecases.docx` in Quire, `Budget.csv` in Ledger and `Notes.md` in Quire, each in a window of its own in the one running process, under the App Sandbox with no denials, and quits cleanly. Not yet used by a person on a Mac (menus, ⌘ keys, saving, Browse, links, the Dock). The network rules (2026-10-16; failing closed, with the probe, the same day): the first macOS run with them (2026-10-17) hung as the Start Center opened, on both systems' path: `net::guard` went on to the page from inside `with_webview`, which holds the web view's lock that `navigate` takes. Fixed; with `guard` compiled for Linux for the check, the probe (which reports `REACHED` there, with no rules), an app's window and refusing either run. Not yet run on macOS or Windows.
- Not yet run against the published site: the update check (the site has no `updater/current.json` yet).
- Installer (2026-10-11), Wine 10.0, `tests/desktop/setup-wine.py`: passes. The dialog shows the pinning hint (Wine has no TaskbarManager); install writes the three files, both shortcuts and the HKCU values; the same version again shows no dialog; uninstall (Enter) leaves no folder, shortcut or value of VibeOffice's and keeps another program's. `cargo test` (Windows target, under Wine): 5 passed. Wine cannot show the app itself (no WebView2), the pin prompt or Default apps.
- `cargo test`, Windows target under Wine 10.0 (2026-10-18, a prefix of its own in `tmp/`): 32 passed, after a fix
  the run found: a save's previous version was never kept on Windows, as `FlushFileBuffers` refuses a handle opened
  only to read (`Store::keep_previous`, "Access denied", printed on stderr alone).
- Windows: built, not yet run. To check, as a user without administrator rights: install (the dialog at 100 % and 200 %), the pin prompt on an up-to-date Windows 11 and the hint on Windows 10, Default apps opening at VibeOffice's page, a double-clicked .docx afterwards; Open with lists VibeOffice for .docx, .xlsx and .pptx with their document icons; an update while the app runs; Uninstall from Apps & features; all three apps, saving, drops.

## Known limits / to do

- Linux: the AppImage adds no menu entry or file types by itself (AppImage tools such as Gear Lever do, from its
  `vibeoffice.desktop`); `AppRun`'s message is in English; x86_64 only.
- GNU builds import `WebView2Loader.dll`, installed beside the exe (Microsoft's static loader is MSVC-only).
- Saving (Saving, above) has run on Linux only: Windows' `ReplaceFileW` and macOS's `replaceItemAtURL:` are type-checked, not yet run.
- macOS: the menu bar has the suite's own items only (New, Open…, Start Center, Help); the apps' commands stay in the
  window's menus. To do before App Review if it asks, or anyway: each app window sends its menus to the host, which
  draws them in the menu bar for the window in front (with Undo and Redo going to the document's own undo, and ⌘
  shortcuts the menu catches running in the page), and the window's own menu bar hides on a Mac.
- macOS: the menu bar is in English whatever the UI's language; the system draws the document icons (the app's icon on a page) rather than Quire's, Ledger's and Lectern's.
- Windows and macOS: incognito web views, `https://tauri.localhost` and the settings served with `look.js` are type-checked, not yet run there (2026-10-17).
- Linux: the web views aren't incognito (Settings and storage, above).
- The Start Center's wording is the web's ("Office suite in your browser", "Recent files are kept in this browser only").
- Code signing (SmartScreen warns about unsigned installers).
- Installing for all users (UAC, Program Files, HKLM), where only administrators can change the app's folder; today it installs per user only.
