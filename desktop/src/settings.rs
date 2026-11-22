//! The settings: `VO.settings` in the pages (public/common/look.js), strings by key as the web keeps them in
//! localStorage, here in ~/.vibeoffice/config.json (store.rs). The pages have no browser storage in the desktop app
//! (host::NO_STORAGE): each gets the settings at the top of its first script, common/look.js (`Served`), so a page
//! that reloads has them as they are now, and every later change another window makes (`vo-settings`).
//!
//! The values are the pages' own data, but for two: `vibeoffice.uiScale`, the zoom of every window (0.5 to 4), and
//! `vibeoffice.look`, which is Paper while the user hasn't picked one.

use std::{
    borrow::Cow,
    collections::BTreeMap,
    ops::RangeInclusive,
    path::Path,
    sync::{Mutex, OnceLock},
    time::Duration,
};

use tauri::{
    utils::assets::{AssetKey, AssetsIter, CspHash},
    AppHandle, Assets, Manager, State, WebviewWindow, Wry,
};

use crate::store::Store;

pub const UI_SCALE: &str = "vibeoffice.uiScale";
/// The lite edition's settings until the user changes them, for its small screen (docs/desktop.md, Lite edition):
/// the Classic look's menus at 75 % (also `ui_scale`), each app without toolbars (none that shows by itself
/// either), rulers or task pane, Quire in Web Layout; nothing checked as you type (Quire's spelling and grammar went
/// over a whole document for minutes on the device's CPU). Each app's options merge with its own defaults.
#[cfg(feature = "lite")]
const LITE: &[(&str, &str)] = &[
    (LOOK, "classic"),
    (UI_SCALE, "0.75"),
    (
        "quire.opts",
        r#"{"taskOpen":false,"startupPane":false,"ruler":false,"vruler":false,"view":"web","spell":false,"toolbars":{"standard":false,"formatting":false,"drawing":false,"tables":false,"reviewing":false,"picture":false,"outlining":false,"headerFooter":false,"mailMerge":false,"wordart":false}}"#,
    ),
    (
        "ledger.opts",
        r#"{"taskOpen":false,"toolbars":{"standard":false,"formatting":false}}"#,
    ),
    (
        "lectern.opts",
        r#"{"taskOpen":false,"startupPane":false,"layoutPaneOnNew":false,"leftOpen":false,"notesH":24,"spell":false,"toolbars":{"standard":false,"formatting":false,"drawing":false}}"#,
    ),
];
const SCALES: RangeInclusive<f64> = 0.5..=4.0;
const LOOK: &str = "vibeoffice.look";
const MAX_VALUE: usize = 1 << 20;
const MAX_ALL: usize = 8 << 20;
/// Changes within this time are written together.
const DELAY: Duration = Duration::from_millis(300);

pub struct Settings {
    store: Store,
    /// the values, and whether a write is due
    values: Mutex<(BTreeMap<String, String>, bool)>,
    /// held while config.json is written, so writes go out in the order their values were taken
    writing: Mutex<()>,
}

/// `vibeoffice.look`, `quire.opts`: an app (or `vibeoffice`), a dot, a name.
fn is_key(key: &str) -> bool {
    let Some((app, name)) = key.split_once('.') else {
        return false;
    };
    app.starts_with(|c: char| c.is_ascii_lowercase())
        && app
            .chars()
            .all(|c| c.is_ascii_lowercase() || c.is_ascii_digit())
        && (1..=64).contains(&name.len())
        && name
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || "_.-".contains(c))
}

fn scale_of(value: &str) -> Option<f64> {
    value.parse().ok().filter(|s| SCALES.contains(s))
}

fn size(values: &BTreeMap<String, String>) -> usize {
    values.iter().map(|(k, v)| k.len() + v.len()).sum()
}

impl Settings {
    /// config.json, or when there is none the UI size an earlier version kept in `old` (settings.json in the
    /// app's config folder). A broken file, a key or value out of bounds: left out.
    pub fn load(store: Store, old: Option<&Path>) -> Settings {
        let mut values = BTreeMap::new();
        match store.read_config(MAX_ALL as u64 * 2) {
            Ok(bytes) => {
                let all: serde_json::Map<String, serde_json::Value> =
                    serde_json::from_slice(&bytes).unwrap_or_default();
                for (key, value) in all {
                    if let serde_json::Value::String(v) = value {
                        if is_key(&key)
                            && v.len() <= MAX_VALUE
                            && size(&values) + key.len() + v.len() <= MAX_ALL
                        {
                            values.insert(key, v);
                        }
                    }
                }
            }
            Err(_) => {
                #[derive(serde::Deserialize)]
                #[serde(rename_all = "camelCase")]
                struct Old {
                    ui_scale: f64,
                }
                if let Some(old) = old
                    .and_then(|p| std::fs::read(p).ok())
                    .and_then(|b| serde_json::from_slice::<Old>(&b).ok())
                    .filter(|o| SCALES.contains(&o.ui_scale))
                {
                    values.insert(UI_SCALE.into(), old.ui_scale.to_string());
                }
            }
        }
        Settings {
            store,
            values: Mutex::new((values, false)),
            writing: Mutex::new(()),
        }
    }

    pub fn ui_scale(&self) -> f64 {
        self.values
            .lock()
            .unwrap()
            .0
            .get(UI_SCALE)
            .and_then(|v| scale_of(v))
            .unwrap_or(if cfg!(feature = "lite") { 0.75 } else { 1.0 })
    }

    /// Every setting as a JSON object, with the desktop app's default look (the lite edition's defaults: `LITE`).
    pub fn json(&self) -> String {
        let mut values = self.values.lock().unwrap().0.clone();
        #[cfg(feature = "lite")]
        for (key, value) in LITE {
            values
                .entry((*key).into())
                .or_insert_with(|| (*value).into());
        }
        values.entry(LOOK.into()).or_insert_with(|| "paper".into());
        serde_json::to_string(&values).unwrap_or_else(|_| "{}".into())
    }

    /// Keeps `value` for `key` (None removes it); true when a write has to be planned.
    fn set(&self, key: &str, value: Option<String>) -> Result<bool, String> {
        if !is_key(key) {
            return Err(format!("not a setting: {key}"));
        }
        if key == UI_SCALE && value.as_deref().is_some_and(|v| scale_of(v).is_none()) {
            return Err(format!(
                "the UI size is from {} to {}",
                SCALES.start(),
                SCALES.end()
            ));
        }
        let mut guard = self.values.lock().unwrap();
        let (values, due) = &mut *guard;
        match value {
            Some(v) => {
                let before = values.get(key).map_or(0, |o| key.len() + o.len());
                if v.len() > MAX_VALUE || size(values) - before + key.len() + v.len() > MAX_ALL {
                    return Err(format!("too large to keep: {key}"));
                }
                values.insert(key.into(), v);
            }
            None => {
                values.remove(key);
            }
        }
        Ok(!std::mem::replace(due, true))
    }

    /// Writes config.json if a change is waiting (after DELAY, and as the app ends).
    pub fn flush(&self) {
        let _writing = self.writing.lock().unwrap();
        let json = {
            let mut guard = self.values.lock().unwrap();
            if !std::mem::replace(&mut guard.1, false) {
                return;
            }
            serde_json::to_vec_pretty(&guard.0)
        };
        if let Err(e) = json
            .map_err(std::io::Error::other)
            .and_then(|j| self.store.write_config(&j))
        {
            eprintln!("config.json not written: {e}");
        }
    }
}

/// A page keeps a setting (VO.settings.set; None removes it). The UI size applies to every window at once; the
/// other windows hear of the change (`vo-settings`), and config.json is written soon after.
#[tauri::command]
pub fn host_set_setting(
    app: AppHandle,
    window: WebviewWindow,
    settings: State<Settings>,
    key: String,
    value: Option<String>,
) -> Result<(), String> {
    if settings.set(&key, value.clone())? {
        let app = app.clone();
        std::thread::spawn(move || {
            std::thread::sleep(DELAY);
            app.state::<Settings>().flush();
        });
    }
    if key == UI_SCALE {
        let scale = settings.ui_scale();
        for w in app.webview_windows().values() {
            let _ = w.set_zoom(scale);
        }
    }
    let detail = serde_json::json!({ "key": key, "value": value });
    for (label, w) in app.webview_windows() {
        if label != window.label() {
            let _ = w.eval(format!(
                "window.dispatchEvent(new CustomEvent('vo-settings', {{ detail: {detail} }}));"
            ));
        }
    }
    Ok(())
}

/// The pages as built into the app, but for common/look.js, the first script of every page: the settings as
/// they are now come at its top (`window.__VO_SETTINGS__`).
pub struct Served {
    pages: Option<Box<dyn Assets<Wry>>>,
    app: OnceLock<AppHandle>,
}

/// Puts `Served` in front of the context's pages.
pub fn serve(context: &mut tauri::Context<Wry>) {
    let none = Served {
        pages: None,
        app: OnceLock::new(),
    };
    let pages = context.set_assets(Box::new(none));
    context.set_assets(Box::new(Served {
        pages: Some(pages),
        app: OnceLock::new(),
    }));
}

impl Assets<Wry> for Served {
    fn setup(&self, app: &tauri::App<Wry>) {
        let _ = self.app.set(app.handle().clone());
        if let Some(pages) = &self.pages {
            pages.setup(app);
        }
    }

    fn get(&self, key: &AssetKey) -> Option<Cow<'_, [u8]>> {
        let file = self.pages.as_ref()?.get(key)?;
        if key.as_ref().trim_start_matches('/') != "common/look.js" {
            return Some(file);
        }
        let Some(settings) = self.app.get().and_then(|a| a.try_state::<Settings>()) else {
            return Some(file);
        };
        let mut out = script(&settings).into_bytes();
        out.extend_from_slice(&file);
        Some(Cow::Owned(out))
    }

    fn iter(&self) -> Box<AssetsIter<'_>> {
        match &self.pages {
            Some(pages) => pages.iter(),
            None => Box::new(std::iter::empty()),
        }
    }

    fn csp_hashes(&self, html_path: &AssetKey) -> Box<dyn Iterator<Item = CspHash<'_>> + '_> {
        match &self.pages {
            Some(pages) => pages.csp_hashes(html_path),
            None => Box::new(std::iter::empty()),
        }
    }
}

/// What a page is given before common/look.js reads it. A build that loads the pages from the dev server
/// (`--no-default-features`) gives it in each window's init script instead, as it was when the window was made.
pub fn script(settings: &Settings) -> String {
    format!("window.__VO_SETTINGS__ = {};\n", settings.json())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn temp(name: &str) -> std::path::PathBuf {
        let dir = std::env::temp_dir().join(format!("vo-settings-{name}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        dir
    }

    #[test]
    fn keys_are_an_app_a_dot_and_a_name() {
        for good in [
            "vibeoffice.look",
            "quire.opts",
            "scout.lastGreetedDay",
            "ledger.acList",
            "a1.b_c-d.e",
        ] {
            assert!(is_key(good), "{good}");
        }
        for bad in [
            "",
            "look",
            "Quire.opts",
            "1x.a",
            "quire.",
            "quire.a b",
            "quire.é",
            "quire/../x.a",
            &format!("a.{}", "x".repeat(65)),
        ] {
            assert!(!is_key(bad), "{bad}");
        }
    }

    #[test]
    fn settings_are_kept_and_read_back() {
        let root = temp("kept");
        let settings = Settings::load(Store::new(root.clone()), None);
        assert!(settings
            .set("quire.opts", Some("{\"a\":1}".into()))
            .unwrap());
        assert!(
            !settings.set("vibeoffice.lang", Some("fr".into())).unwrap(),
            "one write for both"
        );
        settings.set("vibeoffice.lang", None).unwrap();
        settings.set(UI_SCALE, Some("1.5".into())).unwrap();
        settings.flush();
        let again = Settings::load(Store::new(root.clone()), None);
        let json: serde_json::Value = serde_json::from_str(&again.json()).unwrap();
        assert_eq!(json["quire.opts"], "{\"a\":1}");
        assert_eq!(
            json["vibeoffice.look"],
            if cfg!(feature = "lite") {
                "classic"
            } else {
                "paper"
            },
            "the desktop app's look until one is picked"
        );
        assert!(json.get("vibeoffice.lang").is_none());
        assert_eq!(again.ui_scale(), 1.5);
        let _ = std::fs::remove_dir_all(root);
    }

    #[test]
    fn what_is_out_of_bounds_is_refused() {
        let settings = Settings::load(Store::new(temp("bounds")), None);
        assert!(settings.set("look", Some("x".into())).is_err());
        assert!(settings.set(UI_SCALE, Some("9".into())).is_err());
        assert!(settings.set(UI_SCALE, Some("big".into())).is_err());
        assert!(settings
            .set("quire.big", Some("x".repeat(MAX_VALUE + 1)))
            .is_err());
        for i in 0..7 {
            settings
                .set(&format!("quire.k{i}"), Some("x".repeat(MAX_VALUE)))
                .unwrap();
        }
        assert!(
            settings
                .set("quire.k7", Some("x".repeat(MAX_VALUE)))
                .is_err(),
            "8 MB in all"
        );
        settings
            .set("quire.k6", Some("y".repeat(MAX_VALUE)))
            .unwrap(); // replacing one is fine
    }

    #[test]
    fn a_broken_file_or_entry_is_left_out_and_the_old_ui_size_read_once() {
        let root = temp("broken");
        std::fs::create_dir_all(&root).unwrap();
        let old = root.join("settings.json");
        std::fs::write(&old, r#"{"uiScale":2}"#).unwrap();
        assert_eq!(
            Settings::load(Store::new(root.clone()), Some(&old)).ui_scale(),
            2.0
        );
        std::fs::write(
            root.join("config.json"),
            r#"{"vibeoffice.look":"classic","Bad":"x","quire.n":3}"#,
        )
        .unwrap();
        let s = Settings::load(Store::new(root.clone()), Some(&old));
        #[cfg(not(feature = "lite"))]
        assert_eq!(s.json(), r#"{"vibeoffice.look":"classic"}"#);
        let json: serde_json::Value = serde_json::from_str(&s.json()).unwrap();
        assert_eq!(json["vibeoffice.look"], "classic");
        assert!(json.get("Bad").is_none() && json.get("quire.n").is_none());
        assert_eq!(
            s.ui_scale(),
            if cfg!(feature = "lite") { 0.75 } else { 1.0 },
            "config.json is there: settings.json is no longer read"
        );
        std::fs::write(root.join("config.json"), "{not json").unwrap();
        #[cfg(not(feature = "lite"))]
        assert_eq!(
            Settings::load(Store::new(root.clone()), None).json(),
            r#"{"vibeoffice.look":"paper"}"#
        );
        #[cfg(feature = "lite")]
        assert_eq!(
            Settings::load(Store::new(root.clone()), None).json(),
            Settings::load(Store::new(temp("lite-empty")), None).json(),
            "a broken file is as none"
        );
        let _ = std::fs::remove_dir_all(root);
    }

    #[cfg(feature = "lite")]
    #[test]
    fn the_lite_defaults_hold_until_the_user_changes_them() {
        let root = temp("lite");
        let settings = Settings::load(Store::new(root.clone()), None);
        let json: serde_json::Value = serde_json::from_str(&settings.json()).unwrap();
        assert_eq!(json["vibeoffice.look"], "classic");
        assert_eq!(json[UI_SCALE], "0.75");
        let quire: serde_json::Value =
            serde_json::from_str(json["quire.opts"].as_str().unwrap()).unwrap();
        assert_eq!(quire["view"], "web");
        assert_eq!(settings.ui_scale(), 0.75);
        settings.set(UI_SCALE, Some("1".into())).unwrap();
        settings.set("quire.opts", Some("{}".into())).unwrap();
        let json: serde_json::Value = serde_json::from_str(&settings.json()).unwrap();
        assert_eq!(
            (json[UI_SCALE].as_str(), json["quire.opts"].as_str()),
            (Some("1"), Some("{}"))
        );
        assert_eq!(settings.ui_scale(), 1.0);
        settings.flush();
        let written = std::fs::read_to_string(root.join("config.json")).unwrap();
        assert!(
            !written.contains("ledger.opts"),
            "defaults aren't written: {written}"
        );
        let _ = std::fs::remove_dir_all(&root);
    }
}
