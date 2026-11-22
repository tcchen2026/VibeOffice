//! The system's dialogs, on the window that asks: Open, Save and a yes-or-no question. Windows and macOS through
//! rfd (`IFileDialog`, NSOpenPanel as a sheet), whose async dialogs leave the main thread free while the user picks;
//! Linux through GTK's own native chooser on Tauri's GTK thread, parented to the window. rfd's GTK backend makes an
//! unparented chooser on a thread of its own, so it isn't used there.

use std::path::PathBuf;

use tauri::WebviewWindow;

/// One line of a dialog's file type list: its label and the extensions it shows.
pub struct Filter {
    pub name: String,
    pub exts: Vec<String>,
}

impl Filter {
    pub fn new(name: String, exts: &[&str]) -> Filter {
        Filter {
            name,
            exts: exts.iter().map(|e| e.to_string()).collect(),
        }
    }
}

/// The Open dialog: the files chosen, none when cancelled.
#[cfg(any(windows, target_os = "macos"))]
pub async fn open(window: &WebviewWindow, title: &str, filters: Vec<Filter>) -> Vec<PathBuf> {
    pick(
        rfd::AsyncFileDialog::new().set_parent(window),
        title,
        filters,
    )
    .await
}

/// macOS: the Open dialog on no window, the menu bar's File ▸ Open… with none of ours in front (main.rs).
#[cfg(target_os = "macos")]
pub async fn open_alone(title: &str, filters: Vec<Filter>) -> Vec<PathBuf> {
    pick(rfd::AsyncFileDialog::new(), title, filters).await
}

#[cfg(any(windows, target_os = "macos"))]
async fn pick(d: rfd::AsyncFileDialog, title: &str, filters: Vec<Filter>) -> Vec<PathBuf> {
    let mut d = d.set_title(title);
    for f in &filters {
        d = d.add_filter(&f.name, &f.exts);
    }
    d.pick_files()
        .await
        .unwrap_or_default()
        .into_iter()
        .map(|h| h.path().to_path_buf())
        .collect()
}

/// The Save dialog, in `folder` with `name` filled in, the filters' first type chosen: the path picked, and which
/// filter was chosen where the system says (GTK), or none when cancelled. The system asks before replacing a file.
#[cfg(any(windows, target_os = "macos"))]
pub async fn save(
    window: &WebviewWindow,
    title: &str,
    filters: Vec<Filter>,
    folder: Option<PathBuf>,
    name: &str,
) -> Option<(PathBuf, Option<usize>)> {
    let mut d = rfd::AsyncFileDialog::new()
        .set_title(title)
        .set_file_name(name)
        .set_parent(window);
    if let Some(f) = folder {
        d = d.set_directory(f);
    }
    for f in &filters {
        d = d.add_filter(&f.name, &f.exts);
    }
    d.save_file().await.map(|h| (h.path().to_path_buf(), None))
}

/// The folder chooser, opening in `folder`, `text` its message: the folder picked, none when cancelled. The Mac
/// sandbox lets the app into the folder picked (doc::host_allow_access).
#[cfg(target_os = "macos")]
pub async fn folder(
    window: &WebviewWindow,
    text: &str,
    folder: &std::path::Path,
) -> Option<PathBuf> {
    rfd::AsyncFileDialog::new()
        .set_title(text)
        .set_directory(folder)
        .set_parent(window)
        .pick_folder()
        .await
        .map(|h| h.path().to_path_buf())
}

/// A question with Yes and No; whether the answer was Yes.
#[cfg(any(windows, target_os = "macos"))]
pub async fn confirm(window: &WebviewWindow, title: &str, text: &str) -> bool {
    rfd::AsyncMessageDialog::new()
        .set_title(title)
        .set_description(text)
        .set_level(rfd::MessageLevel::Warning)
        .set_buttons(rfd::MessageButtons::YesNo)
        .set_parent(window)
        .show()
        .await
        == rfd::MessageDialogResult::Yes
}

#[cfg(all(unix, not(target_os = "macos")))]
pub async fn open(window: &WebviewWindow, title: &str, filters: Vec<Filter>) -> Vec<PathBuf> {
    let (tx, mut rx) = tauri::async_runtime::channel(1);
    let (w, title) = (window.clone(), title.to_string());
    let shown = window.run_on_main_thread(move || {
        use gtk::prelude::*;
        let parent = w.gtk_window().ok();
        let d = gtk::FileChooserNative::new(
            Some(&title),
            parent.as_ref(),
            gtk::FileChooserAction::Open,
            None,
            None,
        );
        d.set_modal(true);
        d.set_select_multiple(true);
        for f in &filters {
            let filter = gtk::FileFilter::new();
            filter.set_name(Some(&f.name));
            for e in &f.exts {
                filter.add_pattern(&any_case(e));
            }
            d.add_filter(filter);
        }
        // the chooser lives until it answers
        let keep = std::rc::Rc::new(std::cell::RefCell::new(Some(d.clone())));
        d.connect_response(move |d, response| {
            let files = if response == gtk::ResponseType::Accept {
                d.filenames()
            } else {
                Vec::new()
            };
            let _ = tx.try_send(files);
            keep.borrow_mut().take();
        });
        d.show();
    });
    if shown.is_err() {
        return Vec::new();
    }
    rx.recv().await.unwrap_or_default()
}

#[cfg(all(unix, not(target_os = "macos")))]
#[cfg_attr(feature = "lite", allow(dead_code))] // the lite edition asks for a name in the page (doc.rs host_choose)
pub async fn save(
    window: &WebviewWindow,
    title: &str,
    filters: Vec<Filter>,
    folder: Option<PathBuf>,
    name: &str,
) -> Option<(PathBuf, Option<usize>)> {
    let (tx, mut rx) = tauri::async_runtime::channel(1);
    let (w, title, name) = (window.clone(), title.to_string(), name.to_string());
    let shown = window.run_on_main_thread(move || {
        use gtk::prelude::*;
        let parent = w.gtk_window().ok();
        let d = gtk::FileChooserNative::new(
            Some(&title),
            parent.as_ref(),
            gtk::FileChooserAction::Save,
            None,
            None,
        );
        d.set_modal(true);
        d.set_do_overwrite_confirmation(true);
        if let Some(f) = &folder {
            d.set_current_folder(f);
        }
        d.set_current_name(&name);
        let list: Vec<gtk::FileFilter> = filters
            .iter()
            .map(|f| {
                let filter = gtk::FileFilter::new();
                filter.set_name(Some(&f.name));
                for e in &f.exts {
                    filter.add_pattern(&any_case(e));
                }
                d.add_filter(filter.clone());
                filter
            })
            .collect();
        if let Some(first) = list.first() {
            d.set_filter(first);
        }
        let keep = std::rc::Rc::new(std::cell::RefCell::new(Some(d.clone())));
        d.connect_response(move |d, response| {
            let picked = (response == gtk::ResponseType::Accept)
                .then(|| d.filename())
                .flatten()
                .map(|p| {
                    let chosen = d.filter();
                    (p, list.iter().position(|f| Some(f) == chosen.as_ref()))
                });
            let _ = tx.try_send(picked);
            keep.borrow_mut().take();
        });
        d.show();
    });
    if shown.is_err() {
        return None;
    }
    rx.recv().await.flatten()
}

#[cfg(all(unix, not(target_os = "macos")))]
pub async fn confirm(window: &WebviewWindow, title: &str, text: &str) -> bool {
    let (tx, mut rx) = tauri::async_runtime::channel(1);
    let (w, title, text) = (window.clone(), title.to_string(), text.to_string());
    let shown = window.run_on_main_thread(move || {
        use gtk::prelude::*;
        let parent = w.gtk_window().ok();
        let d = gtk::MessageDialog::new(
            parent.as_ref(),
            gtk::DialogFlags::MODAL | gtk::DialogFlags::DESTROY_WITH_PARENT,
            gtk::MessageType::Warning,
            gtk::ButtonsType::YesNo,
            &text,
        );
        d.set_title(&title);
        d.connect_response(move |d, response| {
            let _ = tx.try_send(response == gtk::ResponseType::Yes);
            d.close();
        });
        d.show();
    });
    shown.is_ok() && rx.recv().await.unwrap_or(false)
}

/// A glob that matches an extension in any case (`*.[dD][oO][cC][xX]`): GTK's patterns match case.
#[cfg(all(unix, not(target_os = "macos")))]
fn any_case(ext: &str) -> String {
    let mut p = String::from("*.");
    for c in ext.chars() {
        if c.is_ascii_alphabetic() {
            p.push('[');
            p.push(c.to_ascii_lowercase());
            p.push(c.to_ascii_uppercase());
            p.push(']');
        } else {
            p.push(c);
        }
    }
    p
}
