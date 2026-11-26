//! One window per workbook on the desktop, as Excel has: each window shows one workbook, New and Open in a window
//! already in use open another window, and View › Switch Windows lists them all. Files and links handed to Cells go to
//! the window used last; each window takes what's waiting for it with `take_pending_opens`.

use std::collections::HashMap;
use std::sync::Mutex;
use std::sync::atomic::{AtomicUsize, Ordering};

use tauri::{Emitter, Manager};

/// What waits to be opened, by window label; and which window was used last (it gets files from the system).
#[derive(Default)]
pub struct WindowState {
    pending: Mutex<HashMap<String, Vec<String>>>,
    last_focused: Mutex<String>,
    /// The file each window shows (none for a new workbook), so a file already open is brought forward.
    paths: Mutex<HashMap<String, String>>,
}

#[cfg_attr(mobile, allow(dead_code))]
static NEXT_WINDOW: AtomicUsize = AtomicUsize::new(2);

impl WindowState {
    pub fn queue(&self, label: &str, items: Vec<String>) {
        if let Ok(mut map) = self.pending.lock() {
            map.entry(label.to_string()).or_default().extend(items);
        }
    }

    pub fn take(&self, label: &str) -> Vec<String> {
        self.pending
            .lock()
            .map(|mut map| map.remove(label).unwrap_or_default())
            .unwrap_or_default()
    }

    pub fn focused(&self, label: &str) {
        if let Ok(mut last) = self.last_focused.lock() {
            *last = label.to_string();
        }
    }

    pub fn closed(&self, label: &str) {
        if let Ok(mut map) = self.paths.lock() {
            map.remove(label);
        }
        if let Ok(mut map) = self.pending.lock() {
            map.remove(label);
        }
    }

    #[cfg_attr(mobile, allow(dead_code))]
    fn window_with(&self, path: &str) -> Option<String> {
        let map = self.paths.lock().ok()?;
        map.iter()
            .find(|(_, p)| same_path(p, path))
            .map(|(label, _)| label.clone())
    }
}

/// Paths compare as the system does: without case on Windows and macOS.
fn same_path(a: &str, b: &str) -> bool {
    if cfg!(any(windows, target_os = "macos")) {
        a.eq_ignore_ascii_case(b)
    } else {
        a == b
    }
}

/// The window that gets files and menu commands: the one used last, if it's still open, else any.
pub fn target_window<R: tauri::Runtime>(
    app: &tauri::AppHandle<R>,
) -> Option<tauri::WebviewWindow<R>> {
    let state = app.try_state::<WindowState>();
    let last = state.and_then(|s| s.last_focused.lock().ok().map(|l| l.clone()));
    if let Some(window) = last.as_deref().and_then(|l| app.get_webview_window(l)) {
        return Some(window);
    }
    app.get_webview_window("main")
        .or_else(|| app.webview_windows().into_values().next())
}

pub fn bring_forward<R: tauri::Runtime>(window: &tauri::WebviewWindow<R>) {
    #[cfg(desktop)]
    let _ = window.unminimize();
    let _ = window.show();
    let _ = window.set_focus();
}

/// Hands links and files to a window: the one showing that file, else the one used last.
pub fn deliver_to<R: tauri::Runtime>(app: &tauri::AppHandle<R>, items: Vec<String>) {
    if items.is_empty() {
        return;
    }
    let Some(window) = target_window(app) else {
        return;
    };
    if let Some(state) = app.try_state::<WindowState>() {
        state.queue(window.label(), items);
    }
    let _ = app.emit_to(window.label(), "cells-open", ());
    bring_forward(&window);
}

/// The links and files waiting for this window, oldest first.
#[tauri::command]
pub fn take_pending_opens(
    window: tauri::Window,
    state: tauri::State<'_, WindowState>,
) -> Vec<String> {
    state.take(window.label())
}

/// The window's workbook is now this file (or a new workbook: none).
#[tauri::command]
pub fn window_document(
    window: tauri::Window,
    state: tauri::State<'_, WindowState>,
    path: Option<String>,
) {
    if let Ok(mut map) = state.paths.lock() {
        match path {
            Some(p) => map.insert(window.label().to_string(), p),
            None => map.remove(window.label()),
        };
    }
}

/// A new window for `request` (a workbook's path or a cells:// link), or for a new workbook. When the workbook is
/// already open in a window, that window comes forward instead (and gets the request, to go to a range).
#[cfg(desktop)]
#[tauri::command]
pub fn open_window(
    app: tauri::AppHandle,
    window: tauri::Window,
    state: tauri::State<'_, WindowState>,
    request: Option<String>,
    path: Option<String>,
) -> Result<(), String> {
    if let Some(label) = path.as_deref().and_then(|p| state.window_with(p)) {
        if let Some(existing) = app.get_webview_window(&label) {
            if let Some(r) = request {
                state.queue(&label, vec![r]);
                let _ = app.emit_to(label.as_str(), "cells-open", ());
            }
            bring_forward(&existing);
            return Ok(());
        }
    }
    let label = format!("book-{}", NEXT_WINDOW.fetch_add(1, Ordering::SeqCst));
    if let Some(r) = request {
        state.queue(&label, vec![r]);
    }
    // A little below and to the right of the window it came from, as Excel places a new window.
    let at = window
        .outer_position()
        .ok()
        .zip(window.scale_factor().ok())
        .map(|(p, scale)| (p.x as f64 / scale + 26.0, p.y as f64 / scale + 26.0));
    let mut builder =
        tauri::WebviewWindowBuilder::new(&app, &label, tauri::WebviewUrl::App("index.html".into()))
            .title("Cells")
            .inner_size(1200.0, 800.0)
            .min_inner_size(640.0, 420.0)
            .resizable(true)
            .decorations(false)
            .zoom_hotkeys_enabled(false);
    #[cfg(not(target_os = "macos"))]
    {
        builder = builder.transparent(true).shadow(true);
    }
    #[cfg(target_os = "macos")]
    {
        builder = builder.transparent(true);
    }
    builder = match at {
        Some((x, y)) => builder.position(x, y),
        None => builder.center(),
    };
    let created = builder.build().map_err(|e| e.to_string())?;
    #[cfg(target_os = "linux")]
    crate::pinch::install(&created);
    let _ = created.set_focus();
    Ok(())
}

#[derive(serde::Serialize)]
#[cfg_attr(mobile, allow(dead_code))]
pub struct WindowEntry {
    id: String,
    title: String,
    current: bool,
}

/// Phones and tablets keep their workbooks in one window (the app's own switcher).
#[cfg(mobile)]
#[tauri::command]
pub fn open_window(_request: Option<String>, _path: Option<String>) -> Result<(), String> {
    Err("One window on this device.".into())
}

/// Every Cells window, in the order they were opened, with its title ("Budget.xlsx - Cells").
#[cfg(desktop)]
#[tauri::command]
pub fn list_windows(app: tauri::AppHandle, window: tauri::Window) -> Vec<WindowEntry> {
    let mut list: Vec<WindowEntry> = app
        .webview_windows()
        .into_iter()
        .map(|(label, w)| WindowEntry {
            current: label == window.label(),
            title: w.title().unwrap_or_default(),
            id: label,
        })
        .collect();
    list.sort_by_key(|w| order_of(&w.id));
    list
}

#[cfg(mobile)]
#[tauri::command]
pub fn list_windows() -> Vec<WindowEntry> {
    Vec::new()
}

/// "main" first, then "book-2", "book-3"… by number.
#[cfg_attr(mobile, allow(dead_code))]
fn order_of(label: &str) -> usize {
    if label == "main" {
        0
    } else {
        label
            .strip_prefix("book-")
            .and_then(|n| n.parse().ok())
            .unwrap_or(usize::MAX)
    }
}

/// Switch Windows: brings another Cells window to the front.
#[tauri::command]
pub fn focus_window(app: tauri::AppHandle, id: String) -> Result<(), String> {
    let window = app
        .get_webview_window(&id)
        .ok_or("That window is closed.")?;
    bring_forward(&window);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn queues_by_window() {
        let state = WindowState::default();
        state.queue("main", vec!["/home/alex/a.xlsx".into()]);
        state.queue("book-2", vec!["cells://open?file=/home/alex/b.xlsx".into()]);
        assert_eq!(
            state.take("book-2"),
            vec!["cells://open?file=/home/alex/b.xlsx".to_string()]
        );
        assert!(state.take("book-2").is_empty());
        assert_eq!(state.take("main"), vec!["/home/alex/a.xlsx".to_string()]);
    }

    #[test]
    fn finds_the_window_showing_a_file() {
        let state = WindowState::default();
        state
            .paths
            .lock()
            .unwrap()
            .insert("book-3".into(), "/home/alex/Budget.xlsx".into());
        assert_eq!(
            state.window_with("/home/alex/Budget.xlsx").as_deref(),
            Some("book-3")
        );
        assert_eq!(state.window_with("/home/alex/Other.xlsx"), None);
        state.closed("book-3");
        assert_eq!(state.window_with("/home/alex/Budget.xlsx"), None);
    }

    #[test]
    fn window_order() {
        let mut labels = vec!["book-10", "book-2", "main"];
        labels.sort_by_key(|l| order_of(l));
        assert_eq!(labels, vec!["main", "book-2", "book-10"]);
    }
}
