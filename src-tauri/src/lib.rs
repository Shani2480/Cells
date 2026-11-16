//! The desktop shell: the window, the clipboard, and reading and writing the workbooks the person picks. The
//! workbook engine lives in TypeScript (packages/core and packages/io).

mod ai;
mod menu;

use std::path::{Path, PathBuf};
use std::sync::Mutex;

use tauri::{Emitter, Manager};

/// The window title for a workbook, as Excel shows it ("Book1 - Cells").
pub fn window_title(workbook: &str) -> String {
    if workbook.is_empty() {
        "Cells".to_string()
    } else {
        format!("{workbook} - Cells")
    }
}

/// Spreadsheet files Cells opens and saves. The commands below refuse any other file, so they can't be used to read
/// or overwrite anything else on the machine.
fn is_workbook_extension(ext: &str) -> bool {
    matches!(
        ext,
        "xlsx" | "xlsm" | "xltx" | "xltm" | "csv" | "tsv" | "txt"
    )
}

fn file_extension(path: &str) -> Option<String> {
    let name = path.rsplit(['/', '\\']).next()?;
    let dot = name.rfind('.')?;
    if dot == 0 {
        return None;
    }
    Some(name[dot + 1..].to_ascii_lowercase())
}

fn assert_workbook_path(path: &str) -> Result<PathBuf, String> {
    let trimmed = path.trim();
    if trimmed.is_empty() {
        return Err("Choose a file path.".into());
    }
    let checked = PathBuf::from(trimmed);
    if !checked.is_absolute() {
        return Err("Choose a full file path.".into());
    }
    if checked
        .components()
        .any(|part| matches!(part, std::path::Component::ParentDir))
    {
        return Err("That file path is not allowed.".into());
    }
    match file_extension(trimmed).as_deref() {
        Some(ext) if is_workbook_extension(ext) => Ok(checked),
        _ => Err(
            "Cells can only open and save .xlsx, .xlsm, .xltx, .xltm, .csv, .tsv and .txt files."
                .into(),
        ),
    }
}

fn percent_decode(value: &str) -> Result<String, String> {
    let bytes = value.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            let hex = value
                .get(index + 1..index + 3)
                .and_then(|pair| u8::from_str_radix(pair, 16).ok())
                .ok_or("That file path is not allowed.")?;
            out.push(hex);
            index += 3;
        } else {
            out.push(bytes[index]);
            index += 1;
        }
    }
    String::from_utf8(out).map_err(|_| "That file path is not allowed.".to_string())
}

/// Reads a workbook the person picked (or one from the recent list).
#[tauri::command]
fn read_workbook(path: String) -> Result<tauri::ipc::Response, String> {
    let path = assert_workbook_path(&path)?;
    if !path.is_file() {
        return Err(
            "The file could not be found. It may have been moved, renamed or deleted.".into(),
        );
    }
    let bytes = std::fs::read(&path).map_err(|err| format!("Could not read the file: {err}"))?;
    Ok(tauri::ipc::Response::new(bytes))
}

/// Writes a workbook: beside the destination first, then renamed over it, so a crash mid-write leaves the previous
/// file whole. The bytes come as the raw request body; the path in a header (percent-encoded).
#[tauri::command]
fn write_workbook(request: tauri::ipc::Request<'_>) -> Result<(), String> {
    let tauri::ipc::InvokeBody::Raw(bytes) = request.body() else {
        return Err("The workbook was not sent as bytes.".into());
    };
    let encoded = request
        .headers()
        .get("x-workbook-path")
        .and_then(|value| value.to_str().ok())
        .ok_or("Choose a file path.")?;
    let path = assert_workbook_path(&percent_decode(encoded)?)?;
    if let Some(parent) = path.parent() {
        if !parent.as_os_str().is_empty() && !parent.is_dir() {
            return Err("The folder for that file could not be found.".into());
        }
    }
    let file_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or("Choose a file path.")?;
    let mut temporary = path.clone();
    temporary.set_file_name(format!("~${file_name}.tmp"));
    if let Err(err) = std::fs::write(&temporary, bytes) {
        return Err(format!("Could not save the file: {err}"));
    }
    if let Err(err) = std::fs::rename(&temporary, &path) {
        let _ = std::fs::remove_file(&temporary);
        return Err(format!("Could not save the file: {err}"));
    }
    Ok(())
}

/// Whether a recent file still exists (the Recent list greys out the ones that don't).
#[tauri::command]
fn workbook_exists(path: String) -> bool {
    assert_workbook_path(&path)
        .map(|p| p.is_file())
        .unwrap_or(false)
}

/// The settings file Cells shares with Inkspire (docs/shared-settings.md): `inkspire/settings.json` in the
/// system's config folder (`$XDG_CONFIG_HOME` or `~/.config` on Linux, Application Support on macOS, `%APPDATA%` on
/// Windows). On Android each app has only its own folder, so the file is Cells' own there.
fn shared_settings_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let base = app
        .path()
        .config_dir()
        .map_err(|err| format!("The settings folder could not be found: {err}"))?;
    Ok(outside_flatpak(base, ".config")
        .join("inkspire")
        .join("settings.json"))
}

/// In a Flatpak the config and cache folders point into the app's sandbox (`~/.var/app/<id>/…`), where Inkspire
/// can't see them. The manifest grants the real `~/.config/inkspire` and `~/.cache/inkspire`, so those are used.
fn outside_flatpak(folder: PathBuf, home_folder: &str) -> PathBuf {
    host_folder(
        std::env::var_os("FLATPAK_ID").is_some(),
        std::env::var_os("HOME").map(PathBuf::from),
        folder,
        home_folder,
    )
}

fn host_folder(
    flatpak: bool,
    home: Option<PathBuf>,
    folder: PathBuf,
    home_folder: &str,
) -> PathBuf {
    match home {
        Some(home) if flatpak && home.is_absolute() => home.join(home_folder),
        _ => folder,
    }
}

/// The most the settings file may hold; anything bigger isn't a settings file.
const SETTINGS_LIMIT: u64 = 1024 * 1024;

/// The most a linked-ranges snapshot may hold (20,000 cells of formatted HTML fit well within it).
const SNAPSHOT_LIMIT: u64 = 16 * 1024 * 1024;

fn read_json_file(path: &Path, limit: u64) -> Option<String> {
    let meta = std::fs::metadata(path).ok()?;
    if !meta.is_file() || meta.len() > limit {
        return None;
    }
    std::fs::read_to_string(path).ok()
}

fn read_settings_file(path: &Path) -> Option<String> {
    read_json_file(path, SETTINGS_LIMIT)
}

/// Writes JSON beside the file first and renames it over, so the other app never reads half a file.
fn write_json_file(path: &Path, text: &str, limit: u64) -> Result<(), String> {
    if text.len() as u64 > limit {
        return Err("The file would be too big to save.".into());
    }
    serde_json::from_str::<serde_json::Value>(text)
        .map_err(|_| "The text is not valid JSON.".to_string())?;
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|err| format!("Could not save {}: {err}", path.display()))?;
    }
    let name = path
        .file_name()
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_default();
    let mut temporary = path.to_path_buf();
    temporary.set_file_name(format!("{name}.{}.tmp", std::process::id()));
    std::fs::write(&temporary, text).map_err(|err| format!("Could not save {name}: {err}"))?;
    std::fs::rename(&temporary, path).map_err(|err| {
        let _ = std::fs::remove_file(&temporary);
        format!("Could not save {name}: {err}")
    })
}

fn write_settings_file(path: &Path, text: &str) -> Result<(), String> {
    write_json_file(path, text, SETTINGS_LIMIT)
}

/// The folder of Inkspire's linked ranges (docs/linked-ranges.md), beside the shared settings.
fn links_folder(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    Ok(shared_settings_path(app)?.with_file_name("links"))
}

/// A snapshot's name: 16 lowercase hex digits, so it can't point outside the folder.
fn is_snapshot_name(name: &str) -> bool {
    name.len() == 16
        && name
            .bytes()
            .all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b))
}

#[tauri::command]
fn read_link_wants(app: tauri::AppHandle) -> Result<Option<String>, String> {
    Ok(read_json_file(
        &links_folder(&app)?.join("wanted.json"),
        SETTINGS_LIMIT,
    ))
}

#[tauri::command]
fn write_link_snapshot(app: tauri::AppHandle, name: String, text: String) -> Result<(), String> {
    if !is_snapshot_name(&name) {
        return Err("Not a snapshot name.".into());
    }
    write_json_file(
        &links_folder(&app)?.join(format!("{name}.json")),
        &text,
        SNAPSHOT_LIMIT,
    )
}

/// Whether `path` is a chart's workbook Inkspire copied out for Edit Data: an .xlsx inside `edit_folder` (after
/// following links and `..`), which Cells may then save back without asking.
fn inside_edit_folder(path: &Path, edit_folder: &Path) -> bool {
    let (Ok(file), Ok(folder)) = (path.canonicalize(), edit_folder.canonicalize()) else {
        return false;
    };
    file.is_file()
        && file.starts_with(&folder)
        && file_extension(&file.to_string_lossy()).as_deref() == Some("xlsx")
}

/// Inkspire's edit folder: `<cache dir>/inkspire/edit` (`~/.cache/inkspire/edit` on Linux).
#[tauri::command]
fn is_edit_copy(app: tauri::AppHandle, path: String) -> bool {
    match app.path().cache_dir() {
        Ok(cache) => inside_edit_folder(
            Path::new(&path),
            &outside_flatpak(cache, ".cache")
                .join("inkspire")
                .join("edit"),
        ),
        Err(_) => false,
    }
}

#[tauri::command]
fn read_shared_settings(app: tauri::AppHandle) -> Result<Option<String>, String> {
    Ok(read_settings_file(&shared_settings_path(&app)?))
}

#[tauri::command]
fn write_shared_settings(app: tauri::AppHandle, text: String) -> Result<(), String> {
    write_settings_file(&shared_settings_path(&app)?, &text)
}

/// The file's place for File › Options, with the home folder written as `~`.
#[tauri::command]
fn shared_settings_location(app: tauri::AppHandle) -> Result<String, String> {
    let path = shared_settings_path(&app)?;
    if let Ok(home) = app.path().home_dir() {
        if let Ok(rest) = path.strip_prefix(&home) {
            return Ok(format!("~/{}", rest.to_string_lossy()));
        }
    }
    Ok(path.to_string_lossy().into_owned())
}

/// What the desktop session says about the person and their region, for the first-run setup. Every field is empty
/// when the system does not give it. Taken from Inkspire (src-tauri/src/lib.rs), so both apps start from the same.
#[derive(serde::Serialize, Default)]
struct SystemDefaults {
    user: String,
    full_name: String,
    lang: String,
    lc_all: String,
    lc_time: String,
    lc_numeric: String,
}

/// The display name in a `passwd` entry's GECOS field: its first comma-separated part.
fn gecos_name(entry: &str) -> String {
    entry
        .split(':')
        .nth(4)
        .and_then(|gecos| gecos.split(',').next())
        .unwrap_or("")
        .trim()
        .to_string()
}

/// The name the system holds for the signed-in person: the passwd entry on Linux, the directory entry on macOS and
/// the account's display name on Windows. Empty when it holds none, so the app asks instead of guessing.
fn system_full_name(user: &str) -> String {
    let run = |program: &str, args: &[&str]| -> String {
        let mut command = std::process::Command::new(program);
        command.args(args);
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            command.creation_flags(0x0800_0000);
        }
        command
            .output()
            .map(|output| String::from_utf8_lossy(&output.stdout).trim().to_string())
            .unwrap_or_default()
    };
    #[cfg(target_os = "macos")]
    {
        let _ = user;
        return run("id", &["-F"]);
    }
    #[cfg(windows)]
    {
        let _ = user;
        return run(
            "powershell",
            &[
                "-NoProfile",
                "-Command",
                "Add-Type -AssemblyName System.DirectoryServices.AccountManagement; [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.DisplayName",
            ],
        );
    }
    #[cfg(target_os = "android")]
    {
        let _ = (user, run);
        return String::new();
    }
    #[cfg(all(unix, not(target_os = "macos"), not(target_os = "android")))]
    {
        if user.is_empty() {
            return String::new();
        }
        gecos_name(&run("getent", &["passwd", user]))
    }
}

#[tauri::command]
fn system_defaults() -> SystemDefaults {
    let var = |name: &str| std::env::var(name).unwrap_or_default();
    let user = [var("USER"), var("LOGNAME"), var("USERNAME")]
        .into_iter()
        .find(|name| !name.is_empty())
        .unwrap_or_default();
    let full_name = system_full_name(&user);
    SystemDefaults {
        user,
        full_name,
        lang: var("LANG"),
        lc_all: var("LC_ALL"),
        lc_time: var("LC_TIME"),
        lc_numeric: var("LC_NUMERIC"),
    }
}

/// `cells://` links and workbook files the app was started or called with, until the window takes them (it may not be
/// listening yet when the app starts).
#[derive(Default)]
struct PendingOpens(Mutex<Vec<String>>);

/// What in a launch's arguments the window should open: `cells://` links, and workbook files (relative paths are
/// taken from the folder the launch came from). The first argument is the program itself.
fn openable_arguments(args: &[String], cwd: &Path) -> Vec<String> {
    args.iter()
        .skip(1)
        .filter_map(|arg| {
            if arg.starts_with("cells://") {
                return Some(arg.clone());
            }
            if arg.starts_with('-') {
                return None;
            }
            // The desktop entry's %U: file managers pass local files as file:// URLs.
            if arg.starts_with("file://") {
                let url: tauri::Url = arg.parse().ok()?;
                return opened_files(&[url]).into_iter().next();
            }
            let path = PathBuf::from(arg);
            let full = if path.is_absolute() {
                path
            } else {
                cwd.join(path)
            };
            let text = full.to_string_lossy().into_owned();
            assert_workbook_path(&text).ok().map(|_| text)
        })
        .collect()
}

/// Hands links and files to the window: queued for `take_pending_opens`, announced with `cells-open`, and the
/// window brought to the front (a second launch lands here instead of opening a second Cells).
fn deliver(app: &tauri::AppHandle, items: Vec<String>) {
    if items.is_empty() {
        return;
    }
    if let Some(pending) = app.try_state::<PendingOpens>() {
        if let Ok(mut list) = pending.0.lock() {
            list.extend(items);
        }
    }
    let _ = app.emit_to("main", "cells-open", ());
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.unminimize();
        let _ = window.show();
        let _ = window.set_focus();
    }
}

/// The links and files waiting to be opened, oldest first; the list is empty afterwards.
#[tauri::command]
fn take_pending_opens(pending: tauri::State<'_, PendingOpens>) -> Vec<String> {
    pending
        .0
        .lock()
        .map(|mut list| std::mem::take(&mut *list))
        .unwrap_or_default()
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let builder = tauri::Builder::default();
    // One Cells at a time: a second launch (a cells:// link, a file opened from the file manager) hands its
    // arguments to the running one. This plugin has to come first.
    #[cfg(desktop)]
    let builder = builder.plugin(tauri_plugin_single_instance::init(|app, argv, cwd| {
        deliver(app, openable_arguments(&argv, Path::new(&cwd)));
    }));
    builder
        .plugin(tauri_plugin_deep_link::init())
        .manage(PendingOpens::default())
        .setup(|app| {
            let args: Vec<String> = std::env::args().collect();
            let cwd = std::env::current_dir().unwrap_or_default();
            if let Ok(mut list) = app.state::<PendingOpens>().0.lock() {
                list.extend(openable_arguments(&args, &cwd));
            }
            // The menu bar (macOS shows one for every app; elsewhere the window's ribbon is the menu).
            #[cfg(target_os = "macos")]
            {
                app.set_menu(menu::menu_bar(app.handle())?)?;
                app.on_menu_event(|app, event| {
                    if let Some(id) = menu::command_of(event.id().as_ref()) {
                        let _ = app.emit_to("main", "menu-command", id.to_string());
                    }
                });
            }
            #[cfg(not(target_os = "macos"))]
            let _ = menu::menu_bar::<tauri::Wry>;
            #[cfg(desktop)]
            {
                use tauri_plugin_deep_link::DeepLinkExt;
                // Linux and Windows learn the scheme when the app runs (an AppImage or a dev build is never
                // installed); macOS reads it from the bundle.
                #[cfg(any(target_os = "linux", windows))]
                let _ = app.deep_link().register_all();
                let handle = app.handle().clone();
                app.deep_link().on_open_url(move |event| {
                    deliver(
                        &handle,
                        event.urls().iter().map(|url| url.to_string()).collect(),
                    );
                });
            }
            Ok(())
        })
        .plugin(tauri_plugin_clipboard_manager::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            read_workbook,
            write_workbook,
            workbook_exists,
            read_shared_settings,
            write_shared_settings,
            shared_settings_location,
            system_defaults,
            take_pending_opens,
            read_link_wants,
            write_link_snapshot,
            is_edit_copy,
            ai::ai_generate,
            ai::ai_cancel,
            ai::ai_test,
            ai::ai_key_status,
            ai::ai_save_key
        ])
        .build(tauri::generate_context!())
        .expect("error while running Cells")
        .run(|_app, _event| {
            // macOS hands files opened from the Finder (or dropped on the Dock icon) to the running app as an event,
            // not as arguments. cells:// links come through the deep-link plugin instead.
            #[cfg(target_os = "macos")]
            if let tauri::RunEvent::Opened { urls } = _event {
                deliver(_app, opened_files(&urls));
            }
        });
}

/// The workbooks among `file://` URLs: the ones macOS asks the app to open, and %U arguments on Linux.
fn opened_files(urls: &[tauri::Url]) -> Vec<String> {
    urls.iter()
        .filter(|url| url.scheme() == "file")
        .filter_map(|url| url.to_file_path().ok())
        .map(|path| path.to_string_lossy().into_owned())
        .filter(|path| assert_workbook_path(path).is_ok())
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn titles() {
        assert_eq!(window_title("Book1"), "Book1 - Cells");
        assert_eq!(window_title(""), "Cells");
    }

    #[test]
    fn only_workbooks_by_full_path() {
        assert!(assert_workbook_path("/home/alex/Budget.xlsx").is_ok());
        assert!(assert_workbook_path("/home/alex/data.CSV").is_ok());
        assert!(assert_workbook_path("/home/alex/.bashrc").is_err());
        assert!(assert_workbook_path("/home/alex/script.sh").is_err());
        assert!(assert_workbook_path("Budget.xlsx").is_err());
        assert!(assert_workbook_path("/home/alex/../other/Budget.xlsx").is_err());
        assert!(assert_workbook_path("").is_err());
    }

    #[test]
    fn settings_file_round_trip() {
        let dir = std::env::temp_dir().join(format!("cells-settings-{}", std::process::id()));
        let path = dir.join("inkspire").join("settings.json");
        assert_eq!(read_settings_file(&path), None);
        write_settings_file(&path, r#"{"version":1}"#).unwrap();
        assert_eq!(
            read_settings_file(&path).as_deref(),
            Some(r#"{"version":1}"#)
        );
        assert!(write_settings_file(&path, "not json").is_err());
        assert_eq!(
            read_settings_file(&path).as_deref(),
            Some(r#"{"version":1}"#)
        );
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn snapshots_and_edit_copies() {
        assert!(is_snapshot_name("0123456789abcdef"));
        assert!(!is_snapshot_name("../settings.json"));
        assert!(!is_snapshot_name("0123456789ABCDEF"));
        let dir = std::env::temp_dir().join(format!("cells-edit-{}", std::process::id()));
        let edit = dir.join("inkspire").join("edit");
        std::fs::create_dir_all(&edit).unwrap();
        let copy = edit.join("Chart 1.xlsx");
        std::fs::write(&copy, b"x").unwrap();
        let other = dir.join("Budget.xlsx");
        std::fs::write(&other, b"x").unwrap();
        assert!(inside_edit_folder(&copy, &edit));
        assert!(!inside_edit_folder(&other, &edit));
        assert!(!inside_edit_folder(
            &edit.join("..").join("..").join("Budget.xlsx"),
            &edit
        ));
        assert!(!inside_edit_folder(&edit.join("missing.xlsx"), &edit));
        let snapshot = dir.join("links").join("0123456789abcdef.json");
        write_json_file(&snapshot, r#"{"app":"cells"}"#, SNAPSHOT_LIMIT).unwrap();
        assert_eq!(
            read_json_file(&snapshot, SNAPSHOT_LIMIT).as_deref(),
            Some(r#"{"app":"cells"}"#)
        );
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn launch_arguments() {
        let args: Vec<String> = [
            "cells",
            "--flag",
            "cells://open?file=%2Fhome%2Falex%2FBudget.xlsx",
            "Notes.txt",
            "x.sh",
            "/srv/a.xlsx",
            "file:///home/alex/Q1%20Report.xlsx",
            "file:///home/alex/photo.png",
        ]
        .iter()
        .map(|s| s.to_string())
        .collect();
        assert_eq!(
            openable_arguments(&args, Path::new("/home/alex")),
            vec![
                "cells://open?file=%2Fhome%2Falex%2FBudget.xlsx",
                "/home/alex/Notes.txt",
                "/srv/a.xlsx",
                "/home/alex/Q1 Report.xlsx"
            ]
        );
    }

    #[test]
    fn files_opened_from_the_finder() {
        let urls: Vec<tauri::Url> = [
            "file:///Users/alex/Budget%202026.xlsx",
            "file:///Users/alex/notes.pdf",
            "cells://open?file=%2FUsers%2Falex%2FA.xlsx",
            "file:///Users/alex/data.csv",
        ]
        .iter()
        .map(|u| u.parse().unwrap())
        .collect();
        assert_eq!(
            opened_files(&urls),
            vec!["/Users/alex/Budget 2026.xlsx", "/Users/alex/data.csv"]
        );
    }

    #[test]
    fn shared_folders_leave_the_flatpak_sandbox() {
        let sandboxed = PathBuf::from("/home/alex/.var/app/app.cells.desktop/config");
        let home = Some(PathBuf::from("/home/alex"));
        assert_eq!(
            host_folder(true, home.clone(), sandboxed.clone(), ".config"),
            PathBuf::from("/home/alex/.config")
        );
        assert_eq!(
            host_folder(false, home, sandboxed.clone(), ".config"),
            sandboxed
        );
        assert_eq!(
            host_folder(true, None, sandboxed.clone(), ".config"),
            sandboxed
        );
    }

    #[test]
    fn gecos() {
        assert_eq!(
            gecos_name("alex:x:1000:1000:Alex Smith,,,:/home/alex:/bin/bash"),
            "Alex Smith"
        );
        assert_eq!(gecos_name("alex:x:1000:1000::/home/alex:/bin/bash"), "");
    }

    #[test]
    fn percent_decoding() {
        assert_eq!(
            percent_decode("%2Fhome%2Falex%2Fa%20b.xlsx").unwrap(),
            "/home/alex/a b.xlsx"
        );
        assert!(percent_decode("%zz").is_err());
    }
}
