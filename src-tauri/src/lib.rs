//! The desktop shell: the window, the clipboard, and reading and writing the workbooks the person picks. The
//! workbook engine lives in TypeScript (packages/core and packages/io).

use std::path::{Path, PathBuf};

use tauri::Manager;

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
    Ok(base.join("inkspire").join("settings.json"))
}

/// The most the settings file may hold; anything bigger isn't a settings file.
const SETTINGS_LIMIT: u64 = 1024 * 1024;

fn read_settings_file(path: &Path) -> Option<String> {
    let meta = std::fs::metadata(path).ok()?;
    if !meta.is_file() || meta.len() > SETTINGS_LIMIT {
        return None;
    }
    std::fs::read_to_string(path).ok()
}

/// Writes the settings beside the file first and renames it over, so the other app never reads half a file.
fn write_settings_file(path: &Path, text: &str) -> Result<(), String> {
    if text.len() as u64 > SETTINGS_LIMIT {
        return Err("The settings are too big to save.".into());
    }
    serde_json::from_str::<serde_json::Value>(text)
        .map_err(|_| "The settings are not valid JSON.".to_string())?;
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|err| format!("Could not save the settings: {err}"))?;
    }
    let mut temporary = path.to_path_buf();
    temporary.set_file_name(format!("settings.json.{}.tmp", std::process::id()));
    std::fs::write(&temporary, text)
        .map_err(|err| format!("Could not save the settings: {err}"))?;
    std::fs::rename(&temporary, path).map_err(|err| {
        let _ = std::fs::remove_file(&temporary);
        format!("Could not save the settings: {err}")
    })
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

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
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
            system_defaults
        ])
        .run(tauri::generate_context!())
        .expect("error while running Cells");
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
