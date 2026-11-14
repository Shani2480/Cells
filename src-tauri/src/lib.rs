//! The desktop shell: the window, the clipboard, and reading and writing the workbooks the person picks. The
//! workbook engine lives in TypeScript (packages/core and packages/io).

use std::path::PathBuf;

/// The window title for a workbook, as Excel shows it ("Book1 - Cells").
pub fn window_title(workbook: &str) -> String {
    if workbook.is_empty() { "Cells".to_string() } else { format!("{workbook} - Cells") }
}

/// Spreadsheet files Cells opens and saves. The commands below refuse any other file, so they can't be used to read
/// or overwrite anything else on the machine.
fn is_workbook_extension(ext: &str) -> bool {
    matches!(ext, "xlsx" | "xlsm" | "xltx" | "xltm" | "csv" | "tsv" | "txt")
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
    if checked.components().any(|part| matches!(part, std::path::Component::ParentDir)) {
        return Err("That file path is not allowed.".into());
    }
    match file_extension(trimmed).as_deref() {
        Some(ext) if is_workbook_extension(ext) => Ok(checked),
        _ => Err("Cells can only open and save .xlsx, .xlsm, .xltx, .xltm, .csv, .tsv and .txt files.".into()),
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
        return Err("The file could not be found. It may have been moved, renamed or deleted.".into());
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
    let file_name = path.file_name().and_then(|name| name.to_str()).ok_or("Choose a file path.")?;
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
    assert_workbook_path(&path).map(|p| p.is_file()).unwrap_or(false)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_clipboard_manager::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![read_workbook, write_workbook, workbook_exists])
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
    fn percent_decoding() {
        assert_eq!(percent_decode("%2Fhome%2Falex%2Fa%20b.xlsx").unwrap(), "/home/alex/a b.xlsx");
        assert!(percent_decode("%zz").is_err());
    }
}
