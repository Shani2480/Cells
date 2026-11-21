//! Insert › Screenshot: a picture of (part of) the screen, taken with the system's own tool so the person chooses the
//! area or window and nothing is captured without them. Linux: the XDG desktop portal's interactive screenshot (GNOME,
//! KDE and others draw their own picker). macOS: `screencapture -i`. Elsewhere the webview's screen picker is used
//! instead (getDisplayMedia), so this command says it can't.

/// The screenshot as image bytes, or None when the person cancelled.
#[tauri::command]
pub async fn take_screenshot() -> Result<Option<Vec<u8>>, String> {
    capture().await
}

#[cfg(target_os = "linux")]
async fn capture() -> Result<Option<Vec<u8>>, String> {
    use ashpd::desktop::ResponseError;
    use ashpd::desktop::screenshot::Screenshot;
    let request = Screenshot::request()
        .interactive(true)
        .modal(true)
        .send()
        .await
        .map_err(|e| format!("The screenshot tool isn't available: {e}"))?;
    match request.response() {
        Ok(shot) => {
            let path = file_path(shot.uri().as_str()).ok_or("The screenshot isn't a file.")?;
            let bytes = std::fs::read(&path).map_err(|e| e.to_string())?;
            Ok(Some(bytes))
        }
        Err(ashpd::Error::Response(ResponseError::Cancelled)) => Ok(None),
        Err(e) => Err(format!("The screenshot didn't work: {e}")),
    }
}

/// A file:// URI's path, percent escapes decoded.
#[cfg(target_os = "linux")]
fn file_path(uri: &str) -> Option<std::path::PathBuf> {
    let rest = uri.strip_prefix("file://")?;
    let bytes = rest.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        let hex = (bytes[i] == b'%' && i + 2 < bytes.len())
            .then(|| std::str::from_utf8(&bytes[i + 1..i + 3]).ok())
            .flatten()
            .and_then(|h| u8::from_str_radix(h, 16).ok());
        match hex {
            Some(v) => {
                out.push(v);
                i += 3;
            }
            None => {
                out.push(bytes[i]);
                i += 1;
            }
        }
    }
    use std::os::unix::ffi::OsStringExt;
    Some(std::path::PathBuf::from(std::ffi::OsString::from_vec(out)))
}

#[cfg(test)]
mod tests {
    #[cfg(target_os = "linux")]
    #[test]
    fn decodes_file_uris() {
        assert_eq!(
            super::file_path("file:///home/alex/Pictures/Screenshot%20from%202026.png").unwrap(),
            std::path::PathBuf::from("/home/alex/Pictures/Screenshot from 2026.png")
        );
        assert!(super::file_path("https://example.com/a.png").is_none());
        assert_eq!(
            super::file_path("file:///tmp/a%zz.png").unwrap(),
            std::path::PathBuf::from("/tmp/a%zz.png")
        );
    }
}

#[cfg(target_os = "macos")]
async fn capture() -> Result<Option<Vec<u8>>, String> {
    let path = std::env::temp_dir().join(format!("cells-screenshot-{}.png", std::process::id()));
    let _ = std::fs::remove_file(&path);
    let status = std::process::Command::new("/usr/sbin/screencapture")
        .args(["-i", "-x", "-t", "png"])
        .arg(&path)
        .status()
        .map_err(|e| e.to_string())?;
    if !status.success() || !path.exists() {
        return Ok(None);
    }
    let bytes = std::fs::read(&path).map_err(|e| e.to_string())?;
    let _ = std::fs::remove_file(&path);
    Ok(Some(bytes))
}

#[cfg(not(any(target_os = "linux", target_os = "macos")))]
async fn capture() -> Result<Option<Vec<u8>>, String> {
    Err("unsupported".into())
}
