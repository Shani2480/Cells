//! The desktop shell. Phase 1 keeps it to the window and the clipboard; the workbook engine lives in TypeScript
//! (packages/core), and file access arrives with Phase 4.

/// The window title for a workbook, as Excel shows it ("Book1 - Cells").
pub fn window_title(workbook: &str) -> String {
    if workbook.is_empty() { "Cells".to_string() } else { format!("{workbook} - Cells") }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_clipboard_manager::init())
        .run(tauri::generate_context!())
        .expect("error while running Cells");
}

#[cfg(test)]
mod tests {
    use super::window_title;

    #[test]
    fn titles() {
        assert_eq!(window_title("Book1"), "Book1 - Cells");
        assert_eq!(window_title(""), "Cells");
    }
}
