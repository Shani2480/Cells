//! The macOS menu bar, after Excel for Mac's: Cells, File, Edit, View, Format, Window and Help. Cells' own items send
//! their id (without the `cmd:` prefix) to the window as a `menu-command` event, and `app/src/shell/menu-bar.ts` runs
//! them, so the ribbon's commands do the work. Cut, Copy, Paste, Hide, Minimize and Full Screen are the system's own.
//! The menu is built on every OS (so it is type-checked everywhere) but only installed on macOS.
//!
//! The web view sees a key before the menu does: a shortcut Cells handles (⌘S) runs once, in the page, and the menu
//! only acts on a click or a key the page leaves alone (⌘Q, ⌘H).

use tauri::menu::{Menu, MenuBuilder, MenuItem, MenuItemBuilder, SubmenuBuilder};
use tauri::{AppHandle, Runtime};

/// What marks one of Cells' own items.
pub const PREFIX: &str = "cmd:";

/// Cells' own items: id, label, shortcut (as Tauri writes accelerators; `Cmd` is ⌘, `Ctrl` is ⌃).
const ITEMS: &[(&str, &str, Option<&str>)] = &[
    ("About", "About Cells", None),
    ("Settings", "Settings…", Some("Cmd+,")),
    ("Quit", "Quit Cells", Some("Cmd+Q")),
    ("FileNew", "New Workbook", Some("Cmd+N")),
    ("FileOpen", "Open…", Some("Cmd+O")),
    ("FileClose", "Close", Some("Cmd+W")),
    ("FileSave", "Save", Some("Cmd+S")),
    ("FileSaveAs", "Save As…", Some("Cmd+Shift+S")),
    ("FilePrint", "Print…", Some("Cmd+P")),
    ("Undo", "Undo", Some("Cmd+Z")),
    ("Redo", "Redo", Some("Cmd+Y")),
    ("PasteSpecialDialog", "Paste Special…", Some("Ctrl+Cmd+V")),
    ("SelectAll", "Select All", Some("Cmd+A")),
    ("FindDialogExcel", "Find…", Some("Cmd+F")),
    ("ReplaceDialog", "Replace…", Some("Ctrl+H")),
    ("GoTo", "Go To…", Some("Ctrl+G")),
    ("ToggleRibbon", "Ribbon", Some("Alt+Cmd+R")),
    ("ShowFormulas", "Formulas", Some("Ctrl+Backquote")),
    ("FormatCellsDialog", "Cells…", Some("Cmd+1")),
    ("WhatsNew", "What's New in Cells", None),
];

fn item<R: Runtime>(app: &AppHandle<R>, id: &str) -> tauri::Result<MenuItem<R>> {
    let (_, label, shortcut) = ITEMS
        .iter()
        .find(|(item, _, _)| *item == id)
        .copied()
        .unwrap_or((id, id, None));
    let mut builder = MenuItemBuilder::with_id(format!("{PREFIX}{id}"), label);
    if let Some(shortcut) = shortcut {
        builder = builder.accelerator(shortcut);
    }
    builder.build(app)
}

/// The whole menu bar.
pub fn menu_bar<R: Runtime>(app: &AppHandle<R>) -> tauri::Result<Menu<R>> {
    let cells = SubmenuBuilder::new(app, "Cells")
        .item(&item(app, "About")?)
        .separator()
        .item(&item(app, "Settings")?)
        .separator()
        .services()
        .separator()
        .hide()
        .hide_others()
        .show_all()
        .separator()
        // Cells' own Quit, so unsaved changes are asked about first (the system's ends the app at once).
        .item(&item(app, "Quit")?)
        .build()?;
    let file = SubmenuBuilder::new(app, "File")
        .item(&item(app, "FileNew")?)
        .item(&item(app, "FileOpen")?)
        .separator()
        .item(&item(app, "FileClose")?)
        .item(&item(app, "FileSave")?)
        .item(&item(app, "FileSaveAs")?)
        .separator()
        .item(&item(app, "FilePrint")?)
        .build()?;
    let edit = SubmenuBuilder::new(app, "Edit")
        .item(&item(app, "Undo")?)
        .item(&item(app, "Redo")?)
        .separator()
        .cut()
        .copy()
        .paste()
        .item(&item(app, "PasteSpecialDialog")?)
        .item(&item(app, "SelectAll")?)
        .separator()
        .item(&item(app, "FindDialogExcel")?)
        .item(&item(app, "ReplaceDialog")?)
        .item(&item(app, "GoTo")?)
        .build()?;
    let view = SubmenuBuilder::new(app, "View")
        .item(&item(app, "ToggleRibbon")?)
        .item(&item(app, "ShowFormulas")?)
        .separator()
        .fullscreen()
        .build()?;
    let format = SubmenuBuilder::new(app, "Format")
        .item(&item(app, "FormatCellsDialog")?)
        .build()?;
    let window = SubmenuBuilder::new(app, "Window")
        .minimize()
        .maximize()
        .build()?;
    let help = SubmenuBuilder::new(app, "Help")
        .item(&item(app, "WhatsNew")?)
        .build()?;
    MenuBuilder::new(app)
        .items(&[&cells, &file, &edit, &view, &format, &window, &help])
        .build()
}

/// The id a menu event carries for one of Cells' items, without its prefix.
#[cfg_attr(not(target_os = "macos"), allow(dead_code))]
pub fn command_of(menu_id: &str) -> Option<&str> {
    menu_id.strip_prefix(PREFIX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ids() {
        assert_eq!(command_of("cmd:FileSave"), Some("FileSave"));
        assert_eq!(command_of("quit"), None);
        let mut seen = std::collections::HashSet::new();
        for (id, label, _) in ITEMS {
            assert!(seen.insert(*id), "{id} twice");
            assert!(!label.is_empty());
        }
    }
}
