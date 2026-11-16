# Releasing Cells

One tag builds every installer (`.github/workflows/release.yml`):

| Platform | File                       | Built on                                            | Notes                                                                                        |
| -------- | -------------------------- | --------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| Linux    | `Cells_<v>_amd64.deb`      | Ubuntu 22.04, so it runs on older distributions too | Registers the file types and `cells://`                                                      |
| Linux    | `Cells_<v>_amd64.AppImage` | Ubuntu 22.04                                        | Registers `cells://` the first time it runs; file types only through the desktop's Open With |
| Linux    | `Cells_<v>_x86_64.flatpak` | GNOME 50 runtime, from the `.deb`                   | `packaging/flatpak/`; shares `~/.config/inkspire` and `~/.cache/inkspire` with Inkspire      |
| Windows  | `Cells_<v>_x64-setup.exe`  | NSIS, installs for the current user                 | Registers the file types and `cells://`                                                      |
| macOS    | `Cells_<v>_universal.dmg`  | Apple Silicon and Intel in one app                  | Signed ad hoc, not notarized: the first open needs right-click › Open                        |

Android comes with Phase 6's Android part.

## Steps

1. Raise the version in `package.json`, `src-tauri/tauri.conf.json` and `src-tauri/Cargo.toml`.
2. Add the version's notes at the top of `app/src/about/release-notes.ts`, and the same version and date at the top of
   `<releases>` in `packaging/flatpak/app.cells.desktop.metainfo.xml`. `release-notes.test.ts` fails until all of
   these agree.
3. Commit, then tag `v<version>` and push the tag. The workflow checks the tag against `package.json`, drafts the
   release with the notes (`node scripts/release-notes.mjs <version>`), builds each platform, uploads the files and
   publishes the release when all of them are there.

## File types

`bundle.fileAssociations` in `tauri.conf.json`: `.xlsx`, `.xlsm`, `.xltx` and `.xltm` as Cells' own, `.csv` and
`.tsv` as an alternative (on macOS ranked Alternate, so the Finder keeps another app as their default; on Linux
and Windows the desktop decides which app is the default). Opened files reach the app as
launch arguments on Linux and Windows, handed to the running Cells when there is one, and as an "opened" event on
macOS (`RunEvent::Opened` in `src-tauri/src/lib.rs`).

## Checked so far

- 2026-11-16: the `.deb`, AppImage and Flatpak built locally. The desktop entry (`src-tauri/cells.desktop`, so `Exec`
  takes `%U`) passes `desktop-file-validate` and lists the file types and `cells://`; the metainfo passes
  `appstreamcli validate`. The release build on Xvfb opened files handed to it by a second launch (a path and a
  `file://` URL) and a `cells://` link with a range. The Windows and macOS builds run only in the workflow, which
  hasn't run yet (`verify-list.md`).
