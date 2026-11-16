# Releasing Cells

One tag builds every installer (`.github/workflows/release.yml`):

| Platform | File                            | Built on                                            | Notes                                                                                        |
| -------- | ------------------------------- | --------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| Linux    | `Cells_<v>_amd64.deb`           | Ubuntu 22.04, so it runs on older distributions too | Registers the file types and `cells://`                                                      |
| Linux    | `Cells_<v>_amd64.AppImage`      | Ubuntu 22.04                                        | Registers `cells://` the first time it runs; file types only through the desktop's Open With |
| Linux    | `Cells_<v>_x86_64.flatpak`      | GNOME 50 runtime, from the `.deb`                   | `packaging/flatpak/`; shares `~/.config/inkspire` and `~/.cache/inkspire` with Inkspire      |
| Windows  | `Cells_<v>_x64-setup.exe`       | NSIS, installs for the current user                 | Registers the file types and `cells://`                                                      |
| macOS    | `Cells_<v>_universal.dmg`       | Apple Silicon and Intel in one app                  | Signed ad hoc, not notarized: the first open needs right-click › Open                        |
| Android  | `Cells_<v>_android.apk`, `.aab` | arm64 and x86_64, from `main`                       | Uploaded only when signed (below); id `app.cells`                                            |

## Android signing

Every Android update must be signed with the same key as the version already installed. Cells reuses Inkspire's
key (decided 2026-11-16): the release job reads the secrets `ANDROID_KEY_BASE64` (the `.jks` file, base64) and
`ANDROID_KEY_PASSWORD`, with the key alias `inkspire`, as Inkspire's workflow does. Copy those two secrets from
Inkspire's repository settings into this one before the first Android release. Until then the job builds unsigned,
warns, and uploads nothing. Debug builds (`npm run tauri android build -- --debug`) use the local debug key.

The version code comes from the version (`0.6.0` → 6000), so it rises with every release.

## Building for Android here

`ANDROID_HOME=~/Android/Sdk NDK_HOME=~/Android/Sdk/ndk/<version> npm run tauri android build -- --debug --target
x86_64 --apk` makes an emulator build; `src-tauri/gen/android` is the generated project, kept in git with Cells'
changes (`MainActivity.kt`: system bar and keyboard insets, no page zoom; signing in `app/build.gradle.kts`; Cells'
launcher icons; no Android TV launcher).

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

## Android smoke test

`scripts/android-smoke.sh <emulator serial> <phone|tablet> [screenshot folder]` runs on an emulator only (it clears
Cells' data and puts the status bar in demo mode): it installs the debug build, skips the setup, opens
`packages/io/test/fixtures/excel/budget.xlsx` as "Open with" does, types into D1 with hardware keys, saves with
Ctrl+S and pulls the file back to check the entry is in it. It then swipes the sheet while
`scripts/android-frames.mjs` records frame times inside the web view (through its DevTools, which debug builds
allow), checks that Back closes the File view and then puts Cells in the background still running, and takes light
and dark screenshots (`docs/images/android-*.png` come from it).

The emulators used: `cells_phone` (Pixel 7, 1080 × 2400) and a Pixel Tablet (2560 × 1600), both Android 15 (API 35).

Measured on 2026-11-16, five quick swipes, frame times inside the web view:

| Where                                  | Median  | 99th percentile | Dropped (over 25 ms) |
| -------------------------------------- | ------- | --------------- | -------------------- |
| Phone emulator, host GPU (`-gpu host`) | 16.7 ms | 16.8 ms         | 0%                   |
| Phone emulator, software GPU           | 16.7 ms | 67 ms           | 21%                  |
| Tablet emulator, host GPU              | 16.7 ms | 67 ms           | 26%                  |
| Real tablet (more pixels than the emulator) | 16.7 ms | 16.8 ms         | 0.7%                 |

The emulators' drops come from the emulator: the real tablet, with more pixels, keeps 60 frames a second. A real
phone hasn't been measured yet (verify-list).
