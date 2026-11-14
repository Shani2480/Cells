<p align="center"><img src="design/app-icon.svg" width="96" alt="Cells icon"></p>

# Cells

A local-only spreadsheet with Excel's interface, keyboard and formula behaviour. It never talks to a cloud service.
Desktop first (Linux, then macOS and Windows), with Android tablets and phones on the same engine later.

| Light                                               | Dark                                              |
| --------------------------------------------------- | ------------------------------------------------- |
| ![Cells in light mode](docs/images/cells-light.png) | ![Cells in dark mode](docs/images/cells-dark.png) |

_Screenshots show made-up demo data._

## Status

Phase 1 of 6 (see [PLAN.md](PLAN.md)): the grid, editing, Excel's keyboard navigation, the clipboard, rows and
columns, and a minimal shell. Formulas arrive in Phase 2, the ribbon and formatting in Phase 3, files in Phase 4.

## Development

```sh
npm install
npm run dev            # the app in a browser at http://localhost:1430 (try ?demo=100k)
npm run tauri dev      # the desktop app
npm test               # Vitest: core and app
npm run bench          # core benchmarks on a 100,000 × 20 sheet
npm run lint           # ESLint, Prettier and the package-boundary check
npm run typecheck
npm run tauri build    # .deb and AppImage
```

Linux needs WebKitGTK 4.1 and GTK 3 development packages for the desktop build.

## Layout

| Path                | What                                                                                      |
| ------------------- | ----------------------------------------------------------------------------------------- |
| `packages/core`     | The workbook model, operations, undo, navigation, fill, formatting. Pure TypeScript only. |
| `packages/platform` | Interfaces the shells implement (clipboard, window).                                      |
| `app/`              | The React shell: canvas grid, formula bar, status bar, dialogs.                           |
| `src-tauri/`        | The Tauri 2 desktop shell.                                                                |
| `design/`           | The app icon's master SVG.                                                                |
| `docs/`             | `changes.md`, `next-todo.md`, `verify-list.md`.                                           |

## License

MIT
