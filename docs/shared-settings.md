# Settings shared by Cells and Inkspire

Both apps keep the person's settings in one JSON file, so they are set once (PLAN.md Phase 5, item 1;
`docs/inkspire-todo.md` I1 is Inkspire's side).

## Where

`<config dir>/inkspire/settings.json`, where the config dir is Tauri's `config_dir()`:

| System  | Folder                                        |
| ------- | --------------------------------------------- |
| Linux   | `$XDG_CONFIG_HOME` or `~/.config`             |
| macOS   | `~/Library/Application Support`               |
| Windows | `%APPDATA%` (`C:\Users\alex\AppData\Roaming`) |
| Android | each app's own folder (nothing is shared)     |

The browser build keeps the same text in `localStorage["inkspire.shared-settings"]`.

## What

```json
{
  "version": 1,
  "user": { "name": "Alex Demo", "initials": "AD" },
  "regional": { "locale": "en-GB", "dateFormat": "dd/MM/yyyy", "decimalSeparator": ".", "groupSeparator": "," },
  "appearance": { "theme": "system" },
  "ai": {
    "provider": "ollama",
    "models": { "ollama": "llama3.2", "anthropic": "" },
    "addresses": { "lmstudio": "http://127.0.0.1:1234/v1" }
  },
  "proofing": { "customWords": ["Zorbly", "Q-Ledger"] },
  "apps": {
    "cells": { "setupDone": true, "ai": true, "hints": true, "hintsDismissed": ["quick-analysis", "ai-tools"] },
    "inkspire": { "setupDone": true }
  }
}
```

- `user.name` signs notes, comments and tracked changes; `user.initials` is what Office shows in comment avatars.
- `regional`: a BCP 47 tag, a short date pattern made of `yyyy`, `yy`, `MM`, `M`, `dd` and `d`, and one character
  for each separator. A blank value means "the locale's own" (and a blank locale means the system's).
- `appearance.theme`: `system`, `light` or `dark`. Inkspire, which has no "system" choice, reads `system` as the
  system's current setting.
- `ai`: which AI both apps use, as Inkspire names providers, with a model and an address per provider (blank: the
  default). A map is merged key by key, so one app's change to one provider keeps the others. Keys are never here
  (`docs/ai.md`).
- `proofing.customWords`: words added to the dictionary (Spelling › Add to Dictionary) in either app, kept as
  typed. An app adds its words to the list in the file (never replaces it), so both apps' words stay.
- `apps.<app>`: what only one app uses (its own first-run setup; Cells' `ai`: the AI tools are on; Cells' `hints`:
  feature hints are shown, missing means on; `hintsDismissed`: the ids of the hints already dismissed, replaced as a
  whole, so Show hints again can empty it). An app never
  changes another app's entry.

## Rules for writing it

1. Never write from memory. Read the file, change only the values the person changed, and write it back, so a
   change the other app made meanwhile survives. Keep every key you don't know.
2. Write it whole: to a temporary file in the same folder, then rename it over `settings.json`.
3. Keep a higher `version` you find; write `1` otherwise. A later version may add keys, never change these.
4. Read it again when the window comes back to the front, to pick up the other app's changes.
5. A missing, unreadable or broken file reads as all defaults; the next save writes a fresh one.

Cells: `app/src/settings/shared.ts` (format), `app/src/settings/settings.ts` (store),
`src-tauri/src/lib.rs` (`read_shared_settings`, `write_shared_settings`).
