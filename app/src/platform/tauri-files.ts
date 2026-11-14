import type { FilePort, FileRef } from "@cells/platform"
import { invoke } from "@tauri-apps/api/core"
import { open, save } from "@tauri-apps/plugin-dialog"
import { readFile, writeFile } from "@tauri-apps/plugin-fs"

/**
 * Files on the desktop and Android. Desktop paths go through the shell's own commands (src-tauri/src/lib.rs), which
 * accept spreadsheet files only and save through a temporary file. Android's picker hands out content URIs, which
 * only the fs plugin can read and write.
 */

export const baseName = (path: string) => path.split(/[/\\]/).pop() ?? path

const isContentUri = (path: string) => path.startsWith("content://")

async function readPath(path: string): Promise<Uint8Array> {
  if (isContentUri(path)) return readFile(path)
  return new Uint8Array(await invoke<ArrayBuffer>("read_workbook", { path }))
}

export function tauriFiles(): FilePort {
  return {
    saveInPlace: true,
    remembersFiles: true,
    async pickOpen(types) {
      const path = await open({ title: "Open", multiple: false, directory: false, filters: types })
      if (typeof path !== "string") return null
      return { ref: { name: baseName(path), path }, bytes: await readPath(path) }
    },
    read(ref: FileRef) {
      if (!ref.path) throw new Error("The file has no path.")
      return readPath(ref.path)
    },
    async pickSave(suggestedName, types) {
      const path = await save({ title: "Save As", defaultPath: suggestedName, filters: types })
      if (!path) return null
      // Some dialogs leave the extension off when the person types a bare name.
      const ext = types[0]?.extensions[0]
      const full = ext && !/\.[A-Za-z0-9]{1,5}$/.test(baseName(path)) ? `${path}.${ext}` : path
      return { name: baseName(full), path: full }
    },
    async write(ref, bytes) {
      if (!ref.path) throw new Error("The file has no path.")
      if (isContentUri(ref.path)) return writeFile(ref.path, bytes)
      await invoke("write_workbook", bytes, { headers: { "x-workbook-path": encodeURIComponent(ref.path) } })
    },
    async exists(ref) {
      if (!ref.path) return false
      if (isContentUri(ref.path)) return true
      return invoke<boolean>("workbook_exists", { path: ref.path })
    },
  }
}
