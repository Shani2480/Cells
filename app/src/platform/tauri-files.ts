import type { FilePort, FileRef } from "@cells/platform"
import { invoke } from "@tauri-apps/api/core"
import { getCurrentWebview } from "@tauri-apps/api/webview"
import { open, save } from "@tauri-apps/plugin-dialog"
import { readFile, writeFile } from "@tauri-apps/plugin-fs"
import { androidFiles, fileNameOf, isContentUri, pickOpenDocument, pickSaveDocument } from "./android-files"

/**
 * Files on the desktop and Android. Desktop paths go through the shell's own commands (src-tauri/src/lib.rs), which
 * accept spreadsheet files only and save through a temporary file. On Android the system's document picker hands out
 * content links (android-files.ts), which only the fs plugin can read and write.
 */

export const baseName = fileNameOf

async function readPath(path: string): Promise<Uint8Array> {
  if (isContentUri(path)) return readFile(path)
  return new Uint8Array(await invoke<ArrayBuffer>("read_workbook", { path }))
}

export function tauriFiles(): FilePort {
  const files: FilePort = {
    saveInPlace: true,
    remembersFiles: true,
    async pickOpen(types) {
      const path = androidFiles()
        ? await pickOpenDocument(types.flatMap((type) => type.extensions))
        : await open({ title: "Open", multiple: false, directory: false, filters: types })
      if (typeof path !== "string") return null
      return { ref: { name: baseName(path), path }, bytes: await readPath(path) }
    },
    read(ref: FileRef) {
      if (!ref.path) throw new Error("The file has no path.")
      return readPath(ref.path)
    },
    async pickSave(suggestedName, types) {
      if (androidFiles()) {
        const ext = types[0]?.extensions[0]
        const named = ext && !/\.[A-Za-z0-9]{1,5}$/.test(suggestedName) ? `${suggestedName}.${ext}` : suggestedName
        const uri = await pickSaveDocument(named)
        return uri ? { name: baseName(uri), path: uri } : null
      }
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
      if (isContentUri(ref.path)) return androidFiles()?.exists(ref.path) ?? true
      return invoke<boolean>("workbook_exists", { path: ref.path })
    },
    onDrop(handler) {
      let disposed = false
      let unlisten: (() => void) | undefined
      void getCurrentWebview()
        .onDragDropEvent((event) => {
          if (event.payload.type !== "drop") return
          const path = event.payload.paths[0]
          if (!path) return
          void readPath(path).then(
            (bytes) => handler({ ref: { name: baseName(path), path }, bytes }),
            () => undefined,
          )
        })
        .then((fn) => {
          unlisten = fn
          if (disposed) fn()
        })
      return () => {
        disposed = true
        unlisten?.()
      }
    },
  }
  // Android's picker hands out content links that only last while Cells runs: no lasting path for a link.
  if (!androidFiles()) {
    files.pickPath = async (title) => {
      const path = await open({ title, multiple: false, directory: false })
      return typeof path === "string" ? path : null
    }
  }
  return files
}
