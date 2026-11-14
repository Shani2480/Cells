import type { FilePort, FileRef, FileType } from "@cells/platform"

/**
 * Files in a browser. With the File System Access API (Chromium) Save writes over the file that was opened; without
 * it Open uses a file input and every save downloads a copy. Handles live only as long as the page.
 */

type Handle = {
  name: string
  getFile(): Promise<File>
  createWritable(): Promise<{ write(data: Blob): Promise<void>; close(): Promise<void> }>
}
type PickerWindow = Window & {
  showOpenFilePicker?: (options: unknown) => Promise<Handle[]>
  showSaveFilePicker?: (options: unknown) => Promise<Handle>
}

const MIME: Record<string, string> = {
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
  xltm: "application/vnd.ms-excel.template.macroEnabled.12",
  csv: "text/csv",
  tsv: "text/tab-separated-values",
  txt: "text/plain",
}

const pickerTypes = (types: FileType[]) =>
  types.map((t) => ({
    description: t.name,
    accept: Object.fromEntries(t.extensions.map((e) => [MIME[e] ?? "application/octet-stream", [`.${e}`]])),
  }))

const isAbort = (error: unknown) => error instanceof DOMException && error.name === "AbortError"

function download(bytes: Uint8Array, name: string) {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  const url = URL.createObjectURL(new Blob([copy]))
  const a = document.createElement("a")
  a.href = url
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function inputPick(types: FileType[]): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement("input")
    input.type = "file"
    input.accept = types.flatMap((t) => t.extensions.map((e) => `.${e}`)).join(",")
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.addEventListener("cancel", () => resolve(null))
    input.click()
  })
}

export function webFiles(): FilePort {
  const w = (typeof window !== "undefined" ? window : undefined) as PickerWindow | undefined
  const access = Boolean(w?.showOpenFilePicker && w?.showSaveFilePicker)
  return {
    saveInPlace: access,
    remembersFiles: false,
    async pickOpen(types) {
      if (access) {
        try {
          const [handle] = await w!.showOpenFilePicker!({ multiple: false, types: pickerTypes(types) })
          const file = await handle.getFile()
          return { ref: { name: file.name, handle }, bytes: new Uint8Array(await file.arrayBuffer()) }
        } catch (error) {
          if (isAbort(error)) return null
          throw error
        }
      }
      const file = await inputPick(types)
      return file ? { ref: { name: file.name }, bytes: new Uint8Array(await file.arrayBuffer()) } : null
    },
    async read(ref: FileRef) {
      const handle = ref.handle as Handle | undefined
      if (!handle) throw new Error("This browser can't open the file again; use Open to pick it.")
      return new Uint8Array(await (await handle.getFile()).arrayBuffer())
    },
    async pickSave(suggestedName, types) {
      if (!access) return { name: suggestedName }
      try {
        const handle = await w!.showSaveFilePicker!({ suggestedName, types: pickerTypes(types) })
        return { name: handle.name, handle }
      } catch (error) {
        if (isAbort(error)) return null
        throw error
      }
    },
    async write(ref, bytes) {
      const handle = ref.handle as Handle | undefined
      if (!handle) {
        download(bytes, ref.name)
        return
      }
      const copy = new Uint8Array(bytes.byteLength)
      copy.set(bytes)
      const stream = await handle.createWritable()
      await stream.write(new Blob([copy]))
      await stream.close()
    },
    async exists(ref) {
      return Boolean(ref.handle)
    },
  }
}
