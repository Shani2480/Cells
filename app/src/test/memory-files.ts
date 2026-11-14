import type { FilePort, FileRef } from "@cells/platform"

/** A FilePort over a map, for tests: `nextOpen` / `nextSave` answer the next Open and Save As dialogs. */
export class MemoryFiles implements FilePort {
  readonly saveInPlace = true
  readonly remembersFiles = true
  readonly disk = new Map<string, Uint8Array>()
  nextOpen: string | null = null
  nextSave: string | null = null
  writes: string[] = []
  failWrite: Error | null = null

  async pickOpen() {
    const path = this.nextOpen
    this.nextOpen = null
    if (!path) return null
    return { ref: this.ref(path), bytes: await this.read(this.ref(path)) }
  }

  async read(ref: FileRef) {
    const bytes = ref.path ? this.disk.get(ref.path) : undefined
    if (!bytes) throw new Error(`ENOENT: ${ref.path}`)
    return bytes
  }

  async pickSave() {
    const path = this.nextSave
    this.nextSave = null
    return path ? this.ref(path) : null
  }

  async write(ref: FileRef, bytes: Uint8Array) {
    if (this.failWrite) throw this.failWrite
    this.disk.set(ref.path!, bytes)
    this.writes.push(ref.path!)
  }

  async exists(ref: FileRef) {
    return Boolean(ref.path && this.disk.has(ref.path))
  }

  private dropHandler: ((file: { ref: FileRef; bytes: Uint8Array }) => void) | null = null

  onDrop(handler: (file: { ref: FileRef; bytes: Uint8Array }) => void) {
    this.dropHandler = handler
    return () => {
      this.dropHandler = null
    }
  }

  /** A file dropped on the window. */
  drop(path: string) {
    this.dropHandler?.({ ref: this.ref(path), bytes: this.disk.get(path)! })
  }

  ref(path: string): FileRef {
    return { name: path.split("/").pop()!, path }
  }
}
