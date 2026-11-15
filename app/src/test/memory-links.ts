import type { LinkPort } from "@cells/platform"

/** A LinkPort over memory, for tests: Inkspire's `wanted.json` text, the snapshots written, and its edit folder. */
export class MemoryLinks implements LinkPort {
  wantedText: string | null = null
  readonly published = new Map<string, string>()
  writes = 0
  editFolder = "/home/alex/.cache/inkspire/edit/"

  async wanted() {
    return this.wantedText
  }

  async publish(name: string, text: string) {
    this.writes++
    this.published.set(name, text)
  }

  async isEditCopy(path: string) {
    return path.startsWith(this.editFolder)
  }
}
