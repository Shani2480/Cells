// Prints one version's release notes as Markdown for its GitHub release (default: the newest), or, with --metainfo,
// checks that packaging/flatpak/app.cells.desktop.metainfo.xml lists every version. The notes are the app's own
// (app/src/about/release-notes.ts), so the release page says what What's New says.
//   node scripts/release-notes.mjs [version] > notes.md
import { readFileSync } from "node:fs"
import path from "node:path"

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..")
const { RELEASE_NOTES } = await import(path.join(root, "app/src/about/release-notes.ts"))

if (process.argv[2] === "--metainfo") {
  const xml = readFileSync(path.join(root, "packaging/flatpak/app.cells.desktop.metainfo.xml"), "utf8")
  const listed = [...xml.matchAll(/<release version="([^"]+)" date="([^"]+)"/g)].map((m) => `${m[1]} ${m[2]}`)
  const wanted = RELEASE_NOTES.map((n) => `${n.version} ${n.date}`)
  if (listed.join("\n") !== wanted.join("\n")) {
    console.error(`The metainfo's <releases> should be, newest first:\n${wanted.join("\n")}`)
    process.exit(1)
  }
  console.log("Metainfo releases match the release notes")
} else {
  const version = (process.argv[2] ?? RELEASE_NOTES[0].version).replace(/^v/, "")
  const entry = RELEASE_NOTES.find((n) => n.version === version)
  if (!entry) {
    console.error(`No release notes for ${version}`)
    process.exit(1)
  }
  console.log(`## What's new in Cells ${entry.version}\n\n${entry.notes.map((n) => `- ${n}`).join("\n")}`)
}
