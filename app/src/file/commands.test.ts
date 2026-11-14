import { saveXlsx } from "@cells/io"
import { Workbook } from "@cells/core"
import { beforeEach, describe, expect, it, vi } from "vitest"
import { MemoryFiles } from "@/test/memory-files"
import { FileCommands, type FileUi, type SaveChoice } from "./commands"
import { newDocument, type WorkbookDocument } from "./document"

function setup(answer: SaveChoice = "cancel") {
  const files = new MemoryFiles()
  const shown: WorkbookDocument[] = []
  const ui: FileUi & { problems: string[]; asked: string[]; notices: string[][] } = {
    problems: [],
    asked: [],
    notices: [],
    show: (doc) => shown.push(doc),
    askSave: async (title) => {
      ui.asked.push(title)
      return answer
    },
    problem: (p) => ui.problems.push(p.kind),
    progress: () => {},
    notice: (lines) => ui.notices.push(lines),
    recentChanged: () => {},
  }
  const commands = new FileCommands(files, ui, newDocument())
  return { files, ui, shown, commands }
}

/** Types into A1 and presses Enter, as a person would. */
function enter(doc: WorkbookDocument, text: string) {
  const c = doc.controller
  c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
  c.beginEdit("enter", text)
  c.commitEdit()
}

const budget = async (fileName = "Budget.xlsx") => {
  const wb = new Workbook()
  const s = wb.sheets[0]
  s.cells.set(0, 0, { value: "Rent" })
  s.cells.set(0, 1, { value: 1200 })
  s.cells.set(1, 1, { value: 14400, input: "=B1*12" })
  wb.calc.rebuild()
  return saveXlsx(wb, { fileName })
}

describe("FileCommands", () => {
  beforeEach(() => localStorage.clear())

  it("Save on a new workbook goes to Save As, then saves in place", async () => {
    const { files, commands } = setup()
    const doc = commands.doc
    enter(doc, "5")
    expect(doc.dirty).toBe(true)
    expect(await commands.save()).toBe(false) // Save As cancelled
    files.nextSave = "/home/alex/Sums.xlsx"
    expect(await commands.save()).toBe(true)
    expect(doc.fileName).toBe("Sums.xlsx")
    expect(doc.dirty).toBe(false)
    enter(doc, "6")
    expect(await commands.save()).toBe(true)
    expect(files.writes).toEqual(["/home/alex/Sums.xlsx", "/home/alex/Sums.xlsx"])
    expect(commands.recent[0]).toMatchObject({ name: "Sums.xlsx", path: "/home/alex/Sums.xlsx" })
  })

  it("opens a workbook, replacing an untouched Book without asking", async () => {
    const { files, ui, shown, commands } = setup()
    files.disk.set("/home/alex/Budget.xlsx", await budget())
    files.nextOpen = "/home/alex/Budget.xlsx"
    expect(await commands.open()).toBe(true)
    expect(ui.asked).toEqual([])
    const doc = shown.at(-1)!
    expect(doc.title).toBe("Budget")
    expect(doc.dirty).toBe(false)
    const sheet = doc.controller.wb.sheets[0]
    expect(sheet.cells.get(0, 0)?.value).toBe("Rent")
    expect(sheet.cells.get(1, 1)?.value).toBe(14400)
  })

  it("asks to save changes; Cancel keeps the workbook, Don't Save replaces it", async () => {
    const cancel = setup("cancel")
    enter(cancel.commands.doc, "1")
    expect(await cancel.commands.newWorkbook()).toBe(false)
    expect(cancel.ui.asked).toHaveLength(1)
    expect(cancel.shown).toHaveLength(0)

    const discard = setup("discard")
    enter(discard.commands.doc, "1")
    expect(await discard.commands.close()).toBe(true)
    expect(discard.shown).toHaveLength(1)
    expect(discard.commands.doc.dirty).toBe(false)
  })

  it("Save in the prompt saves before going on, and a failed save stops", async () => {
    const { files, ui, commands } = setup("save")
    enter(commands.doc, "1")
    files.nextSave = "/home/alex/Kept.xlsx"
    expect(await commands.confirmClose()).toBe(true)
    expect(files.disk.has("/home/alex/Kept.xlsx")).toBe(true)

    enter(commands.doc, "2")
    files.failWrite = new Error("EACCES")
    expect(await commands.confirmClose()).toBe(false)
    expect(ui.problems).toEqual(["write"])
  })

  it("names what's wrong with a file that won't open", async () => {
    const { files, ui, commands } = setup()
    files.disk.set("/home/alex/notes.xlsx", new TextEncoder().encode("just text"))
    files.disk.set("/home/alex/old.xls", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0]))
    files.disk.set("/home/alex/broken.xlsx", new Uint8Array([0x50, 0x4b, 3, 4, 0, 0, 0]))
    for (const path of ["/home/alex/notes.xlsx", "/home/alex/broken.xlsx"]) {
      files.nextOpen = path
      expect(await commands.open()).toBe(false)
    }
    expect(await commands.openRecent({ name: "gone.xlsx", path: "/home/alex/gone.xlsx", opened: 0 })).toBe(false)
    expect(ui.problems).toEqual(["type", "damaged", "read"])
  })

  it("a template opens as a new workbook named after it", async () => {
    const { files, shown, commands } = setup()
    files.disk.set("/home/alex/Invoice.xltx", await budget("Invoice.xltx"))
    files.nextOpen = "/home/alex/Invoice.xltx"
    expect(await commands.open()).toBe(true)
    expect(shown.at(-1)!.fileName).toBe("Invoice1.xlsx")
    expect(shown.at(-1)!.ref).toBeNull()
  })

  it("AutoRecover keeps a copy only while there are new changes", async () => {
    const { commands } = setup()
    const keep = vi.spyOn(await import("./recovery"), "keepRecovery")
    expect(await commands.autoRecover()).toBe(false)
    enter(commands.doc, "1")
    expect(await commands.autoRecover()).toBe(true)
    expect(await commands.autoRecover()).toBe(false)
    expect(keep).toHaveBeenCalledTimes(1)
  })

  it("new workbooks are numbered Book2, Book3… after the first", async () => {
    const { shown, commands } = setup()
    expect(commands.doc.fileName).toBe("Book1.xlsx")
    await commands.newWorkbook()
    await commands.close()
    expect(shown.map((d) => d.title)).toEqual(["Book2", "Book3"])
  })
})
