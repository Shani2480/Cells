import { decodeText, defaultImport, saveXlsx, type TextImport } from "@cells/io"
import { Workbook } from "@cells/core"
import { beforeEach, describe, expect, it, vi } from "vitest"
import { MemoryFiles } from "@/test/memory-files"
import { EMBEDDED_SAVE_MS, FileCommands, type FileUi, type SaveChoice } from "./commands"
import { newDocument, type WorkbookDocument } from "./document"

function setup(answer: SaveChoice = "cancel", wizard: Partial<TextImport> | null = {}, ok = true) {
  const files = new MemoryFiles()
  const shown: WorkbookDocument[] = []
  const ui: FileUi & {
    problems: string[]
    asked: string[]
    notices: string[][]
    wizards: string[]
    confirms: string[]
  } = {
    problems: [],
    asked: [],
    notices: [],
    wizards: [],
    confirms: [],
    textImport: async (bytes, fileName) => {
      ui.wizards.push(fileName)
      return wizard ? { ...defaultImport(decodeText(bytes).text, fileName), ...wizard } : null
    },
    confirm: async (message) => {
      ui.confirms.push(message)
      return ok
    },
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

  it("a CSV opens at once, keeps its delimiter and encoding, and saves back the same way", async () => {
    const { files, ui, shown, commands } = setup()
    // Windows-1252, semicolons, as a European export would be.
    const ansi = new Uint8Array([
      ...new TextEncoder().encode("Item;Cost\r\nCaf"),
      0xe9,
      ...new TextEncoder().encode(";3\r\n"),
    ])
    files.disk.set("/home/alex/export.csv", ansi)
    files.nextOpen = "/home/alex/export.csv"
    expect(await commands.open()).toBe(true)
    expect(ui.wizards).toEqual([])
    const doc = shown.at(-1)!
    const sheet = doc.controller.wb.sheets[0]
    expect(sheet.name).toBe("export")
    expect(sheet.cells.get(1, 0)?.value).toBe("Café")
    expect(sheet.cells.get(1, 1)?.value).toBe(3)
    enter(doc, "4")
    expect(await commands.save()).toBe(true)
    expect([...files.disk.get("/home/alex/export.csv")!]).toEqual([
      ...new TextEncoder().encode("4;Cost\r\nCaf"),
      0xe9,
      ...new TextEncoder().encode(";3\r\n"),
    ])
  })

  it("a .txt goes through the Text Import Wizard; Cancel there opens nothing", async () => {
    const cancelled = setup("cancel", null)
    cancelled.files.disk.set("/home/alex/log.txt", new TextEncoder().encode("a\tb\n1\t2\n"))
    cancelled.files.nextOpen = "/home/alex/log.txt"
    expect(await cancelled.commands.open()).toBe(false)
    expect(cancelled.ui.wizards).toEqual(["log.txt"])
    expect(cancelled.shown).toHaveLength(0)

    const { files, shown, commands } = setup("cancel", { columns: ["text"] })
    files.disk.set("/home/alex/log.txt", new TextEncoder().encode("id\tqty\n007\t2\n"))
    files.nextOpen = "/home/alex/log.txt"
    expect(await commands.open()).toBe(true)
    const sheet = shown.at(-1)!.controller.wb.sheets[0]
    expect(sheet.cells.get(1, 0)?.value).toBe("007")
    expect(sheet.cells.get(1, 1)?.value).toBe(2)
  })

  it("saving several sheets as CSV asks first and writes the active sheet", async () => {
    const no = setup("cancel", {}, false)
    no.commands.doc.controller.sheets.insert()
    no.files.nextSave = "/home/alex/out.csv"
    expect(await no.commands.saveAs()).toBe(false)
    expect(no.ui.confirms).toHaveLength(1)
    expect(no.files.writes).toEqual([])

    const { files, commands } = setup()
    const c = commands.doc.controller
    enter(commands.doc, "a,b")
    c.sheets.insert()
    enter(commands.doc, "second")
    files.nextSave = "/home/alex/out.csv"
    expect(await commands.saveAs()).toBe(true)
    expect(new TextDecoder("utf-8", { ignoreBOM: true }).decode(files.disk.get("/home/alex/out.csv")!)).toBe(
      "\ufeffsecond\r\n",
    )
    expect(commands.doc.fileName).toBe("out.csv")
    expect(c.wb.sheets).toHaveLength(2)
  })

  it("AutoRecover keeps a CSV's copy as a whole workbook", async () => {
    const { files, commands } = setup()
    const keep = vi.spyOn(await import("./recovery"), "keepRecovery")
    files.disk.set("/home/alex/list.csv", new TextEncoder().encode("x\n"))
    files.nextOpen = "/home/alex/list.csv"
    await commands.open()
    enter(commands.doc, "y")
    expect(await commands.autoRecover()).toBe(true)
    const kept = keep.mock.calls.at(-1)![0]
    expect(kept.fileName).toBe("list.xlsx")
    expect([...kept.bytes.slice(0, 2)]).toEqual([0x50, 0x4b])
  })

  it("Data › From Text/CSV puts the file on a new sheet after the active one, undone in one step", async () => {
    const { files, ui, commands } = setup("cancel", { delimiters: [";"], columns: ["general", "skip", "text"] })
    files.disk.set("/home/alex/rates.csv", new TextEncoder().encode("1;x;01\n2;y;02\n"))
    files.nextOpen = "/home/alex/rates.csv"
    const c = commands.doc.controller
    expect(await commands.importText()).toBe(true)
    expect(ui.wizards).toEqual(["rates.csv"])
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "rates"])
    expect(c.sheetIndex).toBe(1)
    expect(c.wb.sheets[1].cells.get(1, 0)?.value).toBe(2)
    expect(c.wb.sheets[1].cells.get(1, 1)?.value).toBe("02")
    files.nextOpen = "/home/alex/rates.csv"
    await commands.importText()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "rates", "rates (2)"])
    c.undo()
    c.undo()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet1"])
  })
})

describe("Queries & Connections (linked text files)", () => {
  beforeEach(() => localStorage.clear())
  const csv = (text: string) => new TextEncoder().encode(text)

  async function imported() {
    const t = setup("cancel", { delimiters: [","] })
    t.files.disk.set("/home/alex/sales.csv", csv("Item,Qty\nPens,4\nInk,2\nPaper,9\n"))
    t.files.nextOpen = "/home/alex/sales.csv"
    await t.commands.importText()
    return { ...t, c: t.commands.doc.controller }
  }

  it("an import stays linked to its file, with names from the file", async () => {
    const { c } = await imported()
    expect(c.sheet.queryTables).toMatchObject([
      { name: "sales", connection: "sales", source: "/home/alex/sales.csv", range: { r0: 0, c0: 0, r1: 3, c1: 1 } },
    ])
  })

  it("Refresh reads the file again: new data in, old data beyond it cleared, one step to undo", async () => {
    const { files, commands, c } = await imported()
    files.disk.set("/home/alex/sales.csv", csv("Item,Qty,Price\nPens,5,1.5\n"))
    c.sheet.cells.set(10, 5, { value: "mine" })
    expect(await commands.refreshQuery(c.sheetIndex, "sales")).toBe(true)
    expect(c.sheet.cells.get(1, 1)?.value).toBe(5)
    expect(c.sheet.cells.get(1, 2)?.value).toBe(1.5)
    expect(c.sheet.cells.get(2, 0)).toBeUndefined()
    expect(c.sheet.cells.get(10, 5)?.value).toBe("mine")
    expect(c.sheet.queryTables[0].range).toEqual({ r0: 0, c0: 0, r1: 1, c1: 2 })
    c.undo()
    expect(c.sheet.cells.get(3, 0)?.value).toBe("Paper")
    expect(c.sheet.queryTables[0].range).toEqual({ r0: 0, c0: 0, r1: 3, c1: 1 })
  })

  it("a file that moved is looked for with the Open dialog, and the link follows it", async () => {
    const { files, commands, c } = await imported()
    files.disk.delete("/home/alex/sales.csv")
    expect(await commands.refreshQuery(c.sheetIndex, "sales")).toBe(false) // Open cancelled
    files.disk.set("/home/alex/new/sales.csv", csv("Item\nCards\n"))
    files.nextOpen = "/home/alex/new/sales.csv"
    expect(await commands.refreshAll()).toBe(1)
    expect(c.sheet.cells.get(1, 0)?.value).toBe("Cards")
    expect(c.sheet.queryTables[0].source).toBe("/home/alex/new/sales.csv")
  })

  it("the link moves with inserted rows; Delete keeps the data", async () => {
    const { c } = await imported()
    c.transact([{ kind: "insertRows", sheet: c.sheetIndex, at: 0, count: 2 }], { label: "Insert", before: c.selection })
    expect(c.sheet.queryTables[0].range).toEqual({ r0: 2, c0: 0, r1: 5, c1: 1 })
    c.data.setQueryTable(c.sheetIndex, "sales", null)
    expect(c.sheet.queryTables).toEqual([])
    expect(c.sheet.cells.get(3, 0)?.value).toBe("Pens")
    c.undo()
    expect(c.sheet.queryTables).toHaveLength(1)
  })
})

describe("Edit Data (a chart's workbook from Inkspire)", () => {
  const path = "/home/alex/.cache/inkspire/edit/Report-Chart 1.xlsx"

  it("saves each change back a moment later, without asking or adding it to Recent", async () => {
    const { files, ui, commands } = setup()
    files.disk.set(path, await budget("Report-Chart 1.xlsx"))
    const recent = vi.fn()
    ui.recentChanged = recent
    expect(await commands.openEmbedded(path, "Report.docx")).toBe(true)
    expect(commands.doc.title).toBe("Chart in Report.docx")
    vi.useFakeTimers()
    try {
      enter(commands.doc, "Water")
      expect(files.writes).toEqual([])
      await vi.advanceTimersByTimeAsync(EMBEDDED_SAVE_MS + 50)
      expect(files.writes).toEqual([path])
      expect(commands.doc.dirty).toBe(false)
    } finally {
      vi.useRealTimers()
    }
    expect(recent).not.toHaveBeenCalled()
    // Closing saves what's left, never asks.
    enter(commands.doc, "Gas")
    expect(await commands.close()).toBe(true)
    expect(ui.asked).toEqual([])
    expect(files.writes).toEqual([path, path])
  })

  it("waits while a cell is being typed in", async () => {
    const { files, commands } = setup()
    files.disk.set(path, await budget("Report-Chart 1.xlsx"))
    await commands.openEmbedded(path, "Report.docx")
    vi.useFakeTimers()
    try {
      enter(commands.doc, "Water")
      commands.doc.controller.beginEdit("enter", "Ga")
      await vi.advanceTimersByTimeAsync(EMBEDDED_SAVE_MS * 3)
      expect(files.writes).toEqual([])
      commands.doc.controller.commitEdit()
      await vi.advanceTimersByTimeAsync(EMBEDDED_SAVE_MS + 50)
      expect(files.writes).toEqual([path])
    } finally {
      vi.useRealTimers()
    }
  })

  it("Save As makes it an ordinary workbook again", async () => {
    const { files, ui, commands } = setup("discard")
    files.disk.set(path, await budget("Report-Chart 1.xlsx"))
    await commands.openEmbedded(path, "Report.docx")
    files.nextSave = "/home/alex/Chart data.xlsx"
    expect(await commands.saveAs()).toBe(true)
    expect(commands.doc.title).toBe("Chart data")
    enter(commands.doc, "Water")
    await commands.close()
    expect(ui.asked).toEqual(["Chart data"])
  })
})
