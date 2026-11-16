import { afterEach, describe, expect, it, vi } from "vitest"
import { parseOpenRequest } from "@/links/link"
import {
  fileNameOf,
  installBackHandler,
  mimeTypeOf,
  onOpenedDocuments,
  pickerTypes,
  pickSaveDocument,
} from "./android-files"

const host = globalThis as Record<string, unknown>

afterEach(() => {
  delete host.CellsFiles
  delete host.__cellsPicked
  delete host.__cellsOpened
  delete host.__cellsBack
})

describe("Android documents", () => {
  it("names a file by its path, or by what a document link holds", () => {
    expect(fileNameOf("/home/alex/Q1 Report.xlsx")).toBe("Q1 Report.xlsx")
    expect(fileNameOf("C:\\Users\\alex\\Prices.xlsx")).toBe("Prices.xlsx")
    expect(
      fileNameOf("content://com.android.externalstorage.documents/document/primary%3ADocuments%2FBudget.xlsx"),
    ).toBe("Budget.xlsx")
  })

  it("asks Android for the name when the link has only an id", () => {
    const displayName = vi.fn(() => "Prices.xlsx")
    host.CellsFiles = { displayName }
    const uri = "content://com.android.providers.downloads.documents/document/msf%3A1000000018"
    expect(fileNameOf(uri)).toBe("Prices.xlsx")
    expect(fileNameOf(uri)).toBe("Prices.xlsx")
    expect(displayName).toHaveBeenCalledTimes(1)
  })

  it("offers the picker the types a file manager may label spreadsheets with", () => {
    expect(pickerTypes(["xlsx"])).toEqual([mimeTypeOf("xlsx"), "application/octet-stream"])
    expect(pickerTypes(["csv"])).toEqual(["text/csv", "text/comma-separated-values", "text/plain"])
    expect(mimeTypeOf("XLSM")).toBe("application/vnd.ms-excel.sheet.macroEnabled.12")
  })

  it("saves through the system picker with the file's type", async () => {
    const createDocument = vi.fn((_name: string, _type: string) => {
      ;(host.__cellsPicked as (uri: string | null) => void)("content://docs/document/7")
    })
    host.CellsFiles = { createDocument }
    await expect(pickSaveDocument("Book1.xlsx")).resolves.toBe("content://docs/document/7")
    expect(createDocument).toHaveBeenCalledWith("Book1.xlsx", mimeTypeOf("xlsx"))
    expect(host.__cellsPicked).toBeUndefined()
  })

  it("hands over files from other apps at launch and later", () => {
    const queue = [["content://mail/attachment/3"], [], ["content://drive/doc/9"]]
    host.CellsFiles = { takeOpened: () => JSON.stringify(queue.shift() ?? []) }
    const got: string[][] = []
    const stop = onOpenedDocuments((uris) => got.push(uris))
    expect(got).toEqual([["content://mail/attachment/3"]])
    ;(host.__cellsOpened as () => void)()
    ;(host.__cellsOpened as () => void)()
    expect(got).toEqual([["content://mail/attachment/3"], ["content://drive/doc/9"]])
    stop()
    expect(host.__cellsOpened).toBeUndefined()
    expect(parseOpenRequest("content://mail/attachment/3")).toEqual({ path: "content://mail/attachment/3" })
  })

  it("Back closes what is open with Esc, and says when nothing was", () => {
    let editing = false
    const stop = installBackHandler(() => editing)
    const back = host.__cellsBack as () => boolean
    const keys: string[] = []
    const listen = (event: KeyboardEvent) => keys.push(event.key)
    window.addEventListener("keydown", listen)
    try {
      expect(back()).toBe(false)
      const sheet = document.createElement("div")
      sheet.className = "phone-sheet-layer"
      document.body.append(sheet)
      // Sent to the sheet even while the grid's editor holds the focus and keeps Esc.
      const editor = document.createElement("textarea")
      editor.addEventListener("keydown", (event) => event.preventDefault())
      document.body.append(editor)
      editor.focus()
      let atSheet = false
      sheet.addEventListener("keydown", () => (atSheet = true))
      expect(back()).toBe(true)
      expect(atSheet).toBe(true)
      editor.remove()
      sheet.remove()
      editing = true
      expect(back()).toBe(true)
      expect(keys).toEqual(["Escape", "Escape"])
    } finally {
      window.removeEventListener("keydown", listen)
      stop()
    }
    expect(host.__cellsBack).toBeUndefined()
  })
})
