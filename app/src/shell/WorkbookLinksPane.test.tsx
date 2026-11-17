import { linkKey } from "@cells/core"
import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import type { FileCommands } from "@/file/commands"
import { SheetController } from "@/sheet/controller"
import { WorkbookLinksPane } from "./WorkbookLinksPane"

describe("Workbook Links pane", () => {
  it("lists linked workbooks and ones formulas name without values yet; buttons act on each", async () => {
    const c = new SheetController()
    c.wb.transact(
      [
        {
          kind: "externalLinks",
          links: [
            {
              name: "Price List.xlsx",
              target: "Price List.xlsx",
              absolute: "C:\\Users\\alex\\Price List.xlsx",
              sheets: [{ name: "Prices", cells: new Map([[linkKey(0, 0), 2]]) }],
              names: [],
            },
          ],
        },
        {
          kind: "setCells",
          sheet: 0,
          cells: [
            { row: 0, col: 0, data: { input: "='[Price List.xlsx]Prices'!A1", value: 0 } },
            { row: 1, col: 0, data: { input: "=[Stock.xlsx]Sheet1!B2", value: 0 } },
          ],
        },
      ],
      { label: "Links" },
    )
    c.data.toggleLinks(true)
    const updateLink = vi.fn(async () => true)
    const breakLink = vi.fn(async () => true)
    const files = { updateLink, breakLink, changeLinkSource: vi.fn() } as unknown as FileCommands
    render(<WorkbookLinksPane controller={c} files={files} onDone={() => {}} />)
    expect(screen.getByText("C:\\Users\\alex\\Price List.xlsx")).toBeTruthy()
    expect(screen.getByText("Using saved values")).toBeTruthy()
    expect(screen.getByText("Not read yet")).toBeTruthy()
    expect(screen.getAllByText("Used by 1 formula")).toHaveLength(2)
    fireEvent.click(screen.getByRole("button", { name: "Update values from Price List.xlsx" }))
    expect(updateLink).toHaveBeenCalledWith("Price List.xlsx")
    expect(await screen.findByText("Updated")).toBeTruthy()
    fireEvent.click(screen.getByRole("button", { name: "Break link to Stock.xlsx" }))
    expect(breakLink).toHaveBeenCalledWith("Stock.xlsx")
  })
})
