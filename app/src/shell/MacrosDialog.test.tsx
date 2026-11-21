import { act } from "react"
import { createRoot } from "react-dom/client"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { MacrosDialog } from "./MacrosDialog"

describe("Macros dialog", () => {
  it("lists the workbook's macros, keeps Run off and shows a macro's code", async () => {
    const c = new SheetController()
    c.wb.vba = [
      {
        name: "Tools",
        kind: "module",
        code: 'Sub FillDates()\n  Range("A1").Value = Date\nEnd Sub\nPrivate Sub Helper()\nEnd Sub',
      },
      { name: "ThisWorkbook", kind: "document", code: 'Sub Greet()\n  MsgBox "Hi"\nEnd Sub' },
    ]
    const host = document.createElement("div")
    document.body.append(host)
    const root = createRoot(host)
    await act(async () => root.render(<MacrosDialog controller={c} onClose={() => {}} />))
    const options = [...document.body.querySelectorAll('[role="option"]')].map((o) => o.textContent)
    expect(options).toEqual(["FillDates", "ThisWorkbook.Greet"])
    const button = (t: string) => [...document.body.querySelectorAll("button")].find((b) => b.textContent === t)!
    expect(button("Run").disabled).toBe(true)
    await act(async () => button("Edit").click())
    expect(document.body.querySelector(".macro-code")?.textContent).toContain('Range("A1").Value = Date')
    root.unmount()
  })
})
