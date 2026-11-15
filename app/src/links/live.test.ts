import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { MemoryLinks } from "@/test/memory-links"
import { LiveLinks, parseWants, samePath, snapshotName, type LinkSnapshot } from "./live"

const path = "/home/alex/Q1 Budget.xlsx"

function setup(wants: { file: string; range: string }[]) {
  const port = new MemoryLinks()
  port.wantedText = JSON.stringify({ version: 1, links: wants })
  const controller = new SheetController()
  const s = controller.wb.sheets[0]
  s.cells.set(0, 0, { value: "Rent" })
  s.cells.set(0, 1, { value: 1200 })
  s.cells.set(1, 0, { value: "Water" })
  s.cells.set(1, 1, { value: 45.5 })
  const doc = { path: path as string | null, dirty: false, controller }
  const live = new LiveLinks(port, () => doc, 0)
  const snapshot = () => JSON.parse(port.published.get(snapshotName(path)) ?? "null") as LinkSnapshot | null
  return { port, controller, doc, live, snapshot }
}

describe("live links for Inkspire", () => {
  it("writes what a wanted range holds, keyed as Inkspire asked", async () => {
    const { live, snapshot } = setup([
      { file: path, range: "Sheet1!R1C1:R2C2" },
      { file: path, range: "Nowhere!A1" },
      { file: "/home/alex/Other.xlsx", range: "Sheet1!A1" },
    ])
    expect(await live.refresh()).toBe(true)
    const snap = snapshot()!
    expect(snap).toMatchObject({ app: "cells", version: 1, file: path, saved: true })
    expect(Object.keys(snap.ranges)).toEqual(["Sheet1!R1C1:R2C2", "Nowhere!A1"])
    expect(snap.ranges["Sheet1!R1C1:R2C2"]).toMatchObject({
      sheet: "Sheet1",
      range: "A1:B2",
      text: "Rent\t1200\r\nWater\t45.5\r\n",
      values: [
        ["Rent", 1200],
        ["Water", 45.5],
      ],
    })
    expect((snap.ranges["Sheet1!R1C1:R2C2"] as { html: string }).html).toContain("<table")
    expect(snap.ranges["Nowhere!A1"]).toEqual({ error: "#REF!" })
  })

  it("writes again only when the range changed", async () => {
    const { port, controller, doc, live, snapshot } = setup([{ file: path, range: "Sheet1!B1:B2" }])
    await live.refresh()
    expect(await live.refresh()).toBe(false)
    expect(port.writes).toBe(1)
    controller.goTo({ r0: 0, c0: 1, r1: 0, c1: 1 })
    controller.beginEdit("enter", "1300")
    controller.commitEdit()
    doc.dirty = true
    expect(await live.refresh()).toBe(true)
    expect(snapshot()!.ranges["Sheet1!B1:B2"]).toMatchObject({ values: [[1300], [45.5]] })
    expect(snapshot()!.saved).toBe(false)
  })

  it("writes nothing for a workbook never saved, or one Inkspire doesn't link to", async () => {
    const { port, doc, live } = setup([{ file: "/home/alex/Other.xlsx", range: "Sheet1!A1" }])
    expect(await live.refresh()).toBe(false)
    doc.path = null
    expect(await live.refresh()).toBe(false)
    port.wantedText = "not json"
    expect(await live.refresh()).toBe(false)
    expect(port.writes).toBe(0)
  })

  it("names snapshots and compares paths the same way Inkspire must", () => {
    // FNV-1a 64 of "" and "a" (the published test values).
    expect(snapshotName("")).toBe("cbf29ce484222325")
    expect(snapshotName("a")).toBe("af63dc4c8601ec8c")
    expect(snapshotName(path)).toMatch(/^[0-9a-f]{16}$/)
    expect(samePath("C:\\Users\\alex\\A.xlsx", "c:/users/alex/a.xlsx")).toBe(true)
    expect(samePath("/home/alex/A.xlsx", "/home/alex/a.xlsx")).toBe(false)
    expect(parseWants('{"links":[{"file":"/a.xlsx","range":"A1"},{"file":3}]}')).toEqual([
      { file: "/a.xlsx", range: "A1" },
    ])
  })
})
