import { defaultIconThresholds } from "@cells/core"
import { describe, expect, it, vi } from "vitest"
import { DEFAULT_PROTECTION } from "@cells/core"
import { SheetController } from "./controller"

/** A controller holding a small table: headers in row 1, then Name / Team / Score. */
function table() {
  const c = new SheetController()
  const rows = [
    ["Name", "Team", "Score"],
    ["Ann", "Blue", "5"],
    ["Bob", "Red", "7"],
    ["Cy", "Blue", "9"],
  ]
  rows.forEach((row, r) =>
    row.forEach((text, col) => {
      c.goTo({ r0: r, c0: col, r1: r, c1: col })
      c.beginEdit("enter", text)
      c.commitEdit()
    }),
  )
  return c
}
const column = (c: SheetController, col: number) => [0, 1, 2, 3].map((r) => c.sheet.cells.get(r, col)?.value)

describe("Sort & Filter: sorting", () => {
  it("Sort Z to A on one cell sorts the data around it by its column, keeping the header", () => {
    const c = table()
    c.goTo({ r0: 2, c0: 2, r1: 2, c1: 2 })
    expect(c.data.labels()).toEqual({ ascending: "Sort Smallest to Largest", descending: "Sort Largest to Smallest" })
    c.data.quickSort(true)
    expect(column(c, 0)).toEqual(["Name", "Cy", "Bob", "Ann"])
    expect(column(c, 2)).toEqual(["Score", 9, 7, 5])
    // The whole table is selected afterwards; one Undo puts it back.
    expect(c.selection.ranges[0]).toEqual({ r0: 0, c0: 0, r1: 3, c1: 2 })
    c.undo()
    expect(column(c, 0)).toEqual(["Name", "Ann", "Bob", "Cy"])
  })

  it("text columns say A to Z; a selection with data beside it asks first", () => {
    const c = table()
    c.goTo({ r0: 1, c0: 1, r1: 3, c1: 1 })
    expect(c.data.labels().ascending).toBe("Sort A to Z")
    const warn = vi.fn((then: (expand: boolean) => void) => then(false))
    c.shell = { ...c.shell, sortWarning: warn }
    c.data.quickSort(true)
    expect(warn).toHaveBeenCalledOnce()
    // Only the selection was sorted.
    expect(column(c, 1)).toEqual(["Team", "Red", "Blue", "Blue"])
    expect(column(c, 0)).toEqual(["Name", "Ann", "Bob", "Cy"])
    c.undo()
    c.shell = { ...c.shell, sortWarning: (then) => then(true) }
    c.goTo({ r0: 1, c0: 1, r1: 3, c1: 1 })
    c.data.quickSort(false)
    expect(column(c, 0)).toEqual(["Name", "Ann", "Cy", "Bob"])
  })

  it("merged cells and protected sheets refuse with Excel's messages", () => {
    const c = table()
    const alert = vi.fn()
    c.shell = { ...c.shell, alert }
    c.sheet.merges.add({ r0: 1, c0: 0, r1: 1, c1: 1 })
    c.goTo({ r0: 2, c0: 2, r1: 2, c1: 2 })
    c.data.quickSort(false)
    expect(alert).toHaveBeenLastCalledWith("To do this, all the merged cells need to be the same size.")
    c.sheet.merges.remove([{ r0: 1, c0: 0, r1: 1, c1: 1 }])
    c.sheets.protect({ ...c.sheet.protection, selectLocked: true, selectUnlocked: true } as never, "")
    c.data.quickSort(true)
    expect(alert.mock.calls.at(-1)?.[0]).toMatch(/protected sheet/)
    expect(column(c, 2)).toEqual(["Score", 5, 7, 9])
  })

  it("Custom Sort opens the Sort dialog on the data around the cell", () => {
    const c = table()
    const sortDialog = vi.fn()
    c.shell = { ...c.shell, sortDialog }
    c.goTo({ r0: 1, c0: 0, r1: 1, c1: 0 })
    c.data.openSortDialog()
    expect(sortDialog).toHaveBeenCalledWith({ r0: 0, c0: 0, r1: 3, c1: 2 })
    expect(c.data.target().header).toBe(true)
  })
})

const protectedSheet = (autoFilter: boolean) => ({ ...DEFAULT_PROTECTION, autoFilter })

describe("Sort & Filter: AutoFilter", () => {
  const hidden = (c: SheetController) => [1, 2, 3].map((r) => c.sheet.rows.isHidden(r))

  it("Ctrl+Shift+L turns the filter on for the data around the cell, and off again showing every row", () => {
    const c = table()
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    c.keyReady({ key: "L", primary: true, shift: true, alt: false })
    expect(c.sheet.autoFilter?.range).toEqual({ r0: 0, c0: 0, r1: 3, c1: 2 })
    expect(c.data.filterColumn(1, { kind: "values", values: ["Blue"] })).toBe(true)
    expect(hidden(c)).toEqual([false, true, false])
    expect(c.data.records()).toEqual({ shown: 2, total: 3 })
    c.data.toggleFilter()
    expect(c.sheet.autoFilter).toBeNull()
    expect(hidden(c)).toEqual([false, false, false])
    c.undo()
    expect(c.sheet.autoFilter?.columns[1]).toEqual({ kind: "values", values: ["Blue"] })
    expect(hidden(c)).toEqual([false, true, false])
  })

  it("an empty cell with nothing around it can't be filtered", () => {
    const c = table()
    const alert = vi.fn()
    c.shell.alert = alert
    c.goTo({ r0: 10, c0: 10, r1: 10, c1: 10 })
    c.data.toggleFilter()
    expect(c.sheet.autoFilter).toBeNull()
    expect(alert).toHaveBeenCalled()
  })

  it("Clear shows all rows and keeps the buttons; Reapply takes in rows typed under the data", () => {
    const c = table()
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    c.data.toggleFilter()
    c.data.filterColumn(2, { kind: "custom", conditions: [{ op: "gt", value: "6" }] })
    expect(hidden(c)).toEqual([true, false, false])
    c.data.clearFilters()
    expect(c.sheet.autoFilter?.columns).toEqual({})
    expect(hidden(c)).toEqual([false, false, false])
    c.data.filterColumn(2, { kind: "custom", conditions: [{ op: "gt", value: "6" }] })
    for (const [col, text] of [
      [0, "Di"],
      [1, "Red"],
      [2, "1"],
    ] as const) {
      c.goTo({ r0: 4, c0: col, r1: 4, c1: col })
      c.beginEdit("enter", text)
      c.commitEdit()
    }
    expect(c.sheet.rows.isHidden(4)).toBe(false)
    c.data.reapply()
    expect(c.sheet.autoFilter?.range.r1).toBe(4)
    expect(c.sheet.rows.isHidden(4)).toBe(true)
  })

  it("sorting from a dropdown sorts the filter's data and marks the column", () => {
    const c = table()
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    c.data.toggleFilter()
    c.data.sortFiltered(2, true)
    expect(column(c, 2)).toEqual(["Score", 9, 7, 5])
    expect(c.sheet.autoFilter?.sort).toEqual({ col: 2, descending: true })
    c.undo()
    expect(column(c, 2)).toEqual(["Score", 5, 7, 9])
    expect(c.sheet.autoFilter?.sort).toBeUndefined()
  })

  it("a protected sheet keeps its filter unless Use AutoFilter is allowed", () => {
    const c = table()
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    c.data.toggleFilter()
    c.shell.alert = vi.fn()
    c.sheet.protection = protectedSheet(false)
    expect(c.data.filterColumn(1, { kind: "values", values: ["Red"] })).toBe(false)
    c.sheet.protection = protectedSheet(true)
    expect(c.data.filterColumn(1, { kind: "values", values: ["Red"] })).toBe(true)
    c.data.toggleFilter()
    expect(c.sheet.autoFilter).not.toBeNull()
  })
})

describe("the cell menu's Filter › and Sort ›", () => {
  const shown = (c: SheetController) => [1, 2, 3].filter((r) => !c.sheet.rows.isHidden(r))

  it("filters by the selected cell's value, turning the filter on first, and clears that column", () => {
    const c = table()
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    expect(c.data.filteredHeader()).toBeNull()
    expect(c.data.filterBySelected("value")).toBe(true)
    expect(c.sheet.autoFilter?.range).toEqual({ r0: 0, c0: 0, r1: 3, c1: 2 })
    expect(shown(c)).toEqual([1, 3])
    expect(c.data.filteredHeader()).toBe("Team")
    expect(c.data.clearColumnFilter()).toBe(true)
    expect(shown(c)).toEqual([1, 2, 3])
  })

  it("filters by the selected cell's fill and font colour", () => {
    const c = table()
    c.goTo({ r0: 2, c0: 0, r1: 2, c1: 0 })
    c.format.fill("#FFFF00")
    c.data.filterBySelected("cell")
    expect(shown(c)).toEqual([2])
    c.data.clearColumnFilter()
    c.goTo({ r0: 1, c0: 0, r1: 1, c1: 0 })
    c.data.filterBySelected("font")
    // Automatic font everywhere: every row matches.
    expect(shown(c)).toEqual([1, 2, 3])
  })

  it("filters by the selected cell's icon and puts an icon on top; colour scales count as fills", () => {
    const c = table()
    c.sheet.conditional = [
      {
        id: 1,
        ranges: [{ r0: 1, c0: 2, r1: 3, c1: 2 }],
        rule: {
          kind: "iconSet",
          set: "3Arrows",
          reverse: false,
          thresholds: defaultIconThresholds(3),
          showValue: true,
        },
        stopIfTrue: false,
      },
    ]
    c.goTo({ r0: 1, c0: 2, r1: 1, c1: 2 })
    expect(c.data.activeIcon).toEqual({ set: "3Arrows", index: 0 })
    expect(c.data.sortSelectedColorOnTop("icon")).toBe(true)
    expect(column(c, 0)).toEqual(["Name", "Ann", "Bob", "Cy"])
    c.goTo({ r0: 3, c0: 2, r1: 3, c1: 2 })
    expect(c.data.filterBySelected("icon")).toBe(true)
    expect(shown(c)).toEqual([3])
    c.data.clearColumnFilter()
    c.sheet.conditional = [
      {
        id: 2,
        ranges: [{ r0: 1, c0: 2, r1: 3, c1: 2 }],
        rule: { kind: "cellIs", operator: "greaterThan", values: ["6"], style: { fill: "#FFC7CE" } },
        stopIfTrue: false,
      },
    ]
    c.goTo({ r0: 2, c0: 2, r1: 2, c1: 2 })
    expect(c.data.filterBySelected("cell")).toBe(true)
    expect(shown(c)).toEqual([2, 3])
  })

  it("puts the selected cell's colour on top", () => {
    const c = table()
    c.goTo({ r0: 3, c0: 2, r1: 3, c1: 2 })
    c.format.fill("#FF0000")
    expect(c.data.sortSelectedColorOnTop("cell")).toBe(true)
    expect(column(c, 0)).toEqual(["Name", "Cy", "Ann", "Bob"])
  })
})

describe("Pick From Drop-down List", () => {
  it("lists the column's text entries above and below, A to Z, each once; numbers and formulas aren't offered", () => {
    const c = new SheetController()
    const put = (row: number, text: string) => {
      c.goTo({ r0: row, c0: 0, r1: row, c1: 0 })
      c.beginEdit("enter", text)
      c.commitEdit()
    }
    ;["pear", "Apple", "apple", "12", "=1+1", "fig"].forEach((t, r) => put(r, t))
    put(8, "after the gap")
    const list = c.pickList(6, 0)!
    expect(list.items).toEqual(["Apple", "fig", "pear"])
    list.pick("fig")
    expect(c.sheet.cells.get(6, 0)?.value).toBe("fig")
    expect(c.pickList(20, 0)).toBeNull()
  })
})
