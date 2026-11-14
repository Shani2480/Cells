import { describe, expect, it } from "vitest"
import { MAX_ROWS } from "../src/address"
import { Axis } from "../src/axis"
import { CellStore } from "../src/cell-store"
import { rng } from "./helpers"

describe("cell store", () => {
  it("stores values, inputs and styles sparsely", () => {
    const store = new CellStore()
    store.set(5, 2, { value: "=1", input: "=1", style: 3 })
    expect(store.get(5, 2)).toEqual({ value: "=1", input: "=1", style: 3 })
    store.set(5, 2, { style: 3 })
    expect(store.get(5, 2)).toEqual({ style: 3 })
    expect(store.hasValue(5, 2)).toBe(false)
    expect(store.usedBounds()).toEqual({ r0: 5, r1: 5, c0: 2, c1: 2 })
    store.set(5, 2, undefined)
    expect(store.get(5, 2)).toBeUndefined()
    expect(store.usedBounds()).toBeNull()
    expect(store.usedColumns()).toEqual([])
  })

  it("finds the next value across chunks both ways", () => {
    const store = new CellStore()
    for (const row of [3, 1023, 1024, 70_000, MAX_ROWS - 1]) store.set(row, 0, { value: row })
    const down: number[] = []
    for (let row = -1; (row = store.nextValueRow(0, row, 1)) !== -1;) down.push(row)
    expect(down).toEqual([3, 1023, 1024, 70_000, MAX_ROWS - 1])
    const up: number[] = []
    for (let row = MAX_ROWS; (row = store.nextValueRow(0, row, -1)) !== -1;) up.push(row)
    expect(up).toEqual([MAX_ROWS - 1, 70_000, 1024, 1023, 3])
  })

  it("shifts blocks of rows and columns", () => {
    const store = new CellStore()
    store.set(0, 0, { value: "a" })
    store.set(2, 0, { value: "b" })
    store.set(2, 1, { value: "c" })
    store.shiftRows(1, 2)
    expect(store.value(4, 0)).toBe("b")
    expect(store.value(0, 0)).toBe("a")
    store.shiftRows(4, -1, 0, 0) // column A only
    expect(store.value(3, 0)).toBe("b")
    expect(store.value(4, 1)).toBe("c")
    store.shiftCols(0, 1)
    expect(store.value(0, 1)).toBe("a")
    expect(store.value(4, 2)).toBe("c")
  })
})

describe("axis", () => {
  it("matches a plain array under random edits", () => {
    const random = rng(7)
    const n = 800
    const axis = new Axis(n, 20)
    const sizes = new Array(n).fill(20)
    const hidden = new Array(n).fill(false)
    for (let step = 0; step < 2000; step++) {
      const i = Math.floor(random() * n)
      const roll = random()
      if (roll < 0.4) {
        const size = Math.floor(random() * 60)
        axis.setSize(i, size)
        sizes[i] = size
      } else if (roll < 0.6) {
        axis.setHidden(i, !hidden[i])
        hidden[i] = !hidden[i]
      } else if (roll < 0.7) {
        axis.setSize(i, null)
        sizes[i] = 20
      } else if (roll < 0.8) {
        const k = 1 + Math.floor(random() * 5)
        axis.insert(i, k)
        sizes.splice(i, 0, ...new Array(k).fill(20))
        hidden.splice(i, 0, ...new Array(k).fill(false))
        sizes.length = n
        hidden.length = n
      } else if (roll < 0.9) {
        const k = Math.min(1 + Math.floor(random() * 5), n - i)
        axis.remove(i, k)
        sizes.splice(i, k)
        hidden.splice(i, k)
        sizes.push(...new Array(k).fill(20))
        hidden.push(...new Array(k).fill(false))
      }
      if (step % 200 === 0) {
        let offset = 0
        for (let j = 0; j < n; j++) {
          const drawn = hidden[j] ? 0 : sizes[j]
          expect(axis.offset(j)).toBe(offset)
          expect(axis.size(j)).toBe(drawn)
          if (drawn > 0) {
            expect(axis.indexAt(offset)).toBe(j)
            expect(axis.indexAt(offset + drawn - 0.5)).toBe(j)
          }
          offset += drawn
        }
      }
    }
  })

  it("is O(1) memory until a size changes", () => {
    const axis = new Axis(MAX_ROWS, 20)
    expect(axis.offset(MAX_ROWS)).toBe(MAX_ROWS * 20)
    expect(axis.indexAt(20 * 800_000 + 5)).toBe(800_000)
    expect(axis.overrides(0, MAX_ROWS - 1)).toEqual([])
  })
})
