import { describe, expect, it } from "vitest"
import { fitGroups, type GroupWidths } from "./fit"

const groups: GroupWidths[] = [
  [100, 80, 50],
  [200, 150, 50],
  [120, 120, 50],
]

describe("ribbon fitting", () => {
  it("keeps every group full when there is room", () => {
    expect(fitGroups(groups, 420)).toEqual([0, 0, 0])
  })

  it("makes groups compact from the right first", () => {
    // 420 → 400: the last group has no compact form, the middle one goes compact.
    expect(fitGroups(groups, 400)).toEqual([0, 1, 0])
    expect(fitGroups(groups, 350)).toEqual([1, 1, 0])
  })

  it("then collapses groups from the right", () => {
    expect(fitGroups(groups, 300)).toEqual([1, 1, 2])
    expect(fitGroups(groups, 200)).toEqual([1, 2, 2])
    expect(fitGroups(groups, 10)).toEqual([2, 2, 2])
  })
})
