/**
 * How the ribbon fits a narrow window, as Excel's does: groups give up room from the right. First each group goes
 * compact (small buttons lose their labels, large buttons become small ones), right to left; if that is not enough,
 * groups collapse into a single button that opens the group, again right to left.
 */

/** 0 full, 1 compact, 2 collapsed to one button. */
export type Level = 0 | 1 | 2

/** A group's width at each level, in pixels. */
export type GroupWidths = [full: number, compact: number, collapsed: number]

export function fitGroups(widths: readonly GroupWidths[], available: number): Level[] {
  const levels: Level[] = widths.map(() => 0)
  let total = widths.reduce((sum, w) => sum + w[0], 0)
  for (const target of [1, 2] as const) {
    for (let i = widths.length - 1; i >= 0 && total > available; i--) {
      const next = widths[i][target]
      if (next >= widths[i][levels[i]]) continue
      total -= widths[i][levels[i]] - next
      levels[i] = target
    }
  }
  return levels
}
