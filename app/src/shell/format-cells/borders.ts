import type { Border, Op, Range, Sheet, StyleTable } from "@cells/core"

/**
 * Format Cells › Border: eight lines around and inside the selection. Each is unchanged (undefined), removed (null)
 * or set. The dialog opens with the lines the selection has where they are the same along the whole edge.
 */
export type Edge = "top" | "bottom" | "left" | "right" | "insideH" | "insideV" | "diagUp" | "diagDown"

export type BorderPlan = Partial<Record<Edge, Border | null>>

const same = (a: Border | undefined, b: Border | undefined) =>
  a?.style === b?.style && (a?.color ?? "") === (b?.color ?? "")

/** The selection's current lines (an edge counts when every cell along it has the same line). */
export function readBorders(sheet: Sheet, styles: StyleTable, r: Range): BorderPlan {
  // Big selections: look at the first 200 cells along each edge.
  const r1 = Math.min(r.r1, r.r0 + 199)
  const c1 = Math.min(r.c1, r.c0 + 199)
  const at = (row: number, col: number) => styles.get(sheet.styleAt(row, col))
  const along = (cells: [number, number][], pick: (row: number, col: number) => Border | undefined) => {
    const first = pick(...cells[0])
    return cells.every(([row, col]) => same(pick(row, col), first)) ? (first ?? null) : undefined
  }
  const rowCells = (row: number): [number, number][] => Array.from({ length: c1 - r.c0 + 1 }, (_, i) => [row, r.c0 + i])
  const colCells = (col: number): [number, number][] => Array.from({ length: r1 - r.r0 + 1 }, (_, i) => [r.r0 + i, col])
  const plan: BorderPlan = {
    top: along(rowCells(r.r0), (row, col) => at(row, col).borderTop),
    bottom: along(rowCells(r1), (row, col) => at(row, col).borderBottom),
    left: along(colCells(r.c0), (row, col) => at(row, col).borderLeft),
    right: along(colCells(c1), (row, col) => at(row, col).borderRight),
  }
  if (r1 > r.r0) {
    const inner: [number, number][] = []
    for (let row = r.r0; row < r1; row++) inner.push(...rowCells(row))
    plan.insideH = along(inner, (row, col) => at(row, col).borderBottom)
  }
  if (c1 > r.c0) {
    const inner: [number, number][] = []
    for (let col = r.c0; col < c1; col++) inner.push(...colCells(col))
    plan.insideV = along(inner, (row, col) => at(row, col).borderRight)
  }
  const a = at(r.r0, r.c0)
  plan.diagUp = a.diagonalUp ? (a.borderDiagonal ?? null) : null
  plan.diagDown = a.diagonalDown ? (a.borderDiagonal ?? null) : null
  return plan
}

/** The operations that apply the lines that changed. */
export function borderPlanOps(sheet: number, r: Range, plan: BorderPlan): Op[] {
  const ops: Op[] = []
  const put = (range: Range, patch: Record<string, unknown>) =>
    ops.push({ kind: "format", sheet, ranges: [range], patch })
  if (plan.top !== undefined) put({ ...r, r1: r.r0 }, { borderTop: plan.top })
  if (plan.bottom !== undefined) put({ ...r, r0: r.r1 }, { borderBottom: plan.bottom })
  if (plan.left !== undefined) put({ ...r, c1: r.c0 }, { borderLeft: plan.left })
  if (plan.right !== undefined) put({ ...r, c0: r.c1 }, { borderRight: plan.right })
  if (plan.insideH !== undefined && r.r1 > r.r0) {
    put({ ...r, r1: r.r1 - 1 }, { borderBottom: plan.insideH })
    // The cell below keeps no top line of its own that would fight the one set here.
    put({ ...r, r0: r.r0 + 1 }, { borderTop: null })
  }
  if (plan.insideV !== undefined && r.c1 > r.c0) {
    put({ ...r, c1: r.c1 - 1 }, { borderRight: plan.insideV })
    put({ ...r, c0: r.c0 + 1 }, { borderLeft: null })
  }
  if (plan.diagUp !== undefined || plan.diagDown !== undefined) {
    const line = plan.diagUp ?? plan.diagDown ?? null
    put(r, {
      borderDiagonal: line,
      diagonalUp: plan.diagUp ? true : plan.diagUp === null ? null : undefined,
      diagonalDown: plan.diagDown ? true : plan.diagDown === null ? null : undefined,
    })
  }
  return ops
}
