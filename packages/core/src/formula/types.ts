import type { CellError } from "../values"

/** A value inside a calculation. `null` is a blank cell, which is neither 0 nor "" until something coerces it. */
export type Scalar = number | string | boolean | CellError | null

/** A rectangle of one sheet. */
export type Area = { sheet: number; r0: number; c0: number; r1: number; c1: number }

/** A reference: one or more areas (more than one only from the union operator). */
export type RefValue = { kind: "ref"; areas: Area[]; threeD?: boolean }

/** An array: row-major values. Arrays from ranges can hold blanks. */
export type Matrix = { kind: "matrix"; rows: number; cols: number; data: Scalar[] }

export type Value = Scalar | RefValue | Matrix

export function isRef(value: Value): value is RefValue {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "ref"
}

export function isMatrix(value: Value): value is Matrix {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "matrix"
}

export function matrix(rows: number, cols: number, data: Scalar[]): Matrix {
  return { kind: "matrix", rows, cols, data }
}

export function matrixFrom(rows: Scalar[][]): Matrix {
  return matrix(rows.length, rows[0]?.length ?? 0, rows.flat())
}

export function areaRows(area: Area) {
  return area.r1 - area.r0 + 1
}

export function areaCols(area: Area) {
  return area.c1 - area.c0 + 1
}
