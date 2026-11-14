import type { Node } from "../ast"
import type { EvalContext } from "../context"
import type { Value } from "../types"

/**
 * "scalar": one value; an array or multi-cell range given here lifts the call over its elements (=LEN(A1:A3)
 * spills). "any": passed as it is (references stay references, for SUM, INDEX, ROW…).
 */
export type ParamKind = "scalar" | "any"

export type FnDef = {
  min: number
  max: number
  /** One kind per parameter; the last repeats. Default "any". */
  params?: ParamKind[]
  /** Errors in scalar parameters are passed in rather than returned straight away (ISERROR, IFERROR…). */
  errors?: boolean
  volatile?: boolean
  call?: (args: Value[], ctx: EvalContext) => Value
  /** Branching functions get their arguments unevaluated. */
  lazy?: (args: (() => Value)[], ctx: EvalContext, nodes: Node[]) => Value
}

export const FUNCTIONS = new Map<string, FnDef>()

export function define(name: string, def: FnDef) {
  FUNCTIONS.set(name, def)
}

/** Shorthand for a function whose parameters are all scalars. */
export function scalarFn(
  name: string,
  min: number,
  max: number,
  call: NonNullable<FnDef["call"]>,
  extra: Partial<FnDef> = {},
) {
  define(name, { min, max, params: ["scalar"], call, ...extra })
}

export function isVolatile(name: string) {
  return FUNCTIONS.get(name)?.volatile === true
}
