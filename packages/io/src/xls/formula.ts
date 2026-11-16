import { columnName as colName, quoteSheet } from "@cells/core"
import { Cursor } from "./biff"
import { fixedArgs, functionName, USER_DEFINED } from "./ftab"

/**
 * A BIFF8 formula's parsed tokens (Rgce, [MS-XLS] 2.5.198) back to the text Cells keeps (`=SUM(A1:B2)`). The tokens
 * are in postfix order with explicit parentheses (PtgParen), so writing each operator between its operands gives back
 * the formula as typed. Anything Cells can't spell (a macro function, a token of a kind it doesn't know) throws, and
 * the cell keeps its value without the formula.
 */

export class FormulaSkip extends Error {}

/** Where a 3-D reference (EXTERNSHEET's index) points: sheets of this workbook, or of another one. */
export type Xti = { sheets: [number, number] | null; book: string | null; supbook: number }

export type FormulaContext = {
  sheetNames: readonly string[]
  /** EXTERNSHEET's entries. */
  xti: readonly Xti[]
  /** The workbook's NAME records' names, in order (PtgName is 1-based into them). */
  names: readonly string[]
  /** The names of an EXTERNSHEET entry's book (PtgNameX is 1-based into them): add-in functions, or this book's. */
  externNames: (xti: number) => readonly string[]
  /** The cell the formula belongs to, for the relative references of a shared formula (PtgRefN, PtgAreaN). */
  row: number
  col: number
}

const ERRORS: Record<number, string> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
}

export const errorText = (code: number) => ERRORS[code] ?? "#N/A"

const BINARY: Record<number, string> = {
  0x03: "+",
  0x04: "-",
  0x05: "*",
  0x06: "/",
  0x07: "^",
  0x08: "&",
  0x09: "<",
  0x0a: "<=",
  0x0b: "=",
  0x0c: ">=",
  0x0d: ">",
  0x0e: "<>",
  0x0f: " ",
  0x10: ",",
  0x11: ":",
}

const ROWS = 65536
const COLS = 256

function cellText(row: number, col: number, rowRel: boolean, colRel: boolean) {
  return `${colRel ? "" : "$"}${colName(col)}${rowRel ? "" : "$"}${row + 1}`
}

/** A reference's row and column field (RgceLoc / RgceLocRel): relative bits in the column word. */
function readLoc(c: Cursor, ctx: FormulaContext, relative: boolean) {
  const rowRaw = c.u16()
  const colRaw = c.u16()
  const rowRel = (colRaw & 0x8000) !== 0
  const colRel = (colRaw & 0x4000) !== 0
  let row = rowRaw
  let col = colRaw & 0x3fff
  if (relative) {
    // A shared formula: relative parts are offsets from the cell (16-bit rows, 8-bit columns, both wrapping).
    if (rowRel) row = (((ctx.row + (rowRaw >= 0x8000 ? rowRaw - 0x10000 : rowRaw)) % ROWS) + ROWS) % ROWS
    if (colRel) {
      const d = col & 0xff
      col = (((ctx.col + (d >= 0x80 ? d - 0x100 : d)) % COLS) + COLS) % COLS
    }
  }
  return { row, col, rowRel, colRel }
}

function readArea(c: Cursor, ctx: FormulaContext, relative: boolean) {
  const r0 = c.u16()
  const r1 = c.u16()
  const c0 = c.u16()
  const c1 = c.u16()
  const fix = (rowRaw: number, colRaw: number) => {
    const rowRel = (colRaw & 0x8000) !== 0
    const colRel = (colRaw & 0x4000) !== 0
    let row = rowRaw
    let col = colRaw & 0x3fff
    if (relative) {
      if (rowRel) row = (((ctx.row + (rowRaw >= 0x8000 ? rowRaw - 0x10000 : rowRaw)) % ROWS) + ROWS) % ROWS
      if (colRel) {
        const d = col & 0xff
        col = (((ctx.col + (d >= 0x80 ? d - 0x100 : d)) % COLS) + COLS) % COLS
      }
    }
    return { row, col, rowRel, colRel }
  }
  const a = fix(r0, c0)
  const b = fix(r1, c1)
  // Whole columns and rows, as Excel writes them (A:A, 1:1).
  if (a.row === 0 && b.row === ROWS - 1 && a.rowRel === b.rowRel)
    return `${a.colRel ? "" : "$"}${colName(a.col)}:${b.colRel ? "" : "$"}${colName(b.col)}`
  if (a.col === 0 && b.col === COLS - 1 && a.colRel === b.colRel)
    return `${a.rowRel ? "" : "$"}${a.row + 1}:${b.rowRel ? "" : "$"}${b.row + 1}`
  return `${cellText(a.row, a.col, a.rowRel, a.colRel)}:${cellText(b.row, b.col, b.rowRel, b.colRel)}`
}

/** The sheet prefix of a 3-D reference (`Sheet2!`, `'Q1 Sales'!`, `Sheet1:Sheet3!`, `[Book2]Sheet1!`). */
function sheetPrefix(ctx: FormulaContext, ixti: number): string {
  const x = ctx.xti[ixti]
  if (!x) throw new FormulaSkip(`Unknown sheet reference ${ixti}.`)
  if (x.book !== null) throw new FormulaSkip("A reference to another workbook.")
  if (!x.sheets) return "#REF!"
  const [first, last] = x.sheets
  const a = ctx.sheetNames[first]
  const b = ctx.sheetNames[last]
  if (a === undefined || b === undefined) return "#REF!"
  if (first === last) return quoteSheet(a) + "!"
  // Sheet1:Sheet3 is quoted as one when either needs it.
  const plain = quoteSheet(a) === a && quoteSheet(b) === b
  return plain ? `${a}:${b}!` : `'${a.replace(/'/g, "''")}:${b.replace(/'/g, "''")}'!`
}

const quoteText = (s: string) => `"${s.replace(/"/g, '""')}"`

/** An inline array's values (PtgArray's data after the tokens, [MS-XLS] 2.5.198.8). */
function readArrayData(c: Cursor): string {
  const cols = c.u8() + 1
  const rows = c.u16() + 1
  const out: string[] = []
  for (let r = 0; r < rows; r++) {
    const row: string[] = []
    for (let k = 0; k < cols; k++) {
      const type = c.u8()
      switch (type) {
        case 0x01:
          row.push(String(c.f64()))
          break
        case 0x02:
          row.push(quoteText(c.string()))
          break
        case 0x04:
          row.push(c.u8() ? "TRUE" : "FALSE")
          c.skip(7)
          break
        case 0x10:
          row.push(errorText(c.u8()))
          c.skip(7)
          break
        default:
          c.skip(8)
          row.push("")
      }
    }
    out.push(row.join(","))
  }
  return `{${out.join(";")}}`
}

/** A number as a formula writes it. */
const numberText = (n: number) => String(n)

/** The formula text of `cce` bytes of tokens at the cursor, without the leading "=" (inline arrays' values follow). */
export function decodeFormula(c: Cursor, cce: number, ctx: FormulaContext): string {
  const end = c.pos + cce
  const stack: string[] = []
  const arrays: number[] = []
  const pop = () => {
    const v = stack.pop()
    if (v === undefined) throw new FormulaSkip("The formula's tokens don't add up.")
    return v
  }
  const call = (name: string, argc: number) => {
    const args: string[] = []
    for (let i = 0; i < argc; i++) args.unshift(pop())
    if (name === "") {
      // Function 255: the first argument names the function.
      const fn = args.shift() ?? ""
      stack.push(`${fn.replace(/^_xlfn\./i, "").replace(/^_xlws\./i, "")}(${args.join(",")})`)
    } else stack.push(`${name}(${args.join(",")})`)
  }
  while (c.pos < end) {
    const ptg = c.u8()
    if (BINARY[ptg] !== undefined) {
      const b = pop()
      const a = pop()
      stack.push(a + BINARY[ptg] + b)
      continue
    }
    switch (ptg) {
      case 0x01: // PtgExp: a shared or array formula, resolved by the caller
      case 0x02: // PtgTbl: a data table
        throw new FormulaSkip("A shared formula, array formula or data table.")
      case 0x12:
        stack.push("+" + pop())
        continue
      case 0x13:
        stack.push("-" + pop())
        continue
      case 0x14:
        stack.push(pop() + "%")
        continue
      case 0x15:
        stack.push(`(${pop()})`)
        continue
      case 0x16:
        stack.push("")
        continue
      case 0x17:
        stack.push(quoteText(c.string(1)))
        continue
      case 0x18:
        throw new FormulaSkip("An extended token.")
      case 0x19: {
        const flags = c.u8()
        const data = c.u16()
        if (flags & 0x04) c.skip((data + 1) * 2) // tAttrChoose's jump table
        if (flags & 0x10) stack.push(`SUM(${pop()})`) // tAttrSum
        continue
      }
      case 0x1c:
        stack.push(errorText(c.u8()))
        continue
      case 0x1d:
        stack.push(c.u8() ? "TRUE" : "FALSE")
        continue
      case 0x1e:
        stack.push(String(c.u16()))
        continue
      case 0x1f:
        stack.push(numberText(c.f64()))
        continue
    }
    // Tokens in three classes (reference, value, array): the same token with bits 5–6 set.
    switch (ptg >= 0x20 && ptg < 0x80 ? (ptg & 0x1f) | 0x20 : -1) {
      case 0x20: // PtgArray
        c.skip(7)
        stack.push(`\u0000${arrays.length}\u0000`)
        arrays.push(arrays.length)
        break
      case 0x21: {
        // PtgFunc
        const index = c.u16()
        const name = functionName(index)
        const argc = fixedArgs(index)
        if (name === null || argc === null) throw new FormulaSkip(`Function ${index}.`)
        call(name, argc)
        break
      }
      case 0x22: {
        // PtgFuncVar
        const argc = c.u8() & 0x7f
        const index = c.u16()
        if (index & 0x8000) throw new FormulaSkip("A macro command.")
        if (index === USER_DEFINED) call("", argc)
        else {
          const name = functionName(index)
          if (name === null) throw new FormulaSkip(`Function ${index}.`)
          call(name, argc)
        }
        break
      }
      case 0x23: {
        // PtgName
        const index = c.u16()
        c.skip(2)
        const name = ctx.names[index - 1]
        if (name === undefined) throw new FormulaSkip(`Name ${index}.`)
        stack.push(name)
        break
      }
      case 0x24: {
        const l = readLoc(c, ctx, false)
        stack.push(cellText(l.row, l.col, l.rowRel, l.colRel))
        break
      }
      case 0x25:
        stack.push(readArea(c, ctx, false))
        break
      case 0x26: // PtgMemArea, PtgMemErr, PtgMemNoMem: the expression that follows stands
      case 0x27:
      case 0x28:
        c.skip(6)
        break
      case 0x29: // PtgMemFunc
        c.skip(2)
        break
      case 0x2a:
        c.skip(4)
        stack.push("#REF!")
        break
      case 0x2b:
        c.skip(8)
        stack.push("#REF!")
        break
      case 0x2c: {
        const l = readLoc(c, ctx, true)
        stack.push(cellText(l.row, l.col, l.rowRel, l.colRel))
        break
      }
      case 0x2d:
        stack.push(readArea(c, ctx, true))
        break
      case 0x39: {
        // PtgNameX: a name in another book's list, an add-in function's name most often
        const ixti = c.u16()
        const index = c.u16()
        c.skip(2)
        const name = ctx.externNames(ixti)[index - 1]
        if (name === undefined) throw new FormulaSkip(`External name ${index}.`)
        stack.push(name)
        break
      }
      case 0x3a: {
        const prefix = sheetPrefix(ctx, c.u16())
        const l = readLoc(c, ctx, false)
        stack.push(prefix === "#REF!" ? prefix : prefix + cellText(l.row, l.col, l.rowRel, l.colRel))
        break
      }
      case 0x3b: {
        const prefix = sheetPrefix(ctx, c.u16())
        const area = readArea(c, ctx, false)
        stack.push(prefix === "#REF!" ? prefix : prefix + area)
        break
      }
      case 0x3c: {
        const prefix = sheetPrefix(ctx, c.u16())
        c.skip(4)
        stack.push(prefix === "#REF!" ? prefix : prefix + "#REF!")
        break
      }
      case 0x3d: {
        const prefix = sheetPrefix(ctx, c.u16())
        c.skip(8)
        stack.push(prefix === "#REF!" ? prefix : prefix + "#REF!")
        break
      }
      default:
        throw new FormulaSkip(`Token 0x${ptg.toString(16)}.`)
    }
  }
  c.pos = end
  // Inline arrays' values come after the tokens, in the order the arrays appear.
  let text = stack.length === 1 ? stack[0] : null
  if (text === null) throw new FormulaSkip("The formula's tokens don't add up.")
  for (const n of arrays) text = text.replace(`\u0000${n}\u0000`, readArrayData(c))
  return text
}
