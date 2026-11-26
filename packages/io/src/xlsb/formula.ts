import { columnName as colName, quoteSheet } from "@cells/core"
import { Cursor } from "../xls/biff"
import { errorText, FormulaSkip } from "../xls/formula"
import { fixedArgs, functionName, USER_DEFINED } from "../xls/ftab"

/**
 * An .xlsb formula's parsed tokens (Rgce, [MS-XLSB] 2.2.2) back to formula text as an .xlsx keeps it (no "="). The
 * tokens are BIFF8's with wider fields: 4-byte rows, 2-byte columns carrying the relative bits, 4-byte name indexes,
 * UTF-16 strings, and tables (PtgList). Anything Cells can't spell throws FormulaSkip and the cell keeps its value.
 */

export type XlsbTable = { name: string; columns: string[] }

export type XlsbFormulaContext = {
  sheetNames: readonly string[]
  /** BrtExternSheet's entries: [first sheet, last sheet], or null for another workbook's (or a deleted sheet). */
  xti: readonly ([number, number] | null)[]
  /** The workbook's names in BrtName order (PtgName is 1-based into them). */
  names: readonly string[]
  /** Tables by their id. */
  tables: ReadonlyMap<number, XlsbTable>
  /** The cell the formula belongs to, for a shared formula's relative references. */
  row: number
  col: number
}

/** BIFF12's function numbers past BIFF8's (369 on): the Analysis ToolPak's and Excel 2007's, with fixed counts. */
const MORE_FUNCTIONS: Record<number, [string, number | null]> = {}
;`THAIDAYOFWEEK THAIDIGIT THAIMONTHOFYEAR THAINUMSOUND THAINUMSTRING THAISTRINGLENGTH ISTHAIDIGIT ROUNDBAHTDOWN
ROUNDBAHTUP THAIYEAR RTD CUBEVALUE CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER HEX2BIN HEX2DEC HEX2OCT DEC2BIN DEC2HEX
DEC2OCT OCT2BIN OCT2HEX OCT2DEC BIN2DEC BIN2OCT BIN2HEX IMSUB:2 IMDIV:2 IMPOWER:2 IMABS:1 IMSQRT:1 IMLN:1 IMLOG2:1
IMLOG10:1 IMSIN:1 IMCOS:1 IMEXP:1 IMARGUMENT:1 IMCONJUGATE:1 IMAGINARY:1 IMREAL:1 COMPLEX IMSUM IMPRODUCT SERIESSUM:4
FACTDOUBLE:1 SQRTPI:1 QUOTIENT:2 DELTA GESTEP ISEVEN:1 ISODD:1 MROUND:2 ERF ERFC:1 BESSELJ:2 BESSELK:2 BESSELY:2
BESSELI:2 XIRR XNPV:3 PRICEMAT YIELDMAT INTRATE RECEIVED DISC PRICEDISC YIELDDISC TBILLEQ:3 TBILLPRICE:3 TBILLYIELD:3
PRICE YIELD DOLLARDE:2 DOLLARFR:2 NOMINAL:2 EFFECT:2 CUMPRINC:6 CUMIPMT:6 EDATE:2 EOMONTH:2 YEARFRAC COUPDAYBS COUPDAYS
COUPDAYSNC COUPNCD COUPNUM COUPPCD DURATION MDURATION ODDLPRICE ODDLYIELD ODDFPRICE ODDFYIELD RANDBETWEEN:2 WEEKNUM
AMORDEGRC AMORLINC CONVERT:3 ACCRINT ACCRINTM WORKDAY NETWORKDAYS GCD MULTINOMIAL LCM FVSCHEDULE:2 CUBEKPIMEMBER CUBESET
CUBESETCOUNT:1 IFERROR:2 COUNTIFS SUMIFS AVERAGEIF AVERAGEIFS`
  .split(/\s+/)
  .forEach((entry, i) => {
    const [name, count] = entry.split(":")
    MORE_FUNCTIONS[369 + i] = [name, count ? Number(count) : null]
  })

const fnName = (index: number) => (index >= 369 ? (MORE_FUNCTIONS[index]?.[0] ?? null) : functionName(index))
const fnArgs = (index: number) => (index >= 369 ? (MORE_FUNCTIONS[index]?.[1] ?? null) : fixedArgs(index))

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

const ROWS = 1048576
const COLS = 16384

type Loc = { row: number; col: number; rowRel: boolean; colRel: boolean }

const cellText = (l: Loc) => `${l.colRel ? "" : "$"}${colName(l.col)}${l.rowRel ? "" : "$"}${l.row + 1}`

/** A row and column word: in a shared formula (`relative`) the relative parts are offsets from the cell. */
function fix(rowRaw: number, colRaw: number, ctx: XlsbFormulaContext, relative: boolean): Loc {
  const rowRel = (colRaw & 0x8000) !== 0
  const colRel = (colRaw & 0x4000) !== 0
  let row = rowRaw
  let col = colRaw & 0x3fff
  if (relative) {
    if (rowRel) row = (((ctx.row + (rowRaw | 0)) % ROWS) + ROWS) % ROWS
    if (colRel) col = (((ctx.col + (col >= 0x2000 ? col - 0x4000 : col)) % COLS) + COLS) % COLS
  }
  return { row, col, rowRel, colRel }
}

function readLoc(c: Cursor, ctx: XlsbFormulaContext, relative: boolean): string {
  const row = c.u32()
  return cellText(fix(row, c.u16(), ctx, relative))
}

function readArea(c: Cursor, ctx: XlsbFormulaContext, relative: boolean): string {
  const r0 = c.u32()
  const r1 = c.u32()
  const c0 = c.u16()
  const c1 = c.u16()
  const a = fix(r0, c0, ctx, relative)
  const b = fix(r1, c1, ctx, relative)
  // Whole columns and rows, as Excel writes them (A:A, 1:1).
  if (a.row === 0 && b.row === ROWS - 1 && a.rowRel === b.rowRel && !a.rowRel)
    return `${a.colRel ? "" : "$"}${colName(a.col)}:${b.colRel ? "" : "$"}${colName(b.col)}`
  if (a.col === 0 && b.col === COLS - 1 && a.colRel === b.colRel && !a.colRel)
    return `${a.rowRel ? "" : "$"}${a.row + 1}:${b.rowRel ? "" : "$"}${b.row + 1}`
  return `${cellText(a)}:${cellText(b)}`
}

function sheetPrefix(ctx: XlsbFormulaContext, ixti: number): string {
  if (ixti >= ctx.xti.length) throw new FormulaSkip(`Unknown sheet reference ${ixti}.`)
  const x = ctx.xti[ixti]
  if (!x) throw new FormulaSkip("A reference to another workbook.")
  const [first, last] = x
  const a = ctx.sheetNames[first]
  const b = ctx.sheetNames[last]
  if (a === undefined || b === undefined) return "#REF!"
  if (first === last) return quoteSheet(a) + "!"
  const plain = quoteSheet(a) === a && quoteSheet(b) === b
  return plain ? `${a}:${b}!` : `'${a.replace(/'/g, "''")}:${b.replace(/'/g, "''")}'!`
}

const quoteText = (s: string) => `"${s.replace(/"/g, '""')}"`

/** A string of `count` UTF-16 characters. */
function wide(c: Cursor, count: number): string {
  let s = ""
  for (let i = 0; i < count && c.left >= 2; i++) s += String.fromCharCode(c.u16())
  return s
}

/** An inline array's values (PtgArray's data after the tokens: rows, columns, then each value). */
function readArrayData(c: Cursor): string {
  const rows = c.u32()
  const cols = c.u32()
  const out: string[] = []
  for (let r = 0; r < rows; r++) {
    const row: string[] = []
    for (let k = 0; k < cols; k++) {
      const type = c.u8()
      switch (type) {
        case 0x00:
          row.push(String(c.f64()))
          break
        case 0x01:
          row.push(quoteText(wide(c, c.u16())))
          break
        case 0x02:
          row.push(c.u8() ? "TRUE" : "FALSE")
          break
        case 0x04:
          row.push(errorText(c.u8()))
          break
        default:
          throw new FormulaSkip(`Array value 0x${type.toString(16)}.`)
      }
    }
    out.push(row.join(","))
  }
  return `{${out.join(";")}}`
}

/** A column name in a structured reference: [ ] # ' are escaped with a '. */
const columnText = (name: string) => name.replace(/['[\]#]/g, "'$&")

const ROW_TYPES: Record<number, string[]> = {
  0x01: ["#All"],
  0x02: ["#Headers"],
  0x04: ["#Data"],
  0x08: ["#Totals"],
  0x06: ["#Headers", "#Data"],
  0x0c: ["#Data", "#Totals"],
  0x10: ["#This Row"],
}

/** PtgList ([MS-XLSB] 2.5.97.52): a structured reference to a table. */
function readList(c: Cursor, ctx: XlsbFormulaContext): string {
  c.skip(2) // ixti: this workbook
  const flags = c.u16()
  const id = c.u32()
  const first = c.u16()
  const last = c.u16()
  const table = ctx.tables.get(id)
  if (!table) throw new FormulaSkip(`Table ${id}.`)
  if (flags & 0x1000) return "#REF!"
  const columns = flags & 0x03
  const rows = ROW_TYPES[(flags >> 2) & 0x1f] ?? []
  const parts = rows.map((r) => `[${r}]`)
  const name = (i: number) => {
    const n = table.columns[i]
    if (n === undefined) throw new FormulaSkip("A table column that isn't there.")
    return `[${columnText(n)}]`
  }
  if (columns === 1) parts.push(name(first))
  else if (columns === 2) parts.push(`${name(first)}:${name(last)}`)
  if (!parts.length) return `${table.name}[]`
  if (columns === 1 && !rows.length) return `${table.name}${parts[0]}`
  return `${table.name}[${parts.join(",")}]`
}

/** The formula of `cce` bytes of tokens at the cursor (inline arrays' values follow them). */
export function decodeXlsbFormula(c: Cursor, cce: number, ctx: XlsbFormulaContext): string {
  const end = c.pos + cce
  const stack: string[] = []
  let arrays = 0
  const pop = () => {
    const v = stack.pop()
    if (v === undefined) throw new FormulaSkip("The formula's tokens don't add up.")
    return v
  }
  const call = (name: string, argc: number) => {
    const args: string[] = []
    for (let i = 0; i < argc; i++) args.unshift(pop())
    if (name === "") {
      const fn = args.shift() ?? ""
      stack.push(`${fn}(${args.join(",")})`)
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
      case 0x01:
      case 0x02:
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
        stack.push(quoteText(wide(c, c.u16())))
        continue
      case 0x18: {
        const eptg = c.u8()
        if (eptg === 0x19) stack.push(readList(c, ctx))
        else throw new FormulaSkip(`Extended token 0x${eptg.toString(16)}.`)
        continue
      }
      case 0x19: {
        const flags = c.u8()
        const data = c.u16()
        if (flags & 0x04) c.skip((data + 1) * 2) // tAttrChoose's jump table
        if (flags & 0x10) stack.push(`SUM(${pop()})`) // tAttrSum
        if (flags & 0x40) c.skip(0) // tAttrSpace: spacing only
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
        stack.push(String(c.f64()))
        continue
    }
    switch (ptg >= 0x20 && ptg < 0x80 ? (ptg & 0x1f) | 0x20 : -1) {
      case 0x20: // PtgArray: its values follow the tokens
        c.skip(14)
        stack.push(`\u0000${arrays++}\u0000`)
        break
      case 0x21: {
        const index = c.u16()
        const name = fnName(index)
        const argc = fnArgs(index)
        if (name === null || argc === null) throw new FormulaSkip(`Function ${index}.`)
        call(name, argc)
        break
      }
      case 0x22: {
        const argc = c.u8() & 0x7f
        const index = c.u16()
        if (index & 0x8000) throw new FormulaSkip("A macro command.")
        if (index === USER_DEFINED) call("", argc)
        else {
          const name = fnName(index)
          if (name === null) throw new FormulaSkip(`Function ${index}.`)
          call(name, argc)
        }
        break
      }
      case 0x23: {
        const index = c.u32()
        const name = ctx.names[index - 1]
        if (name === undefined) throw new FormulaSkip(`Name ${index}.`)
        stack.push(name)
        break
      }
      case 0x24:
        stack.push(readLoc(c, ctx, false))
        break
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
        c.skip(6)
        stack.push("#REF!")
        break
      case 0x2b:
        c.skip(12)
        stack.push("#REF!")
        break
      case 0x2c:
        stack.push(readLoc(c, ctx, true))
        break
      case 0x2d:
        stack.push(readArea(c, ctx, true))
        break
      case 0x39: {
        // PtgNameX: a name in another workbook (or an add-in's function)
        c.skip(2)
        const index = c.u32()
        const name = ctx.names[index - 1]
        if (name === undefined) throw new FormulaSkip(`External name ${index}.`)
        stack.push(name)
        break
      }
      case 0x3a: {
        const prefix = sheetPrefix(ctx, c.u16())
        const ref = readLoc(c, ctx, false)
        stack.push(prefix === "#REF!" ? prefix : prefix + ref)
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
        c.skip(6)
        stack.push(prefix === "#REF!" ? prefix : prefix + "#REF!")
        break
      }
      case 0x3d: {
        const prefix = sheetPrefix(ctx, c.u16())
        c.skip(12)
        stack.push(prefix === "#REF!" ? prefix : prefix + "#REF!")
        break
      }
      default:
        throw new FormulaSkip(`Token 0x${ptg.toString(16)}.`)
    }
  }
  c.pos = end
  let text = stack.length === 1 ? stack[0] : null
  if (text === null) throw new FormulaSkip("The formula's tokens don't add up.")
  // After the tokens: rgcb's byte count, then the inline arrays' values in the order the arrays appear.
  if (arrays) {
    c.skip(4)
    for (let n = 0; n < arrays; n++) text = text.replace(`\u0000${n}\u0000`, readArrayData(c))
  }
  return text
}
