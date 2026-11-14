import { tokenize } from "@cells/core"
import { FUTURE_FUNCTIONS } from "./constants"

/**
 * Formula text as a file stores it against Cells' own: the file has no "=", and functions newer than Excel 2007
 * carry a prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.FILTER`, LET's names as `_xlpm.x`).
 */

const PREFIX = /^(?:_xlfn\.|_xlws\.|_xlpm\.)+/i

/** A formula from a file (no "=") as Cells keeps it: "=" and no prefixes. */
export function formulaFromFile(text: string): string {
  if (!/_xl(?:fn|ws|pm)\./i.test(text)) return "=" + text
  let out = "="
  for (const t of tokenize(text)) out += t.kind === "func" || t.kind === "name" ? t.text.replace(PREFIX, "") : t.text
  return out
}

/** A Cells formula ("=…") as a file stores it. */
export function formulaToFile(input: string): string {
  const text = input.startsWith("=") ? input.slice(1) : input
  const tokens = tokenize(text)
  if (!tokens.some((t) => t.kind === "func")) return text
  const params = letNames(tokens)
  let out = ""
  for (const t of tokens) {
    if (t.kind === "func") {
      const prefix = FUTURE_FUNCTIONS[t.text.toUpperCase()]
      out += prefix ? prefix + t.text : t.text
    } else if (t.kind === "name" && params.has(t.text.toUpperCase())) out += "_xlpm." + t.text
    else out += t.text
  }
  return out
}

/** Names LET declares (its 1st, 3rd, … arguments). */
function letNames(tokens: ReturnType<typeof tokenize>): Set<string> {
  const names = new Set<string>()
  // A function name is followed by its own "(" token: the bracket opens that function's argument list.
  const stack: { let: boolean; arg: number }[] = []
  let pendingLet: boolean | null = null
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]
    if (t.kind === "func") pendingLet = /^LET$/i.test(t.text)
    else if (t.kind === "open") {
      stack.push({ let: pendingLet === true, arg: 0 })
      pendingLet = null
    } else if (t.kind === "close") stack.pop()
    else if (t.kind === "comma" && stack.length) stack[stack.length - 1].arg++
    else if (t.kind === "name") {
      const top = stack[stack.length - 1]
      if (top?.let && top.arg % 2 === 0) {
        let j = i + 1
        while (tokens[j]?.kind === "space") j++
        if (tokens[j]?.kind === "comma") names.add(t.text.toUpperCase())
      }
    }
  }
  return names
}
