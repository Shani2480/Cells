import { FUNCTION_INFO, tokenize } from "@cells/core"
import { FUTURE_FUNCTIONS } from "./constants"

/**
 * Formula text as a file stores it against Cells' own: the file has no "=", and functions newer than Excel 2007
 * carry a prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.FILTER`, LET's names and LAMBDA's parameters as `_xlpm.x`, calls
 * through them as `_xlpm.f(2)`).
 */

const PREFIX = /^(?:_xlfn\.|_xlws\.|_xlpm\.|_xleta\.)+/i

/** A formula from a file (no "=") as Cells keeps it: "=" and no prefixes. */
export function formulaFromFile(text: string): string {
  if (!/_xl(?:fn|ws|pm|eta)\./i.test(text)) return "=" + text
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
      const upper = t.text.toUpperCase()
      const prefix = params.has(upper) ? "_xlpm." : FUTURE_FUNCTIONS[upper]
      out += prefix ? prefix + t.text : t.text
    } else if (t.kind === "name" && params.has(t.text.toUpperCase())) out += "_xlpm." + t.text
    else if (t.kind === "name" && FUNCTION_INFO.has(t.text.toUpperCase())) {
      // A built-in function passed by name (=BYROW(A1:C3,SUM)): Excel marks it _xleta.
      out += "_xleta." + t.text
    } else out += t.text
  }
  return out
}

/** Names LET declares (its 1st, 3rd, … arguments) and LAMBDA's parameters (every argument but the last). */
function letNames(tokens: ReturnType<typeof tokenize>): Set<string> {
  const names = new Set<string>()
  // A function name is followed by its own "(" token: the bracket opens that function's argument list.
  const stack: { kind: "let" | "lambda" | null; arg: number }[] = []
  let pending: "let" | "lambda" | null = null
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]
    if (t.kind === "func") pending = /^LET$/i.test(t.text) ? "let" : /^LAMBDA$/i.test(t.text) ? "lambda" : null
    else if (t.kind === "open") {
      stack.push({ kind: pending, arg: 0 })
      pending = null
    } else if (t.kind === "close") stack.pop()
    else if (t.kind === "comma" && stack.length) stack[stack.length - 1].arg++
    else if (t.kind === "name") {
      const top = stack[stack.length - 1]
      // A parameter is a name standing alone as its argument, so a comma follows it.
      if (top?.kind === "lambda" || (top?.kind === "let" && top.arg % 2 === 0)) {
        let j = i + 1
        while (tokens[j]?.kind === "space") j++
        if (tokens[j]?.kind === "comma") names.add(t.text.toUpperCase())
      }
    }
  }
  return names
}
