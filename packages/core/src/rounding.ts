/** Excel rounds the decimal number it shows, not the binary one: ROUND(2.675, 2) is 2.68. */
export function roundDecimal(x: number, digits: number, mode: "half" | "up" | "down" = "half"): number {
  if (!Number.isFinite(x) || x === 0) return x
  digits = Math.trunc(digits)
  if (digits > 15 + 308) return x
  const sign = x < 0 ? -1 : 1
  const abs = Math.abs(x)
  // Work on the 15-significant-digit decimal string, shifting the point by `digits`.
  const [mant, expText] = abs.toExponential(14).split("e")
  const exp = Number(expText) + digits
  const ds = mant.replace(".", "") // 15 digits, the first before the point
  // Value = 0.ds × 10^(exp+1); keep exp+1 digits before the cut.
  const keep = exp + 1
  if (keep <= 0) {
    if (mode === "down") return 0
    if (mode === "up") return (sign * 1) / 10 ** digits
    // Only rounds up when the first dropped digit is 5 or more and it sits right after the cut.
    return keep === 0 && Number(ds[0]) >= 5 ? (sign * 1) / 10 ** digits : 0
  }
  if (keep >= ds.length) return x
  let kept = BigInt(ds.slice(0, keep))
  const dropped = ds.slice(keep)
  const up = mode === "up" ? /[1-9]/.test(dropped) : mode === "half" ? Number(dropped[0]) >= 5 : false
  if (up) kept += 1n
  const result = Number(`${kept}e${-digits}`)
  return sign * result
}
