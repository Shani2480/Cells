/**
 * Sizes along one axis (row heights or column widths) in pixels at 100% zoom, with hidden entries. Offsets and
 * "index at pixel" are O(log n) through a Fenwick tree of differences from the default size. Nothing is allocated
 * until the first size differs from the default, so an untouched sheet costs nothing.
 */
export class Axis {
  readonly count: number
  defaultSize: number
  /** Custom sizes (NaN = default); hidden entries keep their size here for Unhide. */
  private custom: Float64Array | null = null
  private hidden: Uint8Array | null = null
  /** Sizes the app set itself to fit the content (rows Excel heightens for wrapped text or big fonts). */
  private auto: Uint8Array | null = null
  /** Fenwick tree over (drawn size − default), 1-based. */
  private tree: Float64Array | null = null
  private customCount = 0
  private hiddenCount = 0

  constructor(count: number, defaultSize: number) {
    this.count = count
    this.defaultSize = defaultSize
  }

  /** The size the entry has when shown (custom or default). */
  storedSize(index: number): number {
    const size = this.custom?.[index]
    return size === undefined || Number.isNaN(size) ? this.defaultSize : size
  }

  hasCustomSize(index: number) {
    const size = this.custom?.[index]
    return size !== undefined && !Number.isNaN(size)
  }

  /** The drawn size: 0 when hidden. */
  size(index: number): number {
    return this.isHidden(index) ? 0 : this.storedSize(index)
  }

  isHidden(index: number) {
    return this.hidden !== null && this.hidden[index] === 1
  }

  get anyHidden() {
    return this.hiddenCount > 0
  }

  /** Whether the size was fitted to the content rather than set by the user (Excel's customHeight="0"). */
  isAuto(index: number) {
    return this.auto !== null && this.auto[index] === 1
  }

  /** Sets a custom size; null returns the entry to the default. `auto`: fitted to the content, not chosen. */
  setSize(index: number, size: number | null, auto = false) {
    const before = this.size(index)
    this.ensure()
    this.auto![index] = size !== null && auto ? 1 : 0
    const wasCustom = this.hasCustomSize(index)
    this.custom![index] = size === null ? Number.NaN : Math.max(0, size)
    this.customCount += (size === null ? 0 : 1) - (wasCustom ? 1 : 0)
    this.bump(index, this.size(index) - before)
  }

  setHidden(index: number, hidden: boolean) {
    if (this.isHidden(index) === hidden) {
      return
    }
    const before = this.size(index)
    this.ensure()
    this.hidden![index] = hidden ? 1 : 0
    this.hiddenCount += hidden ? 1 : -1
    this.bump(index, this.size(index) - before)
  }

  /** Pixel position of the start of `index` (sum of sizes before it). */
  offset(index: number): number {
    const base = index * this.defaultSize
    if (!this.tree) {
      return base
    }
    let sum = 0
    for (let i = index; i > 0; i -= i & -i) {
      sum += this.tree[i]
    }
    return base + sum
  }

  /** Total size of the whole axis. */
  get total(): number {
    return this.offset(this.count)
  }

  /**
   * The entry whose span contains pixel `px` (the last entry at or before it when hidden entries sit there). Clamped
   * to the axis.
   */
  indexAt(px: number): number {
    if (px <= 0) {
      return this.firstVisibleFrom(0, 1)
    }
    if (!this.tree) {
      return Math.min(this.count - 1, Math.floor(px / this.defaultSize))
    }
    // Largest prefix length `pos` with offset(pos) <= px, by binary lifting.
    let pos = 0
    let acc = 0
    for (let step = highestPowerOfTwo(this.count); step > 0; step >>= 1) {
      const next = pos + step
      if (next <= this.count) {
        const candidate = acc + this.tree[next] + step * this.defaultSize
        if (candidate <= px) {
          pos = next
          acc = candidate
        }
      }
    }
    // offset(pos) <= px < offset(pos + 1) unless pos + 1 is hidden (size 0): skip forward to a shown entry.
    const index = Math.min(pos, this.count - 1)
    return this.isHidden(index) ? this.firstVisibleFrom(index, -1) : index
  }

  /** The first shown entry from `index` going in `dir`, or `index` itself if none is shown that way. */
  firstVisibleFrom(index: number, dir: 1 | -1): number {
    if (!this.hiddenCount) {
      return index
    }
    for (let i = index; i >= 0 && i < this.count; i += dir) {
      if (!this.isHidden(i)) {
        return i
      }
    }
    return index
  }

  /** Inserts `n` default-sized, shown entries at `at`; the last `n` entries fall off the end. */
  insert(at: number, n: number) {
    if (!this.custom) {
      return
    }
    this.custom.copyWithin(at + n, at, this.count - n)
    this.custom.fill(Number.NaN, at, at + n)
    this.hidden!.copyWithin(at + n, at, this.count - n)
    this.hidden!.fill(0, at, at + n)
    this.auto!.copyWithin(at + n, at, this.count - n)
    this.auto!.fill(0, at, at + n)
    this.rebuild()
  }

  /** Removes `n` entries at `at`; default entries come in at the end. */
  remove(at: number, n: number) {
    if (!this.custom) {
      return
    }
    this.custom.copyWithin(at, at + n)
    this.custom.fill(Number.NaN, this.count - n)
    this.hidden!.copyWithin(at, at + n)
    this.hidden!.fill(0, this.count - n)
    this.auto!.copyWithin(at, at + n)
    this.auto!.fill(0, this.count - n)
    this.rebuild()
  }

  /** Entries with a custom size or hidden, in [from, to], for undo records. */
  overrides(from: number, to: number): { index: number; size: number | null; hidden: boolean; auto?: true }[] {
    const out: { index: number; size: number | null; hidden: boolean; auto?: true }[] = []
    const custom = this.custom
    const hiddenFlags = this.hidden
    if (!custom || !hiddenFlags || this.customCount + this.hiddenCount === 0) {
      return out
    }
    const end = Math.min(to, this.count - 1)
    for (let i = from; i <= end; i++) {
      const size = custom[i]
      const hidden = hiddenFlags[i] === 1
      if (hidden || size === size) {
        out.push({
          index: i,
          size: size === size ? size : null,
          hidden,
          ...(this.auto![i] ? { auto: true as const } : {}),
        })
      }
    }
    return out
  }

  private ensure() {
    if (!this.custom) {
      this.custom = new Float64Array(this.count).fill(Number.NaN)
      this.hidden = new Uint8Array(this.count)
      this.auto = new Uint8Array(this.count)
      this.tree = new Float64Array(this.count + 1)
    }
  }

  private bump(index: number, delta: number) {
    if (delta === 0 || !this.tree) {
      return
    }
    for (let i = index + 1; i <= this.count; i += i & -i) {
      this.tree[i] += delta
    }
  }

  /** O(n) Fenwick construction from the custom/hidden arrays. */
  private rebuild() {
    const tree = this.tree!
    this.customCount = 0
    this.hiddenCount = 0
    const custom = this.custom!
    const hidden = this.hidden!
    const base = this.defaultSize
    for (let i = 0; i < this.count; i++) {
      const size = custom[i]
      const isCustom = size === size // not NaN
      if (isCustom) this.customCount++
      if (hidden[i]) this.hiddenCount++
      tree[i + 1] = hidden[i] ? -base : isCustom ? size - base : 0
    }
    for (let i = 1; i <= this.count; i++) {
      const parent = i + (i & -i)
      if (parent <= this.count) {
        tree[parent] += tree[i]
      }
    }
  }
}

function highestPowerOfTwo(n: number) {
  let p = 1
  while (p * 2 <= n) p *= 2
  return p
}
