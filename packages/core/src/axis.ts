import type { AxisEntry } from "./ops"

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
  /** Outline levels (0–7, Data › Group) and the collapsed mark on a group's summary line; allocated on first use. */
  private levels: Uint8Array | null = null
  private collapsedFlags: Uint8Array | null = null
  /** How many entries are at each outline level (index 0 unused), for `maxLevel` without a scan. */
  private levelCounts = new Int32Array(8)

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

  /** The entry's outline level: 0 outside any group, up to 7 (Excel's limit). */
  level(index: number): number {
    return this.levels?.[index] ?? 0
  }

  setLevel(index: number, level: number) {
    const next = Math.max(0, Math.min(MAX_OUTLINE, Math.round(level)))
    const before = this.level(index)
    if (before === next) return
    if (!this.levels) this.levels = new Uint8Array(this.count)
    this.levels[index] = next
    if (before) this.levelCounts[before]--
    if (next) this.levelCounts[next]++
  }

  /** The deepest outline level in use (0: no groups). */
  get maxLevel(): number {
    for (let l = MAX_OUTLINE; l > 0; l--) if (this.levelCounts[l] > 0) return l
    return 0
  }

  /** Whether the group this entry summarises is collapsed (the + on the outline bar). */
  isCollapsed(index: number) {
    return this.collapsedFlags !== null && this.collapsedFlags[index] === 1
  }

  setCollapsed(index: number, collapsed: boolean) {
    if (this.isCollapsed(index) === collapsed) return
    if (!this.collapsedFlags) this.collapsedFlags = new Uint8Array(this.count)
    this.collapsedFlags[index] = collapsed ? 1 : 0
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

  /**
   * Inserts `n` default-sized, shown entries at `at`; the last `n` entries fall off the end. New entries inside a
   * group join it, as in Excel (the lower of the levels on either side).
   */
  insert(at: number, n: number) {
    this.shiftOutline(at, n)
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
    this.shiftOutline(at, -n)
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
  overrides(from: number, to: number): AxisEntry[] {
    const out: AxisEntry[] = []
    const custom = this.custom
    const hiddenFlags = this.hidden
    const levels = this.levels
    const sized = custom && hiddenFlags && this.customCount + this.hiddenCount > 0
    if (!sized && !(levels && this.maxLevel) && !this.collapsedFlags) {
      return out
    }
    const end = Math.min(to, this.count - 1)
    for (let i = from; i <= end; i++) {
      const size = sized ? custom[i] : Number.NaN
      const hidden = sized ? hiddenFlags[i] === 1 : false
      const level = levels ? levels[i] : 0
      const collapsed = this.isCollapsed(i)
      if (hidden || size === size || level || collapsed) {
        out.push({
          index: i,
          size: size === size ? size : null,
          hidden,
          ...(sized && this.auto![i] ? { auto: true as const } : {}),
          ...(level || collapsed ? { level, ...(collapsed ? { collapsed: true as const } : {}) } : {}),
        })
      }
    }
    return out
  }

  /** Moves outline levels and collapsed marks for an insert (n > 0) or a removal (n < 0) at `at`. */
  private shiftOutline(at: number, n: number) {
    for (const flags of [this.levels, this.collapsedFlags]) {
      if (!flags) continue
      if (n > 0) {
        flags.copyWithin(at + n, at, this.count - n)
        const fill = flags === this.levels ? Math.min(at > 0 ? flags[at - 1] : 0, flags[at + n] ?? 0) : 0
        flags.fill(fill, at, at + n)
      } else {
        flags.copyWithin(at, at - n)
        flags.fill(0, this.count + n)
      }
    }
    if (this.levels) {
      this.levelCounts.fill(0)
      for (let i = 0; i < this.count; i++) if (this.levels[i]) this.levelCounts[this.levels[i]]++
    }
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

/** Excel's deepest outline level. */
export const MAX_OUTLINE = 7

function highestPowerOfTwo(n: number) {
  let p = 1
  while (p * 2 <= n) p *= 2
  return p
}
