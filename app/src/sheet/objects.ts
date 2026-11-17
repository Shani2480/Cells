import {
  allObjects,
  copyObject,
  defaultShape,
  defaultTextBox,
  isChart,
  nextObjectId,
  objectName,
  plainText,
  renumber,
  textFromPlain,
  type DrawingObject,
  type SheetGroup,
  type SheetObject,
  type SheetPicture,
  type SheetShape,
} from "@cells/core"
import { anchorOf, chartBox, type ChartBox } from "./charts"
import type { SheetController } from "./controller"

/** Excel's size for a shape put in with a click: one inch (96 pixels) square. */
export const SHAPE_SIZE = 96

/** Excel's names for Align (Shape Format › Arrange › Align). */
export type AlignTo = "left" | "center" | "right" | "top" | "middle" | "bottom" | "distributeH" | "distributeV"

/**
 * Drawing objects on the sheet (Insert › Shapes, Text Box, Pictures, Icons, WordArt; Shape and Picture Format):
 * selecting, putting in, moving, sizing, turning, their text and look, stacking order, grouping. Charts are
 * objects too: selecting one here selects it for the Chart Design tab.
 */
export class ObjectCommands {
  /** The selected objects' ids (Ctrl+click and Shift+click add more), in the order picked. */
  selected: number[] = []
  /** Insert › Shapes waiting for a drag on the sheet (the cursor is a crosshair until then). */
  pending: { geometry: string; textBox?: boolean } | null = null
  /** The object whose text is being typed (a text box right after it's drawn, or a double-click). */
  editing: number | null = null
  /** Home › Find & Select › Selection Pane. */
  paneOpen = false

  constructor(private readonly c: SheetController) {}

  /** Objects on the sheet shown, in drawing order. */
  get list(): SheetObject[] {
    return this.c.sheet.objects
  }

  find(id: number): SheetObject | undefined {
    for (const o of allObjects(this.list)) if (o.id === id) return o
    return undefined
  }

  /** The one object selected (null with none or several). */
  active(): SheetObject | null {
    return this.selected.length === 1 ? (this.find(this.selected[0]) ?? null) : null
  }

  /** The selected shapes, text boxes, pictures and groups (not charts). */
  drawings(): DrawingObject[] {
    return this.selected.map((id) => this.find(id)).filter((o): o is DrawingObject => !!o && !isChart(o))
  }

  /** The one shape selected (Shape Format), or null. */
  shape(): SheetShape | null {
    const o = this.active()
    return o && !isChart(o) && o.kind === "shape" ? o : null
  }

  picture(): SheetPicture | null {
    const o = this.active()
    return o && !isChart(o) && o.kind === "picture" ? o : null
  }

  select(ids: number[] | number | null, add = false) {
    const list = ids === null ? [] : Array.isArray(ids) ? ids : [ids]
    const next = add
      ? [...this.selected.filter((x) => !list.includes(x)), ...list.filter((x) => !this.selected.includes(x))]
      : list
    if (next.length === this.selected.length && next.every((x, i) => x === this.selected[i])) return
    if (next.length && this.c.edit) this.c.commitEdit()
    this.selected = next
    if (this.editing !== null && !next.includes(this.editing)) this.editing = null
    this.c.changed()
  }

  /** Ctrl+A with an object selected, and Select Objects: every object on the sheet. */
  selectAll() {
    this.select(this.list.filter((o) => isChart(o) || !o.hidden).map((o) => o.id))
  }

  // ── Changes ──

  /** The sheet's drawing replaced as one undo step. */
  commit(objects: SheetObject[], label: string): boolean {
    if (this.c.edit) this.c.commitEdit()
    const sel = this.c.selection
    return (
      this.c.transact([{ kind: "objects", sheet: this.c.sheetIndex, objects }], { label, before: sel, after: sel }) !==
      null
    )
  }

  /** Each selected object (or `ids`) changed by `change` on a copy (at any depth in groups), as one undo step. */
  update(label: string, change: (o: SheetObject) => void, ids: readonly number[] = this.selected): boolean {
    if (!ids.length) return false
    const map = (o: SheetObject): SheetObject => {
      let next = o
      if (ids.includes(o.id)) {
        next = copyObject(o)
        change(next)
      }
      if (!isChart(next) && next.kind === "group") {
        const items = next.items.map(map)
        if (items.some((x, i) => x !== (next as SheetGroup).items[i])) next = { ...(next as SheetGroup), items }
      }
      return next
    }
    const objects = this.list.map(map)
    if (objects.every((o, i) => o === this.list[i])) return false
    return this.commit(objects, label)
  }

  private put(object: DrawingObject, label: string): DrawingObject {
    this.commit([...this.list, object], label)
    this.selected = [object.id]
    this.c.changed()
    // Excel shows the object's Format tab once it's in.
    this.c.shell.showTab?.(object.kind === "picture" ? "TabPictureToolsFormat" : "TabDrawingToolsFormat")
    return object
  }

  /** A box in the middle of the window, `w` × `h`. */
  private middle(w: number, h: number): ChartBox {
    const zoom = this.c.zoom
    return {
      x: Math.max(0, this.c.scrollX + (this.c.viewWidth / zoom - w) / 2),
      y: Math.max(0, this.c.scrollY + (this.c.viewHeight / zoom - h) / 2),
      w,
      h,
    }
  }

  private nextId() {
    return nextObjectId(this.c.wb.sheets)
  }

  // ── Insert ──

  /** Insert › Shapes › a shape: waits for a drag (or a click) on the sheet. */
  startShape(geometry: string, textBox = false) {
    if (this.c.edit) this.c.commitEdit()
    this.pending = { geometry, textBox }
    this.select(null)
    this.c.changed()
  }

  cancelPending() {
    if (!this.pending) return
    this.pending = null
    this.c.changed()
  }

  /** The pending shape drawn at a box (a click: Excel's 1-inch shape there). Text boxes start typing. */
  drawPending(box: ChartBox): DrawingObject | null {
    const p = this.pending
    if (!p) return null
    this.pending = null
    const sized =
      box.w < 4 && box.h < 4 ? { ...box, w: p.textBox ? 144 : SHAPE_SIZE, h: p.textBox ? 36 : SHAPE_SIZE } : box
    const o = p.textBox ? this.insertTextBox(sized) : this.insertShape(p.geometry, sized)
    return o
  }

  insertShape(geometry: string, box: ChartBox = this.middle(SHAPE_SIZE, SHAPE_SIZE)): SheetShape {
    const id = this.nextId()
    const shape: SheetShape = {
      kind: "shape",
      id,
      name: objectName({ kind: "shape", geometry }, id),
      anchor: anchorOf(this.c.sheet, box),
      geometry,
      ...defaultShape(geometry),
    }
    return this.put(shape, "Insert Shape") as SheetShape
  }

  insertTextBox(box: ChartBox = this.middle(144, 36)): SheetShape {
    const id = this.nextId()
    const shape: SheetShape = {
      kind: "shape",
      id,
      name: objectName({ kind: "shape", textBox: true }, id),
      anchor: anchorOf(this.c.sheet, box),
      geometry: "rect",
      textBox: true,
      ...defaultTextBox(),
    }
    this.put(shape, "Insert Text Box")
    this.editing = id
    this.c.changed()
    return shape
  }

  /** Insert › Pictures (and Icons): at its own size, scaled down to fit the window, at the active cell. */
  insertPicture(src: string, natural: { w: number; h: number }, extra: Partial<SheetPicture> = {}): SheetPicture {
    const id = this.nextId()
    const zoom = this.c.zoom
    const room = { w: (this.c.viewWidth / zoom) * 0.8, h: (this.c.viewHeight / zoom) * 0.8 }
    const k = Math.min(1, room.w / natural.w, room.h / natural.h)
    const { row, col } = this.c.selection.active
    const x = this.c.sheet.cols.offset(col)
    const y = this.c.sheet.rows.offset(row)
    const picture: SheetPicture = {
      kind: "picture",
      id,
      name: objectName({ kind: "picture", icon: extra.icon }, id),
      anchor: anchorOf(this.c.sheet, { x, y, w: natural.w * k, h: natural.h * k }),
      src,
      natural,
      placement: "move",
      ...extra,
    }
    return this.put(picture, extra.icon ? "Insert Icon" : "Insert Picture") as SheetPicture
  }

  /** Insert › WordArt › a style: "Your text here" in the middle of the window, selected for typing. */
  insertWordArt(look: Pick<SheetShape, "wordArt"> & { color?: SheetShape["fill"] }): SheetShape {
    const id = this.nextId()
    const shape: SheetShape = {
      kind: "shape",
      id,
      name: objectName({ kind: "shape", geometry: "rect" }, id),
      anchor: anchorOf(this.c.sheet, this.middle(320, 72)),
      geometry: "rect",
      fill: { kind: "none" },
      line: { color: null, width: 0 },
      text: {
        paragraphs: [{ align: "ctr", runs: [{ text: "Your text here", size: 36 }] }],
        anchor: "ctr",
        wrap: false,
        autoFit: true,
      },
      wordArt: look.wordArt,
    }
    return this.put(shape, "Insert WordArt") as SheetShape
  }

  // ── Place, size, turn ──

  /** Moves or sizes objects to boxes (pixels at 100%), as one step. */
  place(boxes: Map<number, ChartBox>, label = "Move Object") {
    this.update(
      label,
      (o) => {
        const box = boxes.get(o.id)
        if (box) o.anchor = anchorOf(this.c.sheet, box)
      },
      [...boxes.keys()],
    )
  }

  box(o: SheetObject): ChartBox {
    return chartBox(this.c.sheet, o.anchor)
  }

  /** Arrow keys with objects selected: a nudge of a pixel (Ctrl: more). */
  nudge(dx: number, dy: number) {
    const boxes = new Map<number, ChartBox>()
    for (const id of this.selected) {
      const o = this.find(id)
      if (!o || !this.list.includes(o)) continue
      const b = this.box(o)
      boxes.set(id, { ...b, x: Math.max(0, b.x + dx), y: Math.max(0, b.y + dy) })
    }
    this.place(boxes, "Nudge")
  }

  rotate(degrees: number, absolute = false) {
    this.update("Rotate", (o) => {
      if (isChart(o)) return
      const r = (((absolute ? degrees : (o.rotation ?? 0) + degrees) % 360) + 360) % 360
      if (r) o.rotation = Math.round(r * 100) / 100
      else delete o.rotation
    })
  }

  flip(axis: "h" | "v") {
    this.update("Flip", (o) => {
      if (isChart(o)) return
      const key = axis === "h" ? "flipH" : "flipV"
      if (o[key]) delete o[key]
      else o[key] = true
    })
  }

  /** Shape Format › Size: height and width in pixels (Lock aspect ratio: the other follows). */
  resize(w: number | null, h: number | null, keepRatio = false) {
    const boxes = new Map<number, ChartBox>()
    for (const o of this.drawingsAndCharts()) {
      const b = this.box(o)
      let nw = w ?? b.w
      let nh = h ?? b.h
      if (keepRatio && w !== null && h === null) nh = (b.h * nw) / b.w
      if (keepRatio && h !== null && w === null) nw = (b.w * nh) / b.h
      boxes.set(o.id, { ...b, w: Math.max(1, nw), h: Math.max(1, nh) })
    }
    this.place(boxes, "Size")
  }

  private drawingsAndCharts(): SheetObject[] {
    return this.selected.map((id) => this.list.find((o) => o.id === id)).filter((o): o is SheetObject => !!o)
  }

  // ── Text ──

  startText(id: number) {
    const o = this.find(id)
    if (!o || isChart(o) || o.kind !== "shape" || o.connector) return
    if (!this.selected.includes(id)) this.selected = [id]
    this.editing = id
    this.c.changed()
  }

  /** The text typed into a shape (one undo step); a text box that grows with its text keeps its width. */
  setText(id: number, value: string) {
    this.editing = null
    const o = this.find(id)
    if (!o || isChart(o) || o.kind !== "shape") return this.c.changed()
    if (plainText(o.text) === value) return this.c.changed()
    this.update("Typing", (x) => ((x as SheetShape).text = textFromPlain(value, (x as SheetShape).text)), [id])
  }

  // ── Delete ──

  delete() {
    if (!this.selected.length) return
    const gone = new Set(this.selected)
    const strip = (list: SheetObject[]): SheetObject[] =>
      list
        .filter((o) => !gone.has(o.id))
        .map((o) => (!isChart(o) && o.kind === "group" ? { ...o, items: strip(o.items) } : o))
    this.commit(strip(this.list), this.selected.length === 1 && isChart(this.active()!) ? "Delete Chart" : "Delete")
    this.selected = []
    this.editing = null
    this.c.changed()
  }

  /** Copies of the selected objects a bit lower and to the right (Ctrl+D, Duplicate). */
  duplicate() {
    const picked = this.list.filter((o) => this.selected.includes(o.id) && !isChart(o)) as DrawingObject[]
    if (!picked.length) return
    let id = this.nextId()
    const copies = picked.map((o) => {
      const copy = copyObject(o)
      delete copy.source
      id = renumber(copy, id)
      copy.name = objectName(copy, copy.id)
      const b = this.box(o)
      copy.anchor = anchorOf(this.c.sheet, { ...b, x: b.x + 12, y: b.y + 12 })
      return copy
    })
    this.commit([...this.list, ...copies], "Duplicate")
    this.selected = copies.map((o) => o.id)
    this.c.changed()
  }

  // ── Order ──

  /** Bring Forward / Send Backward (one step), Bring to Front / Send to Back (all the way). */
  order(how: "forward" | "backward" | "front" | "back") {
    const list = [...this.list]
    const picked = list.filter((o) => this.selected.includes(o.id))
    if (!picked.length) return
    let next: SheetObject[]
    if (how === "front") next = [...list.filter((o) => !picked.includes(o)), ...picked]
    else if (how === "back") next = [...picked, ...list.filter((o) => !picked.includes(o))]
    else {
      next = [...list]
      const step = how === "forward" ? 1 : -1
      const order = how === "forward" ? [...picked].reverse() : picked
      for (const o of order) {
        const i = next.indexOf(o)
        const j = i + step
        if (j < 0 || j >= next.length || picked.includes(next[j])) continue
        ;[next[i], next[j]] = [next[j], next[i]]
      }
    }
    if (next.every((o, i) => o === list[i])) return
    const labels = {
      forward: "Bring Forward",
      backward: "Send Backward",
      front: "Bring to Front",
      back: "Send to Back",
    }
    this.commit(next, labels[how])
  }

  /** Moves an object to a place in the list (the Selection Pane's drag). */
  moveTo(id: number, index: number) {
    const list = [...this.list]
    const i = list.findIndex((o) => o.id === id)
    if (i < 0) return
    const [o] = list.splice(i, 1)
    list.splice(Math.max(0, Math.min(list.length, index)), 0, o)
    if (list.every((x, k) => x === this.list[k])) return
    this.commit(list, "Reorder Objects")
  }

  // ── Align ──

  /** Align the selected objects to each other (one: to the sheet's top-left), or spread them evenly. */
  align(to: AlignTo) {
    const items = this.drawingsAndCharts().map((o) => ({ o, b: this.box(o) }))
    if (!items.length) return
    const left = items.length > 1 ? Math.min(...items.map((i) => i.b.x)) : 0
    const top = items.length > 1 ? Math.min(...items.map((i) => i.b.y)) : 0
    const right = Math.max(...items.map((i) => i.b.x + i.b.w))
    const bottom = Math.max(...items.map((i) => i.b.y + i.b.h))
    const boxes = new Map<number, ChartBox>()
    if (to === "distributeH" || to === "distributeV") {
      if (items.length < 3) return
      const horizontal = to === "distributeH"
      const sorted = [...items].sort((a, b) => (horizontal ? a.b.x - b.b.x : a.b.y - b.b.y))
      const size = sorted.reduce((s, i) => s + (horizontal ? i.b.w : i.b.h), 0)
      const span = horizontal ? right - left : bottom - top
      const gap = (span - size) / (sorted.length - 1)
      let at = horizontal ? left : top
      for (const i of sorted) {
        boxes.set(i.o.id, horizontal ? { ...i.b, x: at } : { ...i.b, y: at })
        at += (horizontal ? i.b.w : i.b.h) + gap
      }
    } else
      for (const { o, b } of items) {
        const box = { ...b }
        if (to === "left") box.x = left
        if (to === "right") box.x = right - b.w
        if (to === "center") box.x = (left + right) / 2 - b.w / 2
        if (to === "top") box.y = top
        if (to === "bottom") box.y = bottom - b.h
        if (to === "middle") box.y = (top + bottom) / 2 - b.h / 2
        boxes.set(o.id, box)
      }
    this.place(boxes, "Align")
  }

  // ── Group ──

  /** Group: the selected objects (two or more, not charts) in one group, where the top one was. */
  group() {
    const picked = this.list.filter((o) => this.selected.includes(o.id) && !isChart(o)) as DrawingObject[]
    if (picked.length < 2) return
    const boxes = picked.map((o) => this.box(o))
    const x = Math.min(...boxes.map((b) => b.x))
    const y = Math.min(...boxes.map((b) => b.y))
    const w = Math.max(...boxes.map((b) => b.x + b.w)) - x
    const h = Math.max(...boxes.map((b) => b.y + b.h)) - y
    const id = this.nextId()
    const group: SheetGroup = {
      kind: "group",
      id,
      name: objectName({ kind: "group" }, id),
      anchor: anchorOf(this.c.sheet, { x, y, w, h }),
      child: { w, h },
      items: picked.map((o, i) => {
        const item = copyObject(o)
        delete item.source
        item.box = { x: boxes[i].x - x, y: boxes[i].y - y, w: boxes[i].w, h: boxes[i].h }
        return item
      }),
    }
    const top = Math.max(...picked.map((o) => this.list.indexOf(o)))
    const next: SheetObject[] = []
    this.list.forEach((o, i) => {
      if (i === top) next.push(group)
      else if (!picked.includes(o as DrawingObject)) next.push(o)
    })
    this.commit(next, "Group")
    this.selected = [id]
    this.c.changed()
  }

  /** Ungroup: the group's objects back on the sheet where they show, selected. */
  ungroup() {
    const groups = this.list.filter(
      (o) => this.selected.includes(o.id) && !isChart(o) && o.kind === "group",
    ) as SheetGroup[]
    if (!groups.length) return
    const next: SheetObject[] = []
    const freed: number[] = []
    for (const o of this.list) {
      if (!groups.includes(o as SheetGroup)) {
        next.push(o)
        continue
      }
      const g = o as SheetGroup
      const gb = this.box(g)
      const sx = gb.w / (g.child.w || 1)
      const sy = gb.h / (g.child.h || 1)
      for (const item of g.items) {
        const copy = copyObject(item)
        const b = copy.box ?? { x: 0, y: 0, w: g.child.w, h: g.child.h }
        delete copy.box
        if (!isChart(copy)) delete copy.source
        copy.anchor = anchorOf(this.c.sheet, { x: gb.x + b.x * sx, y: gb.y + b.y * sy, w: b.w * sx, h: b.h * sy })
        if (!isChart(copy) && g.rotation) copy.rotation = ((copy.rotation ?? 0) + g.rotation) % 360
        next.push(copy)
        freed.push(copy.id)
      }
    }
    this.commit(next, "Ungroup")
    this.selected = freed
    this.c.changed()
  }

  // ── Clipboard ──

  /** The objects last copied, kept whole for a paste back into Cells. */
  private clip: { token: string; objects: DrawingObject[] } | null = null

  /**
   * Copy (or Cut) of the selected shapes and pictures: a picture of them for other apps, their text as text, and
   * the objects themselves for a paste in Cells.
   */
  copy(cut: boolean, picture: (objects: DrawingObject[]) => string | null): { text: string; html: string } | null {
    const picked = this.list.filter((o) => this.selected.includes(o.id) && !isChart(o)) as DrawingObject[]
    if (!picked.length) return null
    const token = `${Date.now().toString(36)}-${picked[0].id}`
    this.clip = { token, objects: picked.map((o) => copyObject(o)) }
    const img = picture(picked)
    const html = `<!--cells-objects:${token}-->${img ? `<img src="${img}">` : ""}`
    const text = picked
      .map((o) => (o.kind === "shape" ? plainText(o.text) : ""))
      .filter(Boolean)
      .join("\n")
    if (cut) this.delete()
    return { text, html }
  }

  /** Paste of objects Cells copied: copies at the active cell, keeping their places to each other, selected. */
  paste(html: string | null | undefined): boolean {
    const token = html ? /<!--cells-objects:([^>]+?)-->/.exec(html)?.[1] : undefined
    if (!token || !this.clip || this.clip.token !== token) return false
    const sheet = this.c.sheet
    const boxes = this.clip.objects.map((o) => chartBox(sheet, o.anchor))
    const x0 = Math.min(...boxes.map((b) => b.x))
    const y0 = Math.min(...boxes.map((b) => b.y))
    const at = this.c.selection.active
    const ax = sheet.cols.offset(at.col)
    const ay = sheet.rows.offset(at.row)
    let id = this.nextId()
    const copies = this.clip.objects.map((o, i) => {
      const copy = copyObject(o)
      delete copy.source
      id = renumber(copy, id)
      copy.name = objectName(copy, copy.id)
      copy.anchor = anchorOf(sheet, { ...boxes[i], x: ax + boxes[i].x - x0, y: ay + boxes[i].y - y0 })
      return copy
    })
    if (!this.commit([...this.list, ...copies], "Paste")) return false
    this.selected = copies.map((o) => o.id)
    this.c.changed()
    return true
  }

  // ── Keys ──

  /** Keys while objects are selected (true: handled). */
  key(key: string, primary: boolean, shift: boolean): boolean {
    if (!this.selected.length) {
      if (key === "Escape" && this.pending) return (this.cancelPending(), true)
      return false
    }
    if (key === "Delete" || key === "Backspace") return (this.delete(), true)
    if (key === "Escape") return (this.select(null), true)
    const steps: Record<string, [number, number]> = {
      ArrowLeft: [-1, 0],
      ArrowRight: [1, 0],
      ArrowUp: [0, -1],
      ArrowDown: [0, 1],
    }
    if (steps[key]) {
      const k = primary ? 1 : 8
      this.nudge(steps[key][0] * k, steps[key][1] * k)
      return true
    }
    if (key === "Tab") {
      const shown = this.list.filter((o) => isChart(o) || !o.hidden)
      const i = shown.findIndex((o) => o.id === this.selected[this.selected.length - 1])
      const next = shown[(i + (shift ? -1 : 1) + shown.length) % shown.length]
      if (next) this.select(next.id)
      return true
    }
    if ((key === "Enter" || key === "F2") && !primary) {
      const shape = this.shape()
      if (shape && !shape.connector) this.startText(shape.id)
      return true
    }
    if (primary && key.toLowerCase() === "d") return (this.duplicate(), true)
    if (primary && key.toLowerCase() === "a") return (this.selectAll(), true)
    return false
  }

  // ── Selection Pane ──

  setPane(open = !this.paneOpen) {
    this.paneOpen = open
    this.c.changed()
  }

  setHidden(id: number, hidden: boolean) {
    this.update(
      hidden ? "Hide Object" : "Show Object",
      (o) => {
        if (isChart(o)) return
        if (hidden) o.hidden = true
        else delete o.hidden
      },
      [id],
    )
    if (hidden) this.select(this.selected.filter((x) => x !== id))
  }

  /** Hide All / Show All in the Selection Pane. */
  setAllHidden(hidden: boolean) {
    this.update(
      hidden ? "Hide All" : "Show All",
      (o) => {
        if (isChart(o)) return
        if (hidden) o.hidden = true
        else delete o.hidden
      },
      this.list.filter((o) => !isChart(o)).map((o) => o.id),
    )
    if (hidden) this.select(null)
  }

  /** Alt Text: what a screen reader says for an object, or that it's only decoration (one undo step). */
  setAltText(id: number, text: string, decorative: boolean) {
    const o = this.find(id)
    if (!o) return
    const t = text.trim()
    if ((o.altText ?? "") === t && !!o.decorative === decorative) return
    this.update(
      "Alt Text",
      (x) => {
        if (t) x.altText = t
        else delete x.altText
        if (decorative) x.decorative = true
        else delete x.decorative
      },
      [id],
    )
  }

  rename(id: number, name: string) {
    const clean = name.trim()
    if (!clean) return
    this.update("Rename Object", (o) => (o.name = clean), [id])
  }
}
