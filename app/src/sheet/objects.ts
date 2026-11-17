import {
  allObjects,
  copyObject,
  defaultShape,
  defaultTextBox,
  equationFromLinear,
  isEquation,
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
  type SignatureSetup,
  type EmbeddedFile,
  type Model3D,
  normalRotation,
  copySmartArtNodes,
  smartArtLayout,
  smartArtParts,
  SMARTART_LAYOUTS,
  type SmartArt,
  type SmartArtEdit,
  type SmartArtNode,
  type SmartArtPath,
} from "@cells/core"
import { layoutSmartArt } from "@/drawing/smartart"
import { renderModel } from "@/drawing/model3d"
import { textSize } from "@/drawing/render"
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
  /** The open text editor's way to type at its caret, and to close (set while one is open). */
  textInserter: ((text: string) => void) | null = null
  finishText: (() => void) | null = null
  /** Where the caret was when a shape's editor closed (counted from the end of its text). */
  lastCaret: { id: number; fromEnd: number } | null = null
  /** Where the next editor of that shape puts the caret (absent: all its text selected). */
  caret: { id: number; fromEnd: number } | null = null
  /** Home › Find & Select › Selection Pane. */
  paneOpen = false
  /** SmartArt › Text Pane: open beside the selected SmartArt. */
  textPaneOpen = false
  /** The SmartArt point being typed in (the Text Pane's line), what Add Shape, Promote and Demote act on. */
  smartArtPoint: SmartArtPath | null = null

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
  middle(w: number, h: number): ChartBox {
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

  /**
   * Insert › Signature Line: an unsigned line (`png` is its picture for files) at the active cell, Office's size.
   */
  insertSignatureLine(setup: SignatureSetup, png: string, size: { w: number; h: number }): SheetPicture {
    const id = this.nextId()
    const { row, col } = this.c.selection.active
    const x = this.c.sheet.cols.offset(col)
    const y = this.c.sheet.rows.offset(row)
    const picture: SheetPicture = {
      kind: "picture",
      id,
      name: objectName({ kind: "picture" }, id),
      anchor: anchorOf(this.c.sheet, { x, y, w: size.w, h: size.h }),
      src: png,
      natural: { ...size },
      signature: { ...setup },
      altText: "Microsoft Office Signature Line...",
    }
    return this.put(picture, "Insert Signature Line") as SheetPicture
  }

  /** Insert › Object › Create from File: the file as an icon (`png` its picture for files) at the active cell. */
  insertObject(embed: EmbeddedFile, png: string, size: { w: number; h: number }): SheetPicture {
    const id = this.nextId()
    const { row, col } = this.c.selection.active
    const x = this.c.sheet.cols.offset(col)
    const y = this.c.sheet.rows.offset(row)
    const picture: SheetPicture = {
      kind: "picture",
      id,
      name: `Object ${id}`,
      anchor: anchorOf(this.c.sheet, { x, y, w: size.w, h: size.h }),
      src: png,
      natural: { ...size },
      embed: { ...embed },
      placement: "move",
    }
    return this.put(picture, "Insert Object") as SheetPicture
  }

  /**
   * Insert › 3D Models: the model (as Office sets up a new one) at the active cell, its frame fitted to it. Null
   * where the model can't be drawn (no WebGL).
   */
  async insertModel3D(model: Model3D): Promise<SheetPicture | null> {
    const pic = await renderModel(model)
    if (!pic) return null
    const id = this.nextId()
    const { row, col } = this.c.selection.active
    const x = this.c.sheet.cols.offset(col)
    const y = this.c.sheet.rows.offset(row)
    const picture: SheetPicture = {
      kind: "picture",
      id,
      name: `3D Model ${id}`,
      anchor: anchorOf(this.c.sheet, { x, y, w: pic.size.w, h: pic.size.h }),
      src: pic.src,
      model3d: model,
      placement: "move",
    }
    const out = this.put(picture, "Insert 3D Model") as SheetPicture
    this.c.shell.showTab?.("TabModel3DFormat")
    return out
  }

  /**
   * Insert › SmartArt: a layout with Office's empty points to type in (its own number of them), in the middle of the
   * window at Office's size, the Text Pane open.
   */
  insertSmartArt(layout: string, colors?: string): SheetGroup {
    const id = this.nextId()
    const def = smartArtLayout(layout) ?? SMARTART_LAYOUTS[0]!
    const size = { w: 480, h: 320 }
    const smartArt: SmartArt = {
      layout: def.id,
      nodes: Array.from({ length: def.starts }, () => ({ text: "", children: [] })),
      ...(colors ? { colors } : {}),
    }
    smartArt.parts = smartArtParts(smartArt)
    const group: SheetGroup = {
      kind: "group",
      id,
      name: `Diagram ${id}`,
      anchor: anchorOf(this.c.sheet, this.middle(size.w, size.h)),
      child: size,
      items: layoutSmartArt(smartArt, size, this.c.wb.theme),
      smartArt,
    }
    this.put(group, "Insert SmartArt")
    this.textPaneOpen = true
    this.smartArtPoint = [0]
    this.c.shell.showTab?.("TabSmartArtDesign")
    return group
  }

  /** The selected SmartArt, if one is selected alone. */
  smartArt(): (SheetGroup & { smartArt: SmartArt }) | null {
    const d = this.drawings()
    const g = d.length === 1 && d[0]!.kind === "group" ? (d[0] as SheetGroup) : null
    return g?.smartArt ? (g as SheetGroup & { smartArt: SmartArt }) : null
  }

  /**
   * A SmartArt changed (its text, layout or colours): laid out again in its frame and its parts written anew (an
   * Office layout Cells doesn't know keeps its definition; Office lays it out again when it opens the file).
   */
  setSmartArt(id: number, change: Partial<Pick<SmartArt, "nodes" | "layout" | "colors">>, label: string): boolean {
    this.smartArtPreview = null
    return this.update(
      label,
      (o) => {
        if (isChart(o) || o.kind !== "group" || !o.smartArt) return
        const next: SmartArt = { ...o.smartArt, ...change }
        if (change.nodes) next.nodes = copySmartArtNodes(change.nodes)
        next.parts = smartArtParts(next)
        o.smartArt = next
        const box = this.box(o)
        o.child = { w: box.w, h: box.h }
        o.items = layoutSmartArt(next, o.child, this.c.wb.theme)
      },
      [id],
    )
  }

  /**
   * SmartArt Design › Create Graphic (Add Shape, Promote, Demote, Move Up/Down) and the Text Pane's keys: the tree
   * changed by `edit` from the current point, as one undo step; the changed point becomes the current one.
   */
  editSmartArt(edit: (nodes: SmartArtNode[], path: SmartArtPath | null) => SmartArtEdit, label: string): boolean {
    const sa = this.smartArt()
    if (!sa) return false
    const done = edit(sa.smartArt.nodes, this.smartArtPoint)
    if (!done) return false
    this.smartArtPreview = null
    this.smartArtPoint = done.path.length ? done.path : null
    return this.setSmartArt(sa.id, { nodes: done.nodes }, label)
  }

  /** Typing in the Text Pane: the SmartArt's shapes with the text so far (null: none), not yet a step to undo. */
  smartArtPreview: { id: number; items: SheetGroup["items"] } | null = null

  previewSmartArt(id: number, nodes: SmartArtNode[] | null) {
    const o = this.find(id)
    this.smartArtPreview =
      nodes && o && !isChart(o) && o.kind === "group" && o.smartArt
        ? { id, items: layoutSmartArt({ ...o.smartArt, nodes }, o.child, this.c.wb.theme) }
        : null
    this.c.changed()
  }

  /** SmartArt Design › Convert to Shapes: the SmartArt's shapes as a plain group. */
  convertToShapes(id: number): boolean {
    return this.update(
      "Convert to Shapes",
      (o) => {
        if (isChart(o) || o.kind !== "group") return
        delete o.smartArt
        delete o.source
        const base = this.nextId()
        o.items.forEach((item, i) => {
          if (!isChart(item)) {
            item.id = base + i
            item.name = item.name || objectName(item, item.id)
          }
        })
      },
      [id],
    )
  }

  setTextPane(open = !this.textPaneOpen) {
    this.textPaneOpen = open
    this.c.changed()
  }

  /** The selected 3D model, if one is selected alone. */
  model(): (SheetPicture & { model3d: Model3D }) | null {
    const p = this.picture()
    return p?.model3d ? (p as SheetPicture & { model3d: Model3D }) : null
  }

  /**
   * A 3D model's view changed (turned, a view from the gallery, Reset): drawn anew, its frame fitted to the model
   * around the same middle (`keepSize`: the frame stays, the model scaled into it), as one undo step.
   */
  async setModelView(
    id: number,
    change: Partial<Pick<Model3D, "rot" | "viewport" | "data" | "xml">>,
    label: string,
    keepSize = false,
  ): Promise<boolean> {
    const o = this.find(id)
    if (!o || isChart(o) || o.kind !== "picture" || !o.model3d) return false
    const model: Model3D = { ...o.model3d, ...change }
    if (change.rot) model.rot = normalRotation(change.rot)
    const box = this.box(o)
    // Kept size: the turned model scaled to fit the frame it has.
    const pic = await renderModel(model, keepSize ? { box, fit: true } : {})
    if (!pic) return false
    model.viewport = pic.viewport
    const size = keepSize ? box : pic.size
    const next = { x: box.x + (box.w - size.w) / 2, y: box.y + (box.h - size.h) / 2, w: size.w, h: size.h }
    const src = pic.src
    return this.update(
      label,
      (x) => {
        if (isChart(x) || x.kind !== "picture") return
        x.model3d = model
        x.src = src
        x.anchor = anchorOf(this.c.sheet, { ...next, x: Math.max(0, next.x), y: Math.max(0, next.y) })
      },
      [id],
    )
  }

  /** A 3D model's picture drawn again for the frame it has now (after it was sized), not a step of its own. */
  private async redrawModel(o: SheetPicture & { model3d: Model3D }) {
    const pic = await renderModel(o.model3d, { box: this.box(o) })
    if (!pic || !this.list.includes(o)) return
    o.src = pic.src
    this.c.changed()
  }

  /** Signature Setup on a line already there: new details and the picture drawn from them. */
  setSignature(id: number, setup: SignatureSetup, png: string): boolean {
    return this.update(
      "Signature Setup",
      (o) => {
        if (isChart(o) || o.kind !== "picture") return
        o.signature = { ...setup }
        o.src = png
      },
      [id],
    )
  }

  /**
   * Insert › Equation: an equation in a text box of its own in the middle of the window, typed in its linear form
   * (`linear` empty: a new equation to type).
   */
  insertEquation(linear = ""): SheetShape {
    const id = this.nextId()
    const shape: SheetShape = {
      kind: "shape",
      id,
      name: objectName({ kind: "shape", textBox: true }, id),
      anchor: anchorOf(this.c.sheet, this.middle(160, 30)),
      geometry: "rect",
      textBox: true,
      // As Excel makes them: a white box without an outline (so it reads on a dark sheet too).
      fill: { kind: "solid", color: { theme: "lt1" } },
      line: { color: null, width: 0 },
      text: { ...equationFromLinear(linear), anchor: "t", wrap: false, autoFit: true },
    }
    shape.text!.paragraphs[0]!.runs = [{ text: "", color: { theme: "dk1" } }]
    this.fit(shape, true)
    this.put(shape, "Insert Equation")
    if (!linear) {
      this.editing = id
      this.c.changed()
    }
    return shape
  }

  /** Sizes an equation's box to the equation, keeping its top left (or its middle, for a new one). */
  private fit(shape: SheetShape, centre = false) {
    if (!shape.text) return
    const size = textSize(shape.text, this.c.wb.theme)
    const box = chartBox(this.c.sheet, shape.anchor)
    const x = centre ? box.x + (box.w - size.w) / 2 : box.x
    const y = centre ? box.y + (box.h - size.h) / 2 : box.y
    shape.anchor = anchorOf(this.c.sheet, { x: Math.max(0, x), y: Math.max(0, y), w: size.w, h: size.h })
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
    const sized: number[] = []
    this.update(
      label,
      (o) => {
        const box = boxes.get(o.id)
        if (!box) return
        const was = this.box(o)
        o.anchor = anchorOf(this.c.sheet, box)
        // Cells' own SmartArt is laid out again in its new frame (Office's keep their shapes, scaled).
        if (!isChart(o) && o.kind === "group" && o.smartArt && smartArtLayout(o.smartArt.layout)) {
          o.child = { w: box.w, h: box.h }
          o.items = layoutSmartArt(o.smartArt, o.child, this.c.wb.theme)
        }
        // A 3D model grows with its frame (its view square scaled), and is drawn again for it.
        if (!isChart(o) && o.kind === "picture" && o.model3d && (box.w !== was.w || box.h !== was.h)) {
          o.model3d = { ...o.model3d, viewport: o.model3d.viewport * Math.min(box.w / was.w, box.h / was.h) }
          sized.push(o.id)
        }
      },
      [...boxes.keys()],
    )
    for (const id of sized) {
      const o = this.find(id)
      if (o && !isChart(o) && o.kind === "picture" && o.model3d)
        void this.redrawModel(o as SheetPicture & { model3d: Model3D })
    }
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
    if (isEquation(o)) {
      // An equation typed in its linear form builds up, and its box fits it (Office's equations grow as you type).
      this.update(
        "Typing",
        (x) => {
          const shape = x as SheetShape
          shape.text = equationFromLinear(value, shape.text)
          this.fit(shape)
        },
        [id],
      )
      return
    }
    this.update("Typing", (x) => ((x as SheetShape).text = textFromPlain(value, (x as SheetShape).text)), [id])
  }

  /** The Equation tab's symbols and structures: typed at the caret of the equation's editor, which stays open. */
  insertIntoEquation(id: number, linear: string) {
    if (this.editing === id && this.textInserter) return this.textInserter(linear)
    const o = this.find(id)
    if (!isEquation(o)) return
    const text = plainText(o.text)
    const fromEnd = this.lastCaret?.id === id ? Math.min(this.lastCaret.fromEnd, text.length) : 0
    const pos = text.length - fromEnd
    this.setText(id, text.slice(0, pos) + linear + text.slice(pos))
    // Typing goes on just after what went in, as in Office.
    this.caret = { id, fromEnd }
    this.startText(id)
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
      (o) => this.selected.includes(o.id) && !isChart(o) && o.kind === "group" && !o.smartArt,
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
