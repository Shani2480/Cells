import { copiedSheetCharts, copyChart, nextChartId, type ChartAnchor, type SheetChart } from "./chart"
import { parseLinear, toLinear } from "./math/linear"
import type { MathList } from "./math/model"
import type { Sheet } from "./sheet"

/**
 * Drawing objects on a sheet besides charts: shapes, text boxes, pictures and groups of them (Insert › Illustrations
 * and Text). With the charts they make one list in drawing order, `Sheet.objects` (the last on top), as Excel keeps
 * them in a sheet's drawing.
 *
 * Sizes are pixels at 100% zoom; colours are Excel's: an RGB value or a theme colour with its tint, so a shape keeps
 * following the workbook's theme.
 */

/** A theme colour slot, as DrawingML names them. */
export type ThemeSlot =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink"

/**
 * A colour: `rgb` ("4472C4"), or a theme slot. `lumMod`/`lumOff` (0–1) lighten or darken it as Excel's colour
 * picker rows do (Accent 1, Lighter 40%: lumMod 0.6, lumOff 0.4); `alpha` is its opacity (0–1, absent: solid).
 */
export type ShapeColor = { rgb?: string; theme?: ThemeSlot; lumMod?: number; lumOff?: number; alpha?: number }

export type ShapeFill = { kind: "none" } | { kind: "solid"; color: ShapeColor }

export type DashStyle = "solid" | "dot" | "dash" | "lgDash" | "dashDot" | "sysDot" | "sysDash"
export type LineEnd = "none" | "triangle" | "arrow" | "stealth" | "diamond" | "oval"

/** A shape's outline; null colour: No Outline. Width in points. */
export type ShapeLine = {
  color: ShapeColor | null
  width: number
  dash?: DashStyle
  head?: LineEnd
  tail?: LineEnd
}

export type TextRun = {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  /** Points. */
  size?: number
  font?: string
  color?: ShapeColor
}

export type TextParagraph = {
  align?: "l" | "ctr" | "r" | "just"
  runs: TextRun[]
  /**
   * Insert › Equation: the paragraph is an equation (Office's math paragraph), its lines as math trees. `runs` then
   * holds one run without text for the equation's look (size, colour).
   */
  math?: MathList[]
}

/** Text in a shape or text box: paragraphs, where it sits, and its insets (pixels). */
export type ShapeText = {
  paragraphs: TextParagraph[]
  anchor?: "t" | "ctr" | "b"
  /** Text direction: horizontal, or turned (vert: 90°, vert270: 270°). */
  vertical?: "horz" | "vert" | "vert270"
  wrap?: boolean
  /** Resize shape to fit text. */
  autoFit?: boolean
  insets?: { l: number; t: number; r: number; b: number }
}

/** WordArt's text effects: an outline on the letters, a shadow, a glow, a reflection. */
export type TextEffects = {
  outline?: ShapeLine
  shadow?: boolean
  glow?: ShapeColor
  reflection?: boolean
}

/** A box inside a group, in the group's own pixels (it scales with the group). */
export type ChildBox = { x: number; y: number; w: number; h: number }

type ObjectBase = {
  id: number
  /** "Rectangle 1", "TextBox 2", "Picture 3": Excel's names, shown in the Selection Pane and the Name Box. */
  name: string
  anchor: ChartAnchor
  /** Degrees, clockwise. */
  rotation?: number
  flipH?: boolean
  flipV?: boolean
  /** Selection Pane › the eye: hidden objects aren't drawn or printed. */
  hidden?: boolean
  altText?: string
  decorative?: boolean
  /**
   * What the file had for this object, kept so an object Cells didn't change is written back exactly (effects,
   * custom geometry and other details Cells doesn't model). `sourceKey` is the object as read; when it no longer
   * matches, the object is written from what Cells knows.
   */
  source?: { xml: string; key: string; part: string; rels?: Record<string, string> }
  /** Format › Properties: move and size with cells (absent), move but don't size, or neither. */
  placement?: "move" | "fixed"
  /** Inside a group: where it sits in the group. */
  box?: ChildBox
}

export type SheetShape = ObjectBase & {
  kind: "shape"
  /** DrawingML preset geometry ("rect", "ellipse", "rightArrow", "line"…). */
  geometry: string
  /** Adjust values (the yellow handles), by name ("adj", "adj1"…), as DrawingML's 100000ths. */
  adjust?: Record<string, number>
  /** Insert › Text Box: a shape that's mostly text (no fill or outline by default). */
  textBox?: boolean
  /** A line or arrow (Excel's connector shapes): no fill or text. */
  connector?: boolean
  fill: ShapeFill
  line: ShapeLine
  text?: ShapeText
  /** WordArt: the letters' fill and effects. */
  wordArt?: { fill?: ShapeFill; effects?: TextEffects }
  shadow?: boolean
}

/**
 * Insert › Signature Line: who is asked to sign, as Office's Signature Setup sets it. The line is a picture drawn from
 * these (an X, the line, the signer's name and title); the picture itself (`src`) is what files show.
 */
export type SignatureSetup = {
  /** The line's GUID ("{…}"), which a digital signature would name. */
  id: string
  /** The signature provider's GUID as the file had it (Office's own when absent). */
  provider?: string
  signer: string
  title: string
  email: string
  instructions: string
  allowComments: boolean
  showDate: boolean
}

/** Office's default instructions to the signer. */
export const SIGNING_INSTRUCTIONS = "Before signing this document, verify that the content you are signing is correct."

/**
 * Insert › 3D Models: a model (glTF binary) and the view of it. Office keeps the camera, the model's scale and its
 * lights in the file's `am3d:model3d` element (`xml`, kept and written back); Cells turns and zooms it.
 */
export type Model3D = {
  /** The model as a .glb file. */
  data: Uint8Array
  /** Turned about x, y and z, in degrees (Format 3D Model › Rotation). */
  rot: { x: number; y: number; z: number }
  /** The view's square, pixels across: the model's size on the sheet follows it (Pan & Zoom). */
  viewport: number
  /** The model's `am3d:model3d` element as written (camera, scale, lights). */
  xml?: string
}

/**
 * Insert › Object: a file embedded in the workbook, shown as its icon and name and opened in its own app. `src` is
 * the icon picture files carry.
 */
export type EmbeddedFile = {
  /** The file's name ("Shopping list.txt"): the label under the icon, and the name it opens under. */
  name: string
  /** The file itself; absent for an object Cells can't unpack (then `raw` keeps it). */
  data?: Uint8Array
  /** The OLE class: "Package" for any file, "Word.Document.12" and the like for Office documents. */
  progId: string
  /** The embedding as the file had it, written back as it came (objects of other apps). */
  raw?: { bytes: Uint8Array; ext: string; contentType: string; relType: string }
}

export type SheetPicture = ObjectBase & {
  kind: "picture"
  /** An embedded file (Insert › Object): drawn as its icon and name. */
  embed?: EmbeddedFile
  /** A signature line (Insert › Signature Line): drawn from its setup. */
  signature?: SignatureSetup
  /** A 3D model (Insert › 3D Models): `src` is its picture from its view. */
  model3d?: Model3D
  /** The picture as a data URL ("data:image/png;base64,…"). */
  src: string
  /** An SVG picture (icons): drawn instead of `src`, which is its PNG for apps that can't show SVG. */
  svg?: string
  /** Its size in pixels as it came (Reset Picture). */
  natural?: { w: number; h: number }
  /** Crop from each side, as fractions of the picture (0–1). */
  crop?: { l: number; t: number; r: number; b: number }
  line?: ShapeLine
  /** An icon (Insert › Icons): drawn from SVG, recoloured with Graphics Fill. */
  icon?: boolean
}

export type SheetGroup = ObjectBase & {
  kind: "group"
  /** The group's own size the children's boxes are in. */
  child: { w: number; h: number }
  items: GroupItem[]
  /** A SmartArt graphic: the group is its shapes as laid out, this its text and layout. */
  smartArt?: SmartArt
}

/**
 * Insert › SmartArt: a layout of text (lists, processes, cycles, hierarchies…). The text is a tree, as the Text Pane
 * shows it; the shapes are laid out from it.
 */
export type SmartArt = {
  /** The layout's id ("urn:microsoft.com/office/officeart/2005/8/layout/process1"). */
  layout: string
  /** The text: each point, with the points under it. */
  nodes: SmartArtNode[]
  /** Its colours and style (SmartArt Design › Change Colors, SmartArt Styles), by their ids. */
  colors?: string
  quickStyle?: string
  /**
   * The SmartArt's parts as the file had them (data, layout definition, quick style, colours and Office's drawing of
   * it), written back while it's unchanged.
   */
  parts?: SmartArtParts
}

export type SmartArtNode = { text: string; children: SmartArtNode[] }
export type SmartArtParts = { data: string; layout: string; quickStyle: string; colors: string; drawing?: string }

/** A copy of a SmartArt's text tree. */
export const copySmartArtNodes = (nodes: SmartArtNode[]): SmartArtNode[] =>
  nodes.map((n) => ({ text: n.text, children: copySmartArtNodes(n.children) }))

export type DrawingObject = SheetShape | SheetPicture | SheetGroup
export type GroupItem = DrawingObject | SheetChart
/** Everything in a sheet's drawing, in order. */
export type SheetObject = DrawingObject | SheetChart

export const isChart = (o: SheetObject): o is SheetChart => !("kind" in o)

export function copyObject<T extends SheetObject>(o: T): T {
  if (isChart(o)) return copyChart(o) as T
  const out = {
    ...o,
    anchor: { from: { ...o.anchor.from }, to: { ...o.anchor.to } },
    box: o.box ? { ...o.box } : undefined,
  } as DrawingObject
  if (out.box === undefined) delete out.box
  if (out.kind === "shape") {
    const s = o as SheetShape
    out.fill = copyFill(s.fill)
    out.line = copyLine(s.line)!
    if (s.adjust) out.adjust = { ...s.adjust }
    if (s.text) out.text = copyText(s.text)
    if (s.wordArt)
      out.wordArt = {
        fill: s.wordArt.fill && copyFill(s.wordArt.fill),
        effects: s.wordArt.effects && {
          ...s.wordArt.effects,
          outline: copyLine(s.wordArt.effects.outline),
          glow: s.wordArt.effects.glow && { ...s.wordArt.effects.glow },
        },
      }
  } else if (out.kind === "picture") {
    const p = o as SheetPicture
    if (p.crop) out.crop = { ...p.crop }
    if (p.natural) out.natural = { ...p.natural }
    if (p.line) out.line = copyLine(p.line)
    if (p.signature) out.signature = { ...p.signature }
    if (p.embed) out.embed = { ...p.embed }
    if (p.model3d) out.model3d = { ...p.model3d, rot: { ...p.model3d.rot } }
  } else {
    const g = o as SheetGroup
    out.child = { ...g.child }
    out.items = g.items.map(copyObject)
    if (g.smartArt) out.smartArt = { ...g.smartArt, nodes: copySmartArtNodes(g.smartArt.nodes) }
  }
  return out as T
}

export const copyObjects = (list: readonly SheetObject[]) => list.map(copyObject)

const copyColor = (c: ShapeColor): ShapeColor => ({ ...c })
const copyFill = (f: ShapeFill): ShapeFill => (f.kind === "solid" ? { kind: "solid", color: copyColor(f.color) } : f)
function copyLine(l: ShapeLine | undefined): ShapeLine | undefined {
  return l && { ...l, color: l.color && copyColor(l.color) }
}
export function copyText(t: ShapeText): ShapeText {
  return {
    ...t,
    insets: t.insets && { ...t.insets },
    paragraphs: t.paragraphs.map((p) => ({
      ...p,
      runs: p.runs.map((r) => ({ ...r, color: r.color && copyColor(r.color) })),
      ...(p.math ? { math: JSON.parse(JSON.stringify(p.math)) as MathList[] } : {}),
    })),
  }
}

/** The text of a shape as plain lines (for the Selection Pane, Find, accessibility and copying). */
export function plainText(text: ShapeText | undefined): string {
  if (!text) return ""
  // An equation as its linear text (Office's Linear format), a line each.
  return text.paragraphs
    .map((p) => (p.math ? p.math.map(toLinear).join("\n") : p.runs.map((r) => r.text).join("")))
    .join("\n")
}

/** An equation's text from its linear form: a line per equation line, in the equation's look. */
export function equationFromLinear(value: string, like?: ShapeText): ShapeText {
  const para = like?.paragraphs.find((p) => p.math)
  const look = para?.runs[0] ?? { text: "" }
  const lines = value.split("\n").filter((l) => l.trim())
  if (!lines.length) lines.push("")
  return {
    ...(like ? copyText(like) : {}),
    paragraphs: [
      { ...(para?.align ? { align: para.align } : {}), runs: [{ ...look, text: "" }], math: lines.map(parseLinear) },
    ],
  }
}

/** Insert › Equation: a shape whose text is an equation. */
export const isEquation = (o: SheetObject | undefined | null): o is SheetShape =>
  !!o && !isChart(o) && o.kind === "shape" && !!o.text?.paragraphs.some((p) => p.math)

/** Text as one paragraph per line, every run in the first run's format (typing into a shape). */
export function textFromPlain(value: string, like?: ShapeText): ShapeText {
  const first = like?.paragraphs[0]
  const run = first?.runs[0]
  const format = run ? { ...run, text: "" } : { text: "" }
  return {
    ...(like ? copyText(like) : {}),
    paragraphs: value.split("\n").map((line) => ({
      ...(first?.align ? { align: first.align } : {}),
      runs: line ? [{ ...format, text: line }] : [],
    })),
  }
}

/** The next free object id across the workbook (charts and shapes share them, as in a drawing's cNvPr ids). */
export function nextObjectId(sheets: readonly { objects: readonly SheetObject[] }[]): number {
  let max = 0
  for (const s of sheets) for (const o of allObjects(s.objects)) max = Math.max(max, o.id)
  return max + 1
}

/** Gives an object (and what's in it) new ids from `next` on; returns the next free one. */
export function renumber(o: SheetObject, next: number): number {
  o.id = next++
  if (!isChart(o) && o.kind === "group") for (const item of o.items) next = renumber(item, next)
  return next
}

/**
 * A copied sheet's objects: each gets an id of its own in the workbook (names stay, as in Excel), and its charts
 * read the copy's cells where they read the sheet copied.
 */
export function copiedSheetObjects(
  objects: readonly SheetObject[],
  sheets: readonly Sheet[],
  from: Sheet,
  copyName: string,
): SheetObject[] {
  let id = nextChartId(sheets)
  return objects.map((o) => {
    if (isChart(o)) {
      const [chart] = copiedSheetCharts([o], sheets, from, copyName)
      chart.id = id++
      return chart
    }
    const copy = copyObject(o)
    id = renumber(copy, id)
    return copy
  })
}

/** Excel's names for new objects: "Rectangle 3", "TextBox 1"… numbered by the id, as Excel does. */
export function objectName(
  o: { kind: DrawingObject["kind"]; geometry?: string; textBox?: boolean; icon?: boolean },
  id: number,
): string {
  if (o.kind === "picture") return `${o.icon ? "Graphic" : "Picture"} ${id}`
  if (o.kind === "group") return `Group ${id}`
  if (o.textBox) return `TextBox ${id}`
  return `${GEOMETRY_NAMES[o.geometry ?? "rect"] ?? "Freeform"} ${id}`
}

/** Excel's base names for shapes (as the Selection Pane shows them). */
const GEOMETRY_NAMES: Record<string, string> = {
  rect: "Rectangle",
  roundRect: "Rectangle: Rounded Corners",
  ellipse: "Oval",
  triangle: "Isosceles Triangle",
  rtTriangle: "Right Triangle",
  parallelogram: "Parallelogram",
  trapezoid: "Trapezoid",
  diamond: "Diamond",
  pentagon: "Pentagon",
  hexagon: "Hexagon",
  octagon: "Octagon",
  plus: "Cross",
  star5: "Star: 5 Points",
  star4: "Star: 4 Points",
  rightArrow: "Arrow: Right",
  leftArrow: "Arrow: Left",
  upArrow: "Arrow: Up",
  downArrow: "Arrow: Down",
  leftRightArrow: "Arrow: Left-Right",
  chevron: "Arrow: Chevron",
  homePlate: "Arrow: Pentagon",
  heart: "Heart",
  can: "Cylinder",
  donut: "Circle: Hollow",
  wedgeRectCallout: "Speech Bubble: Rectangle",
  wedgeRoundRectCallout: "Speech Bubble: Rectangle with Corners Rounded",
  wedgeEllipseCallout: "Speech Bubble: Oval",
  cloud: "Cloud",
  line: "Straight Connector",
  straightConnector1: "Straight Arrow Connector",
  flowChartProcess: "Flowchart: Process",
  flowChartDecision: "Flowchart: Decision",
  flowChartTerminator: "Flowchart: Terminator",
  smileyFace: "Smiley Face",
  lightningBolt: "Lightning Bolt",
  sun: "Sun",
  moon: "Moon",
  frame: "Frame",
  snip1Rect: "Rectangle: Single Corner Snipped",
}

/** Objects follow inserted and deleted rows and columns, as charts do (Move and size with cells). */
export function shiftObjects(objects: SheetObject[], axis: "rows" | "cols", at: number, count: number) {
  const key = axis === "rows" ? "row" : "col"
  const offset = axis === "rows" ? "y" : "x"
  for (const o of objects)
    for (const c of [o.anchor.from, o.anchor.to]) {
      if (count > 0) {
        if (c[key] >= at) c[key] += count
      } else if (c[key] >= at - count) c[key] += count
      else if (c[key] >= at) {
        c[key] = at
        c[offset] = 0
      }
    }
}

/** Every object, groups opened, with the group it's in (for finding one by id). */
export function* allObjects(list: readonly SheetObject[]): Generator<SheetObject> {
  for (const o of list) {
    yield o
    if (!isChart(o) && o.kind === "group") yield* allObjects(o.items)
  }
}

/** An Excel-made shape's look: Office's default style (Accent 1 fill, a darker outline, white text). */
export function defaultShape(geometry: string): Pick<SheetShape, "fill" | "line" | "text" | "connector"> {
  if (geometry === "line" || geometry.includes("Connector"))
    return {
      fill: { kind: "none" },
      line: {
        color: { theme: "accent1" },
        width: 0.75,
        ...(geometry === "straightConnector1" ? { tail: "triangle" as const } : {}),
      },
      connector: true,
    }
  return {
    fill: { kind: "solid", color: { theme: "accent1" } },
    line: { color: { theme: "accent1", lumMod: 0.5 }, width: 1 },
    text: { paragraphs: [{ align: "ctr", runs: [{ text: "", color: { theme: "lt1" } }] }], anchor: "ctr" },
  }
}

/** A new text box's look: white fill, a thin dark grey outline, black text at the top. */
export function defaultTextBox(): Pick<SheetShape, "fill" | "line" | "text"> {
  return {
    fill: { kind: "solid", color: { theme: "lt1" } },
    line: { color: { theme: "lt1", lumMod: 0.5 }, width: 0.75 },
    text: { paragraphs: [{ runs: [{ text: "", color: { theme: "dk1" } }] }], anchor: "t", wrap: true },
  }
}

// ── Pictures as data URLs ────────────────────────────────────────────────────────────────────────────────────

const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

/** A picture's bytes as a data URL. */
export function dataUrl(bytes: Uint8Array, mime: string): string {
  const out: string[] = []
  let chunk = ""
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i]
    const b = i + 1 < bytes.length ? bytes[i + 1] : 0
    const c = i + 2 < bytes.length ? bytes[i + 2] : 0
    chunk +=
      B64[a >> 2] +
      B64[((a & 3) << 4) | (b >> 4)] +
      (i + 1 < bytes.length ? B64[((b & 15) << 2) | (c >> 6)] : "=") +
      (i + 2 < bytes.length ? B64[c & 63] : "=")
    if (chunk.length > 8192) {
      out.push(chunk)
      chunk = ""
    }
  }
  out.push(chunk)
  return `data:${mime};base64,${out.join("")}`
}

/** A data URL's type and bytes (null when it isn't a base64 data URL). */
export function dataUrlBytes(url: string): { mime: string; bytes: Uint8Array } | null {
  const m = /^data:([^;,]+);base64,/.exec(url)
  if (!m) return null
  const text = url.slice(m[0].length).replace(/[^A-Za-z0-9+/]/g, "")
  const bytes = new Uint8Array(Math.floor((text.length * 3) / 4))
  let n = 0
  for (let i = 0; i < text.length; i += 4) {
    const v = [0, 1, 2, 3].map((k) => (i + k < text.length ? B64.indexOf(text[i + k]) : 0))
    const triple = (v[0] << 18) | (v[1] << 12) | (v[2] << 6) | v[3]
    bytes[n++] = (triple >> 16) & 255
    if (i + 2 < text.length) bytes[n++] = (triple >> 8) & 255
    if (i + 3 < text.length) bytes[n++] = triple & 255
  }
  return { mime: m[1], bytes: bytes.subarray(0, n) }
}

/** The file extension for a picture type. */
export const PICTURE_TYPES: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/svg+xml": "svg",
  "image/tiff": "tiff",
  "image/x-emf": "emf",
  "image/x-wmf": "wmf",
  "image/webp": "webp",
}
