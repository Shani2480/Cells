/**
 * Insert › SmartArt: the layouts Cells lays out itself, in the diagram layout language Office runs (ECMA-376
 * Part 1, 21.4), so Excel lays them out the same way. Written for Cells, checked by opening them in Excel.
 */

import type { SmartArt, SmartArtNode, SmartArtParts } from "./shapes"

const NS =
  'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

export type SmartArtCategory = "list" | "process" | "cycle" | "matrix"

export type SmartArtLayout = {
  id: string
  /** Its short name in the file's ids ("process"). */
  key: string
  name: string
  category: SmartArtCategory
  description: string
  /** How many points a new one starts with. */
  starts: number
  /** The layout definition (`dgm:layoutDef`). */
  def: string
}

const layoutId = (key: string) => `urn:cells.app/smartart/layout/${key}`

export const SMARTART_LAYOUTS: SmartArtLayout[] = [
  {
    id: layoutId("list"),
    key: "list",
    name: "Basic Block List",
    category: "list",
    description: "Shows blocks of information side by side, wrapping into rows.",
    starts: 5,
    def: `<dgm:layoutDef ${NS} uniqueId="${layoutId("list")}"><dgm:title val="Block List"/><dgm:desc val="Shows blocks of information side by side, wrapping into rows."/><dgm:catLst><dgm:cat type="list" pri="1000"/></dgm:catLst><dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst><dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/><dgm:constr type="w" for="ch" forName="space" refType="w" refFor="ch" refForName="node" fact="0.1"/><dgm:constr type="sp" refType="w" refFor="ch" refForName="space"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/></dgm:constrLst><dgm:ruleLst/><dgm:forEach name="nodes" axis="ch" ptType="node"><dgm:layoutNode name="node"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="rect" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="desOrSelf" ptType="node"/><dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode><dgm:forEach name="gaps" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="space"><dgm:alg type="sp"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:forEach></dgm:layoutNode></dgm:layoutDef>`,
  },
  {
    id: layoutId("process"),
    key: "process",
    name: "Basic Process",
    category: "process",
    description: "Shows the steps of a process, task or workflow in a row, an arrow from each to the next.",
    starts: 3,
    def: `<dgm:layoutDef ${NS} uniqueId="${layoutId("process")}"><dgm:title val="Basic Process"/><dgm:desc val="Shows the steps of a process, task or workflow in a row, an arrow from each to the next."/><dgm:catLst><dgm:cat type="process" pri="1000"/></dgm:catLst><dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst><dgm:alg type="lin"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" op="equ"/><dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" op="equ" fact="0.35"/><dgm:constr type="h" for="ch" forName="sibTrans" op="equ"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/><dgm:constr type="primFontSz" for="des" forName="connectorText" op="equ" val="55"/><dgm:constr type="primFontSz" for="des" forName="connectorText" refType="primFontSz" refFor="ch" refForName="node" op="lte" fact="0.8"/></dgm:constrLst><dgm:ruleLst/><dgm:forEach name="nodes" axis="ch" ptType="node"><dgm:layoutNode name="node"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="roundRect" r:blip=""><dgm:adjLst><dgm:adj idx="1" val="0.1"/></dgm:adjLst></dgm:shape><dgm:presOf axis="desOrSelf" ptType="node"/><dgm:constrLst><dgm:constr type="h" refType="w" fact="0.6"/><dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode><dgm:forEach name="arrows" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="sibTrans" styleLbl="sibTrans2D1"><dgm:alg type="conn"/><dgm:shape type="conn" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="h" refType="w" fact="0.62"/><dgm:constr type="connDist"/><dgm:constr type="begPad" refType="connDist" fact="0.25"/><dgm:constr type="endPad" refType="connDist" fact="0.22"/></dgm:constrLst><dgm:ruleLst/><dgm:forEach name="arrowText" axis="self" ptType="sibTrans"><dgm:layoutNode name="connectorText"><dgm:alg type="tx"><dgm:param type="autoTxRot" val="grav"/></dgm:alg><dgm:shape type="conn" r:blip="" hideGeom="1"><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="lMarg"/><dgm:constr type="rMarg"/><dgm:constr type="tMarg"/><dgm:constr type="bMarg"/></dgm:constrLst><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:layoutNode></dgm:forEach></dgm:forEach></dgm:layoutNode></dgm:layoutDef>`,
  },
  {
    id: layoutId("cycle"),
    key: "cycle",
    name: "Basic Cycle",
    category: "cycle",
    description: "Shows stages in a cycle that repeats, an arrow from each to the next.",
    starts: 5,
    def: `<dgm:layoutDef ${NS} uniqueId="${layoutId("cycle")}"><dgm:title val="Basic Cycle"/><dgm:desc val="Shows stages in a cycle that repeats, an arrow from each to the next."/><dgm:catLst><dgm:cat type="cycle" pri="1000"/></dgm:catLst><dgm:layoutNode name="cycle"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst><dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/></dgm:alg><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node"/><dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.4"/><dgm:constr type="h" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="sibTrans" fact="0.6"/><dgm:constr type="sibSp" refType="w" refFor="ch" refForName="node" fact="0.35"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/></dgm:constrLst><dgm:ruleLst/><dgm:forEach name="nodes" axis="ch" ptType="node"><dgm:layoutNode name="node"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="ellipse" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="desOrSelf" ptType="node"/><dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode><dgm:forEach name="arrows" axis="followSib" ptType="sibTrans" hideLastTrans="0" cnt="1"><dgm:layoutNode name="sibTrans" styleLbl="sibTrans2D1"><dgm:alg type="conn"><dgm:param type="begPts" val="auto"/><dgm:param type="endPts" val="auto"/></dgm:alg><dgm:shape type="conn" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="h" refType="w" fact="0.6"/><dgm:constr type="connDist"/><dgm:constr type="begPad" refType="connDist" fact="0.15"/><dgm:constr type="endPad" refType="connDist" fact="0.15"/></dgm:constrLst><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:forEach></dgm:layoutNode></dgm:layoutDef>`,
  },
  {
    id: layoutId("matrix"),
    key: "matrix",
    name: "Grid Matrix",
    category: "matrix",
    description: "Shows ideas in a grid, two to a row.",
    starts: 4,
    def: `<dgm:layoutDef ${NS} uniqueId="${layoutId("matrix")}"><dgm:title val="Grid Matrix"/><dgm:desc val="Shows ideas in a grid, two to a row."/><dgm:catLst><dgm:cat type="matrix" pri="1000"/></dgm:catLst><dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst><dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/><dgm:param type="bkpt" val="fixed"/><dgm:param type="bkPtFixedVal" val="2"/></dgm:alg><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="1"/><dgm:constr type="w" for="ch" forName="space" refType="w" refFor="ch" refForName="node" fact="0.05"/><dgm:constr type="sp" refType="w" refFor="ch" refForName="space"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/></dgm:constrLst><dgm:ruleLst/><dgm:forEach name="nodes" axis="ch" ptType="node"><dgm:layoutNode name="node"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="rect" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="desOrSelf" ptType="node"/><dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode><dgm:forEach name="gaps" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="space"><dgm:alg type="sp"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:forEach></dgm:layoutNode></dgm:layoutDef>`,
  },
]

export const smartArtLayout = (id: string) => SMARTART_LAYOUTS.find((l) => l.id === id)

/** Office's own layouts Cells lays out as one of its own (the nearest in look), by id or by category. */
const OFFICE_LAYOUTS: Record<string, string> = {
  default: "list",
  process1: "process",
  cycle2: "cycle",
  matrix2: "matrix",
}
const CATEGORY_LAYOUTS: Record<string, string> = {
  list: "list",
  process: "process",
  cycle: "cycle",
  matrix: "matrix",
}

/** The layout Cells draws a SmartArt with: its own, or the nearest to an Office layout it doesn't know. */
export function drawnLayout(id: string, category?: string): SmartArtLayout {
  const own = smartArtLayout(id)
  if (own) return own
  const office = /\/layout\/([^/]+)$/.exec(id)?.[1] ?? ""
  const key = OFFICE_LAYOUTS[office] ?? CATEGORY_LAYOUTS[category ?? ""] ?? "list"
  return SMARTART_LAYOUTS.find((l) => l.key === key)!
}

/** One of SmartArt Design › Change Colors: the fill of each shape, from the theme. */
export type SmartArtColors = {
  id: string
  name: string
  /** The theme colours the shapes take in turn. */
  accents: string[]
  def: string
}

const label = (name: string, fill: string, line: string, text: string) =>
  `<dgm:styleLbl name="${name}"><dgm:fillClrLst meth="repeat">${fill}</dgm:fillClrLst><dgm:linClrLst meth="repeat">${line}</dgm:linClrLst><dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst meth="repeat">${text}</dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl>`
const clr = (val: string, mod = "") => `<a:schemeClr val="${val}">${mod}</a:schemeClr>`

function colorsDef(key: string, name: string, accents: string[]): SmartArtColors {
  const id = `urn:cells.app/smartart/colors/${key}`
  const each = (mod = "") => accents.map((a) => clr(a, mod)).join("")
  const def =
    `<dgm:colorsDef ${NS} uniqueId="${id}"><dgm:title val="${name}"/><dgm:desc val=""/>` +
    `<dgm:catLst><dgm:cat type="${accents.length > 1 ? "colorful" : accents[0]}" pri="10100"/></dgm:catLst>` +
    label("node1", each(), clr("lt1"), clr("lt1")) +
    label("sibTrans2D1", each(`<a:tint val="60000"/>`), each(`<a:tint val="60000"/>`), clr("lt1")) +
    `</dgm:colorsDef>`
  return { id, name, accents, def }
}

/** Change Colors: one colour of the theme for every shape, or the accents in turn. */
export const SMARTART_COLORS: SmartArtColors[] = [
  colorsDef("colorful", "Colorful - Accent Colors", ["accent2", "accent3", "accent4", "accent5", "accent6"]),
  ...[1, 2, 3, 4, 5, 6].map((n) => colorsDef(`accent${n}`, `Colored Fill - Accent ${n}`, [`accent${n}`])),
]

export const DEFAULT_SMARTART_COLORS = "urn:cells.app/smartart/colors/accent1"

/** The colours a SmartArt is drawn in: Cells' own by id, else Office's by the accent in its id ("accent3_2"). */
export function smartArtColors(id: string | undefined): SmartArtColors {
  const own = SMARTART_COLORS.find((c) => c.id === id)
  if (own) return own
  const office = /\/colors\/(accent(\d)|colorful)/.exec(id ?? "")
  if (office?.[2]) return SMARTART_COLORS.find((c) => c.accents[0] === `accent${office[2]}` && c.accents.length === 1)!
  if (office?.[1] === "colorful") return SMARTART_COLORS[0]!
  return SMARTART_COLORS.find((c) => c.id === DEFAULT_SMARTART_COLORS)!
}

/** Office's simple style (no effects): named only, Office fills it in. */
const QUICK_STYLE = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"

const SCENE = `<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>`
const BLACK = `<a:scrgbClr r="0" g="0" b="0"/>`
/** A style label of Office's Simple Fill: the theme's outline (`line` index), fill and text in the light colour. */
const styleLbl = (name: string, line: number) =>
  `<dgm:styleLbl name="${name}">${SCENE}<dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="${line}">${BLACK}</a:lnRef><a:fillRef idx="1">${BLACK}</a:fillRef><a:effectRef idx="0">${BLACK}</a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl>`
/** Office's Simple Fill for the labels Cells' layouts use (a file's style is used as written once it has labels). */
const QUICK_STYLE_DEF = `<dgm:styleDef ${NS} uniqueId="${QUICK_STYLE}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>${SCENE}${styleLbl("node1", 2)}${styleLbl("sibTrans2D1", 0)}</dgm:styleDef>`

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")

/**
 * The data part: a document point, a point for each piece of text (empty ones are Office's "[Text]" places to type)
 * and the connections that make the tree, each with the transitions layouts draw arrows for.
 */
export function smartArtDataXml(
  sa: Pick<SmartArt, "layout" | "nodes" | "colors" | "quickStyle">,
  category = "",
): string {
  const pts: string[] = []
  const cxns: string[] = []
  let next = 1
  const t = (text: string) =>
    `<dgm:t><a:bodyPr/><a:lstStyle/>${(text ? text.split("\n") : [""])
      .map((line) =>
        line
          ? `<a:p><a:r><a:rPr lang="en-US"/><a:t>${esc(line)}</a:t></a:r></a:p>`
          : `<a:p><a:endParaRPr lang="en-US"/></a:p>`,
      )
      .join("")}</dgm:t>`
  pts.push(
    `<dgm:pt modelId="0" type="doc"><dgm:prSet loTypeId="${esc(sa.layout)}" loCatId="${category}" qsTypeId="${esc(sa.quickStyle ?? QUICK_STYLE)}" qsCatId="simple" csTypeId="${esc(sa.colors ?? DEFAULT_SMARTART_COLORS)}" csCatId="accent1" phldr="1"/><dgm:spPr/>${t("")}</dgm:pt>`,
  )
  const add = (parent: number, nodes: SmartArtNode[]) =>
    nodes.forEach((n, i) => {
      const id = next++
      const par = next++
      const sib = next++
      const cxn = next++
      const pr = n.text ? `<dgm:prSet phldrT="[Text]"/>` : `<dgm:prSet phldrT="[Text]" phldr="1"/>`
      pts.push(`<dgm:pt modelId="${id}">${pr}<dgm:spPr/>${t(n.text)}</dgm:pt>`)
      pts.push(`<dgm:pt modelId="${par}" type="parTrans" cxnId="${cxn}"><dgm:prSet/><dgm:spPr/>${t("")}</dgm:pt>`)
      pts.push(`<dgm:pt modelId="${sib}" type="sibTrans" cxnId="${cxn}"><dgm:prSet/><dgm:spPr/>${t("")}</dgm:pt>`)
      cxns.push(
        `<dgm:cxn modelId="${cxn}" srcId="${parent}" destId="${id}" srcOrd="${i}" destOrd="0" parTransId="${par}" sibTransId="${sib}"/>`,
      )
      add(id, n.children)
    })
  add(0, sa.nodes)
  return `<dgm:dataModel ${NS}><dgm:ptLst>${pts.join("")}</dgm:ptLst><dgm:cxnLst>${cxns.join("")}</dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>`
}

/**
 * A SmartArt's parts written by Cells: its data from the text, the layout and colours definitions (Cells' own, or
 * those the file had for a layout Cells doesn't know), and Office's simple style by name. No drawing: Office lays
 * it out when it opens the file.
 */
export function smartArtParts(sa: SmartArt): SmartArtParts {
  const own = smartArtLayout(sa.layout)
  const colors = SMARTART_COLORS.find((c) => c.id === sa.colors)
  return {
    data: smartArtDataXml(sa, own?.category ?? ""),
    layout: own ? own.def : (sa.parts?.layout ?? drawnLayout(sa.layout).def),
    quickStyle: sa.parts?.quickStyle || QUICK_STYLE_DEF,
    colors: colors ? colors.def : (sa.parts?.colors ?? smartArtColors(sa.colors).def),
  }
}

// ── Editing the text tree (SmartArt Design › Create Graphic, and the Text Pane) ──

/** A point by its place in the tree: its index at each level. */
export type SmartArtPath = number[]

const clone = (nodes: SmartArtNode[]): SmartArtNode[] =>
  nodes.map((n) => ({ text: n.text, children: clone(n.children) }))

/** The list a point is in (`nodes` itself or a point's children), and its index there. */
function siblings(nodes: SmartArtNode[], path: SmartArtPath): { list: SmartArtNode[]; i: number } | null {
  let list = nodes
  for (let d = 0; d < path.length - 1; d++) {
    const n = list[path[d]!]
    if (!n) return null
    list = n.children
  }
  const i = path[path.length - 1]!
  return path.length && i >= 0 && i <= list.length ? { list, i } : null
}

export type SmartArtEdit = { nodes: SmartArtNode[]; path: SmartArtPath } | null

/** Add Shape After / Before: an empty point beside this one (at the top level when nothing is chosen). */
export function addSmartArtNode(nodes: SmartArtNode[], path: SmartArtPath | null, before = false): SmartArtEdit {
  const next = clone(nodes)
  const at = path?.length ? path : [Math.max(0, nodes.length - 1)]
  const s = siblings(next, at)
  if (!s) return null
  const i = !nodes.length ? 0 : before ? s.i : s.i + 1
  s.list.splice(i, 0, { text: "", children: [] })
  return { nodes: next, path: [...at.slice(0, -1), i] }
}

/** Removes a point; the points under it move up into its place. */
export function removeSmartArtNode(nodes: SmartArtNode[], path: SmartArtPath): SmartArtEdit {
  const next = clone(nodes)
  const s = siblings(next, path)
  const node = s?.list[s.i]
  if (!s || !node) return null
  s.list.splice(s.i, 1, ...node.children)
  const prev = s.i > 0 ? [...path.slice(0, -1), s.i - 1] : path.length > 1 ? path.slice(0, -1) : []
  return { nodes: next, path: prev }
}

/** Demote: the point goes under the one before it, as its last. */
export function demoteSmartArtNode(nodes: SmartArtNode[], path: SmartArtPath): SmartArtEdit {
  const next = clone(nodes)
  const s = siblings(next, path)
  const node = s?.list[s.i]
  if (!s || !node || s.i === 0) return null
  const parent = s.list[s.i - 1]!
  s.list.splice(s.i, 1)
  parent.children.push(node)
  return { nodes: next, path: [...path.slice(0, -1), s.i - 1, parent.children.length - 1] }
}

/** Promote: the point moves up a level, after its parent; the points after it at its level go under it. */
export function promoteSmartArtNode(nodes: SmartArtNode[], path: SmartArtPath): SmartArtEdit {
  if (path.length < 2) return null
  const next = clone(nodes)
  const s = siblings(next, path)
  const up = siblings(next, path.slice(0, -1))
  const node = s?.list[s.i]
  if (!s || !up || !node) return null
  const after = s.list.splice(s.i)
  node.children.push(...after.slice(1))
  up.list.splice(up.i + 1, 0, node)
  return { nodes: next, path: [...path.slice(0, -2), up.i + 1] }
}

/** Move Up / Move Down: the point (with the points under it) swaps places with the one before or after it. */
export function moveSmartArtNode(nodes: SmartArtNode[], path: SmartArtPath, by: -1 | 1): SmartArtEdit {
  const next = clone(nodes)
  const s = siblings(next, path)
  const j = (s?.i ?? 0) + by
  if (!s || !s.list[s.i] || j < 0 || j >= s.list.length) return null
  ;[s.list[s.i], s.list[j]] = [s.list[j]!, s.list[s.i]!]
  return { nodes: next, path: [...path.slice(0, -1), j] }
}

/** Sets one point's text. */
export function setSmartArtText(nodes: SmartArtNode[], path: SmartArtPath, text: string): SmartArtNode[] {
  const next = clone(nodes)
  const s = siblings(next, path)
  const node = s?.list[s.i]
  if (node) node.text = text
  return next
}

/** Every point with its path and depth, in reading order (the Text Pane's lines). */
export function smartArtLines(nodes: SmartArtNode[]): { node: SmartArtNode; path: SmartArtPath }[] {
  const out: { node: SmartArtNode; path: SmartArtPath }[] = []
  const walk = (list: SmartArtNode[], at: SmartArtPath) =>
    list.forEach((node, i) => {
      out.push({ node, path: [...at, i] })
      walk(node.children, [...at, i])
    })
  walk(nodes, [])
  return out
}
