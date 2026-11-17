import { renderChart } from "@cells/chart"
import { isChart, plainText, type DrawingObject, type SheetChart, type SheetObject, type Theme } from "@cells/core"
import { memo, useEffect, useMemo, useRef, useState } from "react"
import { canvasMeasure, objectAt, objectDocument } from "@/drawing/render"
import { chartBox, type ChartBox } from "@/sheet/charts"
import type { SheetController } from "@/sheet/controller"
import { pngKey, preparePng, readyPng } from "./chart-png"
import { cellAt, colX, rowY, type GridLayout } from "./layout"

/** Excel's handles: the four corners and the middle of each side. */
const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const
type Handle = (typeof HANDLES)[number]

/** A drag starts after this many pixels, so a click only selects. */
const DRAG_START = 4
/** The smallest object a drag makes (pixels at 100%). */
const MIN_SIZE = 8

/** One object's picture: SVG markup at its size in sheet pixels, scaled by the view. */
const Picture = memo(function Picture({
  svg,
  w,
  h,
  className,
}: {
  svg: string
  w: number
  h: number
  className: string
}) {
  return (
    <svg
      className={className}
      width="100%"
      height="100%"
      viewBox={`0 0 ${w} ${h}`}
      preserveAspectRatio="none"
      overflow="visible"
      aria-hidden
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  )
})

function resized(box: ChartBox, handle: Handle, dx: number, dy: number, keepRatio: boolean): ChartBox {
  let { x, y, w, h } = box
  if (handle.includes("w")) {
    x += dx
    w -= dx
  }
  if (handle.includes("e")) w += dx
  if (handle.includes("n")) {
    y += dy
    h -= dy
  }
  if (handle.includes("s")) h += dy
  // Shift on a corner keeps the proportions, as in Excel.
  if (keepRatio && handle.length === 2) {
    const ratio = box.w / box.h
    if (w / h > ratio) w = h * ratio
    else h = w / ratio
    if (handle.includes("w")) x = box.x + box.w - w
    if (handle.includes("n")) y = box.y + box.h - h
  }
  if (w < MIN_SIZE) {
    if (handle.includes("w")) x -= MIN_SIZE - w
    w = MIN_SIZE
  }
  if (h < MIN_SIZE) {
    if (handle.includes("n")) y -= MIN_SIZE - h
    h = MIN_SIZE
  }
  return { x: Math.max(0, x), y: Math.max(0, y), w, h }
}

type Gesture = {
  ids: number[]
  handle: Handle | "move" | "rotate"
  x: number
  y: number
  boxes: Map<number, ChartBox>
  /** The middle of the object turned, on screen (rotate). */
  center?: { x: number; y: number }
  rotation?: number
  moved: boolean
}

/**
 * The sheet's drawing over the cells: charts, shapes, text boxes, pictures and groups in their order. A click picks
 * one (Ctrl or Shift: one more), a drag moves the selected ones, the handles size (Shift keeps the proportions) and
 * the round handle turns (Shift: by 15°). A double-click types into a shape. Insert › Shapes waits here for a drag.
 */
export function ObjectLayer({
  controller,
  layout,
  zoom,
  onDone,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  onDone: () => void
}) {
  const sheet = controller.sheet
  const theme = controller.wb.theme
  const version = controller.version
  const objects = controller.objects
  const [drag, setDragState] = useState<{ boxes: Map<number, ChartBox>; rotation?: number } | null>(null)
  const [draw, setDrawState] = useState<ChartBox | null>(null)
  // The latest drag and drawn box, for the pointer-up that may come before React draws them.
  const dragNow = useRef<typeof drag>(null)
  const drawNow = useRef<ChartBox | null>(null)
  const setDrag = (d: typeof drag) => {
    dragNow.current = d
    setDragState(d)
  }
  const setDraw = (d: ChartBox | null) => {
    drawNow.current = d
    setDrawState(d)
  }
  const gesture = useRef<Gesture | null>(null)
  const drawStart = useRef<{ x: number; y: number } | null>(null)
  const layer = useRef<HTMLDivElement>(null)

  // Each object's picture, drawn again when the workbook changes (a chart's data may have).
  const pictures = useMemo(() => {
    const out = new Map<number, { svg: string; box: ChartBox }>()
    for (const o of sheet.objects) {
      if (!isChart(o) && o.hidden) continue
      const box = chartBox(sheet, o.anchor)
      out.set(o.id, { svg: picture(controller, o, box, theme), box })
    }
    return out
  }, [sheet, version, controller, theme])

  // The selected chart's PNG, ready for Copy (which can't wait for one to be drawn).
  const selectedChart = controller.charts.selected
  const chartPicture = selectedChart === null ? null : pictures.get(selectedChart)
  useEffect(() => {
    controller.charts.picture = (chart, _spec, size) => {
      const svg = chart.id === selectedChart && chartPicture ? chartPicture.svg : null
      return svg ? readyPng(chart.id, pngKey(svg, size.w, size.h)) : null
    }
    if (selectedChart !== null && chartPicture)
      void preparePng(selectedChart, chartPicture.svg, chartPicture.box.w, chartPicture.box.h)
  }, [controller, selectedChart, chartPicture])

  // Copying shapes gives other apps an SVG picture of them.
  useEffect(() => {
    controller.objectPicture = (list) => {
      const boxes = list.map((o) => chartBox(sheet, o.anchor))
      const x = Math.min(...boxes.map((b) => b.x))
      const y = Math.min(...boxes.map((b) => b.y))
      const w = Math.max(...boxes.map((b) => b.x + b.w)) - x
      const h = Math.max(...boxes.map((b) => b.y + b.h)) - y
      const body = list
        .map((o, i) => objectAt(o, { ...boxes[i], x: boxes[i].x - x, y: boxes[i].y - y }, theme, canvasMeasure))
        .join("")
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${body}</svg>`
      return `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`
    }
  }, [controller, sheet, theme])

  const left = layout.rowHeaderWidth
  const top = layout.headerHeight
  // Placed from the cell it starts in, so an object in frozen rows or columns stays put as the rest scrolls.
  const toScreen = (o: SheetObject, box: ChartBox) => {
    const { from } = o.anchor
    return {
      left: (colX(sheet, layout, from.col) + box.x - sheet.cols.offset(from.col) - left) * zoom,
      top: (rowY(sheet, layout, from.row) + box.y - sheet.rows.offset(from.row) - top) * zoom,
      width: box.w * zoom,
      height: box.h * zoom,
    }
  }
  /** A point in the layer (screen pixels) as sheet pixels at 100%. */
  const sheetPoint = (event: { clientX: number; clientY: number }) => {
    const r = layer.current!.getBoundingClientRect()
    const cx = left + (event.clientX - r.left) / zoom
    const cy = top + (event.clientY - r.top) / zoom
    const { row, col } = cellAt(controller, layout, cx, cy)
    return {
      x: sheet.cols.offset(col) + (cx - colX(sheet, layout, col)),
      y: sheet.rows.offset(row) + (cy - rowY(sheet, layout, row)),
    }
  }

  const begin = (event: React.PointerEvent, o: SheetObject, handle: Gesture["handle"]) => {
    if (event.button !== 0 || objects.editing === o.id) return
    event.stopPropagation()
    event.preventDefault()
    const add = event.ctrlKey || event.metaKey || event.shiftKey
    if (handle === "move" && add) {
      objects.select(o.id, true)
      return
    }
    if (!objects.selected.includes(o.id)) objects.select(o.id)
    const ids = handle === "move" ? [...new Set([...objects.selected, o.id])] : [o.id]
    const boxes = new Map<number, ChartBox>()
    for (const id of ids) {
      const p = pictures.get(id)
      if (p) boxes.set(id, p.box)
    }
    let center: Gesture["center"]
    if (handle === "rotate") {
      const frame = (event.currentTarget as HTMLElement).closest(".object-frame")!.getBoundingClientRect()
      center = { x: frame.left + frame.width / 2, y: frame.top + frame.height / 2 }
    }
    gesture.current = {
      ids,
      handle,
      x: event.clientX,
      y: event.clientY,
      boxes,
      center,
      rotation: isChart(o) ? 0 : (o.rotation ?? 0),
      moved: false,
    }
    capture(event)
  }
  const move = (event: React.PointerEvent) => {
    const g = gesture.current
    if (!g) return
    const dx = (event.clientX - g.x) / zoom
    const dy = (event.clientY - g.y) / zoom
    if (!g.moved && Math.hypot(dx, dy) < DRAG_START / zoom) return
    g.moved = true
    if (g.handle === "rotate" && g.center) {
      let angle = (Math.atan2(event.clientY - g.center.y, event.clientX - g.center.x) * 180) / Math.PI + 90
      if (event.shiftKey) angle = Math.round(angle / 15) * 15
      setDrag({ boxes: g.boxes, rotation: (angle + 360) % 360 })
      return
    }
    const boxes = new Map<number, ChartBox>()
    for (const [id, box] of g.boxes) {
      boxes.set(
        id,
        g.handle === "move"
          ? { ...box, x: Math.max(0, box.x + dx), y: Math.max(0, box.y + dy) }
          : resized(box, g.handle as Handle, ...turned(dx, dy, g.rotation ?? 0), event.shiftKey),
      )
    }
    setDrag({ boxes })
  }
  const end = () => {
    const g = gesture.current
    const last = dragNow.current
    gesture.current = null
    if (g?.moved && last) {
      if (g.handle === "rotate" && last.rotation !== undefined) objects.rotate(last.rotation, true)
      else objects.place(last.boxes, g.handle === "move" ? "Move Object" : "Size")
    }
    setDrag(null)
    if (objects.editing === null) onDone()
  }

  const pending = objects.pending
  if (!sheet.objects.length && !pending) return null
  return (
    <div className="object-layer" ref={layer} style={{ left: left * zoom, top: top * zoom }}>
      {sheet.objects.map((o) => {
        const p = pictures.get(o.id)
        if (!p) return null
        const selected = objects.selected.includes(o.id)
        const single = selected && objects.selected.length === 1
        const dragging = drag?.boxes.get(o.id) ?? null
        const rotation = isChart(o) ? 0 : drag?.rotation !== undefined && dragging ? drag.rotation : (o.rotation ?? 0)
        const flips = !isChart(o) && (o.flipH || o.flipV) ? ` scale(${o.flipH ? -1 : 1}, ${o.flipV ? -1 : 1})` : ""
        const editing = objects.editing === o.id && !isChart(o) && o.kind === "shape"
        return (
          <div
            key={o.id}
            className={`object-frame${isChart(o) ? " is-chart chart-frame" : ""}${selected ? " is-selected" : ""}`}
            style={{ ...toScreen(o, p.box), transform: rotation ? `rotate(${rotation}deg)` : undefined }}
            role="img"
            aria-label={o.name}
            data-object={o.id}
            data-chart={isChart(o) ? o.id : undefined}
            onPointerDown={(event) => begin(event, o, "move")}
            onPointerMove={move}
            onPointerUp={end}
            onPointerCancel={() => {
              gesture.current = null
              setDrag(null)
            }}
            onDoubleClick={(event) => {
              if (isChart(o)) return
              event.stopPropagation()
              // A signature line opens its setup (signing needs a digital ID, which Cells doesn't have).
              if (o.kind === "picture" && o.signature) controller.shell.openDialog?.(`signatureSetup:${o.id}`)
              else if (o.kind === "picture" && o.embed) controller.shell.openDialog?.(`openObject:${o.id}`)
              else objects.startText(o.id)
            }}
          >
            <div className="object-picture" style={flips ? { transform: flips.trim() } : undefined}>
              <Picture svg={p.svg} w={p.box.w} h={p.box.h} className={isChart(o) ? "chart-svg" : "object-svg"} />
            </div>
            {editing ? (
              <TextEditor
                id={o.id}
                objects={objects}
                value={plainText(o.text)}
                onDone={(value) => {
                  objects.setText(o.id, value)
                  onDone()
                }}
              />
            ) : null}
            {single && !editing
              ? HANDLES.map((h) => (
                  <span
                    key={h}
                    className={`object-handle chart-handle is-${h}`}
                    onPointerDown={(event) => begin(event, o, h)}
                    onPointerMove={move}
                    onPointerUp={end}
                  />
                ))
              : null}
            {single && !editing && !isChart(o) ? (
              <span
                className="object-rotate"
                title="Rotate"
                onPointerDown={(event) => begin(event, o, "rotate")}
                onPointerMove={move}
                onPointerUp={end}
              />
            ) : null}
            {dragging && drag?.rotation === undefined ? (
              <div
                className="chart-ghost"
                style={{
                  left: (dragging.x - p.box.x) * zoom,
                  top: (dragging.y - p.box.y) * zoom,
                  width: dragging.w * zoom,
                  height: dragging.h * zoom,
                }}
              />
            ) : null}
          </div>
        )
      })}
      {pending ? (
        <div
          className="object-draw"
          onPointerDown={(event) => {
            if (event.button !== 0) return
            event.preventDefault()
            event.stopPropagation()
            const at = sheetPoint(event)
            drawStart.current = at
            setDraw({ ...at, w: 0, h: 0 })
            capture(event)
          }}
          onPointerMove={(event) => {
            const s = drawStart.current
            if (!s) return
            const at = sheetPoint(event)
            let w = at.x - s.x
            let h = at.y - s.y
            // Shift keeps a shape square (a line at 45° steps).
            if (event.shiftKey) {
              const side = Math.max(Math.abs(w), Math.abs(h))
              w = Math.sign(w || 1) * side
              h = Math.sign(h || 1) * side
            }
            setDraw({ x: Math.min(s.x, s.x + w), y: Math.min(s.y, s.y + h), w: Math.abs(w), h: Math.abs(h) })
          }}
          onPointerUp={() => {
            const box = drawNow.current
            drawStart.current = null
            setDraw(null)
            if (box) objects.drawPending(box)
            if (objects.editing === null) onDone()
          }}
        >
          {draw ? (
            <div
              className="chart-ghost"
              style={{
                left: (colX(sheet, layout, 0) - left - sheet.cols.offset(0) + draw.x) * zoom,
                top: (rowY(sheet, layout, 0) - top - sheet.rows.offset(0) + draw.y) * zoom,
                width: draw.w * zoom,
                height: draw.h * zoom,
              }}
            />
          ) : null}
        </div>
      ) : null}
    </div>
  )
}

/** A drag in screen directions, turned into the object's own (a turned shape sizes along its sides). */
function turned(dx: number, dy: number, degrees: number): [number, number] {
  if (!degrees) return [dx, dy]
  const r = (-degrees * Math.PI) / 180
  return [dx * Math.cos(r) - dy * Math.sin(r), dx * Math.sin(r) + dy * Math.cos(r)]
}

/**
 * Typing into a shape: its text, kept when focus leaves or with Escape (Excel keeps what was typed). An equation is
 * typed in its linear form; the Equation tab types at the caret while the editor is open.
 */
function TextEditor({
  id,
  objects,
  value,
  onDone,
}: {
  id: number
  objects: SheetController["objects"]
  value: string
  onDone: (value: string) => void
}) {
  const [text, setText] = useState(value)
  const done = useRef(false)
  const area = useRef<HTMLTextAreaElement>(null)
  const finish = (v: string) => {
    if (done.current) return
    done.current = true
    const el = area.current
    objects.lastCaret = el ? { id, fromEnd: el.value.length - el.selectionEnd } : null
    objects.caret = null
    onDone(v)
  }
  const latest = useRef(finish)
  latest.current = finish
  useEffect(() => {
    objects.textInserter = (insert) => {
      const el = area.current
      if (!el) return
      el.focus()
      el.setRangeText(insert, el.selectionStart, el.selectionEnd, "end")
      setText(el.value)
    }
    objects.finishText = () => latest.current(area.current?.value ?? "")
    // All the text selected, or the caret back where it was after the Equation tab put something in.
    const el = area.current
    if (el) {
      el.focus()
      // (Applied again on a second run of this effect: React runs effects twice while developing.)
      const caret = objects.caret
      if (caret?.id === id) {
        const at = Math.max(0, el.value.length - caret.fromEnd)
        el.setSelectionRange(at, at)
      } else el.select()
    }
    return () => {
      objects.textInserter = null
      objects.finishText = null
    }
  }, [objects, id])
  return (
    <textarea
      ref={area}
      className="object-text-editor"
      aria-label="Shape text"
      value={text}
      onPointerDown={(e) => e.stopPropagation()}
      onChange={(e) => setText(e.target.value)}
      onBlur={() => finish(text)}
      onKeyDown={(e) => {
        e.stopPropagation()
        if (e.key === "Escape") {
          e.preventDefault()
          finish(text)
        }
      }}
    />
  )
}

function picture(controller: SheetController, o: SheetObject, box: ChartBox, theme: Theme): string {
  try {
    if (isChart(o)) return renderChart(controller.charts.liveSpec(o as SheetChart), box.w, box.h)
    return objectDocument(o as DrawingObject, box.w, box.h, theme, canvasMeasure)
  } catch {
    // An object the renderer can't draw shows as an empty frame rather than taking the sheet down.
    return ""
  }
}

/** Keeps a drag's pointer events coming here (a pointer that's already gone can't be captured: the drag still works). */
function capture(event: React.PointerEvent) {
  try {
    ;(event.currentTarget as HTMLElement).setPointerCapture(event.pointerId)
  } catch {
    // No such pointer (a cancelled touch): the events still arrive while the pointer is over the layer.
  }
}
