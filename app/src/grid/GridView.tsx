import { linkScreenTip, MAX_ROWS, rangeOf, type CellAddress, type Range } from "@cells/core"
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from "react"
import { usePlatform } from "@/platform/context"
import { MAX_ZOOM, MIN_ZOOM, type SheetController } from "@/sheet/controller"
import { isSystemShortcut, keyInputFrom } from "@/sheet/keys"
import { AnnotationLayer } from "./AnnotationLayer"
import { ChartLayer } from "./ChartLayer"
import { TraceLayer } from "./TraceLayer"
import { OutlineBar, outlineBarSize } from "./OutlineBar"
import { FilterButtons } from "./FilterButtons"
import { ValidationLayer } from "./ValidationLayer"
import { autofitColumnWidth, columnWidthLabel, fittedRowHeight, lineHeightOf, rowHeightLabel } from "./autofit"
import { drawGrid } from "./draw"
import { lookOf, shownText } from "./cell-look"
import { PAD } from "./draw-cells"
import { FormulaAssist, FormulaHighlight } from "./FormulaAssist"
import { cellAt, colX, computeLayout, hitTest, rangeRect, rowY, type GridLayout, type Hit } from "./layout"
import { CELL_FONT, LIGHT_THEME, readGridTheme, type GridTheme } from "./theme"
import { Scrollbar } from "./Scrollbar"
import {
  distance,
  FlingTracker,
  flingStep,
  keyboardOverlap,
  LONG_PRESS_MS,
  midpoint,
  pinch,
  TAP_SLOP,
  TapTracker,
  type Point,
  type Velocity,
} from "./touch"

const SCROLLBAR = 14

type Drag =
  | { kind: "cells" }
  | { kind: "rows"; anchor: number }
  | { kind: "cols"; anchor: number }
  | { kind: "fill"; source: Range }
  /** Point mode: dragging over cells writes a range reference into the formula. */
  | { kind: "point"; anchor: CellAddress }
  | { kind: "resize-col"; col: number; startX: number; startSize: number; size: number }
  | { kind: "resize-row"; row: number; startY: number; startSize: number; size: number }
  | {
      kind: "pan"
      startX: number
      startY: number
      scrollX: number
      scrollY: number
      moved: boolean
      row: number
      col: number
      /** Where the finger went down (client px), for the long-press menu. */
      clientX: number
      clientY: number
      /** The long press opened the menu: lifting the finger is not a tap. */
      pressed: boolean
    }
  /** Two fingers: zoom follows their distance. */
  | { kind: "pinch"; startDistance: number; startZoom: number }

export type GridContextMenu = { x: number; y: number; target: "cells" | "rows" | "cols" }

/** How long the pointer rests on a link before its ScreenTip shows. */
const LINK_TIP_MS = 600

export function GridView({
  controller,
  onContextMenu,
}: {
  controller: SheetController
  onContextMenu: (menu: GridContextMenu) => void
}) {
  const version = useSyncExternalStore(controller.subscribe, controller.getVersion)
  // Scrolling has its own notice: only the grid and its layers move, the rest of the window doesn't re-render.
  const scrollVersion = useSyncExternalStore(controller.subscribeScroll, controller.getScrollVersion)
  const platform = usePlatform()
  const hostRef = useRef<HTMLDivElement>(null)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const layoutRef = useRef<GridLayout | null>(null)
  const dragRef = useRef<Drag | null>(null)
  const pointerRef = useRef<{ x: number; y: number; ctrl: boolean } | null>(null)
  // Fingers on the canvas (CSS px from its top-left), the long-press timer and the last tap (double tap edits).
  const touchesRef = useRef(new Map<number, Point>())
  const pressTimer = useRef<number | undefined>(undefined)
  // A press on a linked cell: a quick release follows the link; holding it (Excel's "click and hold") only selects.
  const linkPress = useRef<{ row: number; col: number; time: number } | null>(null)
  const linkTipTimer = useRef<number | undefined>(undefined)
  const tapsRef = useRef(new TapTracker())
  const lastTouchRef = useRef(0)
  // A pan's speed, and the glide after a quick one (stopped by the next touch).
  const flingRef = useRef(new FlingTracker())
  const glideRef = useRef<number | undefined>(undefined)
  const stopGlide = () => {
    if (glideRef.current !== undefined) cancelAnimationFrame(glideRef.current)
    glideRef.current = undefined
  }
  const glide = (start: Velocity) => {
    stopGlide()
    let v: Velocity | null = start
    let last = performance.now()
    const frame = (now: number) => {
      if (!v) return
      const step = flingStep(v, Math.min(now - last, 50))
      last = now
      const before = { x: controller.scrollX, y: controller.scrollY }
      controller.scrollTo(before.x + step.dx, before.y + step.dy)
      // At an edge the glide ends there.
      const stuck = controller.scrollX === before.x && controller.scrollY === before.y
      v = stuck ? null : step.next
      glideRef.current = v ? requestAnimationFrame(frame) : undefined
    }
    glideRef.current = requestAnimationFrame(frame)
  }
  useEffect(() => stopGlide, [])
  const themeRef = useRef<GridTheme>(LIGHT_THEME)
  const measureCache = useRef(new Map<string, number>())
  const [size, setSize] = useState({ width: 0, height: 0 })
  const [cursor, setCursor] = useState("cell")
  const [resizeTip, setResizeTip] = useState<{ x: number; y: number; text: string } | null>(null)
  // A link's ScreenTip, shown when the pointer rests on its cell.
  const [linkTip, setLinkTip] = useState<{ x: number; y: number; row: number; col: number; text: string } | null>(null)
  // The note or comment under the pointer; leaving it hides the box after a moment, so the pointer can move into
  // a comment card.
  const [hover, setHover] = useState<{ row: number; col: number } | null>(null)
  const hoverTimer = useRef<number | undefined>(undefined)
  const hoverAt = useCallback((cell: { row: number; col: number } | null) => {
    window.clearTimeout(hoverTimer.current)
    if (cell) {
      setHover((h) => (h && h.row === cell.row && h.col === cell.col ? h : cell))
      return
    }
    hoverTimer.current = window.setTimeout(() => setHover(null), 300)
  }, [])
  const keepHover = useCallback(() => window.clearTimeout(hoverTimer.current), [])
  const hideLinkTip = () => {
    window.clearTimeout(linkTipTimer.current)
    setLinkTip((tip) => (tip ? null : tip))
  }
  const [antsOffset, setAntsOffset] = useState(0)
  const [themeVersion, setThemeVersion] = useState(0)
  const composingRef = useRef(false)

  const measure = useCallback((text: string, font: string = CELL_FONT) => {
    const cache = measureCache.current
    const key = font === CELL_FONT ? text : `${font}\u0000${text}`
    let width = cache.get(key)
    if (width === undefined) {
      const ctx = canvasRef.current?.getContext("2d")
      if (!ctx) return text.length * 7
      const before = ctx.font
      ctx.font = font
      width = ctx.measureText(text).width
      ctx.font = before
      if (cache.size > 20_000) cache.clear()
      cache.set(key, width)
    }
    return width
  }, [])

  // What a cell shows in its column (Increase Decimal reads it) and AutoFit for the ribbon's Format menu.
  useEffect(() => {
    controller.shownText = (row, col) => {
      const value = controller.sheet.cells.value(row, col)
      if (value === undefined) return ""
      const look = lookOf(controller.wb.styles, controller.sheet.styleAt(row, col))
      return shownText(value, look, controller.sheet.cols.size(col), PAD, measure).text
    }
    controller.autofit = (axis) => {
      const indexes = controller.selectedIndexes(axis)
      if (axis === "cols")
        for (const col of indexes) controller.setLineSize("cols", [col], autofitColumnWidth(controller, col, measure))
      else
        for (const row of indexes)
          controller.setLineSize("rows", [row], fittedRowHeight(controller, row, measure), true)
    }
    controller.rowFit = (row) => fittedRowHeight(controller, row, measure)
    controller.columnFit = (col, rows) => autofitColumnWidth(controller, col, measure, rows)
    return () => {
      controller.shownText = null
      controller.autofit = null
      controller.rowFit = null
      controller.columnFit = null
    }
  }, [controller, measure])

  // Size: the canvas fills the host minus the scrollbars.
  useLayoutEffect(() => {
    const host = hostRef.current
    if (!host) return
    const observer = new ResizeObserver(() => {
      const rect = host.getBoundingClientRect()
      setSize({
        width: Math.max(0, Math.floor(rect.width) - SCROLLBAR),
        height: Math.max(0, Math.floor(rect.height) - SCROLLBAR),
      })
    })
    observer.observe(host)
    return () => observer.disconnect()
  }, [])

  // Theme: re-read the tokens when light/dark changes.
  useEffect(() => {
    const update = () => {
      if (hostRef.current) themeRef.current = readGridTheme(hostRef.current)
      setThemeVersion((v) => v + 1)
    }
    update()
    const observer = new MutationObserver(update)
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ["class", "data-theme"] })
    return () => observer.disconnect()
  }, [])

  // Marching ants while something is on our clipboard.
  // Marching ants also run round the reference Point mode is writing.
  const hasClip = controller.clip !== null || Boolean(controller.edit?.point)
  useEffect(() => {
    if (!hasClip) return
    const timer = window.setInterval(() => setAntsOffset((offset) => (offset + 1) % 7), 70)
    return () => window.clearInterval(timer)
  }, [hasClip])

  // The layout of this render: worked out here rather than in the drawing effect, so the layers over the canvas (charts,
  // filter buttons, list arrows) place themselves by the scroll and sizes just drawn, not the ones before.
  const zoomNow = controller.zoom
  const layoutNow = useMemo(
    () => (size.width === 0 ? null : computeLayout(controller, size.width / zoomNow, size.height / zoomNow, measure)),
    // version: the controller's state changed; scrollVersion: the sheet scrolled.
    [controller, version, scrollVersion, size, measure, zoomNow],
  )
  layoutRef.current = layoutNow ?? layoutRef.current

  // Draw on every change (state, size, theme, ants), at most once per frame.
  useLayoutEffect(() => {
    const canvas = canvasRef.current
    if (!canvas || size.width === 0) return
    const dpr = window.devicePixelRatio || 1
    if (canvas.width !== Math.round(size.width * dpr) || canvas.height !== Math.round(size.height * dpr)) {
      canvas.width = Math.round(size.width * dpr)
      canvas.height = Math.round(size.height * dpr)
    }
    // Zoom: the layout is in sheet pixels (the canvas size divided by the zoom); drawing scales back up.
    const zoom = controller.zoom
    const layout = layoutNow
    if (!layout) return
    const viewWidth = size.width / zoom - layout.rowHeaderWidth
    const viewHeight = size.height / zoom - layout.headerHeight
    if (controller.viewWidth !== viewWidth || controller.viewHeight !== viewHeight) {
      controller.viewWidth = viewWidth
      controller.viewHeight = viewHeight
      controller.rowHeaderWidth = layout.rowHeaderWidth
    }
    const ctx = canvas.getContext("2d")
    if (!ctx) return
    const start = import.meta.env.DEV ? performance.now() : 0
    drawGrid({ painter: ctx, layout, controller, theme: themeRef.current, dpr: dpr * zoom, measure, antsOffset })
    // Dev builds record draw times for the frame-time checks (scripts read the "cells:draw" measures).
    if (import.meta.env.DEV) performance.measure("cells:draw", { start, end: performance.now() })
  }, [controller, layoutNow, size, measure, antsOffset, themeVersion])

  // Keep keyboard focus on the hidden input whenever no other control has it.
  const focusInput = useCallback(() => {
    const input = inputRef.current
    if (input && document.activeElement !== input) input.focus({ preventScroll: true })
  }, [])

  useEffect(() => {
    focusInput()
  }, [focusInput])

  // The hidden input holds one selected space when idle, so the browser fires copy/cut for Ctrl+C/X.
  const editing = controller.edit
  useLayoutEffect(() => {
    const input = inputRef.current
    if (!input) return
    if (!editing) {
      if (input.value !== " ") input.value = " "
      if (document.activeElement === input) input.select()
      return
    }
    if (editing.surface === "cell" && !composingRef.current) {
      // The controller changed the text (pointing, F4, a picked function) or the selection: follow it.
      if (input.value !== editing.text) input.value = editing.text
      if (input.selectionStart !== editing.selStart || input.selectionEnd !== editing.selEnd)
        input.setSelectionRange(editing.selStart, editing.selEnd)
    }
    if (editing.surface === "bar" && input.value !== editing.text) {
      input.value = editing.text
    }
  }, [editing])

  // ── Pointer ────────────────────────────────────────────────────────────────────────────────────────────

  /** A pointer position in sheet pixels on the canvas (undoing the zoom). */
  const localPoint = (event: { clientX: number; clientY: number }) => {
    const rect = canvasRef.current!.getBoundingClientRect()
    const zoom = controller.zoom
    return { x: (event.clientX - rect.left) / zoom, y: (event.clientY - rect.top) / zoom }
  }

  const fillTarget = (source: Range, address: CellAddress): Range => {
    const dy = address.row > source.r1 ? address.row - source.r1 : address.row < source.r0 ? source.r0 - address.row : 0
    const dx = address.col > source.c1 ? address.col - source.c1 : address.col < source.c0 ? source.c0 - address.col : 0
    if (dy === 0 && dx === 0) {
      // Back inside the selection: the part beyond the pointer will be cleared.
      if (address.row < source.r1 && source.r1 > source.r0) return { ...source, r1: Math.max(source.r0, address.row) }
      if (address.col < source.c1 && source.c1 > source.c0) return { ...source, c1: Math.max(source.c0, address.col) }
      return source
    }
    if (dy >= dx) {
      return address.row > source.r1 ? { ...source, r1: address.row } : { ...source, r0: address.row }
    }
    return address.col > source.c1 ? { ...source, c1: address.col } : { ...source, c0: address.col }
  }

  const applyDrag = (x: number, y: number, ctrl: boolean) => {
    const drag = dragRef.current
    const layout = layoutRef.current
    if (!drag || !layout) return
    const address = cellAt(controller, layout, x, y)
    switch (drag.kind) {
      case "cells":
        controller.dragTo(address)
        break
      case "rows":
        controller.setSelection(
          {
            ...controller.selection,
            ...selectRowsAnchored(drag.anchor, address.row, controller),
          },
          null,
        )
        break
      case "cols":
        controller.setSelection(
          {
            ...controller.selection,
            ...selectColsAnchored(drag.anchor, address.col, controller),
          },
          null,
        )
        break
      case "fill":
        controller.fillPreview = fillTarget(drag.source, address)
        controller.changed()
        break
      case "point":
        controller.pointAt(drag.anchor, address)
        break
      case "resize-col": {
        const size = Math.max(0, Math.round(drag.startSize + (x - drag.startX)))
        drag.size = size
        setResizeTip({ x, y: 4, text: `Width: ${columnWidthLabel(size)}` })
        controller.changed()
        break
      }
      case "resize-row": {
        const size = Math.max(0, Math.round(drag.startSize + (y - drag.startY)))
        drag.size = size
        setResizeTip({ x: 4, y, text: `Height: ${rowHeightLabel(size)}` })
        controller.changed()
        break
      }
      case "pan": {
        const dx = drag.startX - x
        const dy = drag.startY - y
        if (!drag.moved && Math.hypot(dx, dy) * controller.zoom > TAP_SLOP) {
          drag.moved = true
          window.clearTimeout(pressTimer.current)
        }
        controller.scrollTo(drag.scrollX + dx, drag.scrollY + dy)
        flingRef.current.add(controller.scrollX, controller.scrollY, performance.now())
        break
      }
    }
    void ctrl
  }

  // Auto-scroll while dragging outside the cell area.
  useEffect(() => {
    const timer = window.setInterval(() => {
      const drag = dragRef.current
      const point = pointerRef.current
      const layout = layoutRef.current
      if (!drag || !point || !layout || drag.kind === "pan" || drag.kind.startsWith("resize")) return
      let dx = 0
      let dy = 0
      if (drag.kind !== "rows") {
        if (point.x > layout.width) dx = point.x - layout.width
        else if (point.x < layout.rowHeaderWidth) dx = point.x - layout.rowHeaderWidth
      }
      if (drag.kind !== "cols") {
        if (point.y > layout.height) dy = point.y - layout.height
        else if (point.y < layout.headerHeight) dy = point.y - layout.headerHeight
      }
      if (dx || dy) {
        controller.scrollTo(
          controller.scrollX + Math.sign(dx) * Math.min(400, 8 + dx * dx * 0.2),
          controller.scrollY + Math.sign(dy) * Math.min(800, 8 + dy * dy * 0.2),
        )
        applyDrag(point.x, point.y, point.ctrl)
      }
    }, 40)
    return () => window.clearInterval(timer)
  })

  const cssPoint = (event: { clientX: number; clientY: number }): Point => {
    const rect = canvasRef.current!.getBoundingClientRect()
    return { x: event.clientX - rect.left, y: event.clientY - rect.top }
  }

  const onPointerDown = (event: React.PointerEvent<HTMLCanvasElement>) => {
    const layout = layoutRef.current
    if (!layout) return
    const touch = event.pointerType === "touch"
    stopGlide()
    if (controller.touchHandles !== touch) {
      controller.touchHandles = touch
      controller.changed()
    }
    if (touch) {
      lastTouchRef.current = Date.now()
      touchesRef.current.set(event.pointerId, cssPoint(event))
      if (touchesRef.current.size === 2) {
        // A second finger: pinch to zoom (the first finger's pan or press is over).
        window.clearTimeout(pressTimer.current)
        const [a, b] = [...touchesRef.current.values()]
        dragRef.current = { kind: "pinch", startDistance: distance(a, b), startZoom: controller.zoom }
        canvasRef.current!.setPointerCapture(event.pointerId)
        return
      }
      if (touchesRef.current.size > 2) return
    }
    // A click on the sheet closes an open comment card (Excel).
    if (controller.annotations.open) controller.annotations.closeThread()
    const { x, y } = localPoint(event)
    const hit = hitTest(controller, layout, x, y)
    const shift = event.shiftKey
    const ctrl = platform.os === "macos" ? event.metaKey : event.ctrlKey
    if (event.button === 2) {
      // Right-click inside the selection keeps it; elsewhere it selects what was clicked first.
      handleRightClick(hit)
      return
    }
    if (event.button !== 0) return
    event.preventDefault()
    if (controller.edit) {
      if (hit.zone === "cell" && controller.canPoint) {
        // Point mode: the click writes a reference into the formula; dragging makes it a range.
        const cell = { row: hit.row, col: hit.col }
        const point = controller.edit.point
        const anchor = shift && point ? point.anchor : cell
        controller.pointAt(anchor, cell)
        dragRef.current = { kind: "point", anchor }
        pointerRef.current = { x, y, ctrl }
        canvasRef.current!.setPointerCapture(event.pointerId)
        return
      }
      // Clicking another cell finishes the entry (unless Data Validation keeps it in its cell).
      if (!controller.commitChecked()) return
    }
    focusInput()
    if (touch) {
      canvasRef.current!.setPointerCapture(event.pointerId)
      if (hit.zone === "gripper") {
        // Stretch the selection from the opposite corner.
        const r = controller.selection.ranges[0]
        const anchor = hit.end === "end" ? { row: r.r0, col: r.c0 } : { row: r.r1, col: r.c1 }
        const focus = hit.end === "end" ? { row: r.r1, col: r.c1 } : { row: r.r0, col: r.c0 }
        controller.setSelection({ ...controller.selection, anchor, focus }, null)
        dragRef.current = { kind: "cells" }
        pointerRef.current = { x, y, ctrl: false }
        return
      }
      const cell = hit.zone === "cell" ? hit : { row: 0, col: 0 }
      const pan: Drag = {
        kind: "pan",
        startX: x,
        startY: y,
        scrollX: controller.scrollX,
        scrollY: controller.scrollY,
        moved: false,
        row: "row" in cell ? cell.row : 0,
        col: "col" in cell ? cell.col : 0,
        clientX: event.clientX,
        clientY: event.clientY,
        pressed: false,
      }
      dragRef.current = pan
      flingRef.current = new FlingTracker()
      // Long press: the context menu, for what is under the finger (as a right-click).
      window.clearTimeout(pressTimer.current)
      pressTimer.current = window.setTimeout(() => {
        if (dragRef.current !== pan || pan.moved) return
        pan.pressed = true
        tapsRef.current.reset()
        handleRightClick(hit)
        controller.changed()
        onContextMenu({
          x: pan.clientX,
          y: pan.clientY,
          target: hit.zone === "rowHeader" ? "rows" : hit.zone === "colHeader" ? "cols" : "cells",
        })
      }, LONG_PRESS_MS)
      return
    }
    canvasRef.current!.setPointerCapture(event.pointerId)
    pointerRef.current = { x, y, ctrl }
    switch (hit.zone) {
      case "corner":
        controller.selectEverything()
        break
      case "cell":
        controller.pointerCell({ row: hit.row, col: hit.col }, { shift, ctrl })
        dragRef.current = { kind: "cells" }
        linkPress.current =
          !shift && !ctrl && controller.hyperlinks.target(hit.row, hit.col)
            ? { row: hit.row, col: hit.col, time: Date.now() }
            : null
        break
      case "rowHeader":
        controller.pointerRowHeader(hit.row, { shift, ctrl })
        dragRef.current = { kind: "rows", anchor: shift ? controller.selection.anchor.row : hit.row }
        break
      case "colHeader":
        controller.pointerColHeader(hit.col, { shift, ctrl })
        dragRef.current = { kind: "cols", anchor: shift ? controller.selection.anchor.col : hit.col }
        break
      case "fillHandle":
        dragRef.current = { kind: "fill", source: controller.selection.ranges[0] }
        controller.fillPreview = controller.selection.ranges[0]
        controller.changed()
        break
      case "colResize": {
        const startSize = controller.sheet.cols.size(hit.col)
        dragRef.current = { kind: "resize-col", col: hit.col, startX: x, startSize, size: startSize }
        setResizeTip({ x, y: 4, text: `Width: ${columnWidthLabel(startSize)}` })
        break
      }
      case "rowResize": {
        const startSize = controller.sheet.rows.size(hit.row)
        dragRef.current = { kind: "resize-row", row: hit.row, startY: y, startSize, size: startSize }
        setResizeTip({ x: 4, y, text: `Height: ${rowHeightLabel(startSize)}` })
        break
      }
    }
  }

  const handleRightClick = (hit: Hit) => {
    const selection = controller.selection
    const inSelection = (row: number, col: number) =>
      selection.ranges.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)
    if (controller.edit) controller.commitEdit()
    if (hit.zone === "cell" && !inSelection(hit.row, hit.col))
      controller.pointerCell(hit, { shift: false, ctrl: false })
    if (
      hit.zone === "rowHeader" &&
      !selection.ranges.some((r) => hit.row >= r.r0 && hit.row <= r.r1 && r.c0 === 0 && r.c1 === 16383)
    ) {
      controller.pointerRowHeader(hit.row, { shift: false, ctrl: false })
    }
    if (
      hit.zone === "colHeader" &&
      !selection.ranges.some((r) => hit.col >= r.c0 && hit.col <= r.c1 && r.r0 === 0 && r.r1 === MAX_ROWS - 1)
    ) {
      controller.pointerColHeader(hit.col, { shift: false, ctrl: false })
    }
  }

  const onPointerMove = (event: React.PointerEvent<HTMLCanvasElement>) => {
    const layout = layoutRef.current
    if (!layout) return
    if (event.pointerType === "touch" && touchesRef.current.has(event.pointerId)) {
      touchesRef.current.set(event.pointerId, cssPoint(event))
      const drag = dragRef.current
      if (drag?.kind === "pinch") {
        if (touchesRef.current.size < 2) return
        const [a, b] = [...touchesRef.current.values()]
        const next = pinch({
          startZoom: drag.startZoom,
          startDistance: drag.startDistance,
          distance: distance(a, b),
          mid: midpoint(a, b),
          scroll: { x: controller.scrollX, y: controller.scrollY },
          origin: { x: layout.originX, y: layout.originY },
          zoom: controller.zoom,
          clamp: (z) => Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z)),
        })
        controller.setZoom(next.zoom, true)
        controller.scrollTo(next.scroll.x, next.scroll.y)
        return
      }
    }
    const { x, y } = localPoint(event)
    const ctrl = platform.os === "macos" ? event.metaKey : event.ctrlKey
    pointerRef.current = { x, y, ctrl }
    if (dragRef.current) {
      applyDrag(x, y, ctrl)
      return
    }
    const hit = hitTest(controller, layout, x, y)
    let linked = false
    if (hit.zone === "cell") {
      const merge = controller.sheet.merges.at(hit.row, hit.col)
      const cell = merge ? { row: merge.r0, col: merge.c0 } : { row: hit.row, col: hit.col }
      hoverAt(controller.sheet.annotations.get(cell.row, cell.col) ? cell : null)
      const link = controller.hyperlinks.target(cell.row, cell.col)
      linked = link !== null
      if (!link) hideLinkTip()
      else if (!linkTip || linkTip.row !== cell.row || linkTip.col !== cell.col) {
        window.clearTimeout(linkTipTimer.current)
        setLinkTip(null)
        const text = linkScreenTip(link)
        linkTipTimer.current = window.setTimeout(() => setLinkTip({ x, y, ...cell, text }), LINK_TIP_MS)
      }
    } else {
      hoverAt(null)
      hideLinkTip()
    }
    setCursor(
      linked
        ? "pointer"
        : hit.zone === "colResize"
          ? "col-resize"
          : hit.zone === "rowResize"
            ? "row-resize"
            : hit.zone === "fillHandle"
              ? "crosshair"
              : hit.zone === "colHeader"
                ? "s-resize"
                : hit.zone === "rowHeader"
                  ? "e-resize"
                  : "cell",
    )
  }

  const endDrag = (event: React.PointerEvent<HTMLCanvasElement>) => {
    const wasTouch = touchesRef.current.delete(event.pointerId)
    window.clearTimeout(pressTimer.current)
    const drag = dragRef.current
    // A pinch ends with its first lifted finger; the other finger then does nothing until it lifts.
    if (wasTouch && !drag && touchesRef.current.size) return
    dragRef.current = null
    pointerRef.current = null
    setResizeTip(null)
    const press = linkPress.current
    linkPress.current = null
    if (!drag) return
    if (drag.kind === "cells" && press && event.type === "pointerup" && Date.now() - press.time < LONG_PRESS_MS) {
      const a = controller.selection.active
      const single = controller.selection.ranges.length === 1 && !controller.edit
      if (single && a.row === press.row && a.col === press.col) {
        hideLinkTip()
        controller.hyperlinks.follow(press.row, press.col)
        return
      }
    }
    const ctrl = platform.os === "macos" ? event.metaKey : event.ctrlKey
    if (drag.kind === "fill" && controller.fillPreview) {
      controller.fillTo(controller.fillPreview, ctrl)
    } else if (drag.kind === "resize-col") {
      const cols = resizeTargets(controller, "cols", drag.col)
      controller.setLineSize("cols", cols, drag.size)
    } else if (drag.kind === "resize-row") {
      const rows = resizeTargets(controller, "rows", drag.row)
      controller.setLineSize("rows", rows, drag.size)
    } else if (drag.kind === "pan" && drag.moved) {
      const v =
        event.type === "pointerup" && !touchesRef.current.size ? flingRef.current.release(performance.now()) : null
      if (v) glide(v)
    } else if (drag.kind === "pan" && !drag.moved && !drag.pressed && event.type === "pointerup") {
      controller.pointerCell({ row: drag.row, col: drag.col }, { shift: false, ctrl: false })
      // A tap on a link follows it (a long press opens the menu, with Edit Hyperlink).
      if (controller.hyperlinks.follow(drag.row, drag.col)) {
        tapsRef.current.reset()
        return
      }
      // A second tap on the same cell edits it, with the on-screen keyboard.
      if (tapsRef.current.tap(drag.row, drag.col, Date.now())) {
        controller.beginEdit("edit")
        const input = inputRef.current
        input?.blur()
        requestAnimationFrame(() => input?.focus())
        return
      }
    }
    if (drag.kind === "pinch") {
      controller.changed()
      return
    }
    // Format Painter paints the selection made with this drag.
    if (controller.format.painter && drag.kind !== "fill" && !drag.kind.startsWith("resize")) {
      const ranges = controller.selection.ranges
      controller.format.paint(ranges[ranges.length - 1])
    }
    controller.changed()
    // Pointing from the formula bar keeps typing there.
    if (drag.kind === "point" && controller.edit?.surface === "bar") return
    focusInput()
  }

  const onDoubleClick = (event: React.MouseEvent<HTMLCanvasElement>) => {
    const layout = layoutRef.current
    if (!layout) return
    const { x, y } = localPoint(event)
    const hit = hitTest(controller, layout, x, y)
    if (hit.zone === "cell") {
      controller.beginEdit("edit")
    } else if (hit.zone === "colResize") {
      const cols = resizeTargets(controller, "cols", hit.col)
      for (const col of cols) controller.setLineSize("cols", [col], autofitColumnWidth(controller, col, measure))
    } else if (hit.zone === "rowResize") {
      const rows = resizeTargets(controller, "rows", hit.row)
      for (const row of rows) controller.setLineSize("rows", [row], fittedRowHeight(controller, row, measure), true)
    } else if (hit.zone === "fillHandle") {
      fillDownToData(controller)
    }
  }

  // ── Touchpad pinch (Linux) ─────────────────────────────────────────────────────────────────────────────
  // WebKitGTK would zoom the whole window; the shell catches the pinch (src-tauri/src/pinch.rs) and sends its scale
  // here, and the sheet zooms about the pointer as a touchscreen pinch does.
  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    let pointer: { x: number; y: number } | null = null
    let start: number | null = null
    // The shell sends a step for every touchpad event, often more than once a frame: the latest wins.
    let scale: number | null = null
    let frame = 0
    const onMove = (event: MouseEvent) => {
      const rect = canvas.getBoundingClientRect()
      pointer = { x: event.clientX - rect.left, y: event.clientY - rect.top }
    }
    const onLeave = () => (pointer = null)
    canvas.addEventListener("mousemove", onMove)
    canvas.addEventListener("mouseleave", onLeave)
    const apply = () => {
      frame = 0
      const layout = layoutRef.current
      if (start === null || scale === null || !layout) return
      const rect = canvas.getBoundingClientRect()
      const next = pinch({
        startZoom: start,
        startDistance: 1,
        distance: scale,
        mid: pointer ?? { x: rect.width / 2, y: rect.height / 2 },
        scroll: { x: controller.scrollX, y: controller.scrollY },
        origin: { x: layout.originX, y: layout.originY },
        zoom: controller.zoom,
        clamp: (z) => Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Math.round(z * 100) / 100)),
      })
      controller.setZoom(next.zoom, true)
      controller.scrollTo(next.scroll.x, next.scroll.y)
    }
    const host = window as Window & { __cellsPinch?: (phase: string, scale: number) => void }
    host.__cellsPinch = (phase, value) => {
      if (phase === "begin") {
        start = controller.zoom
        scale = null
      } else if (phase === "change" && start !== null && Number.isFinite(value)) {
        scale = value
        if (!frame) frame = requestAnimationFrame(apply)
      } else if (phase === "end" && start !== null) {
        cancelAnimationFrame(frame)
        apply()
        start = null
        controller.changed()
      }
    }
    return () => {
      canvas.removeEventListener("mousemove", onMove)
      canvas.removeEventListener("mouseleave", onLeave)
      cancelAnimationFrame(frame)
      delete host.__cellsPinch
    }
  }, [controller])

  // ── Wheel ──────────────────────────────────────────────────────────────────────────────────────────────

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const onWheel = (event: WheelEvent) => {
      event.preventDefault()
      if (event.ctrlKey || event.metaKey) {
        // Ctrl+wheel zooms in steps of 15% (Excel).
        if (event.deltaY !== 0)
          controller.setZoom(Math.round((controller.zoom - Math.sign(event.deltaY) * 0.15) * 100) / 100)
        return
      }
      const unit = event.deltaMode === 1 ? 20 : event.deltaMode === 2 ? controller.viewHeight : 1
      let dx = event.deltaX * unit
      let dy = event.deltaY * unit
      if (event.shiftKey && dx === 0) {
        dx = dy
        dy = 0
      }
      controller.scrollTo(controller.scrollX + dx, controller.scrollY + dy)
    }
    canvas.addEventListener("wheel", onWheel, { passive: false })
    return () => canvas.removeEventListener("wheel", onWheel)
  }, [controller])

  // ── On-screen keyboard ────────────────────────────────────────────────────────────────────────────────
  // When the keyboard opens (the visual viewport shrinks) or the view moves, scroll the sheet so the cell being
  // edited stays above it.
  useEffect(() => {
    const vv = window.visualViewport
    if (!vv) return
    const keepVisible = () => {
      const input = inputRef.current
      if (!controller.edit || controller.edit.surface !== "cell" || !input) return
      const overlap = keyboardOverlap(input.getBoundingClientRect().bottom, vv.offsetTop, vv.height)
      if (overlap > 0) controller.scrollTo(controller.scrollX, controller.scrollY + overlap / controller.zoom)
    }
    vv.addEventListener("resize", keepVisible)
    vv.addEventListener("scroll", keepVisible)
    return () => {
      vv.removeEventListener("resize", keepVisible)
      vv.removeEventListener("scroll", keepVisible)
    }
  }, [controller])

  // ── Keyboard and clipboard (on the hidden input) ──────────────────────────────────────────────────────

  const onKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (event.nativeEvent.isComposing || composingRef.current) return
    if (isSystemShortcut(event, platform.os)) return
    const input = keyInputFrom(event, platform.os)
    if (controller.edit) {
      if (controller.keyEditing(input)) event.preventDefault()
      return
    }
    // Copy, cut and paste arrive as clipboard events.
    if (input.primary && !input.alt && !input.shift && ["c", "x", "v"].includes(input.key.toLowerCase())) return
    if (controller.keyReady(input)) event.preventDefault()
  }

  const onInput = (event: React.FormEvent<HTMLTextAreaElement>) => {
    const { value, selectionStart, selectionEnd } = event.currentTarget
    if (!controller.edit) {
      if (composingRef.current) return
      // A printable key replaced the placeholder space: start a new entry with it.
      controller.beginEdit("enter", value)
      return
    }
    controller.setEditText(value, selectionStart, selectionEnd)
  }

  const onSelect = (event: React.SyntheticEvent<HTMLTextAreaElement>) => {
    const edit = controller.edit
    if (edit?.surface !== "cell" || composingRef.current) return
    // A selection event can arrive while the textarea still shows older text (before the controller's change is
    // drawn); its offsets would be stale.
    if (event.currentTarget.value !== edit.text) return
    controller.setEditSelection(event.currentTarget.selectionStart, event.currentTarget.selectionEnd)
  }

  const onCompositionStart = () => {
    composingRef.current = true
    if (!controller.edit) {
      const input = inputRef.current
      if (input) input.value = ""
      controller.beginEdit("enter", "")
    }
  }

  const onCompositionEnd = (event: React.CompositionEvent<HTMLTextAreaElement>) => {
    composingRef.current = false
    const { value, selectionStart, selectionEnd } = event.currentTarget
    controller.setEditText(value, selectionStart, selectionEnd)
  }

  const onCopy = (event: React.ClipboardEvent<HTMLTextAreaElement>, cut: boolean) => {
    if (controller.edit) return
    event.preventDefault()
    const content = controller.copy(cut)
    if (!content) return
    event.clipboardData.setData("text/plain", content.text)
    event.clipboardData.setData("text/html", content.html)
  }

  const onPaste = (event: React.ClipboardEvent<HTMLTextAreaElement>) => {
    if (controller.edit) return
    event.preventDefault()
    controller.paste(event.clipboardData.getData("text/plain"), event.clipboardData.getData("text/html"))
  }

  // ── Editor placement ───────────────────────────────────────────────────────────────────────────────────

  const layout = layoutRef.current
  const zoom = controller.zoom
  const active = editing ? { row: editing.row, col: editing.col } : controller.selection.active
  let editorStyle: React.CSSProperties = { left: 0, top: 0, width: 1, height: 1 }
  if (layout && size.width) {
    const rect = rangeRect(controller, layout, controller.sheet.merges.at(active.row, active.col) ?? rangeOf(active))
    if (editing) {
      // The entry is typed in the cell's own font, colour and fill, as Excel shows it.
      const look = lookOf(controller.wb.styles, controller.sheet.styleAt(active.row, active.col))
      const lineHeight = lineHeightOf(look.fontPx)
      const lines = editing.text.split("\n")
      const textWidth = Math.max(...lines.map((line) => measure(line, look.font))) + 10
      const maxWidth = Math.max(rect.w, layout.width - rect.x)
      editorStyle = {
        left: rect.x * zoom,
        top: rect.y * zoom,
        width: Math.min(maxWidth, Math.max(rect.w + 1, textWidth)),
        height: Math.max(rect.h + 1, lines.length * lineHeight + 4),
        font: look.font,
        lineHeight: `${lineHeight}px`,
        // A formula's text is drawn by the coloured copy underneath; plain entries take the cell's colours.
        ...(look.color && !editing.text.startsWith("=") ? { color: look.color } : {}),
        ...(look.fill && !editing.text.startsWith("=") ? { background: look.fill } : {}),
      }
      // The editor is laid out at 100% and scaled like the cells.
      if (zoom !== 1) Object.assign(editorStyle, { transform: `scale(${zoom})`, transformOrigin: "0 0" })
    } else {
      editorStyle = { left: Math.max(0, rect.x * zoom), top: Math.max(0, rect.y * zoom), width: 1, height: 1 }
    }
  }

  const sheet = controller.sheet
  const used = sheet.cells.usedBounds()
  // The scrollbars cover the scrolling part only (frozen rows and columns don't scroll).
  const minX = controller.minScrollX
  const minY = controller.minScrollY
  const contentWidth =
    Math.max(sheet.cols.offset((used?.c1 ?? 0) + 1), controller.scrollX + controller.paneWidth) +
    controller.paneWidth / 2 -
    minX
  const contentHeight =
    Math.max(sheet.rows.offset((used?.r1 ?? 0) + 1), controller.scrollY + controller.paneHeight) +
    controller.paneHeight / 2 -
    minY

  const outlineW = outlineBarSize(controller, "rows")
  const outlineH = outlineBarSize(controller, "cols")

  return (
    <div
      className="grid-frame"
      style={{ gridTemplateColumns: `${outlineW}px minmax(0, 1fr)`, gridTemplateRows: `${outlineH}px minmax(0, 1fr)` }}
    >
      {outlineW && outlineH ? <div className="outline-corner" /> : null}
      {layoutNow && outlineH ? (
        <OutlineBar controller={controller} layout={layoutNow} zoom={zoom} axis="cols" onDone={focusInput} />
      ) : null}
      {layoutNow && outlineW ? (
        <OutlineBar controller={controller} layout={layoutNow} zoom={zoom} axis="rows" onDone={focusInput} />
      ) : null}
      <div ref={hostRef} className="grid-host" data-testid="grid">
        <canvas
          ref={canvasRef}
          className="grid-canvas"
          style={{ width: size.width, height: size.height, cursor }}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={endDrag}
          onPointerCancel={endDrag}
          onPointerLeave={() => {
            hoverAt(null)
            hideLinkTip()
          }}
          onDoubleClick={onDoubleClick}
          onContextMenu={(event) => {
            event.preventDefault()
            const layout = layoutRef.current
            // A touch long press opens the menu itself (iPadOS sends no contextmenu, Android sends one too).
            if (!layout || Date.now() - lastTouchRef.current < 1500) return
            const { x, y } = localPoint(event)
            const hit = hitTest(controller, layout, x, y)
            handleRightClick(hit)
            onContextMenu({
              x: event.clientX,
              y: event.clientY,
              target: hit.zone === "rowHeader" ? "rows" : hit.zone === "colHeader" ? "cols" : "cells",
            })
          }}
        />
        {editing ? (
          <FormulaHighlight controller={controller} className="cell-editor-highlight" style={editorStyle} />
        ) : null}
        <textarea
          ref={inputRef}
          data-grid-input
          className={
            (editing && editing.surface === "cell"
              ? "cell-editor is-editing"
              : editing
                ? "cell-editor is-mirror"
                : "cell-editor") + (editing?.text.startsWith("=") ? " has-highlight" : "")
          }
          style={editorStyle}
          aria-label="Cell editor"
          // On a touch screen the keyboard shows only while a cell is being edited.
          inputMode={controller.touchHandles && !editing ? "none" : undefined}
          spellCheck={false}
          autoCapitalize="off"
          autoComplete="off"
          autoCorrect="off"
          wrap="off"
          onKeyDown={onKeyDown}
          onInput={onInput}
          onSelect={onSelect}
          onFocus={() => controller.setEditSurface("cell")}
          onCompositionStart={onCompositionStart}
          onCompositionEnd={onCompositionEnd}
          onCopy={(event) => onCopy(event, false)}
          onCut={(event) => onCopy(event, true)}
          onPaste={onPaste}
          readOnly={editing?.surface === "bar"}
        />
        {editing && editing.surface === "cell" && layout ? (
          <FormulaAssist
            controller={controller}
            style={{
              left: Number(editorStyle.left ?? 0),
              top: Number(editorStyle.top ?? 0) + Number(editorStyle.height ?? 0) * zoom + 3,
            }}
            onPick={focusInput}
          />
        ) : null}
        {layout ? <ChartLayer controller={controller} layout={layout} zoom={zoom} onDone={focusInput} /> : null}
        {layout ? <TraceLayer controller={controller} layout={layout} zoom={zoom} /> : null}
        {layout ? <FilterButtons controller={controller} layout={layout} zoom={zoom} onDone={focusInput} /> : null}
        {layout ? <ValidationLayer controller={controller} layout={layout} zoom={zoom} onDone={focusInput} /> : null}
        {layout ? (
          <AnnotationLayer
            controller={controller}
            layout={layout}
            zoom={zoom}
            hover={hover}
            onKeepHover={keepHover}
            onLeaveHover={() => hoverAt(null)}
            onDone={focusInput}
          />
        ) : null}
        {linkTip && !dragRef.current ? (
          <div className="link-tip" role="tooltip" style={{ left: linkTip.x * zoom + 14, top: linkTip.y * zoom + 18 }}>
            {linkTip.text}
          </div>
        ) : null}
        {resizeTip ? (
          <div className="resize-tip" style={{ left: resizeTip.x * zoom + 12, top: resizeTip.y * zoom + 4 }}>
            {resizeTip.text}
          </div>
        ) : null}
        {dragRef.current?.kind === "resize-col" && layout ? (
          <div
            className="resize-guide is-vertical"
            style={{
              left: (colX(sheet, layout, dragRef.current.col) + dragRef.current.size) * zoom,
            }}
          />
        ) : null}
        {dragRef.current?.kind === "resize-row" && layout ? (
          <div
            className="resize-guide is-horizontal"
            style={{
              top: (rowY(sheet, layout, dragRef.current.row) + dragRef.current.size) * zoom,
            }}
          />
        ) : null}
        <Scrollbar
          orientation="vertical"
          length={size.height}
          offset={(layout?.originY ?? 20) * zoom}
          position={controller.scrollY - minY}
          view={controller.paneHeight}
          content={contentHeight}
          onScroll={(y) => controller.scrollTo(controller.scrollX, y + minY)}
          onDone={focusInput}
        />
        <Scrollbar
          orientation="horizontal"
          length={size.width}
          offset={(layout?.originX ?? 40) * zoom}
          position={controller.scrollX - minX}
          view={controller.paneWidth}
          content={contentWidth}
          onScroll={(x) => controller.scrollTo(x + minX, controller.scrollY)}
          onDone={focusInput}
        />
        <div className="grid-scroll-corner" />
      </div>
    </div>
  )
}

function selectRowsAnchored(anchor: number, row: number, controller: SheetController) {
  const r0 = Math.min(anchor, row)
  const r1 = Math.max(anchor, row)
  const ranges = [...controller.selection.ranges.slice(0, -1), { r0, r1, c0: 0, c1: 16383 }]
  return {
    ranges,
    activeRange: ranges.length - 1,
    anchor: { row: anchor, col: 0 },
    focus: { row, col: 16383 },
    active: { row: anchor, col: controller.selection.active.col },
  }
}

function selectColsAnchored(anchor: number, col: number, controller: SheetController) {
  const c0 = Math.min(anchor, col)
  const c1 = Math.max(anchor, col)
  const ranges = [...controller.selection.ranges.slice(0, -1), { r0: 0, r1: MAX_ROWS - 1, c0, c1 }]
  return {
    ranges,
    activeRange: ranges.length - 1,
    anchor: { row: 0, col: anchor },
    focus: { row: MAX_ROWS - 1, col },
    active: { row: controller.selection.active.row, col: anchor },
  }
}

/** Resizing a selected column resizes every selected column (Excel); otherwise just that one. */
function resizeTargets(controller: SheetController, axis: "rows" | "cols", index: number): number[] {
  const lines = controller.selectedLines(axis)
  const whole = controller.selection.ranges.every((range) =>
    axis === "cols" ? range.r0 === 0 && range.r1 === MAX_ROWS - 1 : range.c0 === 0 && range.c1 === 16383,
  )
  if (whole && lines.some(([from, to]) => index >= from && index <= to)) {
    return controller.selectedIndexes(axis)
  }
  return [index]
}

/** Double-clicking the fill handle fills down as far as the data in the next column goes (Excel). */
function fillDownToData(controller: SheetController) {
  const source = controller.selection.ranges[0]
  const sheet = controller.sheet
  const neighbour = [source.c0 - 1, source.c1 + 1].find((col) => col >= 0 && sheet.cells.hasValue(source.r1 + 1, col))
  if (neighbour === undefined) return
  let end = source.r1 + 1
  while (end + 1 < MAX_ROWS && sheet.cells.hasValue(end + 1, neighbour)) end++
  controller.fillTo({ ...source, r1: end }, false)
}
