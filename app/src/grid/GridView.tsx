import { MAX_ROWS, rangeOf, type CellAddress, type Range } from "@cells/core"
import { useCallback, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from "react"
import { usePlatform } from "@/platform/context"
import type { SheetController } from "@/sheet/controller"
import { keyInputFrom } from "@/sheet/keys"
import { autofitColumnWidth, autofitRowHeight, columnWidthLabel, rowHeightLabel } from "./autofit"
import { drawGrid } from "./draw"
import { FormulaAssist, FormulaHighlight } from "./FormulaAssist"
import { cellAt, computeLayout, hitTest, rangeRect, type GridLayout, type Hit } from "./layout"
import { CELL_FONT, LIGHT_THEME, readGridTheme, type GridTheme } from "./theme"
import { Scrollbar } from "./Scrollbar"

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
    }

export type GridContextMenu = { x: number; y: number; target: "cells" | "rows" | "cols" }

export function GridView({
  controller,
  onContextMenu,
}: {
  controller: SheetController
  onContextMenu: (menu: GridContextMenu) => void
}) {
  const version = useSyncExternalStore(controller.subscribe, controller.getVersion)
  const platform = usePlatform()
  const hostRef = useRef<HTMLDivElement>(null)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const layoutRef = useRef<GridLayout | null>(null)
  const dragRef = useRef<Drag | null>(null)
  const pointerRef = useRef<{ x: number; y: number; ctrl: boolean } | null>(null)
  const themeRef = useRef<GridTheme>(LIGHT_THEME)
  const measureCache = useRef(new Map<string, number>())
  const [size, setSize] = useState({ width: 0, height: 0 })
  const [cursor, setCursor] = useState("cell")
  const [resizeTip, setResizeTip] = useState<{ x: number; y: number; text: string } | null>(null)
  const [antsOffset, setAntsOffset] = useState(0)
  const [themeVersion, setThemeVersion] = useState(0)
  const composingRef = useRef(false)

  const measure = useCallback((text: string) => {
    const cache = measureCache.current
    let width = cache.get(text)
    if (width === undefined) {
      const ctx = canvasRef.current?.getContext("2d")
      if (!ctx) return text.length * 7
      ctx.font = CELL_FONT
      width = ctx.measureText(text).width
      if (cache.size > 20_000) cache.clear()
      cache.set(text, width)
    }
    return width
  }, [])

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

  // Draw on every change (state, size, theme, ants), at most once per frame.
  useLayoutEffect(() => {
    const canvas = canvasRef.current
    if (!canvas || size.width === 0) return
    const dpr = window.devicePixelRatio || 1
    if (canvas.width !== Math.round(size.width * dpr) || canvas.height !== Math.round(size.height * dpr)) {
      canvas.width = Math.round(size.width * dpr)
      canvas.height = Math.round(size.height * dpr)
    }
    const layout = computeLayout(controller, size.width, size.height, measure)
    layoutRef.current = layout
    const viewWidth = size.width - layout.rowHeaderWidth
    const viewHeight = size.height - layout.headerHeight
    if (controller.viewWidth !== viewWidth || controller.viewHeight !== viewHeight) {
      controller.viewWidth = viewWidth
      controller.viewHeight = viewHeight
      controller.rowHeaderWidth = layout.rowHeaderWidth
    }
    const ctx = canvas.getContext("2d")
    if (!ctx) return
    const start = import.meta.env.DEV ? performance.now() : 0
    drawGrid({ painter: ctx, layout, controller, theme: themeRef.current, dpr, measure, antsOffset })
    // Dev builds record draw times for the frame-time checks (scripts read the "cells:draw" measures).
    if (import.meta.env.DEV) performance.measure("cells:draw", { start, end: performance.now() })
  }, [controller, version, size, measure, antsOffset, themeVersion])

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

  const localPoint = (event: { clientX: number; clientY: number }) => {
    const rect = canvasRef.current!.getBoundingClientRect()
    return { x: event.clientX - rect.left, y: event.clientY - rect.top }
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
        if (Math.abs(dx) + Math.abs(dy) > 6) drag.moved = true
        controller.scrollTo(drag.scrollX + dx, drag.scrollY + dy)
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

  const onPointerDown = (event: React.PointerEvent<HTMLCanvasElement>) => {
    const layout = layoutRef.current
    if (!layout) return
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
      // Clicking another cell finishes the entry.
      controller.commitEdit()
    }
    focusInput()
    if (event.pointerType === "touch") {
      const cell = hit.zone === "cell" ? hit : { row: 0, col: 0 }
      dragRef.current = {
        kind: "pan",
        startX: x,
        startY: y,
        scrollX: controller.scrollX,
        scrollY: controller.scrollY,
        moved: false,
        row: "row" in cell ? cell.row : 0,
        col: "col" in cell ? cell.col : 0,
      }
      canvasRef.current!.setPointerCapture(event.pointerId)
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
    const { x, y } = localPoint(event)
    const ctrl = platform.os === "macos" ? event.metaKey : event.ctrlKey
    pointerRef.current = { x, y, ctrl }
    if (dragRef.current) {
      applyDrag(x, y, ctrl)
      return
    }
    const hit = hitTest(controller, layout, x, y)
    setCursor(
      hit.zone === "colResize"
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
    const drag = dragRef.current
    dragRef.current = null
    pointerRef.current = null
    setResizeTip(null)
    if (!drag) return
    const ctrl = platform.os === "macos" ? event.metaKey : event.ctrlKey
    if (drag.kind === "fill" && controller.fillPreview) {
      controller.fillTo(controller.fillPreview, ctrl)
    } else if (drag.kind === "resize-col") {
      const cols = resizeTargets(controller, "cols", drag.col)
      controller.setLineSize("cols", cols, drag.size)
    } else if (drag.kind === "resize-row") {
      const rows = resizeTargets(controller, "rows", drag.row)
      controller.setLineSize("rows", rows, drag.size)
    } else if (drag.kind === "pan" && !drag.moved) {
      controller.pointerCell({ row: drag.row, col: drag.col }, { shift: false, ctrl: false })
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
      for (const row of rows) controller.setLineSize("rows", [row], autofitRowHeight(controller, row))
    } else if (hit.zone === "fillHandle") {
      fillDownToData(controller)
    }
  }

  // ── Wheel ──────────────────────────────────────────────────────────────────────────────────────────────

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const onWheel = (event: WheelEvent) => {
      if (event.ctrlKey) return // Zoom (Phase 3).
      event.preventDefault()
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

  // ── Keyboard and clipboard (on the hidden input) ──────────────────────────────────────────────────────

  const onKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (event.nativeEvent.isComposing || composingRef.current) return
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
  const active = editing ? { row: editing.row, col: editing.col } : controller.selection.active
  let editorStyle: React.CSSProperties = { left: 0, top: 0, width: 1, height: 1 }
  if (layout && size.width) {
    const rect = rangeRect(controller, layout, rangeOf(active))
    if (editing) {
      const lines = editing.text.split("\n")
      const textWidth = Math.max(...lines.map((line) => measure(line))) + 10
      const maxWidth = Math.max(rect.w, layout.width - rect.x)
      editorStyle = {
        left: rect.x,
        top: rect.y,
        width: Math.min(maxWidth, Math.max(rect.w + 1, textWidth)),
        height: Math.max(rect.h + 1, lines.length * 18 + 4),
      }
    } else {
      editorStyle = { left: Math.max(0, rect.x), top: Math.max(0, rect.y), width: 1, height: 1 }
    }
  }

  const sheet = controller.sheet
  const used = sheet.cells.usedBounds()
  const contentWidth =
    Math.max(sheet.cols.offset((used?.c1 ?? 0) + 1), controller.scrollX + controller.viewWidth) +
    controller.viewWidth / 2
  const contentHeight =
    Math.max(sheet.rows.offset((used?.r1 ?? 0) + 1), controller.scrollY + controller.viewHeight) +
    controller.viewHeight / 2

  return (
    <div ref={hostRef} className="grid-host" data-testid="grid">
      <canvas
        ref={canvasRef}
        className="grid-canvas"
        style={{ width: size.width, height: size.height, cursor }}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={endDrag}
        onPointerCancel={endDrag}
        onDoubleClick={onDoubleClick}
        onContextMenu={(event) => {
          event.preventDefault()
          const layout = layoutRef.current
          if (!layout) return
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
        className={
          (editing && editing.surface === "cell"
            ? "cell-editor is-editing"
            : editing
              ? "cell-editor is-mirror"
              : "cell-editor") + (editing?.text.startsWith("=") ? " has-highlight" : "")
        }
        style={editorStyle}
        aria-label="Cell editor"
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
            top: Number(editorStyle.top ?? 0) + Number(editorStyle.height ?? 0) + 3,
          }}
          onPick={focusInput}
        />
      ) : null}
      {resizeTip ? (
        <div className="resize-tip" style={{ left: resizeTip.x + 12, top: resizeTip.y + 4 }}>
          {resizeTip.text}
        </div>
      ) : null}
      {dragRef.current?.kind === "resize-col" && layout ? (
        <div
          className="resize-guide is-vertical"
          style={{
            left:
              layout.rowHeaderWidth +
              sheet.cols.offset(dragRef.current.col) -
              controller.scrollX +
              dragRef.current.size,
          }}
        />
      ) : null}
      {dragRef.current?.kind === "resize-row" && layout ? (
        <div
          className="resize-guide is-horizontal"
          style={{
            top:
              layout.headerHeight + sheet.rows.offset(dragRef.current.row) - controller.scrollY + dragRef.current.size,
          }}
        />
      ) : null}
      <Scrollbar
        orientation="vertical"
        length={size.height}
        offset={layout?.headerHeight ?? 20}
        position={controller.scrollY}
        view={controller.viewHeight}
        content={contentHeight}
        onScroll={(y) => controller.scrollTo(controller.scrollX, y)}
        onDone={focusInput}
      />
      <Scrollbar
        orientation="horizontal"
        length={size.width}
        offset={layout?.rowHeaderWidth ?? 40}
        position={controller.scrollX}
        view={controller.viewWidth}
        content={contentWidth}
        onScroll={(x) => controller.scrollTo(x, controller.scrollY)}
        onDone={focusInput}
      />
      <div className="grid-scroll-corner" />
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
