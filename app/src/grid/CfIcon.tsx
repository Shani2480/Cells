import type { IconRef } from "@cells/core"
import { useEffect, useRef } from "react"
import { drawIcon } from "./cf-icons"
import type { CellPainter } from "./draw-cells"

/** One conditional formatting icon, drawn with the grid's own painter (for menus and dialogs). */
export function CfIcon({ icon, size = 16 }: { icon: IconRef; size?: number }) {
  const ref = useRef<HTMLCanvasElement>(null)
  useEffect(() => {
    const canvas = ref.current
    const ctx = canvas?.getContext("2d")
    if (!canvas || !ctx) return
    const dpr = window.devicePixelRatio || 1
    canvas.width = size * dpr
    canvas.height = size * dpr
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    ctx.clearRect(0, 0, size, size)
    drawIcon(ctx as unknown as CellPainter, icon.set, icon.index, 0, 0, size)
  }, [icon.set, icon.index, size])
  return <canvas ref={ref} className="cf-icon" style={{ width: size, height: size, flex: "none" }} aria-hidden />
}
