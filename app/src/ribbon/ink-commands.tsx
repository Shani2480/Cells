import { Eraser, Hand, Highlighter, MousePointer2, PenLine } from "lucide-react"
import { ColorPanel, colorName } from "./colors"
import type { Command, MenuEntry, RibbonEnv } from "./registry"

/** Millimetres to pixels at 100% (96 per inch). */
export const mm = (v: number) => (v * 96) / 25.4

const PEN_SIZES = [0.25, 0.35, 0.5, 0.7, 1, 2, 3.5]
const HIGHLIGHTER_SIZES = [2, 4, 6, 8, 10]

const sizeText = (px: number) => `${Math.round((px * 25.4 * 100) / 96) / 100} mm`

/** A pen's picture: the pen in its colour, as the Draw tab's gallery shows it. */
function penIcon(env: RibbonEnv, slot: number, size: number) {
  const pen = env.c.objects.pens[slot]
  const Icon = pen.kind === "highlighter" ? Highlighter : PenLine
  return (
    <span className="ink-pen-icon" style={{ color: pen.color }}>
      <Icon size={size} strokeWidth={1.75} aria-hidden />
    </span>
  )
}

function penMenu(slot: number): MenuEntry[] {
  return [
    { heading: "Thickness" },
    `InkPen${slot}Sizes`,
    "-",
    {
      panel: (env, close) => (
        <ColorPanel
          theme={env.c.wb.theme}
          current={env.c.objects.pens[slot].color}
          onPick={(hex) => {
            if (hex) env.c.objects.setPenLook(slot, { color: hex })
            close()
          }}
        />
      ),
    },
  ]
}

/** The Draw tab (pens, highlighter, eraser, Draw with Touch) and Review › Ink (Hide Ink, Delete All Ink). */
export function inkCommands(): Command[] {
  const pens = [0, 1, 2].flatMap((slot): Command[] => {
    const sizes = (env: RibbonEnv) => (env.c.objects.pens[slot].kind === "highlighter" ? HIGHLIGHTER_SIZES : PEN_SIZES)
    return [
      {
        id: `InkPen${slot}`,
        label: slot === 2 ? "Highlighter" : "Pen",
        labelFor: (env) => {
          const p = env.c.objects.pens[slot]
          return `${p.kind === "highlighter" ? "Highlighter" : "Pen"}: ${colorName(p.color)}, ${sizeText(p.width)}`
        },
        description:
          slot === 2
            ? "Highlight over the sheet, like a marker pen. Click again to change its colour and thickness."
            : "Write and draw on the sheet with a pen. Click again to change its colour and thickness.",
        keytip: `P${slot + 1}`,
        icon: (env) => penIcon(env, slot, 16),
        largeIcon: (env) => penIcon(env, slot, 32),
        pressed: (env) => env.c.objects.penSlot === slot,
        run: (env) => env.c.objects.usePen(slot),
        menu: penMenu(slot),
      },
      {
        id: `InkPen${slot}Sizes`,
        label: "Thickness",
        menu: (env) => sizes(env).map((s) => `InkPen${slot}Size:${s}`),
      },
      ...[...new Set([...PEN_SIZES, ...HIGHLIGHTER_SIZES])].map((s): Command => ({
        id: `InkPen${slot}Size:${s}`,
        label: `${s} mm`,
        pressed: (env) => Math.abs(env.c.objects.pens[slot].width - mm(s)) < 0.01,
        run: (env) => env.c.objects.setPenLook(slot, { width: mm(s) }),
      })),
    ]
  })
  return [
    {
      id: "InkSelect",
      label: "Select",
      description: "Stop drawing: click cells and objects again, as usual.",
      keytip: "S",
      icon: <MousePointer2 size={16} strokeWidth={1.75} aria-hidden />,
      largeIcon: <MousePointer2 size={32} strokeWidth={1.5} aria-hidden />,
      pressed: (env) => env.c.objects.pen === null,
      run: (env) => env.c.objects.usePen(null),
    },
    {
      id: "InkEraser",
      label: "Eraser",
      description: "Erase ink: every stroke you drag over goes.",
      keytip: "E",
      icon: <Eraser size={16} strokeWidth={1.75} aria-hidden />,
      largeIcon: <Eraser size={32} strokeWidth={1.5} aria-hidden />,
      enabled: (env) => env.c.objects.hasInk(),
      pressed: (env) => env.c.objects.pen?.kind === "eraser",
      run: (env) => env.c.objects.usePen("eraser"),
    },
    ...pens,
    {
      id: "InkDrawWithTouch",
      label: "Draw with Touch",
      description: "Draw with your finger as well as a pen. When it's off, a finger scrolls the sheet.",
      keytip: "T",
      icon: <Hand size={16} strokeWidth={1.75} aria-hidden />,
      largeIcon: <Hand size={32} strokeWidth={1.5} aria-hidden />,
      pressed: (env) => env.c.objects.drawWithTouch,
      run: (env) => {
        env.c.objects.drawWithTouch = !env.c.objects.drawWithTouch
        env.c.changed()
      },
    },
    {
      id: "HideInkSplitButton",
      label: "Hide Ink",
      description: "Hide the ink on the sheets, or delete it.",
      keytip: "H",
      enabled: (env) => env.c.objects.hasInk(true),
      pressed: (env) => env.c.objects.inkHidden,
      run: (env) => env.c.objects.setInkHidden(!env.c.objects.inkHidden),
      menu: ["HideInk", "-", "InkDeleteAllOnSheet", "InkDeleteAll"],
    },
    {
      id: "HideInk",
      label: "Hide Ink",
      description: "Hide the ink on every sheet (it's still there, and saved).",
      keytip: "H",
      enabled: (env) => env.c.objects.hasInk(true),
      pressed: (env) => env.c.objects.inkHidden,
      run: (env) => env.c.objects.setInkHidden(!env.c.objects.inkHidden),
    },
    {
      id: "InkDeleteAllOnSheet",
      label: "Delete All Ink on Sheet",
      description: "Delete the ink on this sheet.",
      keytip: "S",
      enabled: (env) => env.c.objects.hasInk(),
      run: (env) => void env.c.objects.deleteAllInk(false),
    },
    {
      id: "InkDeleteAll",
      label: "Delete All Ink",
      description: "Delete the ink on every sheet of the workbook.",
      keytip: "A",
      enabled: (env) => env.c.objects.hasInk(true),
      run: (env) => void env.c.objects.deleteAllInk(true),
    },
  ]
}
