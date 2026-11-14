import { Tooltip as TooltipPrimitive } from "@base-ui/react/tooltip"
import type { ReactElement, ReactNode } from "react"

// Ported from Inkspire (src/components/ui/tooltip.tsx at acbb14f, office variant only), with its fixed colours
// replaced by the app's tokens so it themes in dark mode.

/** Office tooltips wait a moment before showing, then follow the pointer from control to control at once. */
export function TooltipProvider({ children }: { children: ReactNode }) {
  return (
    <TooltipPrimitive.Provider delay={450} closeDelay={0}>
      {children}
    </TooltipPrimitive.Provider>
  )
}

/** Office's rich tooltip: the name in bold, a one-line description, the shortcut. */
export function Tip({
  label,
  description,
  shortcut,
  note,
  side = "bottom",
  children,
  ...trigger
}: {
  label: string
  description?: string
  shortcut?: string
  /** A second paragraph, e.g. "Not built yet". */
  note?: string
  side?: "top" | "bottom" | "left" | "right"
  children: ReactElement
  /** Props from an outer trigger (a popover's), passed on to the element: <Popover trigger={<Tip>…</Tip>} />. */
  [prop: string]: unknown
}) {
  return (
    <TooltipPrimitive.Root>
      <TooltipPrimitive.Trigger render={children} {...trigger} />
      <TooltipPrimitive.Portal>
        <TooltipPrimitive.Positioner side={side} sideOffset={8} className="tip-positioner">
          <TooltipPrimitive.Popup className="tip">
            <span className="tip-label">
              {label}
              {shortcut ? <span className="tip-shortcut"> ({shortcut})</span> : null}
            </span>
            {description ? <span className="tip-text">{description}</span> : null}
            {note ? <span className="tip-note">{note}</span> : null}
          </TooltipPrimitive.Popup>
        </TooltipPrimitive.Positioner>
      </TooltipPrimitive.Portal>
    </TooltipPrimitive.Root>
  )
}
