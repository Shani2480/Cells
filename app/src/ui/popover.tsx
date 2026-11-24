import { Popover as PopoverPrimitive } from "@base-ui/react/popover"
import type { ReactElement, ReactNode } from "react"
import { cn } from "@/lib/cn"

// Ported from Inkspire (src/components/ui/popover.tsx at acbb14f): Base UI's popover, with the app's popover tokens
// in place of shadcn's classes.

export function Popover({
  open,
  onOpenChange,
  trigger,
  children,
  className,
  side = "bottom",
  align = "start",
}: {
  open: boolean
  onOpenChange: (open: boolean) => void
  trigger: ReactElement
  children: ReactNode
  className?: string
  side?: "bottom" | "right" | "top" | "left"
  align?: "start" | "center" | "end"
}) {
  return (
    <PopoverPrimitive.Root open={open} onOpenChange={(next) => onOpenChange(next)}>
      <PopoverPrimitive.Trigger render={trigger} />
      <PopoverPrimitive.Portal>
        <PopoverPrimitive.Positioner side={side} align={align} sideOffset={1} className="popover-positioner">
          <PopoverPrimitive.Popup className={cn("popover", className)} initialFocus={false} finalFocus={returnFocus}>
            {children}
          </PopoverPrimitive.Popup>
        </PopoverPrimitive.Positioner>
      </PopoverPrimitive.Portal>
    </PopoverPrimitive.Root>
  )
}

/**
 * A closed popover gives the focus back to its button, unless something else has taken it meanwhile: a ribbon
 * command run from a menu puts the keyboard back on the sheet (so Esc then reaches it), and that stays.
 */
function returnFocus(): boolean {
  const active = document.activeElement
  return !active || active === document.body || !!active.closest(".popover")
}
