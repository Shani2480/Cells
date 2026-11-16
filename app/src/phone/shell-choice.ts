import { useEffect, useState } from "react"

/**
 * Which window to show: the phone shell or the full one (title bar, ribbon, sheet tabs). By the screen, not the OS:
 * the phone shell when the window's smaller side is under 600 CSS px (Android's dp) and the main pointer is a finger.
 * An Android tablet gets the full ribbon, a phone in landscape keeps the phone shell, and a tablet's split screen can
 * switch live. Both shells work on the same workbook and dialogs, so a switch loses nothing.
 */
export type ShellKind = "phone" | "full"

/** Android's own line between phones and tablets (smallest width 600 dp). */
export const PHONE_MAX_SIDE = 600

export function shellFor(view: { width: number; height: number; coarse: boolean }): ShellKind {
  return view.coarse && Math.min(view.width, view.height) < PHONE_MAX_SIDE ? "phone" : "full"
}

/**
 * The window's height without the on-screen keyboard: on Android the page shrinks while the keyboard is up, and
 * MainActivity says by how much. Typing on a tablet in landscape must not turn it into a phone.
 */
export function heightWithoutKeyboard(innerHeight: number, keyboardHeight: unknown): number {
  return innerHeight + (typeof keyboardHeight === "number" && keyboardHeight > 0 ? keyboardHeight : 0)
}

function current(): ShellKind {
  if (typeof window === "undefined") return "full"
  const coarse = typeof matchMedia === "function" && matchMedia("(pointer: coarse)").matches
  const keyboard = (window as { __cellsKeyboardHeight?: unknown }).__cellsKeyboardHeight
  return shellFor({ width: window.innerWidth, height: heightWithoutKeyboard(window.innerHeight, keyboard), coarse })
}

/** The shell for the window now, following resizes, rotation and a mouse being plugged in. */
export function useShellKind(): ShellKind {
  const [kind, setKind] = useState(current)
  useEffect(() => {
    const update = () => setKind(current())
    window.addEventListener("resize", update)
    const pointer = typeof matchMedia === "function" ? matchMedia("(pointer: coarse)") : null
    pointer?.addEventListener("change", update)
    return () => {
      window.removeEventListener("resize", update)
      pointer?.removeEventListener("change", update)
    }
  }, [])
  return kind
}
