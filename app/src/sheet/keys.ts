import type { KeyInput } from "./controller"

/** The platform's primary modifier is Ctrl, or ⌘ on macOS. */
export function keyInputFrom(
  event: { key: string; ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; altKey: boolean },
  os: string,
): KeyInput {
  return {
    key: event.key,
    primary: os === "macos" ? event.metaKey : event.ctrlKey,
    shift: event.shiftKey,
    alt: event.altKey,
  }
}
