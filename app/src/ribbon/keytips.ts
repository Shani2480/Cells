import type { CommandRegistry } from "./registry"
import { controlIds, type Tab } from "./tabs"

/**
 * The KeyTips of a tab's controls: each command's own letters (Excel's), and for controls without one, free
 * two-letter codes starting with Y (Excel does the same for controls it has no letter for). `typed` filters to the
 * KeyTips that start with what was typed so far.
 */
export function keytipsFor(tab: Tab, registry: CommandRegistry, typed: string): Map<string, string> {
  const all = new Map<string, string>()
  const used = new Set<string>()
  const ids = [...new Set(controlIds(tab))]
  for (const id of ids) {
    const key = registry.get(id)?.keytip
    if (key && !used.has(key)) {
      all.set(id, key)
      used.add(key)
    }
  }
  let n = 0
  for (const id of ids) {
    if (all.has(id)) continue
    let key: string
    do key = `Y${(n++).toString(36).toUpperCase()}`
    while (used.has(key))
    all.set(id, key)
    used.add(key)
  }
  if (!typed) return all
  return new Map([...all].filter(([, key]) => key.startsWith(typed)))
}
