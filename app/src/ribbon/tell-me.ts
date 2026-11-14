import type { Command, CommandRegistry, RibbonEnv } from "./registry"

/**
 * Search (Alt+Q): commands whose name or description has every typed word, names first. Commands not built yet are
 * found too and shown disabled, as Excel shows unavailable ones.
 */
export function searchCommands(registry: CommandRegistry, env: RibbonEnv, query: string, limit = 12): Command[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean)
  if (!words.length) return []
  const scored: { command: Command; score: number }[] = []
  const seen = new Set<string>()
  for (const command of registry.all()) {
    // Duplicates (the same command on two tabs under two ids) show once, by name.
    const name = (command.labelFor?.(env) ?? command.label).replace(/\.\.\.$/, "")
    const label = name.toLowerCase()
    const text = `${label} ${command.description?.toLowerCase() ?? ""}`
    if (!words.every((w) => text.includes(w))) continue
    if (seen.has(label)) continue
    seen.add(label)
    const score =
      (label === words.join(" ") ? 0 : label.startsWith(words[0]) ? 1 : words.every((w) => label.includes(w)) ? 2 : 3) +
      (registry.isBuilt(command.id) ? 0 : 4)
    scored.push({ command, score })
  }
  return scored
    .sort((a, b) => a.score - b.score || a.command.label.localeCompare(b.command.label))
    .slice(0, limit)
    .map((s) => s.command)
}
