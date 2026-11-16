import { describe, expect, it } from "vitest"
import { allCommands } from "./commands"

describe("the command list", () => {
  it("defines each id once, so no stub replaces a working command", () => {
    const seen = new Map<string, number>()
    for (const command of allCommands()) seen.set(command.id, (seen.get(command.id) ?? 0) + 1)
    expect([...seen].filter(([, n]) => n > 1).map(([id]) => id)).toEqual([])
  })
})
