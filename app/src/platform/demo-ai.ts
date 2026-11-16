import type { AiPort, Platform } from "@cells/platform"

/**
 * A stand-in AI for the browser dev build (`?ai=demo`), so the AI tools can be tried and pictured without a model:
 * canned answers, streamed slowly. Never part of a production build (main.tsx only uses it under `import.meta.env.DEV`).
 */
const demoAi: AiPort = {
  available: true,
  async generate(request, onText, signal) {
    const reply = request.system.includes('"formula"')
      ? '{"formula": "=SUMIF(A2:A6,\\"Rent\\",B2:B6)", "explanation": "Adds the costs in column B on rows where column A says Rent."}'
      : request.system.includes('"chart"')
        ? '{"chart": "column", "title": "Costs by month", "why": "A few categories compared side by side read best as columns."}'
        : request.prompt.startsWith("Explain")
          ? "Works out the total cost for the quarter.\n- SUM adds every number in its range.\n- B2:B6 is the Cost column, January to May.\nIt will miss rows added below row 6."
          : "- Total costs are 6,130 over five months.\n- Rent is the largest cost every month (1,200).\n- Water is the smallest, between 38 and 52.\n- Costs rose 8% from January to May."
    let sent = ""
    for (const word of reply.split(/(?<= )/)) {
      if (signal?.aborted) throw new Error("Stopped.")
      await new Promise((resolve) => setTimeout(resolve, 40))
      sent += word
      onText(word)
    }
    return sent
  },
  test: async () => "Connected to the demo AI.",
  keyStatus: async () => ({ openai: "missing", anthropic: "saved", compatible: "missing", storage: "keyring" }),
  saveKey: async (_provider, key) => (key ? "keyring" : "none"),
}

export const withDemoAi = (platform: Platform): Platform => ({ ...platform, ai: demoAi })
