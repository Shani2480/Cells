import type { AiPort, Platform } from "@cells/platform"

/**
 * A stand-in AI for the browser dev build (`?ai=demo`), so the AI tools can be tried and pictured without a model:
 * canned answers, streamed slowly. Never part of a production build (main.tsx only uses it under `import.meta.env.DEV`).
 */
const demoAi: AiPort = {
  available: true,
  async generate(request, onText, signal) {
    const reply = request.system.includes('"group_dates"')
      ? demoPivot(request.context, request.prompt)
      : request.system.includes('"formula"')
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

/** A PivotTable layout from the columns listed: the first text column by the first number, biggest first. */
function demoPivot(context: string, prompt: string): string {
  const columns = [...context.matchAll(/^- (.+?): (text|numbers|dates)/gm)].map((m) => ({ name: m[1], kind: m[2] }))
  const text = columns.find((c) => c.kind === "text")?.name
  const number = columns.find((c) => c.kind === "numbers")?.name
  const date = columns.find((c) => c.kind === "dates")?.name
  if (!text && !date) return '{"problem": "The data has no column to group by."}'
  const rows = text ? [text] : [date]
  const per = /quarter/i.test(prompt) ? "quarters" : /month/i.test(prompt) ? "months" : null
  return JSON.stringify({
    rows,
    columns: per && date && text ? [date] : [],
    values: number ? [{ field: number, summary: "sum" }] : [{ field: rows[0], summary: "count" }],
    ...(per && date ? { group_dates: { field: date, by: [per, "years"] } } : {}),
    sort: { field: rows[0], by: number ?? rows[0], order: "descending" },
  })
}

export const withDemoAi = (platform: Platform): Platform => ({ ...platform, ai: demoAi })
