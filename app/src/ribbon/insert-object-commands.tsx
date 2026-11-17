import { isChart } from "@cells/core"
import type { Command, RibbonEnv } from "./registry"

/** Insert › Text: Signature Line (and later Object), the objects of Part 8b. */

/** The selected signature line's id, if one is selected. */
function selectedSignature(env: RibbonEnv): number | null {
  const o = env.c.objects.active()
  return o && !isChart(o) && o.kind === "picture" && o.signature ? o.id : null
}

export function insertObjectCommands(): Command[] {
  return [
    {
      id: "SignatureLineInsertMenu",
      label: "Signature Line",
      description: "Insert a signature line that specifies who must sign.",
      keytip: "G",
      menu: ["SignatureLineInsert", "SignatureLineSetup"],
    },
    {
      id: "OleObjectctInsert",
      label: "Object",
      description: "Embed a file in the workbook, shown as an icon you can double-click to open it.",
      keytip: "J",
      run: (env) => env.dialog("object"),
    },
    {
      id: "SignatureLineInsert",
      label: "Microsoft Office Signature Line...",
      description: "Insert a signature line: who should sign, their title, and instructions for them.",
      run: (env) => env.dialog("signatureSetup"),
    },
    {
      id: "SignatureLineSetup",
      label: "Signature Setup...",
      description: "Change who the selected signature line asks to sign. You can also double-click the line.",
      enabled: (env) => selectedSignature(env) !== null,
      run: (env) => {
        const id = selectedSignature(env)
        if (id !== null) env.dialog(`signatureSetup:${id}`)
      },
    },
  ]
}
