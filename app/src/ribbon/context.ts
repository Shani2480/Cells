import { createContext, useContext } from "react"
import type { CommandRegistry, RibbonEnv } from "./registry"

/** What every ribbon control needs: the commands, the env they run in, and the one open menu. */
export type RibbonState = {
  env: RibbonEnv
  registry: CommandRegistry
  /** The id of the control whose menu is open (one at a time, so a KeyTip can open one too). */
  openMenu: string | null
  setOpenMenu: (id: string | null) => void
  /** KeyTips to show on controls, by command id (Alt was pressed and a tab chosen). */
  keytips: Map<string, string> | null
  /** Called after a command runs from the ribbon (hides a collapsed ribbon, returns focus to the sheet). */
  afterRun: () => void
}

export const RibbonCtx = createContext<RibbonState | null>(null)

export function useRibbon(): RibbonState {
  const state = useContext(RibbonCtx)
  if (!state) throw new Error("ribbon controls need a <Ribbon>")
  return state
}
