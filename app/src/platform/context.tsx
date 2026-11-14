import type { Platform } from "@cells/platform"
import { createContext, useContext } from "react"
import { webPlatform } from "./web"

export const PlatformContext = createContext<Platform>(webPlatform)

export function usePlatform() {
  return useContext(PlatformContext)
}
