import React from "react"
import ReactDOM from "react-dom/client"
import { App } from "./App"
import { PlatformContext } from "./platform/context"
import { isTauri, tauriPlatform } from "./platform/tauri"
import { webPlatform } from "./platform/web"
import "./styles/index.css"

// The browser's own context menu never shows over the app (ours does).
window.addEventListener("contextmenu", (event) => {
  if (!(event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement)) event.preventDefault()
})

const web =
  import.meta.env.DEV && new URLSearchParams(window.location.search).get("ai") === "demo"
    ? (await import("./platform/demo-ai")).withDemoAi(webPlatform)
    : webPlatform
const platform = isTauri() ? tauriPlatform() : web

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <PlatformContext.Provider value={platform}>
      <App />
    </PlatformContext.Provider>
  </React.StrictMode>,
)
