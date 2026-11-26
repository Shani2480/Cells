import { tesseractPicture, tesseractRecognizer } from "./pdf/ocr"
import { androidOcr, mlkitPicture, mlkitRecognizer } from "./pdf/ocr-android"
import { setClipboardImageReader } from "./pdf/picture"
import { clipboardImageReader } from "./platform/clipboard-image"
import { setPictureReader } from "./sheet/picture-data"
import { fileName, setFileReader, setRecognitionSwitch, setRecognizer, setWebPort } from "./sheet/get-data"
import { updateSettings, useSettings } from "./settings/settings"
import React from "react"
import ReactDOM from "react-dom/client"
import { App } from "./App"
import { PlatformContext } from "./platform/context"
import { isTauri, tauriPlatform } from "./platform/tauri"
import { webPlatform } from "./platform/web"
import "./styles/index.css"
import "./styles/phone.css"

// The browser's own context menu never shows over the app (ours does).
window.addEventListener("contextmenu", (event) => {
  if (!(event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement)) event.preventDefault()
})

const web =
  import.meta.env.DEV && new URLSearchParams(window.location.search).get("ai") === "demo"
    ? (await import("./platform/demo-ai")).withDemoAi(webPlatform)
    : webPlatform
const platform = isTauri() ? tauriPlatform() : web
setWebPort(platform.web)
setFileReader((path) => platform.files.read({ name: fileName(path), path }))
setClipboardImageReader(clipboardImageReader(isTauri(), platform.os))
// Text recognition for scanned PDFs follows its setting (opt-in); a prompt can turn it on, which saves the setting.
// On Android ML Kit reads them (on the device); elsewhere Tesseract in the web view.
const recognition = (on: boolean) => {
  setRecognizer(on ? (androidOcr() ? mlkitRecognizer : tesseractRecognizer) : null)
  setPictureReader(on ? (androidOcr() ? mlkitPicture : tesseractPicture) : null)
}
recognition(useSettings.getState().ocrEnabled)
useSettings.subscribe((s, before) => s.ocrEnabled !== before.ocrEnabled && recognition(s.ocrEnabled))
setRecognitionSwitch(() => void updateSettings({ ocrEnabled: true }), androidOcr() ? "device" : "computer")

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <PlatformContext.Provider value={platform}>
      <App />
    </PlatformContext.Provider>
  </React.StrictMode>,
)
