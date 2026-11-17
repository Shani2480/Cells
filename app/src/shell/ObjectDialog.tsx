import type { EmbeddedFile } from "@cells/core"
import { useState } from "react"
import { embedDocument, OBJECT_ICON_SIZE } from "@/drawing/embed"
import { svgToPng } from "@/drawing/pictures"
import { usePlatform } from "@/platform/context"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** The OLE class Office gives an embedded file: Office documents as themselves, anything else as a Package. */
export function progIdFor(name: string): string {
  const ext = /\.([^.]+)$/.exec(name)?.[1]?.toLowerCase() ?? ""
  if (ext === "docx") return "Word.Document.12"
  if (ext === "xlsx") return "Excel.Sheet.12"
  if (ext === "pptx") return "PowerPoint.Show.12"
  return "Package"
}

/** The icon picture files carry for an embedded file (its SVG, when there's no canvas to draw a PNG). */
export async function objectPicture(embed: EmbeddedFile, controller: SheetController): Promise<string> {
  const { w, h } = OBJECT_ICON_SIZE
  const markup = embedDocument(embed, w, h, controller.wb.theme)
  try {
    return await svgToPng(markup, w * 2, h * 2)
  } catch {
    return `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(markup)))}`
  }
}

const nameOf = (path: string) => path.split(/[\\/]/).pop() ?? path

/**
 * Insert › Object, Create from File: a file put in the workbook as an icon. Create New needs the other app's object
 * server, which Cells doesn't have; Link to file isn't built yet.
 */
export function ObjectDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const platform = usePlatform()
  const [file, setFile] = useState<{ name: string; path?: string; bytes?: Uint8Array } | null>(null)
  const [problem, setProblem] = useState("")

  const browse = async () => {
    setProblem("")
    if (platform.files.pickPath) {
      const path = await platform.files.pickPath("Browse")
      if (path) setFile({ name: nameOf(path), path })
      return
    }
    const input = document.createElement("input")
    input.type = "file"
    input.onchange = async () => {
      const f = input.files?.[0]
      if (f) setFile({ name: f.name, bytes: new Uint8Array(await f.arrayBuffer()) })
    }
    input.click()
  }

  const accept = async () => {
    if (!file) return
    let data = file.bytes
    if (!data && file.path) {
      try {
        data = await platform.files.read({ name: file.name, path: file.path })
      } catch {
        setProblem(`Cells can't read ${file.name}.`)
        return
      }
    }
    const embed: EmbeddedFile = { name: file.name, progId: progIdFor(file.name), data }
    controller.objects.insertObject(embed, await objectPicture(embed, controller), OBJECT_ICON_SIZE)
    onClose()
  }

  return (
    <Dialog title="Object" width={440} onCancel={onClose} onAccept={() => void accept()}>
      <div className="fc-tabs" role="tablist">
        <button
          type="button"
          role="tab"
          aria-selected={false}
          className="fc-tab"
          disabled
          title="Making a new object needs the other app's object server, which Cells doesn't have."
        >
          Create New
        </button>
        <button type="button" role="tab" aria-selected className="fc-tab">
          Create from File
        </button>
      </div>
      <div className="fc-page object-page" role="tabpanel" aria-label="Create from File">
        <label className="dialog-label" htmlFor="object-file">
          File <u>n</u>ame:
        </label>
        <div className="object-file-row">
          <input
            id="object-file"
            className="dialog-field"
            readOnly
            value={file ? (file.path ?? file.name) : ""}
            placeholder="Choose a file"
          />
          <button type="button" className="dialog-btn" data-autofocus onClick={() => void browse()}>
            <u>B</u>rowse...
          </button>
        </div>
        <div className="object-options">
          <label className="fc-check" title="Not built yet: Cells embeds a copy of the file.">
            <input type="checkbox" checked={false} disabled readOnly />
            <span>
              Lin<u>k</u> to file
            </span>
          </label>
          <label className="fc-check" title="Cells shows embedded files as icons.">
            <input type="checkbox" checked disabled readOnly />
            <span>
              Display as ic<u>o</u>n
            </span>
          </label>
        </div>
        <fieldset className="dialog-group">
          <legend>Result</legend>
          <p className="dialog-note">
            Inserts the contents of the file into your document as an icon, so that you can open it later in the program
            that created it.
          </p>
        </fieldset>
        {problem ? <p className="dialog-note is-error">{problem}</p> : null}
      </div>
    </Dialog>
  )
}
