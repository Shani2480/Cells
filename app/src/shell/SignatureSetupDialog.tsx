import { isChart, SIGNING_INSTRUCTIONS, type SignatureSetup } from "@cells/core"
import { useState } from "react"
import { svgToPng } from "@/drawing/pictures"
import { SIGNATURE_SIZE, signatureDocument, signatureId } from "@/drawing/signature"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * Insert › Signature Line (and Signature Setup on a line already there): who should sign, as Office's dialog asks.
 * OK puts the line at the active cell, or redraws the line being set up.
 */
export function SignatureSetupDialog({
  controller,
  editing,
  onClose,
}: {
  controller: SheetController
  /** The line whose setup changes; absent for a new line. */
  editing?: number
  onClose: () => void
}) {
  const found = editing === undefined ? undefined : controller.objects.find(editing)
  const before = found && !isChart(found) && found.kind === "picture" ? found.signature : undefined
  const [signer, setSigner] = useState(before?.signer ?? "")
  const [title, setTitle] = useState(before?.title ?? "")
  const [email, setEmail] = useState(before?.email ?? "")
  const [instructions, setInstructions] = useState(before?.instructions ?? SIGNING_INSTRUCTIONS)
  const [allowComments, setAllowComments] = useState(before?.allowComments ?? false)
  const [showDate, setShowDate] = useState(before?.showDate ?? true)

  const accept = async () => {
    const setup: SignatureSetup = {
      id: before?.id ?? signatureId(),
      provider: before?.provider,
      signer: signer.trim(),
      title: title.trim(),
      email: email.trim(),
      instructions,
      allowComments,
      showDate,
    }
    const size = SIGNATURE_SIZE
    const markup = signatureDocument(setup, size.w, size.h, controller.wb.theme)
    let png = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(markup)))}`
    try {
      png = await svgToPng(markup, size.w * 2, size.h * 2)
    } catch {
      // Without a canvas the SVG stands in for the picture files carry.
    }
    if (editing !== undefined && before) controller.objects.setSignature(editing, setup, png)
    else controller.objects.insertSignatureLine(setup, png, size)
    onClose()
  }

  const field = (id: string, label: React.ReactNode, value: string, set: (v: string) => void, autofocus = false) => (
    <>
      <label className="dialog-label" htmlFor={id}>
        {label}
      </label>
      <input
        id={id}
        className="dialog-field"
        value={value}
        data-autofocus={autofocus || undefined}
        onChange={(e) => set(e.target.value)}
      />
    </>
  )

  return (
    <Dialog title="Signature Setup" width={400} onCancel={onClose} onAccept={() => void accept()}>
      <div className="signature-setup">
        {field(
          "sig-signer",
          <>
            Suggested <u>s</u>igner (for example, John Doe):
          </>,
          signer,
          setSigner,
          true,
        )}
        {field(
          "sig-title",
          <>
            Suggested signer's <u>t</u>itle (for example, Manager):
          </>,
          title,
          setTitle,
        )}
        {field(
          "sig-email",
          <>
            Suggested signer's <u>e</u>-mail address:
          </>,
          email,
          setEmail,
        )}
        <label className="dialog-label" htmlFor="sig-instructions">
          <u>I</u>nstructions to the signer:
        </label>
        <textarea
          id="sig-instructions"
          className="dialog-field"
          rows={3}
          value={instructions}
          onChange={(e) => setInstructions(e.target.value)}
        />
        <label className="fc-check">
          <input type="checkbox" checked={allowComments} onChange={(e) => setAllowComments(e.target.checked)} />
          <span>
            <u>A</u>llow the signer to add comments in the Sign dialog
          </span>
        </label>
        <label className="fc-check">
          <input type="checkbox" checked={showDate} onChange={(e) => setShowDate(e.target.checked)} />
          <span>
            Show sign <u>d</u>ate in signature line
          </span>
        </label>
      </div>
    </Dialog>
  )
}
