import type { PictureAdjust } from "@cells/core"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

export type FormatPicturePart = "corrections" | "color" | "transparency"

type Field = {
  key: "sharpen" | "bright" | "contrast" | "saturation" | "temperature" | "transparency"
  label: string
  min: number
  max: number
  /** Shown value per stored value (100 for percentages). */
  scale: number
  unit: string
  neutral: number
  corrections: boolean
}

const FIELDS: Record<FormatPicturePart, Field[]> = {
  corrections: [
    { key: "sharpen", label: "Sharpness", min: -100, max: 100, scale: 100, unit: "%", neutral: 0, corrections: true },
    { key: "bright", label: "Brightness", min: -100, max: 100, scale: 100, unit: "%", neutral: 0, corrections: true },
    { key: "contrast", label: "Contrast", min: -100, max: 100, scale: 100, unit: "%", neutral: 0, corrections: true },
  ],
  color: [
    { key: "saturation", label: "Saturation", min: 0, max: 400, scale: 100, unit: "%", neutral: 1, corrections: true },
    {
      key: "temperature",
      label: "Temperature",
      min: 1500,
      max: 11500,
      scale: 1,
      unit: " K",
      neutral: 6500,
      corrections: true,
    },
  ],
  transparency: [
    {
      key: "transparency",
      label: "Transparency",
      min: 0,
      max: 100,
      scale: 100,
      unit: "%",
      neutral: 0,
      corrections: false,
    },
  ],
}

const TITLES: Record<FormatPicturePart, string> = {
  corrections: "Picture Corrections",
  color: "Picture Color",
  transparency: "Picture Transparency",
}

/**
 * Format Picture's Picture Corrections, Picture Color and Picture Transparency: each setting as a slider and a
 * number, as Excel's pane has them; OK applies them as one undo step.
 */
export function FormatPictureDialog({
  controller,
  part,
  onClose,
}: {
  controller: SheetController
  part: FormatPicturePart
  onClose: () => void
}) {
  const picture = controller.objects.adjustable()
  const fields = FIELDS[part]
  const locked = !controller.objects.canCorrect(picture)
  const [values, setValues] = useState(() =>
    Object.fromEntries(
      fields.map((f) => [f.key, Math.round(((picture?.adjust?.[f.key] as number | undefined) ?? f.neutral) * f.scale)]),
    ),
  )
  if (!picture) return null
  const set = (key: string, v: number) => setValues((old) => ({ ...old, [key]: v }))
  const accept = () => {
    const change: Partial<PictureAdjust> = {}
    for (const f of fields) {
      const v = Math.min(f.max, Math.max(f.min, Number(values[f.key]) || 0)) / f.scale
      change[f.key] = v
    }
    void controller.objects.setAdjust(change, TITLES[part])
    onClose()
  }
  return (
    <Dialog title={TITLES[part]} width={360} onCancel={onClose} onAccept={accept}>
      {locked && part !== "transparency" ? (
        <p className="dialog-note">
          Excel's corrections on this picture are kept as they are: its original is in a form Cells can't open.
        </p>
      ) : null}
      <div className="format-picture-fields">
        {fields.map((f, i) => {
          const disabled = locked && f.corrections
          const id = `format-picture-${f.key}`
          return (
            <div key={f.key} className="format-picture-row">
              <label htmlFor={id}>{f.label}</label>
              <input
                type="range"
                aria-label={f.label}
                min={f.min}
                max={f.max}
                step={f.scale === 1 ? 100 : 1}
                value={values[f.key]}
                disabled={disabled}
                onChange={(e) => set(f.key, Number(e.target.value))}
              />
              <span className="format-picture-number">
                <input
                  id={id}
                  className="dialog-field"
                  inputMode="numeric"
                  data-autofocus={i === 0 ? true : undefined}
                  value={values[f.key]}
                  disabled={disabled}
                  onChange={(e) => set(f.key, Number(e.target.value.replace(/[^\d-]/g, "")))}
                />
                {f.unit.trim()}
              </span>
            </div>
          )
        })}
      </div>
    </Dialog>
  )
}
