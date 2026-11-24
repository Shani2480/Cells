import { ARTISTIC_EFFECTS, type ArtisticEffect, type PictureAdjust } from "@cells/core"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

export type FormatPicturePart = "corrections" | "color" | "transparency" | "artistic"

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

const FIELDS: Record<Exclude<FormatPicturePart, "artistic">, Field[]> = {
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
  artistic: "Artistic Effects",
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
  if (part === "artistic") return <ArtisticDialog controller={controller} onClose={onClose} />
  return <AdjustDialog controller={controller} part={part} onClose={onClose} />
}

/** A setting as a slider and a number. */
function SliderRow({
  id,
  label,
  min,
  max,
  step = 1,
  unit,
  value,
  disabled,
  autoFocus,
  onChange,
}: {
  id: string
  label: string
  min: number
  max: number
  step?: number
  unit: string
  value: number
  disabled?: boolean
  autoFocus?: boolean
  onChange: (v: number) => void
}) {
  return (
    <div className="format-picture-row">
      <label htmlFor={id}>{label}</label>
      <input
        type="range"
        aria-label={label}
        min={min}
        max={max}
        step={step}
        value={value}
        disabled={disabled}
        onChange={(e) => onChange(Number(e.target.value))}
      />
      <span className="format-picture-number">
        <input
          id={id}
          className="dialog-field"
          inputMode="numeric"
          data-autofocus={autoFocus ? true : undefined}
          value={value}
          disabled={disabled}
          onChange={(e) => onChange(Number(e.target.value.replace(/[^\d-]/g, "")))}
        />
        {unit.trim()}
      </span>
    </div>
  )
}

/** Artistic Effects Options: the effect, how much of the picture shows through it, and its own setting. */
function ArtisticDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const picture = controller.objects.adjustable()
  const now = picture?.adjust?.artistic
  const [effect, setEffect] = useState<ArtisticEffect | "">(now?.effect ?? "")
  const spec = effect ? ARTISTIC_EFFECTS[effect] : null
  const [trans, setTrans] = useState(() => Math.round((now?.trans ?? spec?.trans ?? 0) * 100))
  const [amount, setAmount] = useState(() => now?.amount ?? spec?.amount ?? 0)
  if (!picture) return null
  const locked = !controller.objects.canCorrect(picture)
  const choose = (next: ArtisticEffect | "") => {
    setEffect(next)
    // Another effect starts from its own defaults, as in Excel.
    const s = next ? ARTISTIC_EFFECTS[next] : null
    setTrans(Math.round((s?.trans ?? 0) * 100))
    setAmount(s?.amount ?? 0)
  }
  const accept = () => {
    const artistic =
      effect && spec
        ? {
            effect,
            ...(spec.trans !== null ? { trans: Math.min(100, Math.max(0, trans)) / 100 } : {}),
            amount: Math.min(spec.max, Math.max(0, amount)),
            ...(now?.last && now.effect === effect ? { last: true } : {}),
          }
        : undefined
    void controller.objects.setAdjust({ artistic }, "Artistic Effects")
    onClose()
  }
  return (
    <Dialog title="Artistic Effects" width={380} onCancel={onClose} onAccept={accept}>
      {locked ? (
        <p className="dialog-note">
          Excel's effects on this picture are kept as they are: its original is in a form Cells can't open, or it has
          more than one effect.
        </p>
      ) : null}
      <div className="format-picture-fields">
        <div className="format-picture-row">
          <label htmlFor="format-picture-effect">Effect</label>
          <select
            id="format-picture-effect"
            className="dialog-field format-picture-select"
            data-autofocus
            value={effect}
            disabled={locked}
            onChange={(e) => choose(e.target.value as ArtisticEffect | "")}
          >
            <option value="">None</option>
            {(Object.keys(ARTISTIC_EFFECTS) as ArtisticEffect[]).map((k) => (
              <option key={k} value={k}>
                {ARTISTIC_EFFECTS[k].label}
              </option>
            ))}
          </select>
        </div>
        {spec && spec.trans !== null ? (
          <SliderRow
            id="format-picture-art-trans"
            label="Transparency"
            min={0}
            max={100}
            unit="%"
            value={trans}
            disabled={locked}
            onChange={setTrans}
          />
        ) : null}
        {spec ? (
          <SliderRow
            id="format-picture-art-amount"
            label={spec.setting}
            min={0}
            max={spec.max}
            unit=""
            value={amount}
            disabled={locked}
            onChange={setAmount}
          />
        ) : null}
      </div>
    </Dialog>
  )
}

function AdjustDialog({
  controller,
  part,
  onClose,
}: {
  controller: SheetController
  part: Exclude<FormatPicturePart, "artistic">
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
        {fields.map((f, i) => (
          <SliderRow
            key={f.key}
            id={`format-picture-${f.key}`}
            label={f.label}
            min={f.min}
            max={f.max}
            step={f.scale === 1 ? 100 : 1}
            unit={f.unit}
            value={values[f.key]}
            disabled={locked && f.corrections}
            autoFocus={i === 0}
            onChange={(v) => set(f.key, v)}
          />
        ))}
      </div>
    </Dialog>
  )
}
