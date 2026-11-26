import { useEffect, useState } from "react"
import type { HeaderPicture } from "@cells/core"
import { pictureSize } from "@/drawing/pictures"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"
import { Check, NumberBox } from "./PageSetupDialog"

type Tab = "size" | "picture"

/** Image control › Color, with Washout (brightness 85%, contrast 15%) as Excel lists it. */
type ColorChoice = HeaderPicture["color"] | "washout"
const WASHOUT = { brightness: 0.85, contrast: 0.15 }

/**
 * Header & Footer › Format Picture, as Excel's: Size (height and width, or scale, with the aspect ratio locked if
 * asked; the picture's own size and Reset) and Picture (cropped from each side; colour, brightness and contrast).
 * Sizes in inches, crops of the picture at its own size.
 */
export function HeaderPictureDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const start = controller.headerFooter.currentPicture()
  const [tab, setTab] = useState<Tab>("size")
  // The picture's own size in points (from its pixels at 96 dpi), once known.
  const [own, setOwn] = useState<{ w: number; h: number } | null>(null)
  useEffect(() => {
    if (!start) return
    let live = true
    void pictureSize(start.src).then(
      (s) => live && setOwn({ w: s.w * 0.75, h: s.h * 0.75 }),
      () => live && setOwn({ w: start.width, h: start.height }),
    )
    return () => {
      live = false
    }
  }, [start])
  const [pic, setPic] = useState<HeaderPicture | null>(start)
  if (!start || !pic) return null

  const crop = pic.crop
  const shownW = (own?.w ?? pic.width) * Math.max(0.01, 1 - crop.left - crop.right)
  const shownH = (own?.h ?? pic.height) * Math.max(0.01, 1 - crop.top - crop.bottom)
  const scaleW = (pic.width / shownW) * 100
  const scaleH = (pic.height / shownH) * 100
  const inches = (pt: number) => pt / 72

  const setSize = (width: number, height: number, from: "w" | "h") => {
    if (pic.lockAspect) {
      // Locked: the other side follows by the same scale.
      if (from === "w") height = (width / shownW) * shownH
      else width = (height / shownH) * shownW
    }
    setPic({ ...pic, width: Math.max(1, width), height: Math.max(1, height) })
  }
  const setCrop = (side: keyof HeaderPicture["crop"], inch: number) => {
    const ownSide = side === "left" || side === "right" ? (own?.w ?? shownW) : (own?.h ?? shownH)
    const next = { ...crop, [side]: Math.min(0.95, Math.max(0, (inch * 72) / ownSide)) }
    // The shown size keeps its scale as the crop changes.
    const kw = pic.width / shownW
    const kh = pic.height / shownH
    const w = (own?.w ?? shownW) * Math.max(0.01, 1 - next.left - next.right) * kw
    const h = (own?.h ?? shownH) * Math.max(0.01, 1 - next.top - next.bottom) * kh
    setPic({ ...pic, crop: next, width: w, height: h })
  }
  const color: ColorChoice =
    pic.color === "automatic" && pic.brightness === WASHOUT.brightness && pic.contrast === WASHOUT.contrast
      ? "washout"
      : pic.color
  const setColor = (c: ColorChoice) =>
    setPic(
      c === "washout"
        ? { ...pic, color: "automatic", ...WASHOUT }
        : { ...pic, color: c, ...(color === "washout" ? { brightness: 0.5, contrast: 0.5 } : {}) },
    )

  return (
    <Dialog
      title="Format Picture"
      width={440}
      onCancel={onClose}
      onAccept={() => {
        controller.headerFooter.formatPicture(pic)
        onClose()
      }}
    >
      <div className="fc-tabs" role="tablist">
        {(["size", "picture"] as const).map((t) => (
          <button
            key={t}
            type="button"
            role="tab"
            aria-selected={tab === t}
            className="fc-tab"
            onClick={() => setTab(t)}
          >
            {t === "size" ? "Size" : "Picture"}
          </button>
        ))}
      </div>
      <div className="fc-page hp-page" role="tabpanel" aria-label={tab === "size" ? "Size" : "Picture"}>
        {tab === "size" ? (
          <div className="ps-column">
            <fieldset className="dialog-group">
              <legend>Size and rotate</legend>
              <div className="ps-row">
                <span className="hp-label">Height:</span>
                <NumberBox
                  label="Height"
                  value={inches(pic.height)}
                  min={0.01}
                  max={22}
                  step={0.01}
                  onChange={(v) => setSize(pic.width, v * 72, "h")}
                />
                <span className="hp-label">Width:</span>
                <NumberBox
                  label="Width"
                  value={inches(pic.width)}
                  min={0.01}
                  max={22}
                  step={0.01}
                  onChange={(v) => setSize(v * 72, pic.height, "w")}
                />
              </div>
            </fieldset>
            <fieldset className="dialog-group">
              <legend>Scale</legend>
              <div className="ps-row">
                <span className="hp-label">Height:</span>
                <NumberBox
                  label="Scale height"
                  value={Math.round(scaleH)}
                  min={1}
                  max={9999}
                  onChange={(v) => setSize(pic.width, (v / 100) * shownH, "h")}
                />
                <span>%</span>
                <span className="hp-label">Width:</span>
                <NumberBox
                  label="Scale width"
                  value={Math.round(scaleW)}
                  min={1}
                  max={9999}
                  onChange={(v) => setSize((v / 100) * shownW, pic.height, "w")}
                />
                <span>%</span>
              </div>
              <Check
                label="Lock aspect ratio"
                checked={pic.lockAspect}
                onChange={(v) => setPic({ ...pic, lockAspect: v })}
              />
            </fieldset>
            <fieldset className="dialog-group">
              <legend>Original size</legend>
              <div className="ps-row">
                <span>
                  {own ? `Height: ${inches(own.h).toFixed(2)}"   Width: ${inches(own.w).toFixed(2)}"` : "Measuring…"}
                </span>
                <button
                  type="button"
                  className="dialog-btn"
                  disabled={!own}
                  onClick={() =>
                    own &&
                    setPic({ ...pic, width: own.w, height: own.h, crop: { top: 0, bottom: 0, left: 0, right: 0 } })
                  }
                >
                  Reset
                </button>
              </div>
            </fieldset>
          </div>
        ) : (
          <div className="ps-column">
            <fieldset className="dialog-group">
              <legend>Crop from</legend>
              <div className="hp-grid">
                {(["left", "right", "top", "bottom"] as const).map((side) => (
                  <label key={side} className="ps-row">
                    <span className="hp-label">{side[0].toUpperCase() + side.slice(1)}:</span>
                    <NumberBox
                      label={`Crop from ${side}`}
                      value={inches(crop[side] * ((side === "left" || side === "right" ? own?.w : own?.h) ?? 0))}
                      min={0}
                      max={22}
                      step={0.01}
                      disabled={!own}
                      onChange={(v) => setCrop(side, v)}
                    />
                  </label>
                ))}
              </div>
            </fieldset>
            <fieldset className="dialog-group">
              <legend>Image control</legend>
              <label className="ps-line">
                <span className="dialog-label">Color:</span>
                <select
                  className="dialog-field"
                  aria-label="Color"
                  value={color}
                  onChange={(e) => setColor(e.target.value as ColorChoice)}
                >
                  <option value="automatic">Automatic</option>
                  <option value="grayscale">Grayscale</option>
                  <option value="blackAndWhite">Black &amp; White</option>
                  <option value="washout">Washout</option>
                </select>
              </label>
              <label className="ps-line">
                <span className="dialog-label">Brightness:</span>
                <span className="ps-row">
                  <NumberBox
                    label="Brightness"
                    value={Math.round(pic.brightness * 100)}
                    min={0}
                    max={100}
                    onChange={(v) => setPic({ ...pic, brightness: v / 100 })}
                  />
                  <span>%</span>
                </span>
              </label>
              <label className="ps-line">
                <span className="dialog-label">Contrast:</span>
                <span className="ps-row">
                  <NumberBox
                    label="Contrast"
                    value={Math.round(pic.contrast * 100)}
                    min={0}
                    max={100}
                    onChange={(v) => setPic({ ...pic, contrast: v / 100 })}
                  />
                  <span>%</span>
                </span>
              </label>
              <div className="ps-row">
                <button
                  type="button"
                  className="dialog-btn"
                  onClick={() =>
                    setPic({
                      ...pic,
                      color: "automatic",
                      brightness: 0.5,
                      contrast: 0.5,
                      crop: { top: 0, bottom: 0, left: 0, right: 0 },
                      width: own ? own.w * (pic.width / shownW) : pic.width,
                      height: own ? own.h * (pic.height / shownH) : pic.height,
                    })
                  }
                >
                  Reset
                </button>
              </div>
            </fieldset>
          </div>
        )}
      </div>
    </Dialog>
  )
}
