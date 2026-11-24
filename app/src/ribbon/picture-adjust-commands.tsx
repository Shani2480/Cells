import {
  ARTISTIC_EFFECTS,
  correctionsOf,
  type ArtisticEffect,
  type DuotoneColor,
  type PictureAdjust,
  type PictureRecolor,
  type SheetPicture,
  type ThemeSlot,
} from "@cells/core"
import { useEffect, useState } from "react"
import { objectSvg } from "@/drawing/render"
import { effectSide } from "@/sheet/objects"
import type { Command, MenuEntry, RibbonEnv } from "./registry"

/**
 * Picture Format › Adjust: Corrections, Color and Transparency, each a gallery of the selected picture as it would
 * look (Excel's presets), with its options in the Format Picture dialog.
 */

// ── Previews ──

const small = new Map<string, Promise<string>>()
const corrected = new Map<string, Promise<string>>()
const keyOf = (src: string) => `${src.length}:${src.slice(-64)}`

const PREVIEW_SIDE = 96

/**
 * The picture small (for the galleries), with corrections applied: made once per picture and setting. `unit` scales
 * an artistic effect to the small copy (pixels per pixel of the picture Office would work on).
 */
function preview(base: string, corrections: PictureAdjust | null, unit: number): Promise<string> {
  const smallKey = keyOf(base)
  let tiny = small.get(smallKey)
  if (!tiny) {
    tiny = import("@/drawing/picture-adjust").then((m) => m.correctPicture(base, {}, { maxSide: PREVIEW_SIDE }))
    small.set(smallKey, tiny)
  }
  if (!corrections) return tiny
  const key = `${smallKey}|${JSON.stringify(corrections)}|${unit}`
  let out = corrected.get(key)
  if (!out) {
    out = tiny.then((t) => import("@/drawing/picture-adjust").then((m) => m.correctPicture(t, corrections, { unit })))
    corrected.set(key, out)
  }
  return out
}

let previewIds = 0

/** One gallery tile: the selected picture with `adjust` instead of its own (cropped and shaped as it is). */
function PictureTile({ env, picture, adjust }: { env: RibbonEnv; picture: SheetPicture; adjust: PictureAdjust }) {
  const base = picture.original ?? picture.src
  // The settings as text: a new object each draw is the same preview.
  const settings = JSON.stringify(correctionsOf(adjust))
  const [src, setSrc] = useState<string | null>(null)
  const [id] = useState(() => --previewIds)
  const box = env.c.objects.box(picture)
  const natural = picture.natural ? Math.max(picture.natural.w, picture.natural.h) : effectSide(box)
  const unit = Math.round((Math.min(PREVIEW_SIDE, natural) / Math.min(natural, effectSide(box))) * 1000) / 1000
  useEffect(() => {
    let live = true
    preview(base, JSON.parse(settings) as PictureAdjust | null, unit)
      .then((s) => live && setSrc(s))
      .catch(() => live && setSrc(picture.src))
    return () => {
      live = false
    }
  }, [base, settings, unit, picture.src])
  const w = 64
  const h = Math.max(24, Math.min(64, Math.round((64 * box.h) / Math.max(1, box.w))))
  const shown: SheetPicture = { ...picture, id, src: src ?? "", adjust, line: undefined }
  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} aria-hidden className="picture-tile-svg">
      {src ? <g dangerouslySetInnerHTML={{ __html: objectSvg(shown, w, h, env.c.wb.theme) }} /> : null}
    </svg>
  )
}

type Preset = { label: string; change: Partial<PictureAdjust>; current: (a: PictureAdjust) => boolean }

/** A gallery section: a heading and its tiles, the one in use marked. */
function section(
  env: RibbonEnv,
  close: () => void,
  heading: string,
  presets: Preset[],
  columns: number,
  undo: string,
  disabled = false,
) {
  const p = env.c.objects.adjustable()
  if (!p) return null
  const now = p.adjust ?? {}
  return (
    <div key={heading}>
      <div className="menu-heading">{heading}</div>
      <div className="picture-gallery-grid" style={{ gridTemplateColumns: `repeat(${columns}, 72px)` }}>
        {presets.map((preset) => (
          <button
            key={preset.label}
            type="button"
            role="menuitem"
            className={`picture-gallery-item${preset.current(now) ? " is-current" : ""}`}
            title={preset.label}
            aria-label={preset.label}
            disabled={disabled}
            onClick={() => {
              close()
              void env.c.objects.setAdjust(preset.change, undo)
            }}
          >
            <PictureTile env={env} picture={p} adjust={{ ...now, ...preset.change }} />
          </button>
        ))}
      </div>
    </div>
  )
}

const near = (a: number | undefined, b: number, neutral = 0) => Math.abs((a ?? neutral) - b) < 1e-6
const pct = (v: number) => `${Math.round(v * 100)}%`

// ── Corrections ──

const SHARPEN = [-0.5, -0.25, 0, 0.25, 0.5]
const STEPS = [-0.4, -0.2, 0, 0.2, 0.4]

const correctionsPanel: MenuEntry = {
  panel: (env, close) => {
    const p = env.c.objects.adjustable()
    const locked = !env.c.objects.canCorrect(p)
    return (
      <div className="picture-gallery" role="menu" aria-label="Corrections">
        {locked ? (
          <div className="picture-gallery-note">
            Excel's corrections on this picture are kept as they are: its original is in a form Cells can't open. Reset
            Picture starts again from the picture as shown.
          </div>
        ) : null}
        {section(
          env,
          close,
          "Sharpen/Soften",
          SHARPEN.map((v) => ({
            label: v < 0 ? `Soften: ${pct(-v)}` : `Sharpen: ${pct(v)}`,
            change: { sharpen: v },
            current: (a) => near(a.sharpen, v),
          })),
          5,
          "Sharpen/Soften",
          locked,
        )}
        {section(
          env,
          close,
          "Brightness/Contrast",
          STEPS.flatMap((contrast) =>
            STEPS.map((bright) => ({
              label: `Brightness: ${bright > 0 ? "+" : ""}${pct(bright)} Contrast: ${contrast > 0 ? "+" : ""}${pct(contrast)}`,
              change: { bright, contrast },
              current: (a: PictureAdjust) => near(a.bright, bright) && near(a.contrast, contrast),
            })),
          ),
          5,
          "Brightness/Contrast",
          locked,
        )}
      </div>
    )
  },
}

// ── Color ──

const SATURATION = [0, 0.33, 0.66, 1, 1.33, 2, 3, 4]
const TONE = [4700, 5300, 5900, 6500, 7200, 7800, 8800, 10000, 11200]
const VARIATIONS: [ThemeSlot, string][] = [
  ["dk2", "Text color 2"],
  ["accent1", "Accent color 1"],
  ["accent2", "Accent color 2"],
  ["accent3", "Accent color 3"],
  ["accent4", "Accent color 4"],
  ["accent5", "Accent color 5"],
  ["accent6", "Accent color 6"],
]
/** Office's Sepia, and its Dark and Light Variations of a theme colour. */
const SEPIA: [DuotoneColor, DuotoneColor] = [{ preset: "black" }, { rgb: "D9C3A5", tint: 0.5, satMod: 1.8 }]
const dark = (theme: ThemeSlot): [DuotoneColor, DuotoneColor] => [
  { theme, shade: 0.45, satMod: 1.35 },
  { preset: "white" },
]
const light = (theme: ThemeSlot): [DuotoneColor, DuotoneColor] => [
  { preset: "black" },
  { theme, tint: 0.45, satMod: 4 },
]
const WASHOUT = { bright: 0.7, contrast: -0.7 }
const sameRecolor = (a: PictureRecolor | undefined, b: PictureRecolor | undefined) =>
  JSON.stringify(a ?? null) === JSON.stringify(b ?? null)
const isWashout = (a: PictureAdjust) => !!a.lum && near(a.lum.bright, 0.7) && near(a.lum.contrast, -0.7)
const recolor = (label: string, r: PictureRecolor | undefined, washout = false): Preset => ({
  label,
  change: { recolor: r, lum: washout ? WASHOUT : undefined },
  current: (a) => sameRecolor(a.recolor, r) && isWashout(a) === washout && (washout || !a.lum),
})

const colorPanel: MenuEntry = {
  panel: (env, close) => {
    const locked = !env.c.objects.canCorrect()
    return (
      <div className="picture-gallery" role="menu" aria-label="Color">
        {section(
          env,
          close,
          "Color Saturation",
          SATURATION.map((v) => ({
            label: `Saturation: ${pct(v)}`,
            change: { saturation: v },
            current: (a) => near(a.saturation, v, 1),
          })),
          8,
          "Color Saturation",
          locked,
        )}
        {section(
          env,
          close,
          "Color Tone",
          TONE.map((v) => ({
            label: `Temperature: ${v} K`,
            change: { temperature: v },
            current: (a) => near(a.temperature, v, 6500),
          })),
          9,
          "Color Tone",
          locked,
        )}
        {section(
          env,
          close,
          "Recolor",
          [
            recolor("No Recolor", undefined),
            recolor("Grayscale", { kind: "grayscale" }),
            recolor("Sepia", { kind: "duotone", colors: SEPIA }),
            recolor("Washout", undefined, true),
            ...[0.25, 0.5, 0.75].map((t) =>
              recolor(`Black and White: ${pct(t)}`, { kind: "blackWhite", threshold: t }),
            ),
            ...VARIATIONS.map(([slot, name]) => recolor(`${name} Dark`, { kind: "duotone", colors: dark(slot) })),
            ...VARIATIONS.map(([slot, name]) => recolor(`${name} Light`, { kind: "duotone", colors: light(slot) })),
          ],
          7,
          "Recolor",
        )}
      </div>
    )
  },
}

// ── Artistic Effects ──

const artisticPanel: MenuEntry = {
  panel: (env, close) => {
    const locked = !env.c.objects.canCorrect()
    return (
      <div className="picture-gallery" role="menu" aria-label="Artistic Effects">
        {locked ? (
          <div className="picture-gallery-note">
            Excel's effects on this picture are kept as they are: its original is in a form Cells can't open, or it has
            more than one effect.
          </div>
        ) : null}
        {section(
          env,
          close,
          "Artistic Effects",
          [
            { label: "None", change: { artistic: undefined }, current: (a: PictureAdjust) => !a.artistic },
            ...(Object.keys(ARTISTIC_EFFECTS) as ArtisticEffect[]).map((effect) => ({
              label: ARTISTIC_EFFECTS[effect].label,
              change: { artistic: { effect } },
              current: (a: PictureAdjust) => a.artistic?.effect === effect,
            })),
          ],
          5,
          "Artistic Effects",
          locked,
        )}
      </div>
    )
  },
}

// ── Transparency ──

const TRANSPARENCY = [0, 0.15, 0.3, 0.5, 0.65, 0.8, 0.95]

const transparencyPanel: MenuEntry = {
  panel: (env, close) => (
    <div className="picture-gallery" role="menu" aria-label="Transparency">
      {section(
        env,
        close,
        "Transparency",
        TRANSPARENCY.map((v) => ({
          label: `Transparency: ${pct(v)}`,
          change: { transparency: v },
          current: (a) => near(a.transparency, v),
        })),
        7,
        "Transparency",
      )}
    </div>
  ),
}

const adjustable = (env: RibbonEnv) => !!env.c.objects.adjustable()

export function pictureAdjustCommands(): Command[] {
  return [
    {
      id: "PictureCorrectionsGallery",
      label: "Corrections",
      description: "Improve the brightness, contrast or sharpness of the picture.",
      keytip: "R",
      enabled: adjustable,
      menu: [correctionsPanel, "-", "PictureCorrectionsOptions"],
    },
    {
      id: "PictureColorGallery",
      label: "Color",
      description: "Change the colour of the picture to improve its quality or match your document's content.",
      keytip: "I",
      enabled: adjustable,
      menu: [colorPanel, "-", "PictureSetTransparentColor", "PictureColorOptions"],
    },
    {
      id: "PictureArtisticEffectsGallery",
      label: "Artistic Effects",
      description: "Add an artistic effect to the picture to make it look more like a sketch or a painting.",
      keytip: "X",
      enabled: adjustable,
      menu: [artisticPanel, "-", "PictureArtisticEffectsOptions"],
    },
    {
      id: "PictureTransparencyGallery",
      label: "Transparency",
      description: "Make the picture see-through, so what's behind it shows.",
      keytip: "T",
      enabled: adjustable,
      menu: [transparencyPanel, "-", "PictureTransparencyOptions"],
    },
    {
      id: "PictureSetTransparentColor",
      label: "Set Transparent Color",
      description: "Click a colour in the picture to make every pixel of that colour see-through.",
      keytip: "S",
      enabled: adjustable,
      run: (env) => env.c.objects.pickColor(),
      pressed: (env) => env.c.objects.pickingColor !== null,
    },
    ...(
      [
        ["PictureCorrectionsOptions", "Picture Corrections Options...", "corrections"],
        ["PictureColorOptions", "Picture Color Options...", "color"],
        ["PictureTransparencyOptions", "Picture Transparency Options...", "transparency"],
        ["PictureArtisticEffectsOptions", "Artistic Effects Options...", "artistic"],
      ] as const
    ).map(([id, label, part]): Command => ({
      id,
      label,
      keytip: "O",
      enabled: adjustable,
      run: (env) => env.dialog(`formatPicture:${part}`),
    })),
  ]
}
