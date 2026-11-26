import { useState } from "react"
import type { Theme, ThemeSlot } from "@cells/core"
import { FONT_NAMES } from "@/ribbon/home-commands"
import { applyThemeColors, applyThemeFonts } from "@/ribbon/theme-commands"
import { THEME_SLOTS, type CustomThemeColors, type CustomThemeFonts } from "@/settings/shared"
import { updateSettings, useSettings } from "@/settings/settings"
import type { SheetController } from "@/sheet/controller"
import { ColorButton } from "./ChartFormatPane"
import { Dialog } from "./Dialog"

const SLOT_LABELS: Record<ThemeSlot, string> = {
  dk1: "Text/Background - Dark 1",
  lt1: "Text/Background - Light 1",
  dk2: "Text/Background - Dark 2",
  lt2: "Text/Background - Light 2",
  accent1: "Accent 1",
  accent2: "Accent 2",
  accent3: "Accent 3",
  accent4: "Accent 4",
  accent5: "Accent 5",
  accent6: "Accent 6",
  hlink: "Hyperlink",
  folHlink: "Followed Hyperlink",
}

/** Office's next free name for a custom set: "Custom 1", "Custom 2"… */
export function nextCustomName(taken: { name: string }[]): string {
  const names = new Set(taken.map((t) => t.name.toLowerCase()))
  let n = 1
  while (names.has(`custom ${n}`)) n++
  return `Custom ${n}`
}

/** A custom list with `item` in it: in the place of the one edited (or of the same name), else at the end. */
export function putInList<T extends { name: string }>(list: T[], item: T, replacing?: string): T[] {
  const old = replacing ?? item.name
  const kept = list.filter((x) => x.name !== item.name || x.name === old)
  const at = kept.findIndex((x) => x.name === old)
  return at < 0 ? [...kept, item] : kept.map((x, i) => (i === at ? item : x))
}

/** The sample in Create New Theme Colors: text, accents as a little chart, and the two link colours. */
function ColorsSample({ colors, dark }: { colors: Theme["colors"]; dark: boolean }) {
  const c = (s: ThemeSlot) => `#${colors[s]}`
  const accents = (["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const).map(c)
  const [bg, fg, bg2, fg2] = dark ? [c("dk2"), c("lt1"), c("dk1"), c("lt2")] : [c("lt1"), c("dk1"), c("lt2"), c("dk2")]
  return (
    <svg className="thc-sample" viewBox="0 0 132 92" width="132" height="92" aria-hidden>
      <rect width="132" height="92" fill={bg} stroke="var(--app-border)" />
      <rect x="6" y="6" width="44" height="22" fill={bg2} />
      <text x="10" y="21" fill={fg} fontSize="11" fontFamily="sans-serif">
        Text
      </text>
      <text x="58" y="21" fill={fg2} fontSize="11" fontFamily="sans-serif">
        Text
      </text>
      {accents.map((fill, i) => (
        <rect
          key={i}
          x={8 + i * 12}
          y={78 - (14 + ((i * 7) % 24))}
          width="9"
          height={14 + ((i * 7) % 24)}
          fill={fill}
        />
      ))}
      <text x="84" y="58" fill={c("hlink")} fontSize="9" fontFamily="sans-serif" textDecoration="underline">
        Hyperlink
      </text>
      <text x="84" y="74" fill={c("folHlink")} fontSize="9" fontFamily="sans-serif" textDecoration="underline">
        Followed
      </text>
    </svg>
  )
}

/**
 * Page Layout › Colors › Customize Colors (Create New Theme Colors), or Edit on a custom set: a colour for each of
 * the theme's twelve places, a sample in light and dark, and a name. Save keeps the set with the person's custom
 * sets (shared with Inkspire) and gives the workbook its colours; Reset goes back to the colours it started with.
 */
export function ThemeColorsDialog({
  controller,
  edit,
  onClose,
}: {
  controller: SheetController
  edit?: string
  onClose: () => void
}) {
  const list = useSettings((s) => s.themeColors)
  const editing = edit ? list.find((x) => x.name === edit) : undefined
  const start = editing?.colors ?? controller.wb.theme.colors
  const [colors, setColors] = useState<Theme["colors"]>({ ...start })
  const [name, setName] = useState(editing?.name ?? nextCustomName(list))
  const save = () => {
    const set: CustomThemeColors = { name: name.trim(), colors }
    if (!set.name) return
    void updateSettings({ themeColors: putInList(list, set, editing?.name) })
    applyThemeColors(controller, set)
    onClose()
  }
  return (
    <Dialog
      title={editing ? "Edit Theme Colors" : "Create New Theme Colors"}
      width={620}
      acceptLabel="Save"
      onCancel={onClose}
      onAccept={save}
      extra={{ label: "Reset", onSelect: () => setColors({ ...start }) }}
    >
      <div className="thc">
        <div className="thc-slots" role="group" aria-label="Theme colors">
          <span className="dialog-label">Theme colors</span>
          {THEME_SLOTS.map((slot) => (
            <div key={slot} className="thc-slot">
              <span>{SLOT_LABELS[slot]}</span>
              <ColorButton
                label={SLOT_LABELS[slot]}
                color={`#${colors[slot]}`}
                theme={controller.wb.theme}
                onPick={(hex) => hex && setColors({ ...colors, [slot]: hex.slice(1).toUpperCase() })}
              />
            </div>
          ))}
        </div>
        <div className="thc-preview">
          <span className="dialog-label">Sample</span>
          <div className="thc-samples">
            <ColorsSample colors={colors} dark={false} />
            <ColorsSample colors={colors} dark />
          </div>
        </div>
      </div>
      <div className="thc-name">
        <label className="dialog-label" htmlFor="theme-colors-name">
          Name:
        </label>
        <input id="theme-colors-name" className="dialog-field" value={name} onChange={(e) => setName(e.target.value)} />
      </div>
    </Dialog>
  )
}

function FontSelect({ id, value, onChange }: { id: string; value: string; onChange: (font: string) => void }) {
  const names = FONT_NAMES.includes(value) ? FONT_NAMES : [value, ...FONT_NAMES]
  return (
    <select id={id} className="dialog-field" value={value} onChange={(e) => onChange(e.target.value)}>
      {names.map((f) => (
        <option key={f} value={f}>
          {f}
        </option>
      ))}
    </select>
  )
}

/**
 * Page Layout › Fonts › Customize Fonts (Create New Theme Fonts), or Edit on a custom pair: the headings font and
 * the body font with a sample, and a name. Save keeps the pair with the person's (shared with Inkspire) and gives
 * the workbook its fonts.
 */
export function ThemeFontsDialog({
  controller,
  edit,
  onClose,
}: {
  controller: SheetController
  edit?: string
  onClose: () => void
}) {
  const list = useSettings((s) => s.themeFonts)
  const editing = edit ? list.find((x) => x.name === edit) : undefined
  const [major, setMajor] = useState(editing?.major ?? controller.wb.theme.fonts.major)
  const [minor, setMinor] = useState(editing?.minor ?? controller.wb.theme.fonts.minor)
  const [name, setName] = useState(editing?.name ?? nextCustomName(list))
  const save = () => {
    const set: CustomThemeFonts = { name: name.trim(), major, minor }
    if (!set.name) return
    void updateSettings({ themeFonts: putInList(list, set, editing?.name) })
    applyThemeFonts(controller, set)
    onClose()
  }
  return (
    <Dialog
      title={editing ? "Edit Theme Fonts" : "Create New Theme Fonts"}
      width={520}
      acceptLabel="Save"
      onCancel={onClose}
      onAccept={save}
    >
      <div className="thf">
        <div className="thf-fields">
          <label className="dialog-label" htmlFor="theme-font-major">
            Heading font:
          </label>
          <FontSelect id="theme-font-major" value={major} onChange={setMajor} />
          <label className="dialog-label" htmlFor="theme-font-minor">
            Body font:
          </label>
          <FontSelect id="theme-font-minor" value={minor} onChange={setMinor} />
        </div>
        <div className="thf-preview">
          <span className="dialog-label">Sample</span>
          <div className="thf-sample" aria-label="Sample">
            <div style={{ fontFamily: `"${major}", sans-serif` }} className="thf-heading">
              Heading
            </div>
            <div style={{ fontFamily: `"${minor}", sans-serif` }}>
              Body text body text body text. Body text body text.
            </div>
          </div>
        </div>
      </div>
      <div className="thc-name">
        <label className="dialog-label" htmlFor="theme-fonts-name">
          Name:
        </label>
        <input id="theme-fonts-name" className="dialog-field" value={name} onChange={(e) => setName(e.target.value)} />
      </div>
    </Dialog>
  )
}
