import {
  copyTheme,
  OFFICE_EFFECTS,
  OFFICE_THEME,
  sameLook,
  type Theme,
  type ThemeEffects,
  type ThemeSlot,
} from "@cells/core"
import { useState } from "react"
import { themeEffectFilter } from "@/drawing/render"
import { updateSettings, useSettings } from "@/settings/settings"
import type { CustomThemeColors, CustomThemeFonts } from "@/settings/shared"
import { ContextMenu, type MenuItem } from "@/shell/ContextMenu"
import { EFFECT_SCHEMES } from "./effects-data"
import type { SheetController } from "@/sheet/controller"
import { COLOR_SETS, FONT_SETS, THEMES } from "./themes-data"
import type { Command, RibbonEnv } from "./registry"

/**
 * Page Layout › Themes: Office's themes, colour sets and font sets (as Excel 365 lists them). Picking one changes the
 * workbook's theme, and with it every theme colour and theme font in cells, tables, shapes and charts, as one step
 * to undo. Browse for Themes reads an Office theme file (.thmx) or a document's theme.
 */

const ORDER: ThemeSlot[] = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
]

const colorsOf = (list: string): Theme["colors"] => {
  const hex = list.split(" ")
  return Object.fromEntries(ORDER.map((slot, i) => [slot, hex[i]])) as Theme["colors"]
}

/** Office's themes, the default first. */
export const OFFICE_THEMES: Theme[] = [
  copyTheme(OFFICE_THEME),
  ...THEMES.map(([name, colors, major, minor]) => ({ name, colors: colorsOf(colors), fonts: { major, minor } })),
]

/** Colour sets: [shown name, theme name written, colours]. */
export const OFFICE_COLOR_SETS: { label: string; name: string; colors: Theme["colors"] }[] = [
  { label: "Office", name: "Office", colors: { ...OFFICE_THEME.colors } },
  ...COLOR_SETS.map(([label, name, colors]) => ({ label, name, colors: colorsOf(colors) })),
]

export const OFFICE_FONT_SETS: { label: string; name: string; major: string; minor: string }[] = [
  { label: "Office", name: "Office", ...OFFICE_THEME.fonts },
  ...FONT_SETS.map(([label, name, major, minor]) => ({ label, name, major, minor })),
]

const shownName = (name: string) => (name === "Office Theme" ? "Office" : name.replace(/ Theme$/, ""))

/** The theme with its colours or fonts replaced (Excel keeps the theme's name and names the set apart). */
function withColors(theme: Theme, set: { name: string; colors: Theme["colors"] }): Theme {
  return { ...copyTheme(theme), colors: { ...set.colors }, colorsName: set.name }
}

function withFonts(theme: Theme, set: { name: string; major: string; minor: string }): Theme {
  return { ...copyTheme(theme), fonts: { major: set.major, minor: set.minor }, fontsName: set.name }
}

/** A colour set (Office's or the person's) given to the workbook's theme, as one step to undo. */
export function applyThemeColors(c: SheetController, set: { name: string; colors: Theme["colors"] }) {
  setTheme(c, withColors(c.wb.theme, set), "Theme Colors")
}

export function applyThemeFonts(c: SheetController, set: { name: string; major: string; minor: string }) {
  setTheme(c, withFonts(c.wb.theme, set), "Theme Fonts")
}

function setTheme(c: SheetController, theme: Theme, label: string) {
  const now = c.wb.theme
  if (
    sameLook(theme, now) &&
    theme.name === now.name &&
    theme.colorsName === now.colorsName &&
    theme.fontsName === now.fontsName &&
    theme.effects?.scheme === now.effects?.scheme
  )
    return
  if (c.edit) c.commitEdit()
  c.transact([{ kind: "theme", theme }], { label, before: c.selection })
}

/** Office's effect schemes, Office's own first (as `undefined` on the theme: what a new workbook has). */
export const OFFICE_EFFECT_SCHEMES: { label: string; effects: ThemeEffects | undefined }[] = [
  { label: "Office", effects: undefined },
  ...EFFECT_SCHEMES.map(([label, scheme]) => ({
    label,
    effects: { name: /\bname="([^"]*)"/.exec(scheme)?.[1] ?? label, scheme },
  })),
]

function withEffects(theme: Theme, effects: ThemeEffects | undefined): Theme {
  const t = copyTheme(theme)
  if (effects) t.effects = { ...effects }
  else delete t.effects
  return t
}

/** An effect scheme's picture: three rounded shapes in the theme's accents with its three effects. */
function EffectsTile({ theme, effects, id }: { theme: Theme; effects: ThemeEffects | undefined; id: string }) {
  const t = withEffects(theme, effects)
  const shapes = ([1, 2, 3] as const).map((n, i) => {
    const color = `#${theme.colors[(["accent1", "accent2", "accent3"] as const)[i]]}`
    const fx = themeEffectFilter(`${id}-${n}`, n, t, color)
    return (
      `${fx ? `<defs>${fx}</defs>` : ""}<rect x="${6 + i * 22}" y="${10 + (i % 2) * 4}" width="18" height="18" rx="4" fill="${color}"` +
      `${fx ? ` filter="url(#${id}-${n})"` : ""}/>`
    )
  })
  return (
    <svg
      className="th-effects"
      width="76"
      height="40"
      viewBox="0 0 76 40"
      aria-hidden
      dangerouslySetInnerHTML={{ __html: shapes.join("") }}
    />
  )
}

const sameColors = (a: Theme["colors"], b: Theme["colors"]) =>
  ORDER.every((slot) => a[slot].toUpperCase() === b[slot].toUpperCase())

/** A theme's picture: "Aa" in its headings font on its background, over its six accents. */
function ThemeTile({ theme }: { theme: Theme }) {
  const c = theme.colors
  return (
    <span className="th-tile" style={{ background: `#${c.lt1}` }} aria-hidden>
      <span className="th-aa" style={{ color: `#${c.dk1}`, fontFamily: `"${theme.fonts.major}", sans-serif` }}>
        Aa
      </span>
      <span className="th-bars">
        {(["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const).map((s) => (
          <span key={s} style={{ background: `#${c[s]}` }} />
        ))}
      </span>
    </span>
  )
}

function Chips({ colors }: { colors: Theme["colors"] }) {
  const shown: ThemeSlot[] = ["dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
  return (
    <span className="th-chips" aria-hidden>
      {shown.map((s) => (
        <span key={s} style={{ background: `#${colors[s]}` }} />
      ))}
    </span>
  )
}

/** A right-click menu on a gallery's custom entries (Edit…, Delete…), drawn over the gallery. */
function useEntryMenu() {
  const [menu, setMenu] = useState<{ x: number; y: number; items: MenuItem[] } | null>(null)
  const open = (items: MenuItem[]) => (e: React.MouseEvent) => {
    e.preventDefault()
    // The entry's menu, not the ribbon's (Add to Quick Access Toolbar…).
    e.stopPropagation()
    setMenu({ x: e.clientX, y: e.clientY, items })
  }
  const shown = menu ? <ContextMenu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} /> : null
  return { open, shown }
}

function confirmDelete(env: RibbonEnv, what: string, then: () => void) {
  env.c.shell.confirm(`Delete these theme ${what}?`, then)
}

/** Page Layout › Colors: the person's sets (Custom) above Office's. */
function ColorsPanel({ env, close }: { env: RibbonEnv; close: () => void }) {
  const custom = useSettings((s) => s.themeColors)
  const menu = useEntryMenu()
  const now = env.c.wb.theme
  const row = (key: string, label: string, set: { name: string; colors: Theme["colors"] }, items?: MenuItem[]) => {
    const current = sameColors(set.colors, now.colors) && (!items || now.colorsName === set.name)
    return (
      <button
        key={key}
        type="button"
        role="menuitemradio"
        aria-checked={current}
        className={`pl-item th-row${current ? " is-current" : ""}`}
        onClick={() => {
          applyThemeColors(env.c, set)
          close()
        }}
        onContextMenu={items ? menu.open(items) : undefined}
      >
        <Chips colors={set.colors} />
        <span>{label}</span>
      </button>
    )
  }
  const customItems = (set: CustomThemeColors): MenuItem[] => [
    { kind: "item", label: "Edit...", onSelect: () => (close(), env.dialog(`themeColors:${set.name}`)) },
    {
      kind: "item",
      label: "Delete...",
      onSelect: () =>
        confirmDelete(env, "colors", () => void updateSettings({ themeColors: custom.filter((x) => x !== set) })),
    },
  ]
  return (
    <div className="pl-list th-list" role="menu" aria-label="Colors">
      {custom.length ? (
        <>
          <div className="th-heading">Custom</div>
          {custom.map((set) => row(`custom:${set.name}`, set.name, set, customItems(set)))}
          <div className="th-heading">Office</div>
        </>
      ) : null}
      {OFFICE_COLOR_SETS.map((set) => row(set.label, set.label, set))}
      {menu.shown}
    </div>
  )
}

/** Page Layout › Fonts: the person's pairs (Custom) above Office's. */
function FontsPanel({ env, close }: { env: RibbonEnv; close: () => void }) {
  const custom = useSettings((s) => s.themeFonts)
  const menu = useEntryMenu()
  const now = env.c.wb.theme
  const row = (key: string, label: string, set: { name: string; major: string; minor: string }, items?: MenuItem[]) => {
    const current =
      set.major === now.fonts.major && set.minor === now.fonts.minor && (!items || now.fontsName === set.name)
    return (
      <button
        key={key}
        type="button"
        role="menuitemradio"
        aria-checked={current}
        className={`pl-item th-font${current ? " is-current" : ""}`}
        onClick={() => {
          applyThemeFonts(env.c, set)
          close()
        }}
        onContextMenu={items ? menu.open(items) : undefined}
      >
        <span className="th-font-aa" style={{ fontFamily: `"${set.major}", sans-serif` }} aria-hidden>
          Aa
        </span>
        <span className="pl-text">
          <strong>{label}</strong>
          <span style={{ fontFamily: `"${set.major}", sans-serif` }}>{set.major}</span>
          <span style={{ fontFamily: `"${set.minor}", sans-serif` }}>{set.minor}</span>
        </span>
      </button>
    )
  }
  const customItems = (set: CustomThemeFonts): MenuItem[] => [
    { kind: "item", label: "Edit...", onSelect: () => (close(), env.dialog(`themeFonts:${set.name}`)) },
    {
      kind: "item",
      label: "Delete...",
      onSelect: () =>
        confirmDelete(env, "fonts", () => void updateSettings({ themeFonts: custom.filter((x) => x !== set) })),
    },
  ]
  return (
    <div className="pl-list th-list" role="menu" aria-label="Fonts">
      {custom.length ? (
        <>
          <div className="th-heading">Custom</div>
          {custom.map((set) => row(`custom:${set.name}`, set.name, set, customItems(set)))}
          <div className="th-heading">Office</div>
        </>
      ) : null}
      {OFFICE_FONT_SETS.map((set) => row(set.label, set.label, set))}
      {menu.shown}
    </div>
  )
}

/** Page Layout › Themes: the themes the person saved (Custom) above Office's. */
function ThemesPanel({ env, close }: { env: RibbonEnv; close: () => void }) {
  const saved = useSettings((s) => s.savedThemes)
  const menu = useEntryMenu()
  const now = env.c.wb.theme
  const tile = (t: Theme, key: string, items?: MenuItem[]) => {
    const current = t.name === now.name && sameLook(t, now)
    return (
      <button
        key={key}
        type="button"
        role="menuitemradio"
        aria-checked={current}
        className={`th-item${current ? " is-current" : ""}`}
        title={shownName(t.name)}
        onClick={() => {
          setTheme(env.c, copyTheme(t), "Theme")
          close()
        }}
        onContextMenu={items ? menu.open(items) : undefined}
      >
        <ThemeTile theme={t} />
        <span className="th-name">{shownName(t.name)}</span>
      </button>
    )
  }
  return (
    <div className="th-panel" role="menu" aria-label="Themes">
      {saved.length ? (
        <>
          <div className="th-heading">Custom</div>
          <div className="th-grid">
            {saved.map((t) =>
              tile(t, `custom:${t.name}`, [
                {
                  kind: "item",
                  label: "Delete...",
                  onSelect: () =>
                    env.c.shell.confirm(
                      "Delete this theme?",
                      () => void updateSettings({ savedThemes: saved.filter((x) => x !== t) }),
                    ),
                },
              ]),
            )}
          </div>
        </>
      ) : null}
      <div className="th-heading">Office</div>
      <div className="th-grid">{OFFICE_THEMES.map((t) => tile(t, t.name))}</div>
      {menu.shown}
    </div>
  )
}

/**
 * Save Current Theme: the workbook's theme as an Office theme file (.thmx), named as the file is; it's also kept
 * with the person's themes (shared with Inkspire), so the Themes gallery lists it under Custom.
 */
async function saveCurrentTheme(env: RibbonEnv) {
  if (!env.files) return
  const theme = env.c.wb.theme
  const { themeFileBytes } = await import("@cells/io")
  const base = theme.name === "Office Theme" ? "Theme1" : shownName(theme.name)
  const name = await env.files.saveOther(`${base}.thmx`, { name: "Office Theme", extensions: ["thmx"] }, (chosen) =>
    themeFileBytes({ ...copyTheme(theme), name: chosen }),
  )
  if (!name) return
  const kept: Theme = { ...copyTheme(theme), name }
  const saved = useSettings.getState().savedThemes
  void updateSettings({ savedThemes: [...saved.filter((t) => t.name !== name), kept] })
  setTheme(env.c, kept, "Theme")
}

/** Browse for Themes: an Office theme file, or a document whose theme to use. */
function browseTheme(env: RibbonEnv) {
  const input = document.createElement("input")
  input.type = "file"
  input.accept = ".thmx,.xlsx,.xlsm,.xltx,.docx,.pptx"
  input.onchange = async () => {
    const file = input.files?.[0]
    if (!file) return
    const { themeFromFile } = await import("@cells/io")
    const theme = await themeFromFile(new Uint8Array(await file.arrayBuffer()))
    if (!theme) {
      env.c.shell.alert(
        "Cells couldn't find a theme in that file. Choose an Office theme (.thmx) or an Office document.",
      )
      return
    }
    setTheme(env.c, theme, "Theme")
  }
  input.click()
}

export function themeCommands(): Command[] {
  return [
    {
      id: "ThemesGallery",
      label: "Themes",
      description:
        "Change the overall design of the workbook, including colors, fonts, and effects. Cells, tables, charts and shapes in theme colors and fonts follow it.",
      keytip: "TH",
      menu: [
        {
          panel: (env, close) => <ThemesPanel env={env} close={close} />,
        },
        "-",
        "ThemeBrowseForThemes",
        "ThemeSaveCurrentTheme",
      ],
    },
    {
      id: "ThemeBrowseForThemes",
      label: "Browse for Themes...",
      description: "Use the theme of an Office theme file (.thmx) or of another workbook, document or presentation.",
      run: browseTheme,
    },
    {
      id: "ThemeSaveCurrentTheme",
      label: "Save Current Theme...",
      description:
        "Save the workbook's theme as an Office theme file (.thmx), to use in other workbooks and documents.",
      enabled: (env) => !!env.files,
      run: (env) => void saveCurrentTheme(env),
    },
    {
      id: "ThemeColorsGallery",
      label: "Colors",
      description:
        "Change the colors of the current theme. Choosing a set of colors changes every theme color used in the workbook.",
      keytip: "TC",
      menu: [
        {
          panel: (env, close) => <ColorsPanel env={env} close={close} />,
        },
        "-",
        "ThemeColorsCustomize",
      ],
    },
    {
      id: "ThemeColorsCustomize",
      label: "Customize Colors...",
      description: "Make a set of theme colors of your own.",
      run: (env) => env.dialog("themeColors"),
    },
    {
      id: "ThemeFontsGallery",
      label: "Fonts",
      description: "Change the fonts of the current theme: one for headings and one for body text.",
      keytip: "TF",
      menu: [
        {
          panel: (env, close) => <FontsPanel env={env} close={close} />,
        },
        "-",
        "ThemeFontsCustomize",
      ],
    },
    {
      id: "ThemeFontsCustomize",
      label: "Customize Fonts...",
      description: "Pick the theme's heading and body fonts and save them as a set of your own.",
      run: (env) => env.dialog("themeFonts"),
    },
    {
      id: "ThemeEffectsGallery",
      label: "Effects",
      description:
        "Change the look of the theme's effects: the shadows, glows and bevels that shapes with a theme effect style get.",
      keytip: "TE",
      menu: [
        {
          panel: (env, close) => {
            const now = env.c.wb.theme.effects?.scheme ?? OFFICE_EFFECTS.scheme
            return (
              <div className="th-effects-grid" role="menu" aria-label="Effects">
                {OFFICE_EFFECT_SCHEMES.map((set, i) => {
                  const current = (set.effects?.scheme ?? OFFICE_EFFECTS.scheme) === now
                  return (
                    <button
                      key={set.label}
                      type="button"
                      role="menuitemradio"
                      aria-checked={current}
                      title={set.label}
                      aria-label={set.label}
                      className={`th-effects-item${current ? " is-current" : ""}`}
                      onClick={() => {
                        setTheme(env.c, withEffects(env.c.wb.theme, set.effects), "Theme Effects")
                        close()
                      }}
                    >
                      <EffectsTile theme={env.c.wb.theme} effects={set.effects} id={`thfx${i}`} />
                      <span>{set.label}</span>
                    </button>
                  )
                })}
              </div>
            )
          },
        },
      ],
    },
  ]
}
