import { copyTheme, OFFICE_THEME, sameLook, type Theme, type ThemeSlot } from "@cells/core"
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
function withColors(theme: Theme, set: (typeof OFFICE_COLOR_SETS)[number]): Theme {
  return { ...copyTheme(theme), colors: { ...set.colors }, colorsName: set.name }
}

function withFonts(theme: Theme, set: (typeof OFFICE_FONT_SETS)[number]): Theme {
  return { ...copyTheme(theme), fonts: { major: set.major, minor: set.minor }, fontsName: set.name }
}

function setTheme(c: SheetController, theme: Theme, label: string) {
  if (sameLook(theme, c.wb.theme) && theme.name === c.wb.theme.name) return
  if (c.edit) c.commitEdit()
  c.transact([{ kind: "theme", theme }], { label, before: c.selection })
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
          panel: (env, close) => {
            const now = env.c.wb.theme
            return (
              <div className="th-panel" role="menu" aria-label="Themes">
                <div className="th-heading">Office</div>
                <div className="th-grid">
                  {OFFICE_THEMES.map((t) => {
                    const current = t.name === now.name && sameLook(t, now)
                    return (
                      <button
                        key={t.name}
                        type="button"
                        role="menuitemradio"
                        aria-checked={current}
                        className={`th-item${current ? " is-current" : ""}`}
                        title={shownName(t.name)}
                        onClick={() => {
                          setTheme(env.c, copyTheme(t), "Theme")
                          close()
                        }}
                      >
                        <ThemeTile theme={t} />
                        <span className="th-name">{shownName(t.name)}</span>
                      </button>
                    )
                  })}
                </div>
              </div>
            )
          },
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
    { id: "ThemeSaveCurrentTheme", label: "Save Current Theme..." },
    {
      id: "ThemeColorsGallery",
      label: "Colors",
      description:
        "Change the colors of the current theme. Choosing a set of colors changes every theme color used in the workbook.",
      keytip: "TC",
      menu: [
        {
          panel: (env, close) => {
            const now = env.c.wb.theme
            return (
              <div className="pl-list th-list" role="menu" aria-label="Colors">
                {OFFICE_COLOR_SETS.map((set) => {
                  const current = sameColors(set.colors, now.colors)
                  return (
                    <button
                      key={set.label}
                      type="button"
                      role="menuitemradio"
                      aria-checked={current}
                      className={`pl-item th-row${current ? " is-current" : ""}`}
                      onClick={() => {
                        setTheme(env.c, withColors(env.c.wb.theme, set), "Theme Colors")
                        close()
                      }}
                    >
                      <Chips colors={set.colors} />
                      <span>{set.label}</span>
                    </button>
                  )
                })}
              </div>
            )
          },
        },
        "-",
        "ThemeColorsCustomize",
      ],
    },
    { id: "ThemeColorsCustomize", label: "Customize Colors..." },
    {
      id: "ThemeFontsGallery",
      label: "Fonts",
      description: "Change the fonts of the current theme: one for headings and one for body text.",
      keytip: "TF",
      menu: [
        {
          panel: (env, close) => {
            const now = env.c.wb.theme
            return (
              <div className="pl-list th-list" role="menu" aria-label="Fonts">
                {OFFICE_FONT_SETS.map((set) => {
                  const current = set.major === now.fonts.major && set.minor === now.fonts.minor
                  return (
                    <button
                      key={set.label}
                      type="button"
                      role="menuitemradio"
                      aria-checked={current}
                      className={`pl-item th-font${current ? " is-current" : ""}`}
                      onClick={() => {
                        setTheme(env.c, withFonts(env.c.wb.theme, set), "Theme Fonts")
                        close()
                      }}
                    >
                      <span className="th-font-aa" style={{ fontFamily: `"${set.major}", sans-serif` }} aria-hidden>
                        Aa
                      </span>
                      <span className="pl-text">
                        <strong>{set.label}</strong>
                        <span style={{ fontFamily: `"${set.major}", sans-serif` }}>{set.major}</span>
                        <span style={{ fontFamily: `"${set.minor}", sans-serif` }}>{set.minor}</span>
                      </span>
                    </button>
                  )
                })}
              </div>
            )
          },
        },
        "-",
        "ThemeFontsCustomize",
      ],
    },
    { id: "ThemeFontsCustomize", label: "Customize Fonts..." },
  ]
}
