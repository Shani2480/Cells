import { MARGIN_PRESETS, PAPER_SIZES, paperOf, type PageMargins } from "@cells/core"
import { pickPicture } from "@/drawing/pictures"
import { marginPreset } from "@/sheet/print"
import { marginUnit } from "@/shell/PageSetupDialog"
import type { Command, RibbonEnv } from "./registry"

/**
 * Page Layout › Page Setup, Scale to Fit and Sheet Options (Print), as Excel 365 has them. The margin and size lists
 * show their measurements in the region's unit.
 */

const MARGIN_NAMES: Record<keyof typeof MARGIN_PRESETS, string> = { normal: "Normal", wide: "Wide", narrow: "Narrow" }

function measure(env: RibbonEnv, inches: number): string {
  const unit = marginUnit(env.c)
  const v = Math.round(inches * unit.perInch * 100) / 100
  return unit.name === "in" ? `${v}"` : `${v} cm`
}

/** One row of the Margins and Size menus: a picture, the name in bold and the measurements under it. */
function RichItem({
  picture,
  name,
  lines,
  current,
  onPick,
}: {
  picture: React.ReactNode
  name: string
  lines: string[]
  current: boolean
  onPick: () => void
}) {
  return (
    <button
      type="button"
      role="menuitemradio"
      aria-checked={current}
      className={`pl-item${current ? " is-current" : ""}`}
      onClick={onPick}
    >
      <span className="pl-picture" data-keep-light aria-hidden>
        {picture}
      </span>
      <span className="pl-text">
        <strong>{name}</strong>
        {lines.map((l) => (
          <span key={l}>{l}</span>
        ))}
      </span>
    </button>
  )
}

/** A page with its margins drawn in, for the Margins menu. */
function MarginPicture({ m }: { m: PageMargins }) {
  const s = 3.4
  return (
    <svg width="34" height="40" viewBox="0 0 34 40">
      <rect x="0.5" y="0.5" width="33" height="39" fill="#fff" stroke="#8a8a8a" />
      <rect
        x={m.left * s + 0.5}
        y={m.top * s + 0.5}
        width={33 - (m.left + m.right) * s}
        height={39 - (m.top + m.bottom) * s}
        fill="none"
        stroke="#3a7bd5"
        strokeDasharray="2 1.5"
      />
    </svg>
  )
}

function PaperPicture({ w, h }: { w: number; h: number }) {
  const k = 36 / Math.max(w, h)
  return (
    <svg width="34" height="40" viewBox="-1 -1 38 38">
      <rect x={(36 - w * k) / 2} y={(36 - h * k) / 2} width={w * k} height={h * k} fill="#fff" stroke="#8a8a8a" />
    </svg>
  )
}

const FIT_OPTIONS = [
  { label: "Automatic", value: "0" },
  ...Array.from({ length: 9 }, (_, i) => ({ label: `${i + 1} page${i ? "s" : ""}`, value: String(i + 1) })),
]

function fitField(axis: "width" | "height"): Command["field"] {
  return {
    value: (env) => {
      const n = env.c.print.setup().fitTo?.[axis] ?? 0
      return n ? `${n} page${n > 1 ? "s" : ""}` : "Automatic"
    },
    commit: (env, text) => {
      const t = text.trim()
      const n = /^auto/i.test(t) ? 0 : parseInt(t, 10)
      if (Number.isFinite(n) && n >= 0 && n <= 32767) env.c.print.setFit(axis, n)
    },
    options: () => FIT_OPTIONS,
    width: 96,
  }
}

export function pageLayoutCommands(): Command[] {
  return [
    {
      id: "PageMarginsGallery",
      label: "Margins",
      description:
        "Set the margin sizes for the entire document or the current section. Choose from several commonly used margin formats, or define your own.",
      keytip: "M",
      menu: [
        {
          panel: (env, close) => {
            const m = env.c.print.setup().margins
            const current = marginPreset(m)
            const lines = (p: PageMargins) => [
              `Top: ${measure(env, p.top)}   Bottom: ${measure(env, p.bottom)}`,
              `Left: ${measure(env, p.left)}   Right: ${measure(env, p.right)}`,
              `Header: ${measure(env, p.header)}   Footer: ${measure(env, p.footer)}`,
            ]
            return (
              <div className="pl-list" role="menu" aria-label="Margins">
                {current === "custom" ? (
                  <RichItem
                    picture={<MarginPicture m={m} />}
                    name="Last Custom Setting"
                    lines={lines(m)}
                    current
                    onPick={close}
                  />
                ) : null}
                {(Object.keys(MARGIN_PRESETS) as (keyof typeof MARGIN_PRESETS)[]).map((k) => (
                  <RichItem
                    key={k}
                    picture={<MarginPicture m={MARGIN_PRESETS[k]} />}
                    name={MARGIN_NAMES[k]}
                    lines={lines(MARGIN_PRESETS[k])}
                    current={current === k}
                    onPick={() => {
                      env.c.print.setMargins(MARGIN_PRESETS[k])
                      close()
                    }}
                  />
                ))}
              </div>
            )
          },
        },
        "-",
        "PageMarginsCustom",
      ],
    },
    {
      id: "PageMarginsCustom",
      label: "Custom Margins...",
      keytip: "A",
      run: (env) => env.dialog("pageSetup:margins"),
    },
    {
      id: "PageOrientationGallery",
      label: "Orientation",
      description: "Switch the pages between portrait and landscape layouts.",
      keytip: "O",
      menu: ["PageOrientationPortrait", "PageOrientationLandscape"],
    },
    {
      id: "PageOrientationPortrait",
      label: "Portrait",
      check: true,
      run: (env) => env.c.print.setOrientation("portrait"),
      pressed: (env) => env.c.print.setup().orientation === "portrait",
    },
    {
      id: "PageOrientationLandscape",
      label: "Landscape",
      check: true,
      run: (env) => env.c.print.setOrientation("landscape"),
      pressed: (env) => env.c.print.setup().orientation === "landscape",
    },
    {
      id: "PageSizeGallery",
      label: "Size",
      description: "Choose a paper size for the current section.",
      keytip: "SZ",
      menu: [
        {
          panel: (env, close) => {
            const current = paperOf(env.c.print.setup(), env.c.print.defaultPaper)
            return (
              <div className="pl-list" role="menu" aria-label="Size">
                {Object.entries(PAPER_SIZES).map(([id, p]) => (
                  <RichItem
                    key={id}
                    picture={<PaperPicture w={p.width} h={p.height} />}
                    name={p.name}
                    lines={[`${measure(env, p.width)} x ${measure(env, p.height)}`]}
                    current={Number(id) === current}
                    onPick={() => {
                      env.c.print.setPaper(Number(id))
                      close()
                    }}
                  />
                ))}
              </div>
            )
          },
        },
        "-",
        "PageSizeMore",
      ],
    },
    {
      id: "PageSizeMore",
      label: "More Paper Sizes...",
      keytip: "M",
      run: (env) => env.dialog("pageSetup:page"),
    },
    {
      id: "PrintAreaMenu",
      label: "Print Area",
      description: "Mark a specific area of the sheet for printing.",
      keytip: "R",
      menu: (env) => ["PrintAreaSet", ...(env.c.print.canAddToPrintArea() ? ["PrintAreaAddTo"] : []), "PrintAreaClear"],
    },
    {
      id: "PrintAreaSet",
      label: "Set Print Area",
      keytip: "S",
      run: (env) => env.c.print.setPrintArea(),
    },
    {
      id: "PrintAreaAddTo",
      label: "Add to Print Area",
      keytip: "A",
      run: (env) => env.c.print.addToPrintArea(),
    },
    {
      id: "PrintAreaClear",
      label: "Clear Print Area",
      keytip: "C",
      run: (env) => env.c.print.clearPrintArea(),
      enabled: (env) => env.c.print.setup().printArea !== null,
    },
    {
      id: "PageBreakMenu",
      label: "Breaks",
      description:
        "Add a page break where you want the next page to start in the printed copy. The page break will be inserted above and to the left of your selection.",
      keytip: "B",
      menu: ["PageBreakInsert", "PageBreakRemove", "PageBreakResetAll"],
    },
    {
      id: "PageBreakInsert",
      label: "Insert Page Break",
      keytip: "I",
      run: (env) => env.c.print.insertBreak(),
      enabled: (env) => env.c.selection.active.row > 0 || env.c.selection.active.col > 0,
    },
    {
      id: "PageBreakRemove",
      label: "Remove Page Break",
      keytip: "R",
      run: (env) => env.c.print.removeBreak(),
      enabled: (env) => env.c.print.hasBreakAtCell(),
    },
    {
      id: "PageBreakResetAll",
      label: "Reset All Page Breaks",
      keytip: "A",
      run: (env) => env.c.print.resetBreaks(),
      enabled: (env) => env.c.print.setup().rowBreaks.length + env.c.print.setup().colBreaks.length > 0,
    },
    {
      id: "SheetBackground",
      label: "Background",
      labelFor: (env) => (env.c.sheet.background ? "Delete Background" : "Background"),
      description: "Choose a picture to show behind the sheet. The background shows on screen only; it isn't printed.",
      keytip: "G",
      run: (env) => {
        if (env.c.sheet.background) return env.c.sheets.setBackground(null)
        void pickPicture().then((p) => {
          if (p) env.c.sheets.setBackground(p.src)
        })
      },
    },
    {
      id: "PrintTitles",
      label: "Print Titles",
      description:
        "Specify rows and columns to repeat on each printed page, such as labeled rows or columns that appear at the top of each page.",
      keytip: "I",
      run: (env) => env.dialog("pageSetup:sheet"),
    },
    {
      id: "PageSetupPageDialog",
      label: "Page Setup",
      description: "Show the Page Setup dialog box.",
      keytip: "SP",
      run: (env) => env.dialog("pageSetup:page"),
    },
    {
      id: "PageSetupSheetDialog",
      label: "Sheet Options",
      description: "Show the Sheet tab of the Page Setup dialog box.",
      keytip: "SO",
      run: (env) => env.dialog("pageSetup:sheet"),
    },
    {
      id: "PageScaleToFitWidth",
      label: "Width",
      description: "Shrink the width of your printout to fit a maximum number of pages.",
      keytip: "W",
      field: fitField("width"),
    },
    {
      id: "PageScaleToFitHeight",
      label: "Height",
      description: "Shrink the height of your printout to fit a maximum number of pages.",
      keytip: "H",
      field: fitField("height"),
    },
    {
      id: "PageScaleToFitScale",
      label: "Scale",
      description:
        "Stretch or shrink your printout to a percentage of its actual size. The maximum width and height must be set to 'Automatic' to use this feature.",
      keytip: "SC",
      enabled: (env) => env.c.print.setup().fitTo === null,
      field: {
        value: (env) => {
          const c = env.c
          const s = c.print.setup()
          // While fitting, Excel shows the scale it will print at.
          return `${s.fitTo ? Math.round(c.print.pages().scale * 100) : s.scale}%`
        },
        commit: (env, text) => {
          const n = parseFloat(text)
          if (Number.isFinite(n)) env.c.print.setScale(n)
        },
        width: 72,
      },
    },
    {
      id: "SheetPrintGridlines",
      label: "Print",
      description: "Print the lines between rows and columns in the sheet to make reading easier.",
      keytip: "P",
      check: true,
      run: (env) => env.c.print.toggleGridlines(),
      pressed: (env) => env.c.print.setup().gridlines,
    },
    {
      id: "SheetPrintHeadings",
      label: "Print",
      description: "Print row and column headings.",
      keytip: "PH",
      check: true,
      run: (env) => env.c.print.toggleHeadings(),
      pressed: (env) => env.c.print.setup().headings,
    },
  ]
}
