import { pickPicture } from "@/drawing/pictures"
import { HF_ELEMENTS, type HFOption, type HFWhere } from "@/sheet/header-footer"
import { headerPresets } from "@/shell/PageSetupDialog"
import type { Command, RibbonEnv } from "./registry"

const editingHF = (env: RibbonEnv) => env.c.headerFooter.editing !== null

/** Header & Footer › Header or Footer: Excel's ready-made ones, as a list. */
function presetMenu(where: HFWhere): Command["menu"] {
  return [
    {
      panel: (env, close) => {
        const presets = headerPresets(env.c.sheet.name, env.c.printFile.fileName || "Book1")
        return (
          <div className="hf-presets" role="menu" aria-label={where === "header" ? "Header" : "Footer"}>
            {presets.map((p) => (
              <button
                key={p.label}
                type="button"
                role="menuitem"
                className="hf-preset"
                onClick={() => {
                  close()
                  env.c.headerFooter.usePreset(where, p.value)
                }}
              >
                {p.label}
              </button>
            ))}
          </div>
        )
      },
    },
  ]
}

/** The Header & Footer tab's elements: [id, label, keytip, what it inserts, description]. */
const ELEMENTS: [string, string, string, string, string][] = [
  ["HFPageNumber", "Page Number", "N", HF_ELEMENTS.pageNumber, "Add the page number to the header or footer."],
  ["HFNumberOfPages", "Number of Pages", "NP", HF_ELEMENTS.pages, "Add the total number of pages."],
  ["HFCurrentDate", "Current Date", "D", HF_ELEMENTS.date, "Add the date it's printed on."],
  ["HFCurrentTime", "Current Time", "T", HF_ELEMENTS.time, "Add the time it's printed at."],
  ["HFFilePath", "File Path", "P", HF_ELEMENTS.path, "Add the workbook's folder and file name."],
  ["HFFileName", "File Name", "F", HF_ELEMENTS.file, "Add the workbook's file name."],
  ["HFSheetName", "Sheet Name", "S", HF_ELEMENTS.sheet, "Add the sheet's name."],
]

/** Header & Footer › Options, as Excel words them. */
const OPTIONS: [string, string, string, HFOption, string][] = [
  [
    "HFDifferentFirstPage",
    "Different First Page",
    "A",
    "differentFirst",
    "Give the first page a header and footer of its own.",
  ],
  [
    "HFDifferentOddEven",
    "Different Odd & Even Pages",
    "V",
    "differentOddEven",
    "Give even pages a header and footer of their own, different from odd pages.",
  ],
  [
    "HFScaleWithDocument",
    "Scale with Document",
    "L",
    "scaleWithDoc",
    "Make the header and footer text bigger or smaller with the sheet when it's printed at another size.",
  ],
  [
    "HFAlignWithMargins",
    "Align with Page Margins",
    "M",
    "alignWithMargins",
    "Line the header and footer up with the page's margins.",
  ],
]

/** Insert › Links and Text and Symbols: Link (Ctrl+K), Header & Footer and Symbol. */
export function linkCommands(): Command[] {
  return [
    {
      id: "InsertLinkGallery",
      label: "Link",
      description:
        "Create a link in your document for quick access to webpages and files, a place in this workbook, or an e-mail address.",
      shortcut: "Ctrl+K",
      keytip: "I",
      run: (env) => env.dialog("insertLink"),
    },
    {
      id: "HeaderFooterInsert",
      label: "Header & Footer",
      description:
        "Edit the header or footer of your document. They repeat at the top and bottom of each printed page.",
      keytip: "H",
      // On the page, in Page Layout view (as Excel does).
      run: (env) => env.c.headerFooter.open("header"),
    },
    // ── Header & Footer (while one is edited) ──
    {
      id: "HFHeaderMenu",
      label: "Header",
      description: "Pick a ready-made header, such as the page number or the sheet's name.",
      keytip: "H",
      enabled: editingHF,
      menu: presetMenu("header"),
    },
    {
      id: "HFFooterMenu",
      label: "Footer",
      description: "Pick a ready-made footer, such as the page number or the sheet's name.",
      keytip: "O",
      enabled: editingHF,
      menu: presetMenu("footer"),
    },
    ...ELEMENTS.map(([id, label, keytip, code, description]): Command => ({
      id,
      label,
      description,
      keytip,
      enabled: editingHF,
      run: (env) => env.c.headerFooter.insert(code),
    })),
    {
      id: "HFPicture",
      label: "Picture",
      description: "Add a picture to the header or footer, such as a logo. It prints where &[Picture] is.",
      keytip: "PC",
      enabled: editingHF,
      run: (env) =>
        void pickPicture().then((p) => {
          if (p) env.c.headerFooter.insertPicture(p)
        }),
    },
    {
      id: "HFFormatPicture",
      label: "Format Picture",
      description: "Size, crop and adjust the colour, brightness and contrast of the header or footer's picture.",
      keytip: "FP",
      enabled: (env) => env.c.headerFooter.currentPicture() !== null,
      run: (env) => env.dialog("headerPicture"),
    },
    {
      id: "HFGoToHeader",
      label: "Go to Header",
      description: "Edit the header of this page.",
      keytip: "GH",
      enabled: (env) => env.c.headerFooter.editing?.where === "footer",
      run: (env) => env.c.headerFooter.goTo("header"),
    },
    {
      id: "HFGoToFooter",
      label: "Go to Footer",
      description: "Edit the footer of this page.",
      keytip: "GF",
      enabled: (env) => env.c.headerFooter.editing?.where === "header",
      run: (env) => env.c.headerFooter.goTo("footer"),
    },
    ...OPTIONS.map(([id, label, keytip, option, description]): Command => ({
      id,
      label,
      description,
      keytip,
      check: true,
      enabled: editingHF,
      run: (env) => env.c.headerFooter.toggleOption(option),
      pressed: (env) => env.c.sheet.pageSetup[option],
    })),
    {
      id: "SymbolInsert",
      label: "Symbol",
      description:
        "Add characters that aren't on your keyboard, such as copyright symbols, arrows and accented letters.",
      keytip: "U",
      run: (env) => env.dialog("symbol"),
    },
  ]
}
