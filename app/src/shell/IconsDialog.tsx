import { icons } from "lucide-react"
import { useMemo, useRef, useState } from "react"
import { svgToPng } from "@/drawing/pictures"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** How many icons the dialog shows at once (a search narrows them). */
const SHOWN = 240
/** An icon's size on the sheet: Excel puts icons in at one inch. */
const ICON_SIZE = 96

/** "ArrowBigDown" → "arrow big down", for searching and tooltips. */
const words = (name: string) => name.replace(/([a-z0-9])([A-Z])/g, "$1 $2").toLowerCase()

/**
 * Insert › Icons: pictures of everyday things to add to a sheet, searchable, several at once. They're Lucide's
 * open-source icons (ISC licence), which come with Cells: nothing is fetched. Each goes in as an SVG picture (with a
 * PNG for apps that can't show SVG), black like Excel's, recoloured later with Graphics Fill.
 */
export function IconsDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [query, setQuery] = useState("")
  const [picked, setPicked] = useState<string[]>([])
  const grid = useRef<HTMLDivElement>(null)
  const names = useMemo(() => Object.keys(icons), [])
  const shown = useMemo(() => {
    const q = query.trim().toLowerCase()
    return (q ? names.filter((n) => words(n).includes(q)) : names).slice(0, SHOWN)
  }, [names, query])

  const insert = async (names: string[]) => {
    for (const name of names) {
      const svg = grid.current?.querySelector(`[data-icon="${name}"] svg`)
      if (!svg) continue
      const clone = svg.cloneNode(true) as SVGSVGElement
      clone.setAttribute("xmlns", "http://www.w3.org/2000/svg")
      clone.setAttribute("width", String(ICON_SIZE))
      clone.setAttribute("height", String(ICON_SIZE))
      clone.removeAttribute("class")
      clone.removeAttribute("aria-hidden")
      const markup = clone.outerHTML.replace(/currentColor/g, "#000000")
      const svgUrl = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(markup)))}`
      let png = svgUrl
      try {
        png = await svgToPng(markup, ICON_SIZE * 2, ICON_SIZE * 2)
      } catch {
        // Without a PNG the SVG stands in (Excel 365 shows it; older apps may not).
      }
      controller.objects.insertPicture(
        png,
        { w: ICON_SIZE, h: ICON_SIZE },
        { svg: svgUrl, icon: true, altText: words(name) },
      )
    }
    onClose()
  }

  return (
    <Dialog
      title="Insert Icons"
      width={560}
      onCancel={onClose}
      onAccept={() => void insert(picked)}
      acceptLabel={picked.length > 1 ? `Insert (${picked.length})` : "Insert"}
    >
      <input
        className="dialog-field"
        aria-label="Search icons"
        placeholder="Search icons"
        data-autofocus
        value={query}
        onChange={(e) => setQuery(e.target.value)}
      />
      <div className="icons-grid" ref={grid} role="listbox" aria-label="Icons" aria-multiselectable>
        {shown.map((name) => {
          const Icon = icons[name as keyof typeof icons]
          const on = picked.includes(name)
          return (
            <button
              key={name}
              type="button"
              role="option"
              aria-selected={on}
              title={words(name)}
              data-icon={name}
              className={`icons-item${on ? " is-selected" : ""}`}
              onClick={() => setPicked((p) => (on ? p.filter((x) => x !== name) : [...p, name]))}
              onDoubleClick={() => void insert(on ? picked : [...picked, name])}
            >
              <Icon size={28} strokeWidth={1.5} aria-hidden />
            </button>
          )
        })}
      </div>
      <p className="dialog-note">
        {shown.length === SHOWN ? `Showing the first ${SHOWN}: search to find more. ` : ""}Icons from Lucide (ISC
        licence).
      </p>
    </Dialog>
  )
}
