import type { CellWrite } from "@cells/core"
import type { SheetController } from "@/sheet/controller"

/**
 * Made-up data for trying the grid at scale: `?demo=100k` (100,000 rows × 20 columns) or `?demo=small`. Nothing
 * here comes from a real file.
 */
export function loadDemo(controller: SheetController, name: string | null) {
  if (!name) return
  const rows = name === "small" ? 30 : name === "1m" ? 1_000_000 : 100_000
  const cols = name === "1m" ? 5 : 20
  const regions = ["North", "South", "East", "West", "Central"]
  const products = ["Widget", "Gadget", "Gizmo", "Doohickey", "Sprocket", "Flange"]
  const cells: CellWrite[] = []
  const headers = ["Order", "Region", "Product", "Units", "Price", "Total"]
  for (let c = 0; c < cols; c++) {
    cells.push({ row: 0, col: c, data: { value: headers[c] ?? `Metric ${c - headers.length + 1}` } })
  }
  for (let r = 1; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let value: string | number
      switch (c) {
        case 0:
          value = `ORD-${String(r).padStart(6, "0")}`
          break
        case 1:
          value = regions[r % regions.length]
          break
        case 2:
          value = products[(r * 7) % products.length]
          break
        case 3:
          value = (r * 37) % 250
          break
        case 4:
          value = Math.round(((r * 13) % 9000) / 3) / 100
          break
        case 5:
          value = Math.round(((r * 37) % 250) * (((r * 13) % 9000) / 3)) / 100
          break
        default:
          value = ((r * (c + 3)) % 1000) / 10
      }
      cells.push({ row: r, col: c, data: { value } })
    }
  }
  controller.wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load", history: false })
}
