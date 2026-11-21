import { displayText, findTable, type DataModel, type ModelRelationship, type Workbook } from "@cells/core"
import { ArrowRight, Boxes, Table2 } from "lucide-react"
import { useMemo, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * Data › Relationships and Manage Data Model: Excel's Data Model as Cells sees it. Excel keeps the model in a part
 * only Excel can write, and trusts only that part, so both show the model and leave changes to Excel.
 */

const EXCEL_ONLY =
  "Excel keeps the Data Model in a form only Excel can change, so it's shown here and changed in Excel."

/** A model table's columns: those of the workbook table it loads, else the ones its relationships name. */
export function modelColumns(wb: Workbook, model: DataModel, table: string): string[] {
  const t = model.tables.find((x) => x.name === table)
  const found = t?.source ? findTable(wb.sheets, t.source) : null
  if (found) return found.table.columns.map((c) => c.name)
  const named = model.relationships.flatMap((r) => [
    ...(r.fromTable === table ? [r.fromColumn] : []),
    ...(r.toTable === table ? [r.toColumn] : []),
  ])
  return [...new Set(named)]
}

const relText = (r: ModelRelationship) => `${r.fromTable} (${r.fromColumn})`
const relatedText = (r: ModelRelationship) => `${r.toTable} (${r.toColumn})`

/** Manage Relationships: the model's relationships, as Excel lists them. */
export function RelationshipsDialog({
  controller,
  onClose,
  onManage,
}: {
  controller: SheetController
  onClose: () => void
  onManage: () => void
}) {
  const model = controller.wb.dataModel
  const list = model?.relationships ?? []
  const [picked, setPicked] = useState(0)
  return (
    <Dialog
      title="Manage Relationships"
      onCancel={onClose}
      onAccept={onClose}
      width={620}
      acceptLabel="Close"
      cancelLabel={null}
    >
      <div className="rel-dialog">
        <div className="rel-table" role="grid" aria-label="Relationships">
          <div className="rel-row rel-head" role="row">
            <span role="columnheader">Status</span>
            <span role="columnheader">Table</span>
            <span role="columnheader">Related Lookup Table</span>
          </div>
          {list.length ? (
            list.map((r, i) => (
              <button
                key={i}
                type="button"
                role="row"
                aria-selected={i === picked}
                className={`rel-row${i === picked ? " is-selected" : ""}`}
                onClick={() => setPicked(i)}
              >
                <span role="gridcell">Active</span>
                <span role="gridcell">{relText(r)}</span>
                <span role="gridcell">{relatedText(r)}</span>
              </button>
            ))
          ) : (
            <p className="dialog-note rel-empty">
              {model
                ? "The Data Model has no relationships."
                : "This workbook has no Data Model. In Excel, tables added to the Data Model can be related here, so a PivotTable can combine them."}
            </p>
          )}
        </div>
        <div className="rel-buttons">
          {["New...", "Edit...", "Activate", "Deactivate", "Delete"].map((label) => (
            <button key={label} type="button" className="dialog-btn" disabled title={EXCEL_ONLY}>
              {label}
            </button>
          ))}
          <button
            type="button"
            className="dialog-btn"
            onClick={() => {
              onClose()
              onManage()
            }}
          >
            Data Model...
          </button>
        </div>
      </div>
      <p className="dialog-note">{EXCEL_ONLY}</p>
    </Dialog>
  )
}

const ROWS = 500

/** Manage Data Model: the model's tables (Data View) and how they relate (Diagram View). */
export function DataModelDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const wb = controller.wb
  const model = wb.dataModel
  const [view, setView] = useState<"data" | "diagram">("diagram")
  const [table, setTable] = useState(model?.tables[0]?.name ?? "")
  return (
    <Dialog title="Data Model" onCancel={onClose} onAccept={onClose} width={900} acceptLabel="Close" cancelLabel={null}>
      {!model || !model.tables.length ? (
        <div className="dm-empty">
          <Boxes size={40} strokeWidth={1.2} aria-hidden />
          <p>
            This workbook has no Data Model. In Excel, a table added to the Data Model (or a PivotTable made with Add
            this data to the Data Model) starts one; Cells then shows its tables and relationships here.
          </p>
        </div>
      ) : (
        <div className="dm">
          <div className="dm-views" role="tablist" aria-label="View">
            <button type="button" role="tab" aria-selected={view === "diagram"} onClick={() => setView("diagram")}>
              Diagram View
            </button>
            <button type="button" role="tab" aria-selected={view === "data"} onClick={() => setView("data")}>
              Data View
            </button>
          </div>
          {view === "diagram" ? (
            <Diagram wb={wb} model={model} onOpen={(name) => (setTable(name), setView("data"))} />
          ) : (
            <DataView wb={wb} model={model} table={table} onTable={setTable} />
          )}
          <p className="dialog-note">{EXCEL_ONLY}</p>
        </div>
      )}
    </Dialog>
  )
}

function DataView({
  wb,
  model,
  table,
  onTable,
}: {
  wb: Workbook
  model: DataModel
  table: string
  onTable: (name: string) => void
}) {
  const t = model.tables.find((x) => x.name === table) ?? model.tables[0]
  const found = t.source ? findTable(wb.sheets, t.source) : null
  const rows = useMemo(() => {
    if (!found) return []
    const sheet = wb.sheets[found.sheet]
    const r = found.table.range
    const first = found.table.header ? r.r0 + 1 : r.r0
    const last = Math.min(found.table.totals ? r.r1 - 1 : r.r1, first + ROWS - 1)
    const out: { text: string; number: boolean }[][] = []
    for (let row = first; row <= last; row++) {
      const line: { text: string; number: boolean }[] = []
      for (let col = r.c0; col <= r.c1; col++)
        line.push({
          text: displayText(sheet.cells.value(row, col), wb.styles.get(sheet.styleAt(row, col)).numberFormat),
          number: typeof sheet.cells.value(row, col) === "number",
        })
      out.push(line)
    }
    return out
  }, [found, wb])
  const total = found
    ? found.table.range.r1 - found.table.range.r0 + (found.table.header ? 0 : 1) - (found.table.totals ? 1 : 0)
    : 0
  return (
    <div className="dm-data">
      <div className="qe-preview dm-grid" aria-label={`${t.name} data`}>
        {found ? (
          <table className="qe-grid">
            <thead>
              <tr>
                {found.table.columns.map((c) => (
                  <th key={c.name}>{c.name}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((line, i) => (
                <tr key={i}>
                  {line.map((v, k) => (
                    <td key={k} className={v.number ? "is-number" : undefined}>
                      {v.text}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <p className="dialog-note">
            {t.name} comes from outside this workbook ({t.connection}); Excel keeps its data in the model.
          </p>
        )}
      </div>
      <div className="dm-tabs" role="tablist" aria-label="Tables">
        {model.tables.map((x) => (
          <button
            key={x.name}
            type="button"
            role="tab"
            aria-selected={x.name === t.name}
            onClick={() => onTable(x.name)}
          >
            {x.name}
          </button>
        ))}
        <span className="dm-count">{found ? `${total.toLocaleString()} rows` : ""}</span>
      </div>
    </div>
  )
}

const BOX_W = 190
const GAP_X = 90
const GAP_Y = 40
const HEAD = 30
const ROW_H = 22

/** The model's tables as boxes, a line from each relationship's column to the related lookup column. */
function Diagram({ wb, model, onOpen }: { wb: Workbook; model: DataModel; onOpen: (name: string) => void }) {
  const boxes = useMemo(() => {
    const perRow = Math.max(1, Math.min(4, Math.ceil(Math.sqrt(model.tables.length))))
    let y = 0
    const out: { name: string; source: string | null; columns: string[]; x: number; y: number; h: number }[] = []
    for (let i = 0; i < model.tables.length; i += perRow) {
      const row = model.tables.slice(i, i + perRow).map((t, k) => {
        const columns = modelColumns(wb, model, t.name)
        return {
          name: t.name,
          source: t.source,
          columns,
          x: k * (BOX_W + GAP_X),
          y,
          h: HEAD + columns.length * ROW_H + 8,
        }
      })
      out.push(...row)
      y += Math.max(...row.map((b) => b.h)) + GAP_Y
    }
    return out
  }, [wb, model])
  const width = Math.max(...boxes.map((b) => b.x + BOX_W)) + 2
  const height = Math.max(...boxes.map((b) => b.y + b.h)) + 2
  const anchor = (table: string, column: string, toward: number) => {
    const b = boxes.find((x) => x.name === table)
    if (!b) return null
    const i = Math.max(0, b.columns.indexOf(column))
    const right = toward > b.x + BOX_W / 2
    return { x: right ? b.x + BOX_W : b.x, y: b.y + HEAD + i * ROW_H + ROW_H / 2, right }
  }
  return (
    <div className="dm-diagram" aria-label="Diagram">
      <svg
        width={width + 40}
        height={height + 40}
        viewBox={`-20 -20 ${width + 40} ${height + 40}`}
        role="img"
        aria-label="Tables and relationships"
      >
        {model.relationships.map((r, i) => {
          const to = boxes.find((b) => b.name === r.toTable)
          const from = boxes.find((b) => b.name === r.fromTable)
          if (!to || !from) return null
          const a = anchor(r.fromTable, r.fromColumn, to.x + BOX_W / 2)
          const b = anchor(r.toTable, r.toColumn, from.x + BOX_W / 2)
          if (!a || !b) return null
          const bend = Math.max(30, Math.abs(b.x - a.x) / 2)
          const d = `M${a.x} ${a.y} C${a.x + (a.right ? bend : -bend)} ${a.y}, ${b.x + (b.right ? bend : -bend)} ${b.y}, ${b.x} ${b.y}`
          return (
            <g key={i} className="dm-link">
              <title>{`${relText(r)} → ${relatedText(r)}`}</title>
              <path d={d} />
              <text x={a.x + (a.right ? 6 : -6)} y={a.y - 5} textAnchor={a.right ? "start" : "end"}>
                *
              </text>
              <text x={b.x + (b.right ? 6 : -6)} y={b.y - 5} textAnchor={b.right ? "start" : "end"}>
                1
              </text>
            </g>
          )
        })}
        {boxes.map((b) => {
          const keys = new Set(
            model.relationships.flatMap((r) => [
              ...(r.fromTable === b.name ? [r.fromColumn] : []),
              ...(r.toTable === b.name ? [r.toColumn] : []),
            ]),
          )
          return (
            <g
              key={b.name}
              className="dm-box"
              transform={`translate(${b.x} ${b.y})`}
              onDoubleClick={() => onOpen(b.name)}
              tabIndex={0}
              role="button"
              aria-label={`${b.name}, ${b.columns.length} columns`}
              onKeyDown={(e) => e.key === "Enter" && onOpen(b.name)}
            >
              <rect className="dm-box-body" width={BOX_W} height={b.h} rx={6} />
              <path
                className="dm-box-head"
                d={`M0 6a6 6 0 0 1 6-6h${BOX_W - 12}a6 6 0 0 1 6 6v${HEAD - 6}h-${BOX_W}z`}
              />
              <text className="dm-box-title" x={12} y={20}>
                {b.name}
              </text>
              {b.columns.map((c, i) => (
                <text key={c} className={`dm-box-col${keys.has(c) ? " is-key" : ""}`} x={14} y={HEAD + i * ROW_H + 15}>
                  {c}
                </text>
              ))}
            </g>
          )
        })}
      </svg>
      <p className="dm-legend">
        <Table2 size={14} aria-hidden /> Double-click a table to see its data. <ArrowRight size={14} aria-hidden /> Each
        line runs from a table's column (*, many) to the lookup table's column (1).
      </p>
    </div>
  )
}
