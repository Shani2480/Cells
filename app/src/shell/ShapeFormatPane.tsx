import { colorCss, isChart, type LineEnd, type ShapeColor, type SheetObject, type SheetShape } from "@cells/core"
import { X } from "lucide-react"
import { shapeColorOf } from "@/ribbon/colors"
import { DASH_NAMES } from "@/ribbon/object-commands"
import type { SheetController } from "@/sheet/controller"
import { Check, Choice, ColorButton, NumberInput, Row, Section } from "./ChartFormatPane"

const ARROWS: [LineEnd, string][] = [
  ["none", "None"],
  ["triangle", "Arrow"],
  ["arrow", "Open Arrow"],
  ["stealth", "Stealth Arrow"],
  ["diamond", "Diamond Arrow"],
  ["oval", "Oval Arrow"],
]

const INCH = 96
const round = (n: number, places = 2) => Math.round(n * 10 ** places) / 10 ** places

/**
 * Excel's Format Shape pane for the selected shapes: Shape Options (fill, line, shadow, size and how it moves with
 * cells) and Text Options (the text's colour, where it sits, its direction, margins, wrapping and the shape fitting
 * it). It shows the first selected shape; every change applies to all of them and is one undo step.
 */
export function ShapeFormatPane({ controller, onDone }: { controller: SheetController; onDone: () => void }) {
  const objects = controller.objects
  const shapes = objects.drawings().filter((o): o is SheetShape => o.kind === "shape")
  const close = () => {
    objects.closeFormatPane()
    onDone()
  }
  const first = shapes[0]
  if (!first) return null
  const theme = controller.wb.theme
  const tab = objects.formatPaneTab
  const lines = shapes.every((s) => s.connector)
  const hex = (c: ShapeColor | null | undefined) => (c ? colorCss(c, theme).hex : undefined)
  const change = (label: string, f: (s: SheetShape) => void) =>
    objects.update(
      label,
      (o: SheetObject) => {
        if (!isChart(o) && o.kind === "shape") f(o)
      },
      shapes.map((s) => s.id),
    )
  const box = objects.box(first)
  const fill = first.fill
  const line = first.line
  const text = first.text
  const ins = text?.insets ?? { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }
  const runColor = text?.paragraphs.flatMap((p) => p.runs).find((r) => r.color)?.color
  return (
    <aside
      className="comments-pane chart-format-pane shape-format-pane"
      aria-label="Format Shape"
      onKeyDown={(e) => {
        if (e.key === "Escape") {
          e.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Format Shape</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.75} />
        </button>
      </header>
      <div className="fc-tabs sfp-tabs" role="tablist">
        {(
          [
            ["shape", "Shape Options"],
            ["text", "Text Options"],
          ] as const
        ).map(([id, label]) => (
          <button
            key={id}
            type="button"
            role="tab"
            className="fc-tab"
            aria-selected={tab === id}
            disabled={id === "text" && lines}
            onClick={() => objects.openFormatPane(id)}
          >
            {label}
          </button>
        ))}
      </div>
      <div className="cfp-body">
        {tab === "shape" || lines ? (
          <>
            {!lines ? (
              <Section title="Fill">
                <Choice
                  name="sfp-fill"
                  value={fill.kind}
                  options={[
                    ["none", "No fill"],
                    ["solid", "Solid fill"],
                  ]}
                  onChange={(v) =>
                    change("Shape Fill", (s) => {
                      s.fill =
                        v === "none"
                          ? { kind: "none" }
                          : { kind: "solid", color: s.fill.kind === "solid" ? s.fill.color : { theme: "accent1" } }
                    })
                  }
                />
                {fill.kind === "solid" ? (
                  <>
                    <Row label="Color">
                      <ColorButton
                        label="Fill color"
                        color={hex(fill.color)}
                        theme={theme}
                        onPick={(c) =>
                          change("Shape Fill", (s) => {
                            s.fill = c ? { kind: "solid", color: shapeColorOf(c) } : { kind: "none" }
                          })
                        }
                      />
                    </Row>
                    <Row label="Transparency">
                      <NumberInput
                        label="Fill transparency"
                        unit="%"
                        min={0}
                        max={100}
                        value={Math.round((1 - (fill.color.alpha ?? 1)) * 100)}
                        onCommit={(v) =>
                          change("Shape Fill", (s) => {
                            if (s.fill.kind !== "solid") return
                            const alpha = 1 - (v ?? 0) / 100
                            s.fill = { kind: "solid", color: { ...s.fill.color, alpha: alpha < 1 ? alpha : undefined } }
                            if (s.fill.color.alpha === undefined) delete s.fill.color.alpha
                          })
                        }
                      />
                    </Row>
                  </>
                ) : null}
              </Section>
            ) : null}
            <Section title="Line">
              <Choice
                name="sfp-line"
                value={line.color ? "solid" : "none"}
                options={[
                  ["none", "No line"],
                  ["solid", "Solid line"],
                ]}
                onChange={(v) =>
                  change("Shape Outline", (s) => {
                    s.line = {
                      ...s.line,
                      color: v === "none" ? null : (s.line.color ?? { theme: "accent1", lumMod: 0.5 }),
                      width: s.line.width || 0.75,
                    }
                  })
                }
              />
              {line.color ? (
                <>
                  <Row label="Color">
                    <ColorButton
                      label="Line color"
                      color={hex(line.color)}
                      theme={theme}
                      onPick={(c) =>
                        change("Shape Outline", (s) => {
                          s.line = { ...s.line, color: c ? shapeColorOf(c) : null }
                        })
                      }
                    />
                  </Row>
                  <Row label="Width">
                    <NumberInput
                      label="Line width"
                      unit="pt"
                      min={0.25}
                      max={1584}
                      step={0.25}
                      value={line.width}
                      onCommit={(v) =>
                        change("Outline Weight", (s) => {
                          s.line = { ...s.line, width: v ?? 0.75 }
                        })
                      }
                    />
                  </Row>
                  <Row label="Dash type">
                    <select
                      className="dialog-field"
                      aria-label="Dash type"
                      value={line.dash ?? "solid"}
                      onChange={(e) =>
                        change("Outline Dash", (s) => {
                          s.line = { ...s.line, dash: e.target.value as (typeof DASH_NAMES)[number][0] }
                        })
                      }
                    >
                      {DASH_NAMES.map(([d, name]) => (
                        <option key={d} value={d}>
                          {name}
                        </option>
                      ))}
                    </select>
                  </Row>
                  {lines
                    ? (["head", "tail"] as const).map((end) => (
                        <Row key={end} label={end === "head" ? "Begin arrow type" : "End arrow type"}>
                          <select
                            className="dialog-field"
                            aria-label={end === "head" ? "Begin arrow type" : "End arrow type"}
                            value={line[end] ?? "none"}
                            onChange={(e) =>
                              change("Arrows", (s) => {
                                const v = e.target.value as LineEnd
                                s.line = { ...s.line, [end]: v === "none" ? undefined : v }
                                if (v === "none") delete s.line[end]
                              })
                            }
                          >
                            {ARROWS.map(([v, name]) => (
                              <option key={v} value={v}>
                                {name}
                              </option>
                            ))}
                          </select>
                        </Row>
                      ))
                    : null}
                </>
              ) : null}
            </Section>
            {!lines ? (
              <Section title="Effects">
                <Check
                  label="Shadow"
                  checked={!!first.shadow}
                  onChange={(on) =>
                    change("Shape Effects", (s) => {
                      if (on) s.shadow = true
                      else delete s.shadow
                    })
                  }
                />
              </Section>
            ) : null}
            <Section title="Size">
              <Row label="Height">
                <NumberInput
                  label="Shape height"
                  unit="in"
                  min={0}
                  step={0.01}
                  value={round(box.h / INCH)}
                  onCommit={(v) => v !== undefined && objects.resize(null, v * INCH)}
                />
              </Row>
              <Row label="Width">
                <NumberInput
                  label="Shape width"
                  unit="in"
                  min={0}
                  step={0.01}
                  value={round(box.w / INCH)}
                  onCommit={(v) => v !== undefined && objects.resize(v * INCH, null)}
                />
              </Row>
              {!lines ? (
                <Row label="Rotation">
                  <NumberInput
                    label="Rotation"
                    unit="°"
                    min={-3600}
                    max={3600}
                    value={round(first.rotation ?? 0, 1)}
                    onCommit={(v) => objects.rotate((((v ?? 0) % 360) + 360) % 360, true)}
                  />
                </Row>
              ) : null}
            </Section>
            <Section title="Properties">
              <Choice
                name="sfp-placement"
                value={first.placement ?? "both"}
                options={[
                  ["both", "Move and size with cells"],
                  ["move", "Move but don't size with cells"],
                  ["fixed", "Don't move or size with cells"],
                ]}
                onChange={(v) =>
                  change("Properties", (s) => {
                    if (v === "both") delete s.placement
                    else s.placement = v as "move" | "fixed"
                  })
                }
              />
            </Section>
          </>
        ) : (
          <>
            <Section title="Text Fill">
              <Row label="Color">
                <ColorButton
                  label="Text color"
                  color={hex(runColor)}
                  automatic
                  theme={theme}
                  onPick={(c) =>
                    change("Text Fill", (s) => {
                      if (!s.text) return
                      const color = c ? shapeColorOf(c) : undefined
                      if (s.wordArt?.fill && color) s.wordArt = { ...s.wordArt, fill: { kind: "solid", color } }
                      for (const p of s.text.paragraphs)
                        for (const r of p.runs) {
                          if (color) r.color = color
                          else delete r.color
                        }
                    })
                  }
                />
              </Row>
            </Section>
            <Section title="Text Box">
              <Row label="Vertical alignment">
                <select
                  className="dialog-field"
                  aria-label="Vertical alignment"
                  value={text?.anchor ?? "t"}
                  onChange={(e) =>
                    change("Alignment", (s) => {
                      s.text = { ...(s.text ?? { paragraphs: [{ runs: [] }] }), anchor: e.target.value as "t" }
                    })
                  }
                >
                  <option value="t">Top</option>
                  <option value="ctr">Middle</option>
                  <option value="b">Bottom</option>
                </select>
              </Row>
              <Row label="Text direction">
                <select
                  className="dialog-field"
                  aria-label="Text direction"
                  value={text?.vertical ?? "horz"}
                  onChange={(e) =>
                    change("Text Direction", (s) => {
                      const v = e.target.value as "horz" | "vert" | "vert270"
                      s.text = { ...(s.text ?? { paragraphs: [{ runs: [] }] }), vertical: v === "horz" ? undefined : v }
                      if (v === "horz") delete s.text.vertical
                    })
                  }
                >
                  <option value="horz">Horizontal</option>
                  <option value="vert">Rotate all text 90°</option>
                  <option value="vert270">Rotate all text 270°</option>
                </select>
              </Row>
              <Check
                label="Resize shape to fit text"
                checked={!!text?.autoFit}
                onChange={(on) =>
                  change("Autofit", (s) => {
                    s.text = { ...(s.text ?? { paragraphs: [{ runs: [] }] }), autoFit: on || undefined }
                    if (!on) delete s.text.autoFit
                  })
                }
              />
              {(["l", "r", "t", "b"] as const).map((side) => (
                <Row key={side} label={`${{ l: "Left", r: "Right", t: "Top", b: "Bottom" }[side]} margin`}>
                  <NumberInput
                    label={`${{ l: "Left", r: "Right", t: "Top", b: "Bottom" }[side]} margin`}
                    unit="in"
                    min={0}
                    step={0.01}
                    value={round(ins[side] / INCH)}
                    onCommit={(v) =>
                      change("Margins", (s) => {
                        const now = s.text?.insets ?? { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }
                        s.text = {
                          ...(s.text ?? { paragraphs: [{ runs: [] }] }),
                          insets: { ...now, [side]: (v ?? 0) * INCH },
                        }
                      })
                    }
                  />
                </Row>
              ))}
              <Check
                label="Wrap text in shape"
                checked={text?.wrap !== false}
                onChange={(on) =>
                  change("Wrap Text", (s) => {
                    s.text = { ...(s.text ?? { paragraphs: [{ runs: [] }] }), wrap: on ? undefined : false }
                    if (on) delete s.text.wrap
                  })
                }
              />
            </Section>
          </>
        )}
      </div>
    </aside>
  )
}
