import { describe, expect, it } from "vitest"
import { isChart, type SheetShape } from "@cells/core"
import { objectSvg, themeEffectFilter } from "@/drawing/render"
import { applyShapeStyle, SHAPE_STYLES } from "@/drawing/shape-styles"
import { SheetController } from "@/sheet/controller"
import { OFFICE_EFFECT_SCHEMES } from "./theme-commands"

const glowEdge = OFFICE_EFFECT_SCHEMES.find((s) => s.label === "Glow Edge")!.effects!

describe("Shape Styles and Theme Effects", () => {
  it("has Excel's 6 rows of styles in dark 1 and the six accents, named as Excel names them", () => {
    expect(SHAPE_STYLES).toHaveLength(6)
    expect(SHAPE_STYLES.every((row) => row.length === 7)).toBe(true)
    expect(SHAPE_STYLES[0][1].name).toBe("Colored Outline - Accent 1")
    expect(SHAPE_STYLES[5][6]).toMatchObject({ name: "Intense Effect - Accent 6", effect: 3 })
    expect(OFFICE_EFFECT_SCHEMES.map((s) => s.label).slice(0, 3)).toEqual([
      "Office",
      "Office 2007 - 2010",
      "Subtle Solids",
    ])
    expect(OFFICE_EFFECT_SCHEMES).toHaveLength(15)
  })

  it("a style sets the shape's fill, outline, text colour and theme effect; its own shadow goes", () => {
    const c = new SheetController()
    const o = c.objects.insertShape("rect")
    c.objects.update("Shape Style", (x) => {
      if (!isChart(x) && x.kind === "shape") {
        x.shadow = true
        applyShapeStyle(x, SHAPE_STYLES[5][2])
      }
    })
    const s = c.sheet.objects.find((x) => x.id === o.id) as SheetShape
    expect(s.fill).toEqual({ kind: "solid", color: { theme: "accent2" } })
    expect(s.line.color).toBeNull()
    expect(s.text!.paragraphs[0].runs[0].color).toEqual({ theme: "lt1" })
    expect(s.themeEffect).toBe(3)
    expect(s.shadow).toBeUndefined()
    c.undo()
    expect((c.sheet.objects[0] as SheetShape).themeEffect).toBeUndefined()
  })

  it("draws a theme effect from the workbook's effects: Office's intense shadow, nothing for its subtle one", () => {
    const c = new SheetController()
    const o = { ...c.objects.insertShape("rect"), themeEffect: 3 as const }
    const svg = objectSvg(o, 100, 60, c.wb.theme)
    expect(svg).toContain(`<filter id="o${o.id}fx"`)
    expect(svg).toContain('stdDeviation="3"')
    expect(themeEffectFilter("x", 1, c.wb.theme, "#123456")).toBe("")
    // Frosted Glass's intense effect is a glow in the shape's own colour (its bevel shows in Excel only).
    const frosted = OFFICE_EFFECT_SCHEMES.find((s) => s.label === "Frosted Glass")!.effects!
    const glow = themeEffectFilter("x", 3, { ...c.wb.theme, effects: frosted }, "#123456")
    expect(glow).toContain("feMorphology")
    expect(glow).toContain('flood-color="#123456"')
    // Extreme Shadow's is a long, soft shadow.
    const extreme = OFFICE_EFFECT_SCHEMES.find((s) => s.label === "Extreme Shadow")!.effects!
    expect(themeEffectFilter("x", 3, { ...c.wb.theme, effects: extreme }, "#123456")).toContain('dy="12"')
  })

  it("picking Effects changes the theme as one step to undo", () => {
    const c = new SheetController()
    c.transact([{ kind: "theme", theme: { ...c.wb.theme, effects: glowEdge } }], { label: "Theme Effects" })
    expect(c.wb.theme.effects?.name).toBe("Glow Edge")
    c.undo()
    expect(c.wb.theme.effects).toBeUndefined()
  })
})
