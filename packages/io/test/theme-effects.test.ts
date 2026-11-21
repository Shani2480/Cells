// Theme Effects (Page Layout › Themes › Effects) and shapes' theme effects (Shape Styles), in xlsx.
import { describe, expect, it } from "vitest"
import { effectLooks, OFFICE_EFFECTS, Workbook, type SheetShape } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

/** Office's Glow Edge and Inset format schemes, as their .eftx files have them. */
const GLOW_EDGE =
  '<a:fmtScheme name="Glow Edge"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="55000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="50800" dist="25400" algn="bl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst/><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="brightRoom" dir="tl"><a:rot lat="0" lon="0" rev="1800000"/></a:lightRig></a:scene3d><a:sp3d contourW="10160" prstMaterial="dkEdge"><a:bevelT w="38100" h="50800" prst="angle"/><a:contourClr><a:schemeClr val="phClr"><a:shade val="40000"/><a:satMod val="150000"/></a:schemeClr></a:contourClr></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
const INSET =
  '<a:fmtScheme name="Inset"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="true"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="20000"/><a:satMod val="180000"/><a:lumMod val="98000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="30000"/><a:satMod val="260000"/><a:lumMod val="84000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="100000"/><a:satMod val="110000"/><a:lumMod val="100000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5040000" scaled="true"/></a:gradFill><a:gradFill rotWithShape="true"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="75000"/><a:satMod val="120000"/><a:lumMod val="90000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="false"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="15875" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="22225" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="50800" dist="25400" dir="5400000" rotWithShape="false"><a:srgbClr val="000000"><a:alpha val="28000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="tl"><a:rot lat="0" lon="0" rev="20400000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="50800" h="12700" prst="softRound"/></a:sp3d></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="44450" dist="50800" dir="5400000" sx="96000" rotWithShape="false"><a:srgbClr val="000000"><a:alpha val="34000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="tl"><a:rot lat="0" lon="0" rev="20400000"/></a:lightRig></a:scene3d><a:sp3d contourW="15875"><a:bevelT w="101600" h="25400" prst="softRound"/><a:contourClr><a:schemeClr val="phClr"><a:shade val="30000"/></a:schemeClr></a:contourClr></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'

const shape = (id: number, themeEffect?: 1 | 2 | 3): SheetShape => ({
  kind: "shape",
  id,
  name: `Rectangle ${id}`,
  anchor: { from: { col: 1, row: 1, x: 0, y: 0 }, to: { col: 4, row: 5, x: 0, y: 0 } },
  geometry: "rect",
  fill: { kind: "solid", color: { theme: "accent2" } },
  line: { color: null, width: 1 },
  ...(themeEffect ? { themeEffect } : {}),
})

const text = async (bytes: Uint8Array, re: RegExp) =>
  new TextDecoder().decode((await readZip(bytes)).find((e) => re.test(e.name))!.data)

describe("theme effects", () => {
  it("reads the effect styles Cells draws: Office's third is a soft drop shadow", () => {
    const [subtle, moderate, intense] = effectLooks(OFFICE_EFFECTS.scheme)
    expect(subtle.shadow).toBeUndefined()
    expect(moderate.shadow).toBeUndefined()
    expect(intense.shadow).toMatchObject({ dx: 0, dy: 2, blur: 6, alpha: 0.63, color: "#000000" })
    expect(effectLooks(GLOW_EDGE)[2].other).toContain("bevel")
    expect(effectLooks(INSET)[2].shadow?.dy).toBeGreaterThan(0)
  })

  it("writes the scheme picked and a shape's theme effect, and reads both back", async () => {
    const wb = new Workbook(["Sheet1"])
    wb.theme = { ...wb.theme, effects: { name: "Glow Edge", scheme: GLOW_EDGE } }
    wb.sheets[0].objects.push(shape(1, 3), shape(2))
    const saved = await saveXlsx(wb)
    expect(await text(saved, /theme\/theme1\.xml$/)).toContain('<a:fmtScheme name="Glow Edge">')
    const drawing = await text(saved, /drawings\/drawing1\.xml$/)
    expect(drawing.match(/<xdr:style>/g)).toHaveLength(1)
    expect(drawing).toContain('<a:effectRef idx="3"><a:schemeClr val="accent2"/></a:effectRef>')
    const back = (await openXlsx(saved, "fx.xlsx")).workbook
    expect(back.theme.effects).toEqual({ name: "Glow Edge", scheme: GLOW_EDGE })
    const [a, b] = back.sheets[0].objects as SheetShape[]
    expect(a.themeEffect).toBe(3)
    expect(b.themeEffect).toBeUndefined()
  })

  it("a theme Cells writes whole gets Office's own format scheme, read back as the default", async () => {
    const wb = new Workbook(["Sheet1"])
    wb.theme = { ...wb.theme, colors: { ...wb.theme.colors, accent1: "336699" } }
    const saved = await saveXlsx(wb)
    expect(await text(saved, /theme\/theme1\.xml$/)).toContain(OFFICE_EFFECTS.scheme)
    expect((await openXlsx(saved, "new.xlsx")).workbook.theme.effects).toBeUndefined()
  })
})
