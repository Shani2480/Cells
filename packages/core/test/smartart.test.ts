import { describe, expect, it } from "vitest"
import {
  addSmartArtNode,
  demoteSmartArtNode,
  drawnLayout,
  moveSmartArtNode,
  promoteSmartArtNode,
  removeSmartArtNode,
  smartArtColors,
  smartArtDataXml,
  smartArtLines,
  type SmartArtNode,
} from "../src"

const tree = (...texts: (string | [string, SmartArtNode[]])[]): SmartArtNode[] =>
  texts.map((t) => (typeof t === "string" ? { text: t, children: [] } : { text: t[0], children: t[1] }))
const show = (nodes: SmartArtNode[]) =>
  smartArtLines(nodes).map((l) => `${"-".repeat(l.path.length - 1)}${l.node.text}`)

describe("SmartArt text tree", () => {
  it("adds points after or before the current one", () => {
    expect(show(addSmartArtNode(tree("A", "B"), [0])!.nodes)).toEqual(["A", "", "B"])
    expect(addSmartArtNode(tree("A", "B"), [0], true)!.path).toEqual([0])
    expect(addSmartArtNode([], null)!.nodes).toHaveLength(1)
  })

  it("demotes under the point before, promotes taking the later ones along", () => {
    const d = demoteSmartArtNode(tree("A", "B", "C"), [1])!
    expect(show(d.nodes)).toEqual(["A", "-B", "C"])
    expect(d.path).toEqual([0, 0])
    expect(demoteSmartArtNode(tree("A"), [0])).toBeNull()
    const nested = tree(["A", tree("x", "y", "z")])
    const p = promoteSmartArtNode(nested, [0, 1])!
    expect(show(p.nodes)).toEqual(["A", "-x", "y", "-z"])
    expect(p.path).toEqual([1])
  })

  it("moves and removes points (the points under a removed one move up)", () => {
    expect(show(moveSmartArtNode(tree("A", "B"), [1], -1)!.nodes)).toEqual(["B", "A"])
    expect(moveSmartArtNode(tree("A", "B"), [1], 1)).toBeNull()
    expect(show(removeSmartArtNode(tree(["A", tree("x")], "B"), [0])!.nodes)).toEqual(["x", "B"])
  })

  it("writes the data part as a tree of points with their transitions", () => {
    const xml = smartArtDataXml({ layout: "urn:cells.app/smartart/layout/list", nodes: tree(["A & B", tree("c")], "") })
    expect(xml).toContain("<a:t>A &amp; B</a:t>")
    expect(xml.match(/type="sibTrans"/g)).toHaveLength(3)
    expect(xml).toMatch(/<dgm:cxn modelId="\d+" srcId="1" destId="5" srcOrd="0"/)
    // An empty point is Office's place to type.
    expect(xml).toContain('phldr="1"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr')
  })

  it("draws Office layouts and colours it doesn't have as the nearest of its own", () => {
    expect(drawnLayout("urn:microsoft.com/office/officeart/2005/8/layout/process1").key).toBe("process")
    expect(drawnLayout("urn:x/unknown", "cycle").key).toBe("cycle")
    expect(drawnLayout("urn:x/unknown").key).toBe("list")
    expect(smartArtColors("urn:microsoft.com/office/officeart/2005/8/colors/accent3_2").accents).toEqual(["accent3"])
    expect(smartArtColors("urn:microsoft.com/office/officeart/2005/8/colors/colorful1").accents).toHaveLength(5)
  })
})
