import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { readZip } from "@cells/ooxml"
import { packageObject, unpackObject, writeCompoundFile } from "../src/ole"
import { CompoundFile } from "../src/xls/cfb"

const bytes = (n: number, seed = 7) => Uint8Array.from({ length: n }, (_, i) => (i * seed + 3) & 0xff)

describe("embedded files (Insert › Object)", () => {
  it("writes compound files that read back, small streams in the mini stream and large ones in sectors", () => {
    const streams = [
      { name: "Small", data: bytes(100) },
      { name: "Empty", data: new Uint8Array() },
      { name: "Large", data: bytes(70_000, 13) },
      { name: "Edge", data: bytes(4095, 5) },
    ]
    const cf = new CompoundFile(writeCompoundFile(streams))
    for (const s of streams) expect(cf.stream(s.name)).toEqual(s.data)
  })

  it("wraps a file as Office's Package object and unwraps Excel's own", async () => {
    const data = new TextEncoder().encode("Shopping list\nApples\n")
    expect(unpackObject(packageObject("List.txt", data))).toEqual({ name: "List.txt", data })
    const big = bytes(200_000)
    expect(unpackObject(packageObject("Photo.bin", big))?.data).toEqual(big)

    const zip = await readZip(new Uint8Array(readFileSync(join(__dirname, "fixtures/excel/oleobject.xlsx"))))
    const bin = zip.find((e) => e.name === "xl/embeddings/oleObject1.bin")!.data
    const inner = unpackObject(bin)!
    expect(inner.name).toBe("note.txt")
    expect(new TextDecoder().decode(inner.data)).toBe("Shopping list\nApples\nBread\n")
  })
})
