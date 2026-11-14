import { describe, expect, it } from "vitest"
import { modernHashMatches, passwordMatches, sha512 } from "../src"

const hex = (b: Uint8Array) => [...b].map((x) => x.toString(16).padStart(2, "0")).join("")
const utf8 = (s: string) => new TextEncoder().encode(s)

describe("sha512", () => {
  it("matches the FIPS 180-4 test vectors", () => {
    expect(hex(sha512(utf8("abc")))).toBe(
      "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f",
    )
    expect(hex(sha512(utf8("")))).toBe(
      "cf83e1357eefb8bdf1542850d66d8007d620e4050b5715dc83f4a921d36ce9ce47d0d13c5d85f2b0ff8318d2877eec2f63b931bd47417a81a538327af927da3e",
    )
    expect(
      hex(
        sha512(
          utf8(
            "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu",
          ),
        ),
      ),
    ).toBe(
      "8e959b75dae313da8cf4f72814fc143f8f7779c6eb9f7fa17299aeadb6889018501d289e4900f7e4331b99dec4b5433ac7d329eeb6dd26545e96e55b874be909",
    )
  })
})

describe("Excel 365 password hashes", () => {
  it("checks a SHA-512 hash with salt and spin count", async () => {
    // Made with the same algorithm through Node's crypto, for the password "cells".
    const { createHash } = await import("node:crypto")
    const salt = Buffer.from("0123456789abcdef")
    let h = createHash("sha512")
      .update(Buffer.concat([salt, Buffer.from("cells", "utf16le")]))
      .digest()
    for (let i = 0; i < 1000; i++) {
      const n = Buffer.alloc(4)
      n.writeUInt32LE(i)
      h = createHash("sha512")
        .update(Buffer.concat([h, n]))
        .digest()
    }
    const hash = { algorithm: "SHA-512", hash: h.toString("base64"), salt: salt.toString("base64"), spinCount: 1000 }
    expect(modernHashMatches(hash, "cells")).toBe(true)
    expect(modernHashMatches(hash, "Cells")).toBe(false)
    expect(passwordMatches({ hash }, "cells")).toBe(true)
  })
})
