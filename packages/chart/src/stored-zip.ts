// From Inkspire (src/document/docx/zip.ts): a stored zip written at once, for a chart's embedded workbook.
const ZIP_FIXED_DOS_STAMP = 0x00210000
const CRC_TABLE = new Uint32Array(256)
for (let index = 0; index < 256; index += 1) {
  let value = index
  for (let bit = 0; bit < 8; bit += 1) {
    value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
  }
  CRC_TABLE[index] = value
}

/** Three more tables, so four bytes are folded in per step ("slicing by 4"): about three times faster on large parts. */
const CRC_SLICES = new Uint32Array(256 * 3)
for (let index = 0; index < 256; index += 1) {
  let value = CRC_TABLE[index]!
  for (let slice = 0; slice < 3; slice += 1) {
    value = CRC_TABLE[value & 0xff]! ^ (value >>> 8)
    CRC_SLICES[slice * 256 + index] = value
  }
}

function crc32(data: Uint8Array) {
  let crc = 0xffffffff
  let index = 0
  const stop = data.length - 3
  for (; index < stop; index += 4) {
    crc ^= data[index]! | (data[index + 1]! << 8) | (data[index + 2]! << 16) | (data[index + 3]! << 24)
    crc =
      CRC_SLICES[512 + (crc & 0xff)]! ^
      CRC_SLICES[256 + ((crc >>> 8) & 0xff)]! ^
      CRC_SLICES[(crc >>> 16) & 0xff]! ^
      CRC_TABLE[crc >>> 24]!
  }
  for (; index < data.length; index += 1) {
    crc = CRC_TABLE[(crc ^ data[index]!) & 0xff]! ^ (crc >>> 8)
  }
  return (crc ^ 0xffffffff) >>> 0
}

/**
 * A small zip with every entry stored, written at once: a chart's embedded workbook, built while a document is
 * turned into WordprocessingML, which has no await.
 */
export function storedZip(entries: { name: string; data: Uint8Array }[]): Uint8Array {
  const encoder = new TextEncoder()
  const chunks: Uint8Array[] = []
  const central: Uint8Array[] = []
  let offset = 0
  const time = ZIP_FIXED_DOS_STAMP & 0xffff
  const date = ZIP_FIXED_DOS_STAMP >>> 16
  for (const entry of entries) {
    const name = encoder.encode(entry.name)
    const crc = crc32(entry.data)
    const local = new DataView(new ArrayBuffer(30))
    local.setUint32(0, 0x04034b50, true)
    local.setUint16(4, 20, true)
    local.setUint16(10, time, true)
    local.setUint16(12, date, true)
    local.setUint32(14, crc, true)
    local.setUint32(18, entry.data.length, true)
    local.setUint32(22, entry.data.length, true)
    local.setUint16(26, name.length, true)
    const record = new DataView(new ArrayBuffer(46))
    record.setUint32(0, 0x02014b50, true)
    record.setUint16(4, 20, true)
    record.setUint16(6, 20, true)
    record.setUint16(12, time, true)
    record.setUint16(14, date, true)
    record.setUint32(16, crc, true)
    record.setUint32(20, entry.data.length, true)
    record.setUint32(24, entry.data.length, true)
    record.setUint16(28, name.length, true)
    record.setUint32(42, offset, true)
    chunks.push(new Uint8Array(local.buffer), name, entry.data)
    central.push(new Uint8Array(record.buffer), name)
    offset += 30 + name.length + entry.data.length
  }
  const size = central.reduce((sum, chunk) => sum + chunk.length, 0)
  const end = new DataView(new ArrayBuffer(22))
  end.setUint32(0, 0x06054b50, true)
  end.setUint16(8, entries.length, true)
  end.setUint16(10, entries.length, true)
  end.setUint32(12, size, true)
  end.setUint32(16, offset, true)
  const all = [...chunks, ...central, new Uint8Array(end.buffer)]
  const bytes = new Uint8Array(all.reduce((sum, chunk) => sum + chunk.length, 0))
  let at = 0
  for (const chunk of all) {
    bytes.set(chunk, at)
    at += chunk.length
  }
  return bytes
}
