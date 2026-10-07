/**
 * Writing a zip archive: each entry deflated, a fixed timestamp (1980-01-01) so the same parts give the same
 * bytes, ASCII names. No ZIP64: a package of 4 GB or more is refused.
 */
import { concat, deflate } from './platform.js'

let table: Uint32Array | undefined

function crc32(data: Uint8Array): number {
  if (!table) {
    table = new Uint32Array(256)
    for (let n = 0; n < 256; n++) {
      let c = n
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      table[n] = c >>> 0
    }
  }
  let c = 0xffffffff
  for (let i = 0; i < data.length; i++) c = table[(c ^ data[i]) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

/** DOS date of 1980-01-01, the earliest a zip can hold. */
const DATE = 0x21

export async function zip(files: { name: string; data: Uint8Array }[]): Promise<Uint8Array> {
  const parts: Uint8Array[] = []
  const central: Uint8Array[] = []
  let offset = 0
  const packs = await Promise.all(files.map((f) => deflate(f.data)))
  for (const [k, f] of files.entries()) {
    if (!/^[\x20-\x7e]+$/.test(f.name)) throw new Error(`xlsx: part name ${f.name} is not ASCII`)
    const name = Uint8Array.from(f.name, (c) => c.charCodeAt(0))
    const packed = packs[k]
    const crc = crc32(f.data)
    // the fields a local header and a central directory entry share, from `at` on
    const common = (v: DataView, at: number) => {
      v.setUint16(at, 20, true)
      v.setUint16(at + 4, 8, true)
      v.setUint16(at + 8, DATE, true)
      v.setUint32(at + 10, crc, true)
      v.setUint32(at + 14, packed.length, true)
      v.setUint32(at + 18, f.data.length, true)
      v.setUint16(at + 22, name.length, true)
    }
    const local = new Uint8Array(30 + name.length)
    const lv = new DataView(local.buffer)
    lv.setUint32(0, 0x04034b50, true)
    common(lv, 4)
    local.set(name, 30)
    const dir = new Uint8Array(46 + name.length)
    const dv = new DataView(dir.buffer)
    dv.setUint32(0, 0x02014b50, true)
    dv.setUint16(4, 20, true)
    common(dv, 6)
    dv.setUint32(42, offset, true)
    dir.set(name, 46)
    parts.push(local, packed)
    central.push(dir)
    offset += local.length + packed.length
  }
  const dirSize = central.reduce((n, d) => n + d.length, 0)
  if (offset + dirSize > 0xffffffff || files.length > 0xffff) throw new Error('xlsx: the package needs ZIP64, which is not supported')
  const end = new Uint8Array(22)
  const ev = new DataView(end.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(8, files.length, true)
  ev.setUint16(10, files.length, true)
  ev.setUint32(12, dirSize, true)
  ev.setUint32(16, offset, true)
  return concat([...parts, ...central, end])
}
