/** Zip archives written by hand, for what the reader refuses: entries as given, and a central directory pointing at them. */
import { concat } from '../../src/xlsx/zip/platform.js'

export interface RawEntry {
  name: string
  /** The bytes stored: deflated ones with `method` 8. */
  data: Uint8Array
  /** 0, stored (the default), or 8, deflated. */
  method?: number
  /** The size the entry says it inflates to: the data's own by default. */
  usize?: number
  flags?: number
}

const enc = new TextEncoder()

/** A local header (30 bytes and the name) or a record of the central directory (46 and the name, its fields from the flags on 2 bytes further). */
function record(sig: number, fixed: number, e: RawEntry, name: string): [Uint8Array, DataView] {
  const n = enc.encode(name)
  const out = new Uint8Array(fixed + n.length)
  const v = new DataView(out.buffer)
  const at = fixed === 46 ? 2 : 0
  v.setUint32(0, sig, true)
  v.setUint16(6 + at, e.flags ?? 0, true)
  v.setUint16(8 + at, e.method ?? 0, true)
  v.setUint32(18 + at, e.data.length, true)
  v.setUint32(22 + at, e.usize ?? e.data.length, true)
  v.setUint16(26 + at, n.length, true)
  out.set(n, fixed)
  return [out, v]
}

/** The archive of `entries`; `dir` lists the central directory's records by name and the entry each points at (each entry once, in order, by default). */
export function rawZip(entries: RawEntry[], dir = entries.map((e, at) => ({ name: e.name, at }))): Uint8Array {
  const parts: Uint8Array[] = []
  const offsets: number[] = []
  let size = 0
  for (const e of entries) {
    offsets.push(size)
    const [local] = record(0x04034b50, 30, e, e.name)
    parts.push(local, e.data)
    size += local.length + e.data.length
  }
  let dirSize = 0
  for (const { name, at } of dir) {
    const [central, v] = record(0x02014b50, 46, entries[at], name)
    v.setUint32(42, offsets[at], true)
    parts.push(central)
    dirSize += central.length
  }
  const end = new Uint8Array(22)
  const ev = new DataView(end.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(8, dir.length, true)
  ev.setUint16(10, dir.length, true)
  ev.setUint32(12, dirSize, true)
  ev.setUint32(16, size, true)
  return concat([...parts, end])
}
