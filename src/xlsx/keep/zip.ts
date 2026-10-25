/** The entries of the original file as it holds them, for the written file to take over those that did not change. */
import { entries, LOCAL } from '../zip/read.js'
import type { EntryRecord } from '../zip/write.js'

/** The signature a data descriptor may start with. */
const DESCRIPTOR = 0x08074b50

/** Each entry's local header, data and data descriptor, and its central directory record, by name in lower case. */
export function records(bytes: Uint8Array): Map<string, EntryRecord> {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const out = new Map<string, EntryRecord>()
  for (const e of entries(bytes)) {
    if (e.offset + 30 > v.byteLength || v.getUint32(e.offset, true) !== LOCAL) throw new Error(`xlsx: no local header for ${e.name}`)
    let end = e.offset + 30 + v.getUint16(e.offset + 26, true) + v.getUint16(e.offset + 28, true) + e.csize
    // the sizes and CRC-32 after the data (flag 3), with or without their signature
    if (e.flags & 8) end += end + 4 <= v.byteLength && v.getUint32(end, true) === DESCRIPTOR ? 16 : 12
    if (end > v.byteLength) throw new Error(`xlsx: entry ${e.name} is cut short`)
    out.set(e.name.toLowerCase(), { local: bytes.subarray(e.offset, end), central: bytes.subarray(e.central[0], e.central[1]) })
  }
  return out
}
