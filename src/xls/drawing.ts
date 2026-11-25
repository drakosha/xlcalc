import type { Payload } from '../xlsb/records.js'
import { concat } from '../xlsx/zip/platform.js'

function isAutoFilterDropdown(object: Payload): boolean {
  const data = object.copy()
  const common = data.u16()
  const size = data.u16()
  if (common !== 0x15 || size < 2 || data.section(size).u16() !== 0x14) return false
  while (data.remaining >= 4) {
    const type = data.u16()
    const length = data.u16()
    if (type === 0x13) {
      // FtLbsData uses the length field for continuation metadata. Its own formula has a byte count.
      data.skip(data.u16() + 4)
      const flags = data.u16()
      return (flags & 1) !== 0 && flags >>> 8 === 3
    }
    if (!type) break
    data.skip(length)
  }
  return false
}

/** A hidden AutoFilter arrow is a hidden OfficeArt shape anchored to its header cell. */
export function hiddenDropdown(parts: Uint8Array[], object: Payload): { row: number; col: number } | undefined {
  if (object.remaining < 6 || !isAutoFilterDropdown(object)) return undefined
  const bytes = parts.length === 1 ? parts[0] : concat(parts)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let hidden = false
  let position: { row: number; col: number } | undefined
  for (let at = 0; at + 8 <= view.byteLength; ) {
    const flags = view.getUint16(at, true)
    const type = view.getUint16(at + 2, true)
    const size = view.getUint32(at + 4, true)
    at += 8
    if ((flags & 15) === 15) {
      // The drawing's parent containers can include several BIFF records.
      if (type === 0xf004) {
        hidden = false
        position = undefined
      }
      continue
    }
    if (size > view.byteLength - at) break
    if (type === 0xf00b || type === 0xf122) {
      const count = Math.min(flags >>> 4, Math.floor(size / 6))
      for (let i = 0; i < count; i++) {
        const offset = at + i * 6
        if (view.getUint16(offset, true) !== 0x03bf) continue
        const properties = view.getUint32(offset + 2, true)
        if (properties & 0x20000) hidden = (properties & 2) !== 0
      }
    } else if (type === 0xf010 && size >= 18) {
      position = { col: view.getUint16(at + 2, true), row: view.getUint16(at + 6, true) }
    }
    at += size
  }
  return hidden ? position : undefined
}
