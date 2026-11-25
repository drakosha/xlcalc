import { LOCALE_PROPERTY } from '../xlsx/property.js'
import { Payload } from '../xlsb/records.js'

const USER_DEFINED = [0x05, 0xd5, 0xcd, 0xd5, 0x9c, 0x2e, 0x1b, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2b, 0x2c, 0xf9, 0xae]

/** [MS-OLEPS] PropertySetStream, Dictionary and TypedPropertyValue. */
export function localeProperty(bytes: Uint8Array | undefined): string | undefined {
  if (!bytes || bytes.length < 48) return undefined
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const word = (at: number) => view.getUint16(at, true)
  const long = (at: number) => view.getUint32(at, true)
  if (word(0) !== 0xfffe) return undefined
  const sets = long(24)
  if (sets < 1 || sets > 2 || 28 + sets * 20 > bytes.length) return undefined

  let start: number | undefined
  for (let i = 0; i < sets; i++) {
    const at = 28 + i * 20
    if (USER_DEFINED.every((value, position) => bytes[at + position] === value)) {
      start = long(at + 16)
      break
    }
  }
  if (start === undefined || start + 8 > bytes.length) return undefined
  const size = long(start)
  const count = long(start + 4)
  const header = 8 + 8 * count
  if (size < header || start + size > bytes.length) return undefined
  const first = start
  const end = start + size
  const offsetOf = (id: number): number | undefined => {
    for (let at = first + 8; at < first + header; at += 8) {
      if (long(at) !== id) continue
      const offset = long(at + 4)
      return offset >= header && offset < size ? first + offset : undefined
    }
    return undefined
  }

  const codePageAt = offsetOf(1)
  const wide = codePageAt !== undefined && codePageAt + 6 <= end && word(codePageAt) === 2 && word(codePageAt + 4) === 1200
  const dictionary = offsetOf(0)
  if (dictionary === undefined || dictionary + 4 > end) return undefined
  const entries = long(dictionary)
  for (let i = 0, at = dictionary + 4; i < entries && at + 8 <= end; i++) {
    const id = long(at)
    const length = long(at + 4)
    const string = at + 8
    const width = wide ? 2 : 1
    const lengthBytes = length * width
    if (string + lengthBytes > end) return undefined
    // Both the property name and valid BCP 47 tags are ASCII in non-Unicode code pages.
    const name = length === LOCALE_PROPERTY.length + 1
      ? new Payload(view, 'Workbook properties', string, string + lengthBytes, 0, 'xls')
      : undefined
    const text = name && (wide ? name.wide(length) : name.narrow(length))
    if (text === LOCALE_PROPERTY + '\0') {
      const value = offsetOf(id)
      if (value === undefined || value + 8 > end) return undefined
      const type = word(value)
      if (type !== 0x1e && type !== 0x1f) return undefined
      const length = long(value + 4)
      const unicode = type === 0x1f || wide
      const count = type === 0x1f ? length * 2 : length
      if (value + 8 + count > end || (unicode && count % 2)) return undefined
      const payload = new Payload(view, 'Workbook properties', value + 8, value + 8 + count, type, 'xls')
      const tag = unicode ? payload.wide(count / 2) : payload.narrow(count)
      return tag.replace(/\0+$/, '')
    }
    const padded = wide ? Math.ceil(lengthBytes / 4) * 4 : lengthBytes
    at = string + padded
  }
  return undefined
}
