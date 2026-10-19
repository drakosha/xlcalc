/**
 * Reading a zip archive (an OPC package): the central directory lists the entries, stored or deflated. Part
 * names compare case-insensitively, as OPC requires; ZIP64 and encrypted entries are refused, and so are entries whose
 * bytes overlap and parts past the bounds of `ReadXlsxOptions.limits` once inflated.
 */
import type { XlsxLimits } from '../types.js'
import { inflate, text } from './platform.js'

export interface Archive {
  /** The entries' names, in the archive's order. */
  readonly names: readonly string[]
  /** The bytes of an entry by name in any case; undefined when there is none. */
  bytes(name: string): Promise<Uint8Array | undefined>
  /** The text of an entry (UTF-8, line ends as XML reads them), undefined when there is none. */
  text(name: string): Promise<string | undefined>
}

interface Entry {
  name: string
  method: number
  csize: number
  usize: number
  offset: number
}

/** The signatures of a zip's records, which the writer writes: the end of the central directory, an entry of it, a local header. */
export const EOCD = 0x06054b50
export const CENTRAL = 0x02014b50
export const LOCAL = 0x04034b50

/** The bounds of a file once inflated by default: 256 MiB a part, 1 GiB the parts read together. */
export const ENTRY_BYTES = 2 ** 28
export const TOTAL_BYTES = 2 ** 30

function fail(why: string): never {
  throw new Error(`xlsx: not a readable zip archive: ${why}`)
}

/** The end of the central directory: searched backwards over a comment of up to 65,535 bytes. */
function endRecord(v: DataView): number {
  for (let i = v.byteLength - 22; i >= Math.max(0, v.byteLength - 22 - 0xffff); i--) if (v.getUint32(i, true) === EOCD) return i
  fail('no end of central directory')
}

function entries(bytes: Uint8Array): Entry[] {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const end = endRecord(v)
  const count = v.getUint16(end + 10, true)
  let at = v.getUint32(end + 16, true)
  if (count === 0xffff || at === 0xffffffff) fail('ZIP64 archives are not supported')
  const out: Entry[] = []
  for (let k = 0; k < count; k++) {
    if (at + 46 > v.byteLength || v.getUint32(at, true) !== CENTRAL) fail('broken central directory')
    const flags = v.getUint16(at + 8, true)
    const nameLen = v.getUint16(at + 28, true)
    const name = text(bytes.subarray(at + 46, at + 46 + nameLen))
    if (flags & 1) fail(`encrypted entry ${name}`)
    out.push({ name, method: v.getUint16(at + 10, true), csize: v.getUint32(at + 20, true), usize: v.getUint32(at + 24, true), offset: v.getUint32(at + 42, true) })
    at += 46 + nameLen + v.getUint16(at + 30, true) + v.getUint16(at + 32, true)
  }
  return out
}

/** The bytes of an entry in the file, its local header and data; without a local header, at least 30 bytes and its data. */
function span(v: DataView, e: Entry): number {
  const local = e.offset + 30 <= v.byteLength && v.getUint32(e.offset, true) === LOCAL
  return 30 + e.csize + (local ? v.getUint16(e.offset + 26, true) + v.getUint16(e.offset + 28, true) : 0)
}

export function openArchive(bytes: Uint8Array, limits: XlsxLimits = {}): Archive {
  const list = entries(bytes)
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  // records over the same bytes let a small file inflate many times over (a zip bomb's way)
  const byOffset = [...list].sort((a, b) => a.offset - b.offset)
  for (let i = 1; i < byOffset.length; i++) {
    const [a, b] = [byOffset[i - 1], byOffset[i]]
    if (a.offset + span(v, a) > b.offset) fail(`entries ${a.name} and ${b.name} overlap`)
  }
  const byName = new Map(list.map((e) => [e.name.toLowerCase(), e]))
  const { entryBytes = ENTRY_BYTES, totalBytes = TOTAL_BYTES } = limits
  let spent = 0
  const bytesOf = async (name: string): Promise<Uint8Array | undefined> => {
    const e = byName.get(name.toLowerCase())
    if (!e) return undefined
    if (v.getUint32(e.offset, true) !== LOCAL) fail(`no local header for ${e.name}`)
    const start = e.offset + 30 + v.getUint16(e.offset + 26, true) + v.getUint16(e.offset + 28, true)
    const raw = bytes.subarray(start, start + e.csize)
    if (e.method !== 0 && e.method !== 8) fail(`entry ${e.name} uses compression method ${e.method}`)
    // a stored part is its bytes in the file; a deflated one is refused as it inflates past the size it declares
    const size = e.method === 0 ? raw.length : e.usize
    if (size > entryBytes) fail(`entry ${e.name} is larger than ${entryBytes} bytes (limits.entryBytes)`)
    if ((spent += size) > totalBytes) fail(`the parts read are larger than ${totalBytes} bytes together (limits.totalBytes)`)
    if (e.method === 0) return raw
    const out = await inflate(raw, e.usize)
    if (!out) fail(`entry ${e.name} inflates to more than the ${e.usize} bytes it declares`)
    if (out.length !== e.usize) fail(`entry ${e.name} inflates to ${out.length} bytes, not ${e.usize}`)
    return out
  }
  return {
    names: list.map((e) => e.name),
    bytes: bytesOf,
    async text(name) {
      const b = await bytesOf(name)
      if (!b) return undefined
      // XML's end-of-line handling: CRLF and a lone CR are a line feed (Excel writes CRLF inside text)
      const s = text(b)
      return s.includes('\r') ? s.replace(/\r\n?/g, '\n') : s
    },
  }
}
