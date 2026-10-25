/**
 * Reading a zip archive (an OPC package): the central directory lists the entries, stored or deflated. Part
 * names compare case-insensitively, as OPC requires; ZIP64 and encrypted entries are refused, and so are two entries of
 * one name, entries whose bytes overlap, are cut short or do not match their CRC-32, and parts past the bounds of
 * `ReadXlsxOptions.limits` once inflated.
 */
import type { XlsxLimits } from '../types.js'
import { xmlText } from '../xml/text.js'
import { inflate, text } from './platform.js'

export interface Archive {
  /** The entries' names, in the archive's order. */
  readonly names: readonly string[]
  /** The bytes of an entry by name in any case; undefined when there is none. */
  bytes(name: string): Promise<Uint8Array | undefined>
  /** The text of an entry as XML reads it (`xmlText`), undefined when there is none. */
  text(name: string): Promise<string | undefined>
}

/** An entry of the central directory. */
export interface Entry {
  name: string
  flags: number
  method: number
  crc: number
  csize: number
  usize: number
  offset: number
  /** Where its central directory record lies. */
  central: [number, number]
}

/** The signatures of a zip's records, which the writer writes: the end of the central directory, an entry of it, a local header. */
export const EOCD = 0x06054b50
export const CENTRAL = 0x02014b50
export const LOCAL = 0x04034b50

/**
 * The bounds of a file once inflated by default: 512 MiB a part (about what a JavaScript string holds), 1 GiB the parts
 * read together, and a part over 16 MiB no more than 200 times its size in the file (xlsx parts inflate 5 to 40 times).
 */
export const ENTRY_BYTES = 2 ** 29
export const TOTAL_BYTES = 2 ** 30
export const RATIO = 200
const RATIO_FROM = 2 ** 24

let tables: Uint32Array | undefined

/** The CRC-32 of bytes, as a zip stores it for each entry, continued from `crc` of the bytes before them: eight bytes a step by eight tables (slicing-by-8). */
export function crc32(data: Uint8Array, crc = 0): number {
  const t = (tables ??= crcTables())
  let c = ~crc
  let i = 0
  for (const end = data.length - 7; i < end; i += 8) {
    const a = (data[i] | (data[i + 1] << 8) | (data[i + 2] << 16) | (data[i + 3] << 24)) ^ c
    const b = data[i + 4] | (data[i + 5] << 8) | (data[i + 6] << 16) | (data[i + 7] << 24)
    c =
      t[1792 + (a & 0xff)] ^ t[1536 + ((a >>> 8) & 0xff)] ^ t[1280 + ((a >>> 16) & 0xff)] ^ t[1024 + (a >>> 24)] ^
      t[768 + (b & 0xff)] ^ t[512 + ((b >>> 8) & 0xff)] ^ t[256 + ((b >>> 16) & 0xff)] ^ t[b >>> 24]
  }
  for (; i < data.length; i++) c = t[(c ^ data[i]) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

/** The tables of the CRC-32 polynomial: the first a byte's, each next one byte further on. */
function crcTables(): Uint32Array {
  const t = new Uint32Array(8 * 256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[n] = c
  }
  for (let n = 0; n < 256; n++) for (let k = 1; k < 8; k++) t[k * 256 + n] = (t[(k - 1) * 256 + n] >>> 8) ^ t[t[(k - 1) * 256 + n] & 0xff]
  return t
}

function fail(why: string, cause?: unknown): never {
  throw new Error(`xlsx: not a readable zip archive: ${why}`, cause === undefined ? undefined : { cause })
}

/** The end of the central directory: searched backwards over a comment of up to 65,535 bytes. */
function endRecord(v: DataView): number {
  for (let i = v.byteLength - 22; i >= Math.max(0, v.byteLength - 22 - 0xffff); i--) if (v.getUint32(i, true) === EOCD) return i
  fail('no end of central directory record')
}

/** The entries the central directory lists, in its order. */
export function entries(bytes: Uint8Array): Entry[] {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const end = endRecord(v)
  const count = v.getUint16(end + 10, true)
  let at = v.getUint32(end + 16, true)
  if (count === 0xffff || at === 0xffffffff) fail('ZIP64 archives are not supported')
  const out: Entry[] = []
  for (let k = 0; k < count; k++) {
    if (at + 46 > v.byteLength || v.getUint32(at, true) !== CENTRAL) fail('corrupt central directory')
    const flags = v.getUint16(at + 8, true)
    const nameLen = v.getUint16(at + 28, true)
    const name = text(bytes.subarray(at + 46, at + 46 + nameLen))
    if (flags & 1) fail(`encrypted entry ${name}`)
    const next = at + 46 + nameLen + v.getUint16(at + 30, true) + v.getUint16(at + 32, true)
    out.push({ name, flags, method: v.getUint16(at + 10, true), crc: v.getUint32(at + 16, true), csize: v.getUint32(at + 20, true), usize: v.getUint32(at + 24, true), offset: v.getUint32(at + 42, true), central: [at, next] })
    at = next
  }
  return out
}

/** The bytes of an entry in the file, its local header and data; without a local header, at least 30 bytes and its data. */
function span(v: DataView, e: Entry): number {
  const local = e.offset + 30 <= v.byteLength && v.getUint32(e.offset, true) === LOCAL
  return 30 + e.csize + (local ? v.getUint16(e.offset + 26, true) + v.getUint16(e.offset + 28, true) : 0)
}

/** An entry's bytes, once their CRC-32 is the one the archive records: damaged bytes that still inflate are refused. */
function checked(e: Entry, data: Uint8Array, crc = crc32(data)): Uint8Array {
  if (crc !== e.crc) fail(`entry ${e.name} is damaged (its CRC-32 does not match)`)
  return data
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
  // OPC part names are unique in any case: a second entry of one would be a part another reader may take instead
  const byName = new Map<string, Entry>()
  for (const e of list) {
    const key = e.name.toLowerCase()
    if (byName.has(key)) fail(`entries ${byName.get(key)!.name} and ${e.name} have the same name`)
    byName.set(key, e)
  }
  const { entryBytes = ENTRY_BYTES, totalBytes = TOTAL_BYTES, ratio = RATIO } = limits
  let spent = 0
  const bytesOf = async (name: string): Promise<Uint8Array | undefined> => {
    const e = byName.get(name.toLowerCase())
    if (!e) return undefined
    if (e.offset + 30 > v.byteLength || v.getUint32(e.offset, true) !== LOCAL) fail(`no local header for ${e.name}`)
    const start = e.offset + 30 + v.getUint16(e.offset + 26, true) + v.getUint16(e.offset + 28, true)
    if (start + e.csize > v.byteLength) fail(`entry ${e.name} is cut short`)
    const raw = bytes.subarray(start, start + e.csize)
    if (e.method !== 0 && e.method !== 8) fail(`entry ${e.name} uses the unsupported compression method ${e.method}`)
    if (e.method === 0 && e.usize !== e.csize) fail(`entry ${e.name} is stored in ${e.csize} bytes but says it holds ${e.usize}`)
    // a stored part is its bytes in the file; a deflated one is refused as it inflates past the size it declares
    const size = e.method === 0 ? raw.length : e.usize
    if (size > entryBytes) fail(`entry ${e.name} is larger than ${entryBytes} bytes (limits.entryBytes)`)
    if ((spent += size) > totalBytes) fail(`the parts read add up to more than ${totalBytes} bytes (limits.totalBytes)`)
    if (e.method === 0) return checked(e, raw)
    if (e.usize > RATIO_FROM && e.usize > ratio * e.csize) fail(`entry ${e.name} inflates to ${e.usize} bytes from ${e.csize}, more than ${ratio} times its size (limits.ratio)`)
    // a platform without the codec throws as it is; damaged data fails the stream
    // the CRC-32 of each chunk as it comes, while the next one inflates
    let crc = 0
    const inflating = inflate(raw, e.usize, (chunk) => (crc = crc32(chunk, crc)))
    let out: Uint8Array | undefined
    try {
      out = await inflating
    } catch (x) {
      fail(`entry ${e.name} is not valid deflate data`, x)
    }
    if (!out) fail(`entry ${e.name} inflates to more than the ${e.usize} bytes it declares`)
    if (out.length !== e.usize) fail(`entry ${e.name} inflates to ${out.length} bytes, not ${e.usize}`)
    return checked(e, out, crc)
  }
  return {
    names: list.map((e) => e.name),
    bytes: bytesOf,
    async text(name) {
      const b = await bytesOf(name)
      if (!b) return undefined
      return xmlText(text(b), name)
    },
  }
}
