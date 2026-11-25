/** BIFF records have a fixed four-byte header. Offsets are relative to the Workbook stream. */
import { formatAddress } from '../model.js'
import { concat } from '../xlsx/zip/platform.js'
import { Payload } from '../xlsb/records.js'
import type { LegacyVersion } from './encoding.js'

export interface RecordInfo {
  type: number
  offset: number
  payload: Payload
}

export const BOF = 0x0809
export const EOF = 0x000a
export const CONTINUE = 0x003c

export const BUILTIN_NAMES: readonly string[] = [
  'Consolidate_Area', 'Auto_Open', 'Auto_Close', 'Extract', 'Database', 'Criteria', 'Print_Area',
  'Print_Titles', 'Recorder', 'Data_Form', 'Auto_Activate', 'Auto_Deactivate', 'Sheet_Title', '_FilterDatabase',
]

export function isBof(type: number): boolean {
  return type === BOF || type === 0x0009 || type === 0x0209 || type === 0x0409
}

export function biffVersion(type: number, declared: number): LegacyVersion | 8 | undefined {
  if (type === 0x0009) return 2
  if (type === 0x0209) return 3
  if (type === 0x0409) return 4
  if (type !== BOF) return undefined
  const version = declared >>> 8
  switch (version) {
    case 2: case 3: case 4: case 5: return version
    case 0: return 5
    case 6: return 8
    default: return undefined
  }
}

export function substreamType(declaredVersion: number, type: number): number {
  // Some unwrapped BIFF5 worksheet streams write the worksheet type with its bytes reversed.
  return declaredVersion === 0 && type === 0x1000 ? 0x0010 : type
}

export function* records(bytes: Uint8Array, start: number, end: number): Generator<RecordInfo> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let at = start
  while (at < end) {
    if (at + 4 > end) throw new Error(`xls: Workbook at offset ${at}: incomplete record header`)
    const type = view.getUint16(at, true)
    const size = view.getUint16(at + 2, true)
    if (at + 4 + size > end) throw new Error(`xls: Workbook at offset ${at}: record ${type} extends beyond its substream`)
    yield { type, offset: at, payload: new Payload(view, 'Workbook', at + 4, at + 4 + size, type, 'xls') }
    at += 4 + size
  }
}

export function combinedPayload(parts: Uint8Array[], part: string, type: number): Payload {
  const bytes = parts.length === 1 ? parts[0] : concat(parts)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return new Payload(view, part, 0, bytes.byteLength, type, 'xls')
}

export function unicode(payload: Payload, length = payload.u16()): string {
  return payload.u8() & 1 ? payload.wide(length) : payload.narrow(length)
}

/** Only the continuation of character data has an encoding byte. Formatting runs and phonetic data are raw. */
export class Continued {
  constructor(private part: Payload, private readonly next: () => Payload | undefined) {}

  private advance(): void {
    this.part = this.next() ?? this.part.fail('string needs another Continue record')
  }

  hasData(): boolean {
    while (!this.part.remaining) {
      const next = this.next()
      if (!next) return false
      this.part = next
    }
    return true
  }

  private current(): Payload {
    if (!this.hasData()) this.part.fail('string needs another Continue record')
    return this.part
  }

  u8(): number { return this.current().u8() }

  u16(): number {
    if (this.current().remaining >= 2) return this.part.u16()
    return this.u8() | (this.u8() << 8)
  }

  i16(): number { return (this.u16() << 16) >> 16 }

  u32(): number {
    if (this.current().remaining >= 4) return this.part.u32()
    return (this.u16() | (this.u16() << 16)) >>> 0
  }

  skip(size: number): void {
    while (size) {
      const part = this.current()
      const available = Math.min(part.remaining, size)
      part.skip(available)
      size -= available
    }
  }

  text(length: number, wide: boolean): string {
    const chunks: string[] = []
    while (length) {
      const units = Math.min(length, Math.floor(this.part.remaining / (wide ? 2 : 1)))
      chunks.push(wide ? this.part.wide(units) : this.part.narrow(units))
      length -= units
      if (!length) break
      this.part.done()
      this.advance()
      wide = Boolean(this.part.u8() & 1)
    }
    return chunks.join('')
  }

  fail(reason: string): never { return this.part.fail(reason) }

}

export function substream(
  bytes: Uint8Array, start: number, end: number, begin: (payload: Payload) => void,
): Generator<RecordInfo> {
  const source = records(bytes, start, end)
  const first = source.next().value
  if (!first || !isBof(first.type)) throw new Error(`xls: expected BOF at offset ${start}`)
  begin(first.payload)
  return source
}

export function address(row: number, col: number, payload: Payload): string {
  if (col > 255) payload.fail(`column ${col} is outside a binary worksheet`)
  return formatAddress(row, col)
}

export function reference(payload: Payload, short = false): string {
  const top = payload.u16()
  const bottom = payload.u16()
  const first = short ? payload.u8() : payload.u16()
  const last = short ? payload.u8() : payload.u16()
  if (top > bottom || first > last) payload.fail('range ends before it starts')
  const start = address(top, first, payload)
  const end = address(bottom, last, payload)
  return start === end ? start : start + ':' + end
}
