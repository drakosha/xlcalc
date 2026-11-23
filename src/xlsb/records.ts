/** BIFF12 record headers and bounded readers for record payloads ([MS-XLSB] 2.1.4). */

function invalid(part: string, offset: number, reason: string): never {
  throw new Error(`xlsb: ${part} at offset ${offset}: ${reason}`)
}

interface RecordInfo {
  type: number
  size: number
  payload: Payload
}

export function* records(bytes: Uint8Array, part: string): IterableIterator<RecordInfo> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let pos = 0
  const byte = () => {
    if (pos === bytes.length) invalid(part, pos, 'incomplete record header')
    return bytes[pos++]
  }
  while (pos < bytes.length) {
    const offset = pos
    const first = byte()
    let type = first & 0x7f
    if (first & 0x80) {
      type |= (byte() & 0x7f) << 7
      if (type < 128) invalid(part, offset, 'overlong record type')
    }
    let size = 0
    for (let shift = 0; shift < 28; shift += 7) {
      const digit = byte()
      size |= (digit & 0x7f) << shift
      // The spec ignores the high bit of the fourth size byte.
      if (!(digit & 0x80)) break
    }
    const end = pos + size
    if (end > bytes.length) invalid(part, offset, `record ${type} declares ${size} bytes beyond the end of the part`)
    yield { type, size, payload: new Payload(view, part, pos, end, type) }
    pos = end
  }
}

/** Offsets printed in errors are relative to the part, including for a sliced input buffer. */
export class Payload {
  constructor(
    private readonly view: DataView,
    private readonly part: string,
    private pos: number,
    private readonly end: number,
    private readonly type: number,
  ) {}

  get remaining(): number {
    return this.end - this.pos
  }

  fail(reason: string): never {
    return invalid(this.part, this.pos, `record ${this.type}: ${reason}`)
  }

  private next(size: number): number {
    if (size < 0 || size > this.remaining) this.fail(`expected ${size} bytes, ${this.remaining} left`)
    const start = this.pos
    this.pos += size
    return start
  }

  skip(size: number): void {
    this.next(size)
  }

  section(size: number): Payload {
    const start = this.next(size)
    return new Payload(this.view, this.part, start, start + size, this.type)
  }

  u8(): number {
    return this.view.getUint8(this.next(1))
  }

  u16(): number {
    return this.view.getUint16(this.next(2), true)
  }

  i16(): number {
    return this.view.getInt16(this.next(2), true)
  }

  u32(): number {
    return this.view.getUint32(this.next(4), true)
  }

  i32(): number {
    return this.view.getInt32(this.next(4), true)
  }

  f64(): number {
    return this.view.getFloat64(this.next(8), true)
  }

  /** UTF-16LE code units. TextDecoder would replace unpaired surrogates before the XML writer could escape them. */
  wide(length = this.u32()): string {
    const start = this.next(length * 2)
    const chunks: string[] = []
    for (let i = 0; i < length; i += 4096) {
      const units: number[] = []
      for (let j = i; j < Math.min(i + 4096, length); j++) units.push(this.view.getUint16(start + j * 2, true))
      chunks.push(String.fromCharCode(...units))
    }
    return chunks.join('')
  }

  nullableWide(): string | undefined {
    const length = this.u32()
    return length === 0xffffffff ? undefined : this.wide(length)
  }

  done(): void {
    if (this.remaining) this.fail(`${this.remaining} unexpected trailing bytes`)
  }
}
