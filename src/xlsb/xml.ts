import type { XlsxLimits } from '../xlsx/types.js'
import { text, utf8Into } from '../xlsx/zip/platform.js'
import { ENTRY_BYTES, TOTAL_BYTES } from '../xlsx/zip/read.js'

export interface Budget {
  spent: number
}

const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g
const nonAscii = /[^\x00-\x7f]/

/**
 * UTF-8 cannot hold a lone surrogate, and the encoder would put U+FFFD in its place. Element text reaches here escaped
 * already (`escapeText`); anything else gets the same `_xHHHH_`, which the reader decodes back to the surrogate.
 */
const encodable = (s: string): string =>
  nonAscii.test(s) ? s.replace(LONE_SURROGATE, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase()}_`) : s

/**
 * XML of a binary part, written as UTF-8 bytes and decoded once at the end. A sheet's XML comes in hundreds of thousands
 * of short pieces; kept as strings until a join, each piece's own header doubled the text's weight on the heap. The
 * bytes are counted as they are written, so a binary part cannot expand past the limits before it is decoded.
 */
export class Xml {
  private bytes = new Uint8Array(4096)
  private size = 0

  constructor(
    private readonly part: string,
    private readonly limits: XlsxLimits,
    private readonly budget: Budget,
  ) {}

  add(piece: string): void {
    const s = encodable(piece)
    // a UTF-16 code unit takes at most 3 bytes of UTF-8 (a surrogate pair, two units, takes 4)
    this.reserve(s.length * 3)
    const { written } = utf8Into(s, this.bytes.subarray(this.size))
    this.grow(written)
    const totalBytes = this.limits.totalBytes ?? TOTAL_BYTES
    if ((this.budget.spent += written) > totalBytes) throw new Error(`xlsb: ${this.part} produces XML past the ${totalBytes} byte total (limits.totalBytes)`)
  }

  fragment(): Xml {
    return new Xml(this.part, this.limits, this.budget)
  }

  /** Another piece's XML, already counted against the total when it was added there. */
  append(other: Xml): void {
    this.reserve(other.size)
    this.bytes.set(other.bytes.subarray(0, other.size), this.size)
    this.grow(other.size)
  }

  get empty(): boolean {
    return this.size === 0
  }

  private reserve(bytes: number): void {
    const need = this.size + bytes
    if (need <= this.bytes.length) return
    let capacity = this.bytes.length * 2
    while (capacity < need) capacity *= 2
    const bigger = new Uint8Array(capacity)
    bigger.set(this.bytes.subarray(0, this.size))
    this.bytes = bigger
  }

  private grow(bytes: number): void {
    const entryBytes = this.limits.entryBytes ?? ENTRY_BYTES
    if ((this.size += bytes) > entryBytes) throw new Error(`xlsb: ${this.part} produces more than ${entryBytes} bytes of XML (limits.entryBytes)`)
  }

  text(): string {
    return text(this.bytes.subarray(0, this.size))
  }
}
