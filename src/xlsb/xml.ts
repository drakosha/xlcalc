import type { XlsxLimits } from '../xlsx/types.js'
import { ENTRY_BYTES, TOTAL_BYTES } from '../xlsx/zip/read.js'

export interface Budget {
  spent: number
}

const nonAscii = /[^\x00-\x7f]/

function utf8Bytes(text: string): number {
  if (!nonAscii.test(text)) return text.length
  let bytes = 0
  for (const character of text) {
    const code = character.codePointAt(0)!
    bytes += code < 0x80 ? 1 : code < 0x800 ? 2 : code < 0x10000 ? 3 : 4
  }
  return bytes
}

/** Counts XML as it is emitted so a binary part cannot expand past the limits before it is joined. */
export class Xml {
  private readonly chunks: string[] = []
  private size = 0

  constructor(
    private readonly part: string,
    private readonly limits: XlsxLimits,
    private readonly budget: Budget,
  ) {}

  add(text: string): void {
    const bytes = utf8Bytes(text)
    const entryBytes = this.limits.entryBytes ?? ENTRY_BYTES
    if ((this.size += bytes) > entryBytes) throw new Error(`xlsb: ${this.part} produces more than ${entryBytes} bytes of XML (limits.entryBytes)`)
    const totalBytes = this.limits.totalBytes ?? TOTAL_BYTES
    if ((this.budget.spent += bytes) > totalBytes) throw new Error(`xlsb: ${this.part} produces XML past the ${totalBytes} byte total (limits.totalBytes)`)
    this.chunks.push(text)
  }

  text(): string {
    return this.chunks.join('')
  }
}
