import type { Payload } from '../xlsb/records.js'
import type { LegacyEncoding } from './encoding.js'
import type { RecordInfo } from './records.js'

export interface PrintMarkup {
  properties: string
  settings: string
}

export type ReadPrintMarkup = (records: Iterable<RecordInfo>, encoding: LegacyEncoding | undefined, sheet: string) => PrintMarkup

/** The leading fields of BIFF's Setup record, also present in its shorter historical layouts. */
export function setupFields(payload: Payload) {
  const paper = payload.u16()
  const scale = payload.u16()
  const first = payload.i16()
  const width = payload.u16()
  const height = payload.u16()
  const flags = payload.u16()
  return { paper, scale, first, width, height, flags }
}

export class Printing {
  fit = false
  private header = false
  private rows = false
  private columns = false
  private setup?: { flags: number; scale: number; width: number; height: number }

  take(type: number, payload: Payload, encoding?: LegacyEncoding): void {
    if (type === 0x0014 || type === 0x0015) {
      if (payload.remaining) this.header ||= (encoding ? payload.u8() : payload.u16()) > 0
    } else if (type === 0x089c) {
      payload.skip(28)
      const flags = payload.u16()
      const lengths = [payload.u16(), payload.u16(), payload.u16(), payload.u16()]
      this.header ||= Boolean(flags & 3) || lengths.some(Boolean)
    } else if (type === 0x001b) this.rows ||= payload.u16() > 0
    else if (type === 0x001a) this.columns ||= payload.u16() > 0
    else if (type === 0x00a1 && payload.remaining >= 12) {
      this.setup = setupFields(payload)
    }
  }

  /** Supplies the print records that the base SpreadsheetML reader reports as outside the workbook model. */
  markers(): string {
    let xml = (this.header ? '<headerFooter/>' : '') + (this.rows ? '<rowBreaks/>' : '') + (this.columns ? '<colBreaks/>' : '')
    const setup = this.setup
    if (setup) {
      const configured = !(setup.flags & 4)
      let attrs = configured && !(setup.flags & (2 | 64)) ? ' orientation="landscape"' : ''
      if (this.fit) attrs += ` fitToWidth="${setup.width}" fitToHeight="${setup.height}"`
      else if (configured && setup.scale !== 100) attrs += ` scale="${setup.scale}"`
      if (attrs) xml += `<pageSetup${attrs}/>`
    }
    return xml
  }
}
