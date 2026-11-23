import { escapeAttr } from '../xlsx/xml/text.js'
import { records } from './records.js'
import type { Xml } from './xml.js'

export const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

export interface Part {
  bytes: Uint8Array
  path: string
  xml: Xml
  pending(): void
}

export function workbook({ bytes, path, xml, pending }: Part): void {
  xml.add(`<workbook ${NS}><sheets>`)
  for (const { type, payload } of records(bytes, path)) {
    if (type === 156) {
      const state = payload.u32()
      const id = payload.u32()
      const rel = payload.nullableWide()
      const name = payload.wide()
      payload.done()
      if (state > 2) payload.fail(`invalid sheet state ${state}`)
      const hidden = state ? ` state="${state === 1 ? 'hidden' : 'veryHidden'}"` : ''
      xml.add(`<sheet name="${escapeAttr(name)}" sheetId="${id}"${rel === undefined ? '' : ` r:id="${escapeAttr(rel)}"`}${hidden}/>`)
    }
  }
  xml.add('</sheets></workbook>')
  pending()
}

export function worksheet({ bytes, path, xml, pending }: Part): void {
  for (const _record of records(bytes, path)) {
    // At this stage the framing is checked and the worksheet data is reported as pending.
  }
  xml.add(`<worksheet ${NS}><sheetData/></worksheet>`)
  pending()
}
