import { escapeAttr } from '../xlsx/xml/text.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'

const REGIONS = [
  'wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn',
  'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe',
  'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell',
  'firstSubtotalColumn', 'secondSubtotalColumn', 'thirdSubtotalColumn',
  'firstSubtotalRow', 'secondSubtotalRow', 'thirdSubtotalRow', 'blankRow',
  'firstColumnSubheading', 'secondColumnSubheading', 'thirdColumnSubheading',
  'firstRowSubheading', 'secondRowSubheading', 'thirdRowSubheading',
  'pageFieldLabels', 'pageFieldValues',
]

interface Region {
  kind: number
  size: number
  dxf: number
}

interface TableStyle {
  name: string
  flags: number
  count: number
  regions: Region[]
}

export class TableStyles {
  private table = ''
  private pivot = ''
  private readonly styles: TableStyle[] = []

  take(type: number, payload: Payload): boolean {
    if (type === 0x088e) {
      payload.skip(16)
      const tableCount = payload.u16()
      const pivotCount = payload.u16()
      this.table = payload.wide(tableCount)
      this.pivot = payload.wide(pivotCount)
    } else if (type === 0x088f) {
      payload.skip(12)
      const flags = payload.u16()
      const count = payload.u32()
      const name = payload.wide(payload.u16())
      this.styles.push({ name, flags, count, regions: [] })
    } else if (type === 0x0890) {
      payload.skip(12)
      const style = this.styles.at(-1)
      if (style && style.regions.length < style.count) {
        style.regions.push({ kind: payload.u32(), size: payload.u32(), dxf: payload.u32() })
      }
    } else return false
    return true
  }

  write(xml: Xml, dxfIndex: (id: number) => number | undefined): void {
    if (!this.styles.length && !this.table && !this.pivot) return
    const table = this.table ? ` defaultTableStyle="${escapeAttr(this.table)}"` : ''
    const pivot = this.pivot ? ` defaultPivotStyle="${escapeAttr(this.pivot)}"` : ''
    xml.add(`<tableStyles${table}${pivot}>`)
    for (const style of this.styles) {
      const flags = (style.flags & 4 ? '' : ' table="0"') + (style.flags & 2 ? '' : ' pivot="0"')
      xml.add(`<tableStyle name="${escapeAttr(style.name)}"${flags}>`)
      for (const { kind, size, dxf } of style.regions) {
        const type = REGIONS[kind]
        const index = dxfIndex(dxf)
        if (!type || index === undefined) continue
        const width = kind >= 5 && kind <= 8 && size > 1 ? ` size="${size}"` : ''
        xml.add(`<tableStyleElement type="${type}" dxfId="${index}"${width}/>`)
      }
      xml.add('</tableStyle>')
    }
    xml.add('</tableStyles>')
  }
}
