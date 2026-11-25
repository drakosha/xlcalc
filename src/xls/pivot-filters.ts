import { escapeAttr } from '../xlsx/xml/text.js'
import { FILTER_TYPES } from '../xlsb/pivot-table.js'
import type { Payload } from '../xlsb/records.js'
import { unicode } from './records.js'

const STRINGS = new Map([
  [0x2f, 'name'],
  [0x39, 'description'],
  [0x3a, 'stringValue1'],
  [0x3b, 'stringValue2'],
])

interface Filter {
  field?: number
  type?: string
  measure?: number
  criteria?: string
  strings: Map<string, { length: number; text: string }>
}

function criteria(payload: Payload, kind: string | undefined): string | undefined {
  payload.skip(6)
  const custom = payload.u32()
  const count = payload.u32()
  if (custom === 3) {
    payload.skip(4) // TOP10FT, also identified by the filter type.
    const top = payload.u16() & 1
    const limit = payload.f64()
    return Number.isFinite(limit) ? `<top10 top="${top}" val="${limit}"/>` : undefined
  }
  if (count > 2) return undefined
  const rules: string[] = []
  for (let i = 0; i < count; i++) {
    const valueType = payload.u8()
    payload.skip(1) // The comparison operator is part of the PivotTable filter type.
    const value = payload.section(8)
    if (valueType === 4) {
      const number = value.f64()
      if (!Number.isFinite(number) || (kind?.startsWith('date') && Math.abs(number) >= 100_000_000)) return undefined
      rules.push(`<customFilter val="${number}"/>`)
    } else if (!kind?.startsWith('caption')) return undefined
  }
  return rules.length ? '<customFilters>' + rules.join('') + '</customFilters>' : ''
}

/** A PivotTable's label, value and date filters, in the order of its SXAddl filter groups. */
export class PivotFilters {
  private filter?: Filter
  readonly entries: string[] = []
  lost = false

  take(property: number, payload: Payload, fieldCount: number): void {
    if (property === 0) {
      this.lost ||= this.filter !== undefined
      this.filter = { strings: new Map() }
      return
    }
    const filter = this.filter
    if (!filter) return
    if (property === 0x38) {
      payload.skip(6)
      filter.field = payload.u32()
      payload.skip(4)
      const kind = payload.u32()
      filter.type = FILTER_TYPES[kind]
      payload.skip(4)
      const measure = payload.i32()
      if ((kind <= 3 || (kind >= 0x12 && kind <= 0x19)) && measure >= 0) filter.measure = measure
    } else if (property === 0x3c) filter.criteria = criteria(payload, filter.type)
    else if (property === 0xff) {
      const { field, type, measure, criteria, strings } = filter
      this.filter = undefined
      if (field === undefined || field >= fieldCount || !type || criteria === undefined) {
        this.lost = true
        return
      }
      let attrs = ''
      for (const [name, { length, text }] of strings) {
        if (text.length !== length) {
          this.lost = true
          return
        }
        attrs += ` ${name}="${escapeAttr(text)}"`
      }
      const value = measure === undefined ? '' : ` iMeasureFld="${measure}"`
      this.entries.push(`<filter fld="${field}" type="${type}"${value}${attrs}>${criteria}</filter>`)
    } else {
      const name = STRINGS.get(property)
      if (!name) return
      const length = payload.u32()
      payload.skip(2)
      const text = (filter.strings.get(name)?.text ?? '') + unicode(payload)
      filter.strings.set(name, { length, text })
    }
  }

  get incomplete(): boolean { return this.lost || this.filter !== undefined }
}
