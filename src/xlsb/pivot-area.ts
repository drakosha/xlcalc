import { AXIS, SUBTOTAL_ATTR } from '../xlsx/pivot-names.js'
import { range } from './fields.js'
import type { Payload } from './records.js'
import type { Xml } from './xml.js'

export const PIVOT_AGGREGATES = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'] as const
export const pivotBit = (value: number, index: number): number => (value >>> index) & 1

const AREA_TYPES = ['none', 'normal', 'data', 'all', 'origin', 'button', 'topRight']
const AXES: Readonly<Partial<Record<number, string>>> = { 1: AXIS.rows, 2: AXIS.columns, 4: AXIS.pages, 8: AXIS.values }

/** Pivot area records shared by formats, conditional formatting and the sort scopes of fields. */
export function pivotAreaRecord(xml: Xml, type: number, payload: Payload): boolean {
  let text: string
  switch (type) {
    case 247: case 1161:
      if (type === 1161) payload.skip(4)
      text = area(payload)
      break
    case 248: case 1162: text = '</pivotArea>'; break
    case 249: case 1163: {
      if (type === 1163) payload.skip(4)
      text = `<references count="${payload.u32()}">`
      break
    }
    case 250: case 1164: text = '</references>'; break
    case 251: case 1165:
      if (type === 1165) payload.skip(4)
      text = reference(payload)
      break
    case 252: case 1166: text = '</reference>'; break
    case 382: case 1167:
      if (type === 1167) payload.skip(4)
      text = `<x v="${payload.u32()}"/>`
      break
    default: return false
  }
  payload.done()
  xml.add(text)
  return true
}

function area(payload: Payload): string {
  const field = payload.i32()
  const kind = AREA_TYPES[payload.u8()]
  const flags = payload.u8()
  const axes = payload.u16()
  const axis = AXES[axes & 15]
  const fieldAttr = field === -1 ? '' : ` field="${field}"`
  const axisAttrs = axis ? ` axis="${axis}"${axis === AXIS.values ? '' : ` fieldPosition="${(axes >>> 4) & 255}"`}` : ''
  const offset = flags & 0x40 ? ` offset="${range(payload)}"` : ''
  return `<pivotArea type="${kind}" dataOnly="${pivotBit(flags, 0)}" labelOnly="${pivotBit(flags, 1)}"`
    + ` grandRow="${pivotBit(flags, 2)}" grandCol="${pivotBit(flags, 3)}" cacheIndex="${pivotBit(flags, 4)}"`
    + ` outline="${pivotBit(flags, 5)}" collapsedLevelsAreSubtotals="${pivotBit(flags, 7)}"${fieldAttr}${axisAttrs}${offset}>`
}

function reference(payload: Payload): string {
  const field = payload.i32()
  const count = payload.u32()
  const flags = payload.u16()
  const selection = payload.u8()
  const subtotals = PIVOT_AGGREGATES.map((fn, i) => pivotBit(flags, i + 2) ? ` ${SUBTOTAL_ATTR[fn]}Subtotal="1"` : '').join('')
  return `<reference field="${field}" count="${count}" selected="${pivotBit(selection, 0)}" defaultSubtotal="${pivotBit(flags, 1)}"${subtotals}>`
}
