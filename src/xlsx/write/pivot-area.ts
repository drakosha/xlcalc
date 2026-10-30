/**
 * A part of a PivotTable as a file holds it (`pivotArea`), for formats, conditional formats, calculated items and
 * sorting by a value: its kind and flags, and the items it picks on each field. The value fields' pseudo-field is -2,
 * which a reference's unsigned `field` writes as 4294967294.
 */
import type { PivotArea, PivotAreaReference, PivotAxisField, PivotFunction } from '../../contract.js'

/** A field of an area: a cache field's index, or -2 for the value fields. */
export const fieldIndex = (f: PivotAxisField) => (f === 'values' ? -2 : f)

const AXIS = { rows: 'axisRow', columns: 'axisCol', pages: 'axisPage', values: 'axisValues' } as const

/** The attribute of a reference's subtotal by function (`sumSubtotal`), and the item type of a subtotal (`t="sum"`). */
export const SUBTOTAL_ATTR: Record<PivotFunction, string> = { sum: 'sum', count: 'countA', average: 'avg', max: 'max', min: 'min', product: 'product', countNums: 'count', stdDev: 'stdDev', stdDevp: 'stdDevP', var: 'var', varp: 'varP' }

const flag = (name: string, v: boolean | undefined, dflt: boolean) => (v === undefined || v === dflt ? '' : ` ${name}="${v ? 1 : 0}"`)

function referenceXml(r: PivotAreaReference): string {
  const field = r.field === 'values' ? 4294967294 : r.field
  const items = r.items ?? []
  const subtotals = (r.subtotals ?? []).map((s) => ` ${s === 'subtotal' ? 'default' : SUBTOTAL_ATTR[s]}Subtotal="1"`).join('')
  const attrs = ` field="${field}"${items.length ? ` count="${items.length}"` : ''}${flag('selected', r.selected, true)}${flag('byPosition', r.byPosition, false)}${flag('relative', r.relative, false)}${subtotals}`
  return items.length ? `<reference${attrs}>${items.map((i) => `<x v="${i}"/>`).join('')}</reference>` : `<reference${attrs}/>`
}

/** The area's element, `pivotArea` unless named otherwise. */
export function pivotAreaXml(a: PivotArea, tag = 'pivotArea'): string {
  const attrs =
    (a.field === undefined ? '' : ` field="${fieldIndex(a.field)}"`) +
    (a.type === undefined || a.type === 'normal' ? '' : ` type="${a.type}"`) +
    flag('dataOnly', a.dataOnly, true) +
    flag('labelOnly', a.labelOnly, false) +
    flag('grandRow', a.grandRow, false) +
    flag('grandCol', a.grandColumn, false) +
    flag('cacheIndex', a.cacheIndex, false) +
    flag('outline', a.outline, true) +
    (a.offset === undefined ? '' : ` offset="${a.offset}"`) +
    flag('collapsedLevelsAreSubtotals', a.collapsedLevelsAreSubtotals, false) +
    (a.axis === undefined ? '' : ` axis="${AXIS[a.axis]}"`) +
    (a.fieldPosition === undefined ? '' : ` fieldPosition="${a.fieldPosition}"`)
  const refs = a.references ?? []
  return refs.length ? `<${tag}${attrs}><references count="${refs.length}">${refs.map(referenceXml).join('')}</references></${tag}>` : `<${tag}${attrs}/>`
}
