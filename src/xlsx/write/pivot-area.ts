/**
 * A part of a PivotTable as a file holds it (`pivotArea`), for formats, conditional formats, calculated items and
 * sorting by a value: its kind and flags, and the items it picks on each field. The value fields' pseudo-field is -2,
 * which a reference's unsigned `field` writes as 4294967294.
 */
import type { PivotArea, PivotAreaReference, PivotAxisField } from '../../contract.js'
import { AXIS, SUBTOTAL_ATTR } from '../pivot-names.js'
import { give } from './xml.js'

/** A field of an area: a cache field's index, or -2 for the value fields. */
export const fieldIndex = (f: PivotAxisField) => (f === 'values' ? -2 : f)

function referenceXml(r: PivotAreaReference): string {
  const field = r.field === 'values' ? 4294967294 : r.field
  const items = r.items ?? []
  const subtotals = (r.subtotals ?? []).map((s) => ` ${s === 'subtotal' ? 'default' : SUBTOTAL_ATTR[s]}Subtotal="1"`).join('')
  const attrs = ` field="${field}"${items.length ? ` count="${items.length}"` : ''}${give(r, 'selected:t byPosition:f relative:f')}${subtotals}`
  return items.length ? `<reference${attrs}>${items.map((i) => `<x v="${i}"/>`).join('')}</reference>` : `<reference${attrs}/>`
}

/** The area's element, `pivotArea` unless named otherwise. */
export function pivotAreaXml(a: PivotArea, tag = 'pivotArea'): string {
  const attrs =
    (a.field === undefined ? '' : ` field="${fieldIndex(a.field)}"`) +
    (a.type === undefined || a.type === 'normal' ? '' : ` type="${a.type}"`) +
    give(a, 'dataOnly:t labelOnly:f grandRow:f grandCol=grandColumn:f cacheIndex:f outline:t offset:n collapsedLevelsAreSubtotals:f') +
    (a.axis === undefined ? '' : ` axis="${AXIS[a.axis]}"`) +
    give(a, 'fieldPosition:n')
  const refs = a.references ?? []
  return refs.length ? `<${tag}${attrs}><references count="${refs.length}">${refs.map(referenceXml).join('')}</references></${tag}>` : `<${tag}${attrs}/>`
}
