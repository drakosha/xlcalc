/**
 * Excel 365's time grouping (`groupDates` of `addPivotTable` and `editPivotTable`): a field of dates checked in the
 * PivotTable Fields list, or dragged to the rows or the columns, is grouped by periods as the Group dialog groups it.
 * Excel does it from its field list alone, not when a macro places the field; here the operation asks for it.
 */
import type { PivotCacheField, PivotGroupBy, PivotRef, PivotTable, PivotValue } from '../contract.js'
import { caseKey } from '../collate.js'
import { sayPivot } from '../locale/pivot-messages.js'
import { pivotCacheEdit } from '../pivot/hook.js'
import type { Applied } from './geometry.js'
import { sheetIndex } from './rewrite.js'
import type { ValuesIn } from './values.js'

/** A field of dates (and empty cells) of the source, not grouped. */
function datesOnly(f: PivotCacheField): boolean {
  const types = f.facts?.types ?? []
  return types.includes('date') && types.every((t) => t === 'date' || t === 'blank') && !f.group && f.formula === undefined
}

const yearOf = (v: PivotValue | number | undefined) => (v !== null && typeof v === 'object' && 'date' in v ? v.date.slice(0, 4) : undefined)

/** The periods Excel groups a field of dates by: months alone within a year, months, quarters and years over several. */
function periodsOf(f: PivotCacheField): Exclude<PivotGroupBy, 'numbers'>[] {
  const years = new Set([...(f.items ?? []), f.facts?.min, f.facts?.max].map(yearOf).filter((y) => y !== undefined))
  return years.size > 1 ? ['months', 'quarters', 'years'] : ['months']
}

/**
 * The PivotTable `ref` laid out (`r`) with each field of dates it has on its rows or columns and `before` had not
 * grouped, as `groupPivotField` groups it; refused (`pivot`) without `xlcalc/pivot` when there is one.
 */
export function groupedDates(r: Applied, ref: PivotRef, before: PivotTable | undefined, values: ValuesIn): Applied {
  if (!r.ok) return r
  const was = new Set([...(before?.rows ?? []), ...(before?.columns ?? [])])
  const t = r.book.sheets[sheetIndex(r.book, ref.sheet)].pivotTables!.find((p) => caseKey(p.name) === caseKey(ref.name))!
  const cache = r.book.pivotCaches!.find((c) => c.id === t.cache)!
  const fields = [...(t.rows ?? []), ...(t.columns ?? [])].filter((f): f is number => f !== 'values' && !was.has(f) && datesOnly(cache.fields[f]))
  if (!fields.length) return r
  const edit = pivotCacheEdit()
  if (!edit) return { ok: false, code: 'pivot', message: sayPivot(r.book.locale, { key: 'pivotNotLoaded' }) }
  let out = r
  for (const field of fields) {
    const next = edit(out.book, { op: 'groupPivotField', sheet: ref.sheet, pivot: ref.name, field, group: { by: 'dates', periods: periodsOf(cache.fields[field]) } }, values, values.today?.() ?? 0)
    if (!next.ok) return next
    out = { ...next, geometry: [...out.geometry, ...next.geometry] }
  }
  return out
}
