/**
 * Excel 365's time grouping (`groupDates` of `addPivotTable` and `editPivotTable`): a field of dates checked in the
 * PivotTable Fields list is grouped by periods as the Group dialog groups it, its dates kept after the periods and the
 * periods collapsed. Excel does it from its field list alone, not when a macro places the field (`Orientation`,
 * `AddFields`); here the operation asks for it.
 */
import type { PivotCacheField, PivotGroupBy, PivotRef, PivotTable } from '../contract.js'
import { caseKey } from '../collate.js'
import { sayPivot } from '../locale/pivot-messages.js'
import { pivotCacheEdit, pivotRefresh } from '../pivot/hook.js'
import type { Applied } from './geometry.js'
import { sheetIndex } from './rewrite.js'
import type { ValuesIn } from './values.js'

/** A field of dates (and empty cells) of the source, not grouped. */
function datesOnly(f: PivotCacheField): boolean {
  const types = f.facts?.types ?? []
  return types.includes('date') && types.every((t) => t === 'date' || t === 'blank') && !f.group && f.formula === undefined
}

/** The calendar years and months (`2025-03`) of a field's dates. */
function spanOf(f: PivotCacheField): { years: number; months: number } {
  const dates = [...(f.items ?? []), f.facts?.min, f.facts?.max].flatMap((v) => (v !== null && typeof v === 'object' && 'date' in v ? [v.date] : []))
  return { years: new Set(dates.map((d) => d.slice(0, 4))).size, months: new Set(dates.map((d) => d.slice(0, 7))).size }
}

/** The periods Excel groups a field of dates by: years, quarters and months over several years, months and days over several months of one, none within a month. */
function periodsOf(f: PivotCacheField): Exclude<PivotGroupBy, 'numbers'>[] {
  const { years, months } = spanOf(f)
  return years > 1 ? ['months', 'quarters', 'years'] : months > 1 ? ['days', 'months'] : []
}

/**
 * The PivotTable with the dates of `base` after the fields of their periods, as the field list leaves them, and the
 * periods' items collapsed, so that the PivotTable shows the coarsest period alone.
 */
function withDatesLast(t: PivotTable, base: number, periods: number[]): PivotTable {
  const axis = (list?: PivotTable['rows']) => {
    if (!list) return list
    const at = list.findIndex((f) => typeof f === 'number' && periods.includes(f))
    if (at < 0) return list
    const out = list.filter((f) => f !== base)
    out.splice(out.findLastIndex((f) => typeof f === 'number' && periods.includes(f)) + 1, 0, base)
    return out
  }
  const fields = t.fields.map((f, i) => (periods.includes(i) && f.items ? { ...f, items: f.items.map((x) => ({ ...x, collapsed: true })) } : f))
  return { ...t, rows: axis(t.rows), columns: axis(t.columns), fields }
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
  const refresh = pivotRefresh()!
  let out = r
  for (const field of fields) {
    const periods = periodsOf(cache.fields[field])
    if (!periods.length) continue
    const grouped = edit(out.book, { op: 'groupPivotField', sheet: ref.sheet, pivot: ref.name, field, group: { by: 'dates', periods } }, values, values.today?.() ?? 0)
    if (!grouped.ok) return grouped
    // the fields of the periods are the last of the cache
    const made = grouped.book.pivotCaches!.find((c) => c.id === t.cache)!.fields.flatMap((f, i) => (f.group?.base === field && f.group.range ? [i] : []))
    const s = sheetIndex(grouped.book, ref.sheet)
    const sheets = [...grouped.book.sheets]
    sheets[s] = { ...sheets[s], pivotTables: sheets[s].pivotTables!.map((p) => (caseKey(p.name) === caseKey(ref.name) ? withDatesLast(p, field, made) : p)) }
    const next = refresh({ ...grouped.book, sheets }, [ref], values, values.today?.() ?? 0, false)
    if (!next.ok) return next
    out = { ...next, geometry: [...out.geometry, ...grouped.geometry, ...next.geometry] }
  }
  return out
}
