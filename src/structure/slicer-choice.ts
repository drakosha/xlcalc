/**
 * A slicer's and a timeline's choice as the PivotTables they filter hold it, as Excel saves it: a slicer's items not
 * chosen are its field's hidden items on each PivotTable (on an axis or not), a timeline's dates a `dateBetween` filter
 * of its field. The PivotTables' hidden items are what a refresh and a calculation read; the slicer cache's choice says
 * the same.
 */
import type { BookInput, PivotCache, PivotFilter, PivotRef, PivotTable, SlicerCache, TimelineCache, TimelineLevel } from '../contract.js'
import { caseKey } from '../collate.js'
import { groupedItems, withListed } from '../pivot/items.js'
import { sheetKey } from '../syntax.js'

/** The index of a cache's field by its name, −1 for none. */
export const fieldIndex = (cache: PivotCache, name: string): number => cache.fields.findIndex((f) => caseKey(f.name) === caseKey(name))

/** The same PivotTable: sheet and name in any case. */
export const sameRef = (a: PivotRef, b: PivotRef) => sheetKey(a.sheet) === sheetKey(b.sheet) && caseKey(a.name) === caseKey(b.name)

/** The book with each PivotTable of `refs` given by `f`. */
function withPivots(book: BookInput, refs: readonly PivotRef[], f: (p: PivotTable) => PivotTable): BookInput {
  const named = (s: string, p: PivotTable) => refs.some((r) => sameRef(r, { sheet: s, name: p.name }))
  const sheets = book.sheets.map((s) => (s.pivotTables?.some((p) => named(s.name, p)) ? { ...s, pivotTables: s.pivotTables.map((p) => (named(s.name, p) ? f(p) : p)) } : s))
  return { ...book, sheets }
}

/**
 * The book with the slicer cache's field listed as items in its pivot cache (Excel lists a field it makes a slicer of),
 * the field's index, and the cache.
 */
export function slicerField(book: BookInput, sc: Pick<SlicerCache, 'field' | 'cache'>): { book: BookInput; field: number; cache: PivotCache } | null {
  const pc = book.pivotCaches?.find((c) => c.id === sc.cache)
  const f = pc ? fieldIndex(pc, sc.field) : -1
  if (!pc || f < 0) return null
  const listed = withListed(pc, [f])
  if (listed === pc) return { book, field: f, cache: pc }
  return { book: { ...book, pivotCaches: book.pivotCaches!.map((c) => (c === pc ? listed : c)) }, field: f, cache: listed }
}

/** A PivotTable's field `f` with an item for each of the cache field's `n` items, the ones it lacks after its own in the cache's order. */
function itemized(p: PivotTable, f: number, n: number): PivotTable['fields'][number] {
  const pf = p.fields[f] ?? {}
  const known = new Set((pf.items ?? []).flatMap((x) => (x.item === undefined ? [] : [x.item])))
  const added = Array.from({ length: n }, (_, k) => k).filter((k) => !known.has(k))
  return added.length ? { ...pf, items: [...(pf.items ?? []), ...added.map((item) => ({ item }))] } : pf
}

/** The items a new slicer cache of field `f` starts with: chosen unless hidden on the first PivotTable it filters, as Excel takes a filter already there. */
export function itemsShown(book: BookInput, ref: PivotRef, f: number, n: number): NonNullable<SlicerCache['items']> {
  const p = book.sheets.find((s) => sheetKey(s.name) === sheetKey(ref.sheet))?.pivotTables?.find((q) => caseKey(q.name) === caseKey(ref.name))
  const items = p?.fields[f]?.items ?? []
  // a page field's one item shown is the one chosen
  const page = p?.pages?.find((x) => x.field === f)?.item
  const shown = page !== undefined ? items[page]?.item : undefined
  const hidden = new Set(items.flatMap((x) => (x.item !== undefined && (shown !== undefined ? x.item !== shown : x.hidden) ? [x.item] : [])))
  return Array.from({ length: n }, (_, item) => (hidden.has(item) ? { item, selected: false } : { item }))
}

/** The book with the slicer cache's choice as the hidden items of its field on the PivotTables `refs` (its own by default). */
export function withSlicerChoice(book: BookInput, sc: SlicerCache, refs: readonly PivotRef[] = sc.pivots): BookInput {
  const at = slicerField(book, sc)
  if (!at || !refs.length) return book
  const n = groupedItems(at.cache.fields[at.field])?.length ?? 0
  const off = new Set((sc.items ?? []).flatMap((x) => (x.selected === false ? [x.item] : [])))
  return withPivots(
    at.book,
    refs,
    (p) => {
      const pf = itemized(p, at.field, n)
      const items = pf.items?.map((x) => {
        const hide = x.item !== undefined && off.has(x.item)
        if (!!x.hidden === hide) return x
        const { hidden: _, ...rest } = x
        return hide ? { ...rest, hidden: true } : rest
      })
      const fields = p.fields.slice()
      // a page field shows the items chosen, Select Multiple Items, in place of its one item
      const page = p.pages?.some((x) => x.field === at.field && x.item !== undefined)
      fields[at.field] = { ...pf, ...(items ? { items } : {}), ...(page ? { multipleItems: true } : {}) }
      return { ...p, fields, ...(page ? { pages: p.pages!.map((x) => (x.field === at.field ? { field: x.field } : x)) } : {}) }
    },
  )
}

/** A filter of dates by day, not by label or value: the kind a timeline replaces on its field. */
const byDate = (f: PivotFilter) => !/^(caption|value)/.test(f.type) && !['count', 'percent', 'sum'].includes(f.type)

/** A timeline's dates as Excel filters its field by them: by whole days, between the two, or on the one day when they are one (the set slicer). */
function timelineFilter(f: number, name: string, s: NonNullable<TimelineCache['selection']>): PivotFilter {
  const [start, end] = [s.start.slice(0, 10), s.end.slice(0, 10)]
  return start === end ? { field: f, type: 'dateEqual', args: [start], name } : { field: f, type: 'dateBetween', args: [start, end], name }
}

/** The book with the timeline's dates as a date filter of its field on the PivotTables `refs` (its own by default), none without a selection; the field stays without items, as Excel keeps it. */
export function withTimelineChoice(book: BookInput, tc: TimelineCache, refs: readonly PivotRef[] = tc.pivots): BookInput {
  const pc = book.pivotCaches?.find((c) => c.id === tc.cache)
  const f = pc ? fieldIndex(pc, tc.field) : -1
  if (!pc || f < 0 || !refs.length) return book
  const own = tc.selection ? timelineFilter(f, pc.fields[f].name, tc.selection) : null
  return withPivots(
    book,
    refs,
    (p) => {
      const filters = [...(p.filters ?? []).filter((x) => x.field !== f || !byDate(x)), ...(own ? [own] : [])]
      const { filters: _, ...rest } = p
      return filters.length ? { ...rest, filters } : rest
    },
  )
}

/**
 * The level a timeline's dates are chosen at, as Excel sets it for a range of dates: the widest whose periods they span
 * whole, from the first day of one to the last day of another (2025-01-01 to 2025-06-30 quarters, 2024-03-01 to
 * 2024-12-31 months).
 */
export function selectionLevel(s: NonNullable<TimelineCache['selection']>): TimelineLevel {
  const [, m1, d1] = s.start.slice(0, 10).split('-').map(Number)
  // the day after the end, which starts a period when the end closes one
  const [y, m, d] = s.end.slice(0, 10).split('-').map(Number)
  const next = new Date(Date.UTC(y, m - 1, d + 1))
  const [m2, d2] = [next.getUTCMonth() + 1, next.getUTCDate()]
  if (d1 !== 1 || d2 !== 1) return 'days'
  if (m1 === 1 && m2 === 1) return 'years'
  return (m1 - 1) % 3 === 0 && (m2 - 1) % 3 === 0 ? 'quarters' : 'months'
}

/** The dates a timeline of a field spans, as Excel bounds it: from the first day of the year of the field's first date to the first day of the year after its last. */
export function timelineBounds(pc: PivotCache, f: number): TimelineCache['bounds'] {
  const facts = pc.fields[f]?.facts
  const dates = [facts?.min, facts?.max, ...(pc.fields[f]?.items ?? [])].flatMap((v) => (v !== null && typeof v === 'object' && 'date' in v ? [v.date] : []))
  if (!dates.length) return undefined
  dates.sort()
  return { start: `${dates[0].slice(0, 4)}-01-01T00:00:00`, end: `${String(Number(dates[dates.length - 1].slice(0, 4)) + 1).padStart(4, '0')}-01-01T00:00:00` }
}
