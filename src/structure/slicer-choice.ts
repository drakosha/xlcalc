/**
 * A slicer's and a timeline's choice as the PivotTables they filter hold it, as Excel saves it: a slicer's items not
 * chosen are its field's hidden items on each PivotTable (on an axis or not), a timeline's dates a `dateBetween` filter
 * of its field. The PivotTables' hidden items are what a refresh and a calculation read; the slicer cache's choice says
 * the same.
 */
import type { BookInput, PivotCache, PivotFilter, PivotRef, PivotTable, SlicerCache, TimelineCache } from '../contract.js'
import { caseKey } from '../collate.js'
import { withListed } from '../pivot/items.js'
import { sheetKey } from '../syntax.js'

/** The index of a cache's field by its name, −1 for none. */
export const fieldIndex = (cache: PivotCache, name: string): number => cache.fields.findIndex((f) => caseKey(f.name) === caseKey(name))

const sameRef = (a: PivotRef, b: PivotRef) => sheetKey(a.sheet) === sheetKey(b.sheet) && caseKey(a.name) === caseKey(b.name)

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
  const hidden = new Set((p?.fields[f]?.items ?? []).flatMap((x) => (x.hidden && x.item !== undefined ? [x.item] : [])))
  return Array.from({ length: n }, (_, item) => (hidden.has(item) ? { item, selected: false } : { item }))
}

/** The book with the slicer cache's choice as the hidden items of its field on the PivotTables `refs` (its own by default). */
export function withSlicerChoice(book: BookInput, sc: SlicerCache, refs: readonly PivotRef[] = sc.pivots): BookInput {
  const at = slicerField(book, sc)
  if (!at || !refs.length) return book
  const n = at.cache.fields[at.field].items?.length ?? 0
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
      fields[at.field] = { ...pf, ...(items ? { items } : {}) }
      return { ...p, fields }
    },
  )
}

/** A filter of dates by day, not by label or value: the kind a timeline replaces on its field. */
const byDate = (f: PivotFilter) => !/^(caption|value)/.test(f.type) && !['count', 'percent', 'sum'].includes(f.type)

/** The book with the timeline's dates as a `dateBetween` filter of its field on the PivotTables `refs` (its own by default), none without a selection. */
export function withTimelineChoice(book: BookInput, tc: TimelineCache, refs: readonly PivotRef[] = tc.pivots): BookInput {
  const pc = book.pivotCaches?.find((c) => c.id === tc.cache)
  const f = pc ? fieldIndex(pc, tc.field) : -1
  if (!pc || f < 0 || !refs.length) return book
  const own: PivotFilter | null = tc.selection ? { field: f, type: 'dateBetween', args: [tc.selection.start, tc.selection.end] } : null
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
