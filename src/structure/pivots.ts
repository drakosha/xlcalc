/**
 * The workbook's PivotTables through operations of the structure: a sheet renamed gives its new name to the pivot
 * caches whose source is on it and to the slicers' and timelines' PivotTables on it.
 */
import type { BookInput, PivotCache, PivotRef, PivotSource } from '../contract.js'
import { caseKey } from '../collate.js'
import { checkBook } from '../input.js'
import { pivotCheck, pivotRefresh } from '../pivot/hook.js'
import type { Applied, Geometry } from './geometry.js'
import { sheetIndex } from './rewrite.js'
import type { ValuesIn } from './values.js'

/** A source's sheet renamed; a source in another workbook keeps its sheet. */
function renamedSource(s: PivotSource, rename: (name: string) => string): PivotSource {
  if (s.type === 'range' && s.book === undefined && rename(s.sheet) !== s.sheet) return { ...s, sheet: rename(s.sheet) }
  if (s.type === 'consolidation' && s.ranges.some((r) => r.book === undefined && r.sheet !== undefined && rename(r.sheet) !== r.sheet))
    return { ...s, ranges: s.ranges.map((r) => (r.book === undefined && r.sheet !== undefined ? { ...r, sheet: rename(r.sheet) } : r)) }
  return s
}

const renamedRefs = (refs: PivotRef[], rename: (name: string) => string): PivotRef[] => (refs.some((r) => rename(r.sheet) !== r.sheet) ? refs.map((r) => ({ ...r, sheet: rename(r.sheet) })) : refs)

/** The book's caches, slicer and timeline caches with a sheet's new name, `rename` giving it; the book itself when none names it. */
export function renamePivotSheets(book: BookInput, rename: (name: string) => string): BookInput {
  if (!book.pivotCaches && !book.slicerCaches && !book.timelineCaches) return book
  const out = { ...book }
  if (book.pivotCaches) out.pivotCaches = book.pivotCaches.map((c): PivotCache => {
    const source = renamedSource(c.source, rename)
    return source === c.source ? c : { ...c, source }
  })
  if (book.slicerCaches) out.slicerCaches = book.slicerCaches.map((c) => ({ ...c, pivots: renamedRefs(c.pivots, rename) }))
  if (book.timelineCaches) out.timelineCaches = book.timelineCaches.map((c) => ({ ...c, pivots: renamedRefs(c.pivots, rename) }))
  return out
}

/** The patch's fields set on `target`: a value replaces, null removes. */
export function patched<T extends object>(target: T, patch: Record<string, unknown>): T {
  const out: Record<string, unknown> = { ...(target as Record<string, unknown>) }
  for (const [k, v] of Object.entries(patch)) {
    if (v === null) delete out[k]
    else if (v !== undefined) out[k] = v
  }
  return out as T
}

/**
 * The PivotTables `refs` laid out again by `xlcalc/pivot` (`fromSource`: their caches refreshed from the sources); without
 * it, their caches marked `stale`, the cells left.
 */
export function relaid(book: BookInput, refs: PivotRef[], values: ValuesIn, fromSource: boolean, geometry: Geometry = []): Applied {
  const refresh = pivotRefresh()
  const r = refresh?.(book, refs, values, values.today?.() ?? 0, fromSource)
  if (r && !r.ok) return r
  // laid out, unless the cache keeps no rows to lay it out from
  if (r && r.book !== book) return { ...r, geometry: [...geometry, ...r.geometry] }
  const caches = new Set(refs.flatMap((r) => book.sheets[sheetIndex(book, r.sheet)]?.pivotTables?.filter((p) => caseKey(p.name) === caseKey(r.name)).map((p) => p.cache) ?? []))
  const pivotCaches = book.pivotCaches?.map((c) => (caches.has(c.id) && !c.stale ? { ...c, stale: true } : c))
  return { ok: true, book: pivotCaches ? { ...book, pivotCaches } : book, geometry, stale: refs }
}

/** The book checked as input, the full check of PivotTables too when `xlcalc/pivot` registered it. */
export function checkedPivots(book: BookInput, path = 'op'): BookInput {
  const out = checkBook(book, path)
  pivotCheck()?.(out, path)
  return out
}
