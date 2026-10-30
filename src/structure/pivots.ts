/**
 * The workbook's PivotTables through operations of the structure: a sheet renamed gives its new name to the pivot
 * caches whose source is on it and to the slicers' and timelines' PivotTables on it.
 */
import type { BookInput, PivotCache, PivotRef, PivotSource } from '../contract.js'

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
