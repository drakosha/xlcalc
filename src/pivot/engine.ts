/**
 * The refresh `xlcalc/pivot` registers (`PivotRefresh`): the caches of the PivotTables asked for refreshed from their
 * sources (or taken as saved), every PivotTable over them calculated again, all of them or none: a refusal of one is
 * the refusal of the whole operation.
 */
import type { BookInput, PivotCache, PivotTable } from '../contract.js'
import { caseKey } from '../collate.js'
import { EPOCH_1904 } from '../dates.js'
import { listsOf } from '../locale/lists.js'
import { say } from '../locale/messages.js'
import type { Applied } from '../structure/geometry.js'
import type { ValuesIn } from '../structure/values.js'
import { computePivot } from './compute.js'
import type { PivotRefresh } from './hook.js'
import { refreshCache } from './refresh.js'
import { layoutPivot } from './layout.js'
import { placePivots, type Placement } from './place.js'
import { sourceGrid } from './source.js'

/** A PivotTable of the book and its sheet's index. */
interface Spot {
  s: number
  table: PivotTable
}

const refused = (book: BookInput, m: Parameters<typeof say>[1]): Applied => ({ ok: false, code: 'pivot', message: say(book.locale, m) })

export const refreshPivots: PivotRefresh = (book, refs, values: ValuesIn, today, fromSource, blocked) => {
  const spots: Spot[] = book.sheets.flatMap((sh, s) => (sh.pivotTables ?? []).map((table) => ({ s, table })))
  const caches = new Map((book.pivotCaches ?? []).map((c) => [c.id, c]))
  let wanted: Set<number>
  if (refs) {
    wanted = new Set()
    for (const ref of refs) {
      const spot = spots.find((x) => caseKey(book.sheets[x.s].name) === caseKey(ref.sheet) && caseKey(x.table.name) === caseKey(ref.name))
      if (!spot) return refused(book, { key: 'pivotUnknown', sheet: ref.sheet, pivot: ref.name })
      wanted.add(spot.table.cache)
    }
  } else wanted = new Set(caches.keys())
  const epoch = book.dateSystem === 1904 ? EPOCH_1904 : 0
  const options = { today, date1904: epoch > 0, lists: listsOf(book) }
  const placed: Placement[] = []
  const newCaches = new Map<number, PivotCache>()
  for (const id of wanted) {
    const cache = caches.get(id)
    const over = spots.filter((x) => x.table.cache === id)
    if (!cache || !over.length) continue
    let tables = over.map((x) => x.table)
    // laid out again, the cache is no longer stale
    const { stale: _, ...kept } = cache
    let next: PivotCache = kept
    if (fromSource) {
      const first = over[0].table.name
      const grid = cache.enableRefresh === false ? 'disabled' : sourceGrid(book, cache.source, values)
      if (typeof grid === 'string') {
        if (!refs) continue
        return refused(book, grid === 'not_found' ? { key: 'pivotSourceGone', pivot: first } : { key: 'pivotSourceKept', pivot: first, why: grid === 'other_book' ? 'outside' : grid === 'not_refreshed' ? 'kind' : 'disabled' })
      }
      const r = refreshCache(cache, tables, grid, epoch)
      if (typeof r === 'string') {
        if (refs) return refused(book, { key: 'pivotFieldName', pivot: first })
        continue
      }
      next = r.cache
      tables = r.tables
    } else if (!cache.records) continue
    over.forEach((spot, k) => {
      const result = computePivot(tables[k], next, options)
      next = result.cache
      placed.push({ sheet: book.sheets[spot.s].name, table: { ...result.table, name: spot.table.name }, laid: layoutPivot(result, book) })
    })
    newCaches.set(id, next)
  }
  if (!newCaches.size) return { ok: true, book, geometry: [] }
  return placePivots({ ...book, pivotCaches: (book.pivotCaches ?? []).map((c) => newCaches.get(c.id) ?? c) }, placed, blocked)
}
