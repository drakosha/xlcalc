/**
 * The place of `xlcalc/pivot`'s refresh: a session's PivotTable operations refresh through it (`refreshPivots`, and
 * the PivotTable an add, edit or move changed), and without it they leave the cells and mark the cache `stale`. Its
 * registration (`registerPivot`) is the subpath's.
 */
import type { BookInput, PivotRef } from '../contract.js'
import type { Applied } from '../structure/geometry.js'
import type { ValuesIn } from '../structure/values.js'

/**
 * Lays out the PivotTables `refs` again and every PivotTable over their caches, every cache that can be without `refs`:
 * with `fromSource`, the caches refreshed from the sources' computed values (`values`), as `refreshPivots` and
 * `addPivotTable` do; without it, from the cache as saved, as Excel lays out an edit, a move, a slicer's choice. Relative
 * date filters go by the session's day `today`.
 */
export type PivotRefresh = (book: BookInput, refs: PivotRef[] | undefined, values: ValuesIn, today: number, fromSource: boolean) => Applied

let refresh: PivotRefresh | null = null

/** The refresh registered, or null. */
export const pivotRefresh = (): PivotRefresh | null => refresh

/** Registers the refresh of `xlcalc/pivot`, the last one registered used. */
export function setPivotRefresh(f: PivotRefresh): void {
  refresh = f
}
