/**
 * The place of `xlcalc/pivot`'s refresh: a session's PivotTable operations refresh through it (`refreshPivots`, and
 * the PivotTable an add, edit or move changed), and without it they leave the cells and mark the cache `stale`. Its
 * registration (`registerPivot`) is the subpath's.
 */
import type { BookInput, PivotRef } from '../contract.js'
import type { Applied } from '../structure/geometry.js'

/** Refreshes the caches of the PivotTables `refs` and every PivotTable over them, every cache that can be without `refs`. */
export type PivotRefresh = (book: BookInput, refs?: PivotRef[]) => Applied

let refresh: PivotRefresh | null = null

/** The refresh registered, or null. */
export const pivotRefresh = (): PivotRefresh | null => refresh

/** Registers the refresh of `xlcalc/pivot`, the last one registered used. */
export function setPivotRefresh(f: PivotRefresh): void {
  refresh = f
}
