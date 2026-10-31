/**
 * The place of `xlcalc/pivot`'s refresh: a session's PivotTable operations refresh through it (`refreshPivots`, and
 * the PivotTable an add, edit or move changed), and without it they leave the cells and mark the cache `stale`. Its
 * registration (`registerPivot`) is the subpath's.
 */
import type { BookInput, CellInput, PivotRef, Refusal, StructureOp } from '../contract.js'
import type { Applied } from '../structure/geometry.js'
import type { ValuesIn } from '../structure/values.js'

/**
 * Lays out the PivotTables `refs` again and every PivotTable over their caches, every cache that can be without `refs`:
 * with `fromSource`, the caches refreshed from the sources' computed values (`values`), as `refreshPivots` and
 * `addPivotTable` do; without it, from the cache as saved, as Excel lays out an edit, a move, a slicer's choice. Relative
 * date filters go by the session's day `today`. A PivotTable laid over cells it cannot take spills, as a refresh in
 * Excel 365 (`spill`, the default), or is refused, as a new or moved one (`refuse`).
 */
export type PivotRefresh = (book: BookInput, refs: PivotRef[] | undefined, values: ValuesIn, today: number, fromSource: boolean, blocked?: 'spill' | 'refuse', labelSort?: readonly number[]) => Applied

let sortThreads = 2
let callSortThreads: number | undefined

/** Excel's sorting-thread profile for the registered engine. */
export const pivotSortThreads = (): number => callSortThreads ?? sortThreads

/** The registered default, independent of an operation being replayed. */
export const registeredPivotSortThreads = (): number => sortThreads

/** A worker call keeps its profile across yields; registration may still change the default for later calls. */
export async function withPivotSortThreads<T>(threads: number, run: () => Promise<T>): Promise<T> {
  const prior = callSortThreads
  callSortThreads = threads
  try { return await run() } finally { callSortThreads = prior }
}

/** Sets the sorting-thread profile, already checked by registerPivot. */
export function setPivotSortThreads(n: number): void { sortThreads = n }

let refresh: PivotRefresh | null = null

/** The refresh registered, or null. */
export const pivotRefresh = (): PivotRefresh | null => refresh

/** Registers the refresh of `xlcalc/pivot`, the last one registered used; null takes it away (a test's session without it). */
export function setPivotRefresh(f: PivotRefresh | null): void {
  refresh = f
}

/** The full check of PivotTables' shape (`pivot/check.ts`), which `xlcalc/pivot` registers for every check of a workbook. */
export type PivotCheck = (book: BookInput, path: string) => void

let check: PivotCheck | null = null

/** The full check registered, or null. */
export const pivotCheck = (): PivotCheck | null => check

/** Registers the full check of PivotTables, the last one registered used. */
export function setPivotCheck(f: PivotCheck): void {
  check = f
}

/**
 * An entry into a PivotTable's cell as `xlcalc/pivot` takes it (`pivot/entry.ts`): an edit of the PivotTable's
 * definition, nothing, or a refusal; null for a cell of no PivotTable. Without it, every entry there is refused.
 */
export type PivotEntryHook = (book: BookInput, sheet: string, cell: string, input: CellInput | null) => { ok: true; op: StructureOp | null } | Refusal<'pivot'> | null

let entry: PivotEntryHook | null = null

/** The entry registered, or null. */
export const pivotEntryHook = (): PivotEntryHook | null => entry

/** Registers the entry into PivotTables' cells, the last one registered used. */
export function setPivotEntry(f: PivotEntryHook): void {
  entry = f
}

/** The operations at the level of a pivot cache, which `xlcalc/pivot` applies (`pivot/cache-ops.ts`): grouping, calculated fields and items. */
export type PivotCacheOp = Extract<StructureOp, { op: 'groupPivotField' | 'ungroupPivotField' | 'renamePivotGroup' | 'setPivotCalculatedField' | 'setPivotCalculatedItem' }>

/** Applies a cache's operation and lays out every PivotTable of the cache again; the day `today` as in `PivotRefresh`. */
export type PivotCacheEdit = (book: BookInput, op: PivotCacheOp, values: ValuesIn, today: number) => Applied

let cacheEdit: PivotCacheEdit | null = null

/** The cache's operations registered, or null. */
export const pivotCacheEdit = (): PivotCacheEdit | null => cacheEdit

/** Registers the cache's operations of `xlcalc/pivot`, the last one registered used. */
export function setPivotCacheEdit(f: PivotCacheEdit): void {
  cacheEdit = f
}
