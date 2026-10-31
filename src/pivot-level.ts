/**
 * The PivotTable engine of `xlcalc/pivot` that `registerPivot` loaded: the refresh a session's PivotTable operations
 * take, the full check of the workbook's PivotTables, an entry into their cells and the operations on their caches (`pivot/hook.ts`). A pool's threads
 * load the same engine (`worker/thread.ts`).
 */
import type { PivotEngine, PivotOptions } from './contract.js'
import { object, onlyOptions } from './input/basic.js'
import { InputError } from './input-error.js'
import { pivotSortThreads, setPivotSortThreads, pivotRefresh, setPivotCacheEdit, setPivotCheck, setPivotEntry, setPivotRefresh, type PivotCacheEdit, type PivotCheck, type PivotEntryHook, type PivotRefresh } from './pivot/hook.js'

/** The engine's parts besides its kind, which the subpath gives and the contract does not show. */
export interface PivotEngineParts extends PivotEngine {
  refresh: PivotRefresh
  check: PivotCheck
  entry: PivotEntryHook
  edit: PivotCacheEdit
}

const listeners = new Set<() => void>()

/** The settings a pool sends to its threads. */
export const pivotOptions = (): PivotOptions => ({ sortThreads: pivotSortThreads() })

/** Whether the engine is loaded. */
export const pivotLoaded = (): boolean => pivotRefresh() !== null

/**
 * Loads the PivotTable engine of `xlcalc/pivot` for every session from now on, a pool's threads too: `import {
 * pivotEngine } from 'xlcalc/pivot'`, then `registerPivot(pivotEngine)`. `sortThreads` selects the Excel sorting profile (2 by default),
 * including for already open sessions on their next refresh; it creates no threads.
 */
export function registerPivot(engine: PivotEngine, options?: PivotOptions): void {
  const parts = engine as PivotEngineParts
  if (typeof engine !== 'object' || engine === null || engine.kind !== 'xlcalc-pivot' || typeof parts.refresh !== 'function' || typeof parts.check !== 'function' || typeof parts.entry !== 'function' || typeof parts.edit !== 'function') throw new InputError('engine', "expected pivotEngine of 'xlcalc/pivot'")
  const o = options === undefined ? {} : object(options, 'options', 'an options object')
  onlyOptions(o, 'options', ['sortThreads'])
  const threads = o.sortThreads === undefined ? 2 : o.sortThreads
  if (typeof threads !== 'number' || !Number.isSafeInteger(threads) || threads < 1) throw new InputError('options.sortThreads', 'expected a positive integer')
  if (pivotRefresh() === parts.refresh && pivotSortThreads() === threads) return
  setPivotSortThreads(threads)
  setPivotRefresh(parts.refresh)
  setPivotCheck(parts.check)
  setPivotEntry(parts.entry)
  setPivotCacheEdit(parts.edit)
  for (const f of listeners) f()
}

/** Hears the engine being loaded; the answer stops it. */
export function onPivotLoaded(f: () => void): () => void {
  listeners.add(f)
  return () => listeners.delete(f)
}
