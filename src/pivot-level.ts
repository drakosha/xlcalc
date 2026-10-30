/**
 * The PivotTable engine of `xlcalc/pivot` that `registerPivot` loaded: the refresh a session's PivotTable operations
 * take (`pivot/hook.ts`). A pool's threads load the same engine (`worker/thread.ts`).
 */
import type { PivotEngine } from './contract.js'
import { InputError } from './input-error.js'
import { pivotRefresh, setPivotRefresh, type PivotRefresh } from './pivot/hook.js'

/** The engine's parts besides its kind, which the subpath gives and the contract does not show. */
export interface PivotEngineParts extends PivotEngine {
  refresh: PivotRefresh
}

const listeners = new Set<() => void>()

/** Whether the engine is loaded. */
export const pivotLoaded = (): boolean => pivotRefresh() !== null

/**
 * Loads the PivotTable engine of `xlcalc/pivot` for every session from now on, a pool's threads too: `import {
 * pivotEngine } from 'xlcalc/pivot'`, then `registerPivot(pivotEngine)`.
 */
export function registerPivot(e: PivotEngine): void {
  const parts = e as PivotEngineParts
  if (typeof e !== 'object' || e === null || e.kind !== 'xlcalc-pivot' || typeof parts.refresh !== 'function') throw new InputError('engine', "expected pivotEngine of 'xlcalc/pivot'")
  if (pivotRefresh() === parts.refresh) return
  setPivotRefresh(parts.refresh)
  for (const f of listeners) f()
}

/** Hears the engine being loaded; the answer stops it. */
export function onPivotLoaded(f: () => void): () => void {
  listeners.add(f)
  return () => listeners.delete(f)
}
