/**
 * Which groups of `xlcalc/functions` are loaded, without the engine: what a pool tells its threads (`worker/pool.ts`)
 * and what the engine's check asks of a function (`registry.ts` loads the groups' definitions).
 */
import type { FunctionGroupName } from '../contract.js'

export const GROUP_NAMES: readonly FunctionGroupName[] = ['engineering', 'finance', 'lambda', 'statistics', 'xml']

const loaded = new Set<FunctionGroupName>()
const listeners = new Set<() => void>()
let version = 0
/** The group of each function of a group, once the light entry gave the manifest (`registry.ts`, `useManifest`). */
let groupOf: ReadonlyMap<string, FunctionGroupName> | undefined

export const isLoaded = (g: FunctionGroupName): boolean => loaded.has(g)

/** Marks groups loaded and tells the listeners. */
export function markLoaded(groups: readonly FunctionGroupName[]): void {
  for (const g of groups) loaded.add(g)
  version++
  for (const f of listeners) f()
}

/** Whether the group of each function of a group was given (`useGroupOf`). */
export const grouped = (): boolean => groupOf !== undefined

/** The group of each function of a group. */
export function useGroupOf(m: ReadonlyMap<string, FunctionGroupName>): void {
  groupOf = m
}

/** The group of a function whose group is not loaded, where the manifest is given: a call of it is refused (`function_not_loaded`). */
export function notLoaded(name: string): FunctionGroupName | undefined {
  const g = groupOf?.get(name)
  return g && !loaded.has(g) ? g : undefined
}

/** The groups a pool's threads load: those loaded where the light entry is, else all of them, as the package root has. */
export const groupsWanted = (): FunctionGroupName[] => (groupOf ? GROUP_NAMES.filter((g) => loaded.has(g)) : [...GROUP_NAMES])

/** How many times groups were loaded: a session's state built before then is built again (`session.ts`). */
export const functionsVersion = (): number => version

/** Hears groups being loaded; the answer stops it. */
export function onFunctions(f: () => void): () => void {
  listeners.add(f)
  return () => listeners.delete(f)
}
