/**
 * The groups of `xlcalc/functions` loaded into the engine's table (`functions.ts`), and what it knows of those not
 * loaded: the light entry gives their functions' arguments (`useManifest`), so that a call of one is `#NAME?` naming its
 * group, its arguments counted as the loaded functions' are. A session built before a group was loaded builds its state
 * again at its next call (`session.ts`, `ready`); a pool's threads load the groups the host has (`worker/pool.ts`).
 */
import { aritiesOf, useGroupArities, type ArityDef } from '../arity.js'
import type { FunctionGroup, FunctionGroupName } from '../contract.js'
import { addFunctions } from '../functions.js'
import { InputError } from '../input-error.js'
import type { LoadedGroup } from './group.js'
import { GROUP_NAMES, grouped, isLoaded, markLoaded, useGroupOf } from './state.js'

/** The declared arguments of each group's functions (`manifest.ts`). */
export type Manifest = Readonly<Record<FunctionGroupName, Readonly<Record<string, ArityDef>>>>

/** The manifest of the groups, given by the light entry and a pool's thread: what a call of a group not loaded is checked by. */
export function useManifest(m: Manifest): void {
  if (grouped()) return
  useGroupOf(new Map(GROUP_NAMES.flatMap((g) => Object.keys(m[g]).map((f) => [f, g] as const))))
  useGroupArities(aritiesOf(Object.assign({}, ...GROUP_NAMES.map((g) => m[g]))))
}

/** Loads groups the package made (unchecked); a group loaded already is passed over. */
export function loadGroups(groups: readonly FunctionGroup[]): void {
  const added = groups.filter((g, i) => !isLoaded(g.name) && groups.findIndex((h) => h.name === g.name) === i)
  if (!added.length) return
  for (const g of added) addFunctions(g.name, (g as LoadedGroup).defs)
  markLoaded(added.map((g) => g.name))
}

const isGroup = (g: unknown): g is LoadedGroup =>
  typeof g === 'object' && g !== null && (g as LoadedGroup).kind === 'xlcalc-functions' && GROUP_NAMES.includes((g as LoadedGroup).name) && typeof (g as LoadedGroup).defs === 'object'

/** Checks the groups a host gives and loads them (`core.ts`). */
export function registerFunctions(...groups: FunctionGroup[]): void {
  groups.forEach((g, i) => {
    if (!isGroup(g)) throw new InputError(`groups[${i}]`, "expected a group of 'xlcalc/functions' (engineering, finance, lambda, statistics, xml)")
  })
  loadGroups(groups)
}
