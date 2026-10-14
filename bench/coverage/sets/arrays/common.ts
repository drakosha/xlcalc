/** Shared parts of the case sets: the `_xlfn.` prefixes of formulas and the case builder. */
import type { Val } from '../../formulajs.js'
import type { Probe } from '../../probes.js'
import { XLFN } from '../../signatures.js'

const NAMES = new RegExp(`(?<![\\w.])(${Object.keys(XLFN).sort((a, b) => b.length - a.length).map((n) => n.replace(/\./g, '\\.')).join('|')})\\(`, 'g')

/** A formula in the engine's spelling as Excel stores it: newer functions get their prefix (`XLFN`). */
export const xl = (formula: string) => formula.replace(NAMES, (_, name: string) => `${XLFN[name]}(`)

/** Adds cases of one group: ids `<group>#<n>`, `fn` the function checked (a group name for operators and constants). */
export type Add = (group: string, fn: string, ...formulas: string[]) => void

export function collector(): { out: Probe[]; add: Add } {
  const out: Probe[] = []
  const count = new Map<string, number>()
  const add: Add = (group, fn, ...formulas) => {
    for (const f of formulas) {
      const n = (count.get(group) ?? 0) + 1
      count.set(group, n)
      out.push({ id: `${group}#${n}`, group, fn, formula: xl(f) })
    }
  }
  return { out, add }
}

/** A collector of cases by group alone: `add(group, formula)` puts a case with no function and the formula as written (`<group>#<n>`). */
export function byGroup(): { out: Probe[]; add: (group: string, formula: string) => void } {
  const out: Probe[] = []
  const count = new Map<string, number>()
  const add = (group: string, formula: string) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}#${n}`, group, formula })
  }
  return { out, add }
}

/** A collector of cases by function: `add(fn, formula)` puts a case in the group of its function (`<fn>#<n>`); `prefixed` false keeps the formulas as written. */
export function byFunction(prefixed = true): { out: Probe[]; add: (fn: string, formula: string) => void } {
  const out: Probe[] = []
  const count = new Map<string, number>()
  const add = (fn: string, formula: string) => {
    const n = (count.get(fn) ?? 0) + 1
    count.set(fn, n)
    out.push({ id: `${fn}#${n}`, group: fn, fn, formula: prefixed ? xl(formula) : formula })
  }
  return { out, add }
}

/** The probes of `[function, formula]` cases, each in the group of its function (`<function>#<n>`). */
export function probesOf(cases: readonly (readonly [string, string, ...unknown[]])[]): Probe[] {
  const { out, add } = collector()
  for (const [fn, f] of cases) add(fn, fn, f)
  return out
}

/**
 * A held-out set of `[function, formula, answer predicted]` cases: its probes (`probesOf`) and the predicted answers by
 * id, a case without a prediction left out.
 */
export const heldCases = (cases: readonly (readonly [string, string, Val?])[]): HeldSet => heldGroups(cases.map(([fn, f, v]) => [fn, fn, f, v]))

/** A held-out set's probes and the answers predicted by case id. */
export interface HeldSet {
  buildProbes: () => Probe[]
  PREDICTED: Readonly<Record<string, Val>>
}

/** A held-out set of `[group, function, formula, answer predicted]` cases, as `heldCases`. */
export function heldGroups(cases: readonly (readonly [string, string, string, Val?])[]): HeldSet {
  const buildProbes = () => {
    const { out, add } = collector()
    for (const [group, fn, f] of cases) add(group, fn, f)
    return out
  }
  return { buildProbes, PREDICTED: Object.fromEntries(buildProbes().flatMap((p, i) => (cases[i][3] ? [[p.id, cases[i][3]]] : []))) }
}
