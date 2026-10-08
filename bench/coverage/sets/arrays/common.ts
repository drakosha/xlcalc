/** Shared parts of the case sets: the `_xlfn.` prefixes of formulas and the case builder. */
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
