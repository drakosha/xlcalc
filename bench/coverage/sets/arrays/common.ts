/** Shared parts of the array set: the `_xlfn.` prefixes and the case builder. */
import type { Probe } from '../../probes.js'
import { XLFN } from '../../signatures.js'

/** Functions newer than Excel 2007, stored with a prefix (the measurement's list and the newer ones). */
const PREFIX: Record<string, string> = {
  ...XLFN,
  ...Object.fromEntries(
    ['ARRAYTOTEXT', 'CHOOSECOLS', 'CHOOSEROWS', 'FORECAST.LINEAR', 'LET', 'PERCENTOF', 'RANDARRAY', 'SEQUENCE', 'SORTBY', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTSPLIT', 'TOCOL', 'TOROW', 'VALUETOTEXT', 'WRAPCOLS', 'WRAPROWS', 'XLOOKUP', 'XMATCH'].map((n) => [n, `_xlfn.${n}`]),
  ),
  FILTER: '_xlfn._xlws.FILTER',
}

const NAMES = new RegExp(`(?<![\\w.])(${Object.keys(PREFIX).sort((a, b) => b.length - a.length).map((n) => n.replace(/\./g, '\\.')).join('|')})\\(`, 'g')

/** A formula in the engine's spelling as Excel stores it: newer functions get their prefix. */
export const xl = (formula: string) => formula.replace(NAMES, (_, name: string) => `${PREFIX[name]}(`)

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
