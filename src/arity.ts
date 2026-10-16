/**
 * The arguments a function takes, as the check of a formula needs them, without its implementation: how many, by its
 * declared kinds and the group that repeats (`FnDef.args`, `rep`, `max`), and which it lets a formula leave out
 * (`notOmitted`). The engine's spec of a function (`args.ts`, `specOf`) takes them from here, and so does the check of
 * typed text without the engine (`function-arity.ts`).
 */
import type { Counts } from './locale/messages.js'
import { FUNCTION_ARITY } from './function-arity.js'

/** Argument positions: those `at`, and from `from` on every one, or the odd or the even ones. */
export interface Positions {
  at?: readonly number[]
  from?: number
  odd?: boolean
  even?: boolean
}

export function holds(p: Positions | undefined, i: number): boolean {
  if (!p) return false
  if (p.at?.includes(i)) return true
  if (p.from === undefined || i < p.from) return false
  return p.odd ? i % 2 === 1 : p.even ? i % 2 === 0 : true
}

/** A function's declared arguments (`FnDef`'s fields of the same names). */
export interface ArityDef {
  args: string
  rep?: number
  max?: number
  notOmitted?: Positions
}

/** Excel's limit on the number of arguments. */
const MAX_ARGS = 255

/** How many arguments a function takes, and which it lets a formula leave out. */
export interface Arity {
  min: number
  max: number
  /** Whether a call with `k` arguments is one Excel loads: within min…max, repeated pairs complete (SUMIFS 3, 5, …; SORTBY 2, 3, 5, …). */
  takes: (k: number) => boolean
  /** The counts it takes, for a message (SUMIFS 3 to 255, odd). */
  counts: Counts
  /** Whether Excel loads a formula that leaves out the argument at `i`. */
  omits: (i: number) => boolean
}

/** The kinds of `def.args` with their `?`, and the arity they give. */
export function arityOf(def: ArityDef): Arity {
  const kinds = def.args.split(' ').filter(Boolean)
  const rep = def.rep ?? 0
  const fixed = kinds.length - rep
  const min = kinds.filter((k) => !k.endsWith('?')).length
  let max = rep ? (def.max ?? MAX_ARGS) : kinds.length
  // past the declared kinds, repeated groups come whole (an optional one inside the first group may be left out)
  const whole = (k: number) => k <= kinds.length || (k - fixed) % rep === 0
  if (rep && !whole(max)) max--
  const counts: Counts = rep === 2 ? { min, max, parity: fixed % 2 ? 'odd' : 'even', or: min < kinds.length } : { min, max }
  return { min, max, takes: (k) => k >= min && k <= max && whole(k), counts, omits: (i) => !holds(def.notOmitted, i) }
}

const ARITIES = new Map<string, Arity>()

/** The arity of a whitelisted function by its English name, from its declared arguments; undefined for another name. */
export function arityFor(name: string): Arity | undefined {
  let a = ARITIES.get(name)
  if (!a && name in FUNCTION_ARITY) ARITIES.set(name, (a = arityOf(FUNCTION_ARITY[name])))
  return a
}

/**
 * Whitelisted functions that are syntax rather than calls: LET's names are bound by the parser
 * (`syntax/parse.ts`) and its values computed by the walk (`exec/walk.ts`).
 */
export const SYNTAX: readonly string[] = ['LET']

/** A whitelisted function, a call or syntax. */
export const isFunction = (name: string): boolean => name in FUNCTION_ARITY || SYNTAX.includes(name)
