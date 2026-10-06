/**
 * Random formulas over the whitelist, deterministic by seed: a call of a random function with as many
 * arguments as the engine's spec allows, each from the function's typical call or drawn by its kind
 * (`pools.ts`), a nested call or an operator expression, up to a depth of nested calls; some formulas are
 * read through a function that takes an array, some join two calls with an operator.
 */
import { readFileSync } from 'node:fs'
import { DEFS, FUNCTION_NAMES, FUNCTIONS } from '../../../../src/functions.js'
import { mulberry32 } from '../../../../src/synth.js'
import { renderArg, type Case } from '../../cases.js'
import EXCEL_FUNCTIONS from '../../excel-functions.json' with { type: 'json' }
import { BASE, splitArgs } from '../arrays/lift.js'
import { FUZZ_BASE, isRefOnly, isSize, KIND_CATEGORY, KIND_MIX, POOLS, type Mix } from './pools.js'

/** Volatile functions: their answers change between Excel and the engine. */
export const VOLATILE: ReadonlySet<string> = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY'])

const NAMES = FUNCTION_NAMES.filter((f) => !VOLATILE.has(f))
const CATEGORY = new Map(EXCEL_FUNCTIONS.map((x) => [x.name, x.category]))
const ARRAY_FNS = NAMES.filter((f) => DEFS[f].array)
const OPS = ['+', '-', '*', '/', '^', '&', '=', '<>', '<', '>', '<=', '>=']
/** Readers of a result that may be an array: all its values as text, a sum, its size, its first value. */
const WRAPPERS: readonly ((x: string) => string)[] = [(x) => `TEXTJOIN(";",FALSE,${x})`, (x) => `SUM(${x})`, (x) => `ROWS(${x})`, (x) => `COLUMNS(${x})`, (x) => `INDEX(${x},1,1)`]
/** Kinds whose argument takes ranges and arrays whole: a nested call there is often a function with an array answer. */
const WHOLE = 'NAYZVGMRDp'

/** Arguments of every function's typical call (the measurement's first typical case, else the array set's or the fuzz's base). */
function typicalArgs(): Map<string, string[]> {
  const cases = JSON.parse(readFileSync(new URL('../../../../docs/coverage-formulajs/cases.json', import.meta.url), 'utf8')) as { cases: Case[] }
  const out = new Map<string, string[]>()
  for (const c of cases.cases) if (c.kind === 'typical' && !out.has(c.fn)) out.set(c.fn, c.args.map(renderArg))
  for (const [fn, call] of Object.entries({ ...BASE, ...FUZZ_BASE })) if (!out.has(fn)) out.set(fn, splitArgs(call.slice(fn.length + 1, -1)))
  return out
}

const TYPICAL = typicalArgs()

export interface Batch {
  seed: number
  count: number
  /** Most levels of calls in a formula (1 is a call with plain arguments). */
  depth: number
}

export interface Fuzzed {
  /** The function of the outer call. */
  fn: string
  formula: string
}

/** Each formula draws from its own stream (the batch's seed and its number), so a change to the generator changes only the formulas it touches. */
export function fuzz(b: Batch): Fuzzed[] {
  return Array.from({ length: b.count }, (_, k) => new Gen(b.seed ^ Math.imul(k + 1, 0x9e3779b1)).formula(b.depth))
}

class Gen {
  private readonly rnd: () => number

  constructor(seed: number) {
    this.rnd = mulberry32(seed)
  }

  private chance = (p: number) => this.rnd() < p
  private one = <T>(xs: readonly T[]): T => xs[Math.floor(this.rnd() * xs.length)]

  private mix(m: Mix): string {
    let r = this.rnd() * m.reduce((s, [, w]) => s + w, 0)
    for (const [pool, w] of m) if ((r -= w) < 0) return this.one(POOLS[pool])
    return this.one(POOLS[m[0][0]])
  }

  formula(maxDepth: number): Fuzzed {
    const depth = 1 + Math.floor(this.rnd() * maxDepth)
    const fn = this.one(NAMES)
    const call = this.call(fn, depth)
    if (this.chance(0.1)) return { fn, formula: `${call}${this.one(OPS)}${this.call(this.one(NAMES), depth)}` }
    return { fn, formula: this.chance(0.25) ? this.one(WRAPPERS)(call) : call }
  }

  /** A call of `fn` whose arguments nest calls up to `depth - 1` levels. */
  private call(fn: string, depth: number): string {
    const base = TYPICAL.get(fn)
    const n = base && this.chance(0.6) ? base.length : this.count(fn)
    const kinds = DEFS[fn].args.split(' ').filter(Boolean).map((k) => k[0])
    const rep = DEFS[fn].rep ?? 0
    const kindAt = (i: number) => (i < kinds.length ? kinds[i] : kinds[kinds.length - rep + ((i - kinds.length + rep) % rep)])
    const args: string[] = []
    for (let i = 0; i < n; i++) {
      if (isSize(fn, i)) args.push(this.one(POOLS.size))
      else if (isRefOnly(fn, i)) args.push(this.one([...POOLS.range, ...POOLS.cellNum]))
      else if (base && i < base.length && this.chance(0.5)) args.push(base[i])
      else args.push(this.arg(kindAt(i), depth - 1))
    }
    return `${fn}(${args.join(',')})`
  }

  /** A number of arguments within the spec: the required ones and up to three more, of the right parity. */
  private count(fn: string): number {
    const { min, max, parity } = FUNCTIONS[fn]
    // ROW() and COLUMN() answer the formula's own cell, which differs between Excel's parts and the engine's book
    const lo = Math.max(min, fn === 'ROW' || fn === 'COLUMN' ? 1 : 0)
    let n = lo + Math.floor(this.rnd() * (Math.min(3, max - lo) + 1))
    if ((parity === 'odd' && n % 2 === 0) || (parity === 'even' && n % 2 === 1)) n = n + 1 <= max ? n + 1 : n - 1
    // Excel loads SORTBY only with each further by-array followed by its order: 2 arguments or an odd number
    if (fn === 'SORTBY' && n > 2 && n % 2 === 0) n++
    return n
  }

  /** An argument of a kind: a nested call or an operator expression while `depth` allows, else a value of its mix. */
  private arg(kind: string, depth: number): string {
    if (depth > 0 && this.chance(0.35)) return this.chance(0.25) ? this.operation(kind, depth) : this.call(this.nested(kind), depth)
    return this.mix(KIND_MIX[kind])
  }

  /** Two operands and an operator; the first operand of the argument's kind, now and then negated or a percentage. */
  private operation(kind: string, depth: number): string {
    const operand = (m: Mix) => (this.chance(0.3) ? this.call(this.one(NAMES), depth) : this.mix(m))
    let a = operand(KIND_MIX[kind])
    if (this.chance(0.1)) a = `-${a}`
    if (this.chance(0.05)) a = `${a}%`
    return `${a}${this.one(OPS)}${operand(KIND_MIX.a)}`
  }

  /** The function of a nested call: one of the categories that suit the kind, an array function where ranges go, or any. */
  private nested(kind: string): string {
    const cats = KIND_CATEGORY[kind]
    if (cats && this.chance(0.6)) return this.one(NAMES.filter((f) => cats.includes(CATEGORY.get(f) ?? '')))
    if (WHOLE.includes(kind) && this.chance(0.5)) return this.one(ARRAY_FNS)
    return this.one(NAMES)
  }
}
