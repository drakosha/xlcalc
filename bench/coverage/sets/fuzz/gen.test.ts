/**
 * The random formulas: what Excel must get to stay safe (size arguments small numeric literals, references
 * where it takes only references, arities it accepts) and determinism by seed.
 */
import { describe, expect, it } from 'vitest'
import { FUNCTIONS } from '../../../../src/functions.js'
import { parse, walk, type Ast } from '../../../../src/syntax.js'
import { engineFormula } from '../../../../test/excel-cases.js'
import { SIZE_LIMIT } from '../../signatures.js'
import { BATCHES, buildProbes } from '../fuzz.js'
import { fuzz, VOLATILE } from './gen.js'
import { isRefOnly, isSize } from './pools.js'

const probes = buildProbes()
const calls: Extract<Ast, { k: 'fn' }>[] = []
for (const p of probes) walk(parse(engineFormula(p.formula)), (n) => void (n.k === 'fn' && calls.push(n)))

const literal = (a: Ast): number | null => (a.k === 'num' ? a.v : a.k === 'un' && a.op === '-' && a.a.k === 'num' ? -a.a.v : null)

describe('random formulas', () => {
  it('as many as the batches ask, with unique ids', () => {
    expect(probes.length).toBe(BATCHES.reduce((s, b) => s + b.count, 0))
    expect(new Set(probes.map((p) => p.id)).size).toBe(probes.length)
  })

  it('a size argument is a small numeric literal', () => {
    const bad = calls.filter((c) => c.args.some((a, i) => isSize(c.name, i) && !(Math.abs(literal(a) ?? Infinity) <= SIZE_LIMIT)))
    expect(bad).toEqual([])
  })

  it('where Excel takes only a reference, a reference', () => {
    expect(calls.filter((c) => c.args.some((a, i) => isRefOnly(c.name, i) && a.k !== 'ref'))).toEqual([])
  })

  it('every call is whitelisted, not volatile, with an argument count the function takes', () => {
    const bad = calls.filter((c) => {
      const s = FUNCTIONS[c.name]
      const n = c.args.length
      return !s || VOLATILE.has(c.name) || n < s.min || n > s.max || (s.parity === 'odd' && n % 2 === 0) || (s.parity === 'even' && n % 2 === 1)
    })
    expect(bad.map((c) => c.name)).toEqual([])
    expect(calls.filter((c) => c.name === 'SORTBY' && c.args.length > 2 && c.args.length % 2 === 0)).toEqual([])
    // ROW() and COLUMN() answer the case's own row, which differs between Excel's parts and the engine's book
    expect(calls.filter((c) => (c.name === 'ROW' || c.name === 'COLUMN') && !c.args.length)).toEqual([])
  })

  it('the same seed gives the same formulas', () => {
    expect(fuzz(BATCHES[0])).toEqual(fuzz(BATCHES[0]))
  })
})
