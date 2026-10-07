/**
 * The largest number of arguments of every whitelisted function with repeated arguments is Excel's: the `arity`
 * set has a call at each limit with Excel's answer (`test/excel/cases/arity.json`), and Excel refused a book with
 * one more argument, or two where they come in pairs (`bench/coverage/sets/arity.ts`, `LIMITS`).
 */
import { describe, expect, it } from 'vitest'
import { call, limit, SHAPES } from '../bench/coverage/sets/arity.js'
import { DEFS, FUNCTIONS } from '../src/functions.js'
import { validateFormula } from '../src/index.js'

describe('argument counts as Excel loads them', () => {
  it('the set covers every function with repeated arguments', () => {
    expect(Object.keys(DEFS).filter((f) => DEFS[f].rep && !(f in SHAPES))).toEqual([])
  })

  for (const fn of Object.keys(SHAPES))
    it(`${fn}: at most ${limit(fn)}`, () => {
      expect(FUNCTIONS[fn].max).toBe(limit(fn))
      expect(validateFormula(call(fn, limit(fn))!).ok).toBe(true)
      const over = call(fn, limit(fn) + 1) ?? call(fn, limit(fn) + 2)
      if (over) expect(validateFormula(over)).toMatchObject({ ok: false, code: 'syntax' })
    })
})
