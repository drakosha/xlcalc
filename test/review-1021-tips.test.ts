/**
 * The helpers of a grid's formula editor, from the review of 21.10: a function's tip and description read its name as
 * the engine does (`wordKey`: `gauß` is no GAUSS, as Excel computes #NAME? for it, the sets `name-case-held`). With
 * closures and eval for what the engine computes.
 */
import { describe, expect, it } from 'vitest'
import { callAt, compute, functionSignature, validateFormula } from '../src/index.js'
import { functionDescription } from '../src/function-text/index.js'
import { ruRU } from '../src/locales/ru-RU.js'

describe('the tip of a name the engine does not know', () => {
  const names = ['gauß', 'ﬁnd', 'addreß', 'ſum', 'ıf']

  it.each(['closures', 'eval'] as const)('is none where the engine computes #NAME? (%s)', (execution) => {
    const cells = Object.fromEntries(names.map((n, i) => [`A${i + 1}`, { formula: `${n}(1)` }]))
    const r = compute({ sheets: [{ name: 'S', cells }] }, { execution })
    for (const [i, n] of names.entries()) {
      expect(r.sheets[0].cells[`A${i + 1}`].value, n).toEqual({ error: 'NAME' })
      expect(validateFormula(`${n}(1)`)).toMatchObject({ ok: false, code: 'unknown_function' })
      expect([functionSignature(n), functionSignature(n, ruRU), functionDescription(n)], n).toEqual([null, null, null])
      expect(callAt(`=${n}(1`)?.fn, n).not.toMatch(/^(GAUSS|FIND|ADDRESS|SUM|IF)$/)
    }
  })

  it('is the function of a name in any case, English or the locale\'s', () => {
    expect([functionSignature('sum')?.fn, functionSignature('суммесли', ruRU)?.fn, callAt('=Sum(1')?.fn, callAt('=суммесли(1', undefined, { locale: ruRU })?.fn]).toEqual(['SUM', 'SUMIF', 'SUM', 'SUMIF'])
    expect(functionDescription('vlookup')).toBe(functionDescription('VLOOKUP'))
  })
})
