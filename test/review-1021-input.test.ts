/** Input the contract types but did not check: the values of external functions, refused at the boundary as InputError. */
import { describe, expect, it } from 'vitest'
import { compute, createSession, InputError, type BookInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const BALANCE = 'BALANCE("x")'
/** The path and message of the InputError `f` throws. */
function refusal(f: () => unknown): { path: string; message: string } | string {
  try {
    f()
    return 'no refusal'
  } catch (e) {
    return e instanceof InputError ? { path: e.path, message: e.message } : String(e)
  }
}

const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: BALANCE }, A2: { formula: 'A1+1' } } }] }

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  const withValue = (v: unknown) => ({ clock, execution, externals: { functions: ['BALANCE'], values: { [BALANCE]: v } } }) as never

  it.each([
    ['an unknown error code', { error: 'BOGUS' }],
    ['a prototype key as an error code', { error: 'toString' }],
    ['a date input', { date: '2026-01-01' }],
    ['an array', [1, 2]],
    ['a Date', new Date(0)],
    ['undefined', undefined],
    ['a function', () => 1],
    ['a BigInt', 5n],
  ])('a value of an external function: %s is refused', (_, v) => {
    const path = `opts.externals.values[${JSON.stringify(BALANCE)}]`
    expect(refusal(() => compute(book, withValue(v)))).toEqual({ path, message: expect.stringMatching(/expected a number, text, a boolean, null or \{ error: code \}/) })
    expect(() => createSession(book, withValue(v))).toThrow(InputError)
  })

  it('a value of an external function: a CellValue computes, a non-finite number still gives #NUM!', () => {
    const a2 = (v: unknown) => compute(book, withValue(v)).sheets[0].cells.A2.value
    expect([a2(3), a2('7'), a2(true), a2(null), a2({ error: 'DIV0' })]).toEqual([4, 8, 2, 1, { error: 'DIV0' }])
    expect([a2(NaN), a2(Infinity)]).toEqual([{ error: 'NUM' }, { error: 'NUM' }])
  })
})
