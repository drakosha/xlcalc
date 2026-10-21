/** Input the contract types but did not check: the values of external functions and a cell's meta, refused at the boundary as InputError. */
import { describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, InputError, type BookInput } from '../src/index.js'

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

  it("a cell's meta that is cyclic or holds a BigInt is refused, in compute, a session's edits and on the same thread", async () => {
    const cyclic: Record<string, unknown> = { row: 1 }
    cyclic.self = { back: cyclic }
    for (const [meta, why] of [[cyclic, /cycle/], [{ id: 10n }, /BigInt/], [[1, { n: 2n }], /BigInt/]] as const) {
      const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1, meta: meta as never } } }] }
      expect(refusal(() => compute(b, { clock, execution }))).toEqual({ path: expect.stringMatching(/^book\.sheets\[0\]\.cells\["A1"\]\.meta/), message: expect.stringMatching(why) })
      expect(() => createSession(b, { clock, execution })).toThrow(InputError)
      const s = createSession(book, { clock, execution })
      expect(() => s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 1, meta: meta as never } }])).toThrow(InputError)
      s.close()
      const calc = createInlineCalc()
      await expect(calc.compute(b, { clock, execution })).rejects.toThrow(InputError)
      calc.close()
    }
    // JSON meta, shared but not cyclic, stays as it was
    const shared = { tag: 'x' }
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { value: 1, meta: { a: shared, b: shared } }, A2: { value: 2, meta: 'm' } } }] }, { clock, execution })
    expect(r.sheets[0].cells.A1.meta).toEqual({ a: shared, b: shared })
  })
})
