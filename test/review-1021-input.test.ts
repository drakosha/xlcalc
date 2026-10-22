/**
 * Input the contract types but did not check: the values of external functions, a cell's meta and a value that is not a
 * finite number, refused at the boundary as InputError.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, InputError, type BookInput, type CellInput } from '../src/index.js'
import { ASYNC_CALCS, closeCalcs, KINDS, openerFor } from './session-kinds.js'

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

describe.each(['closures', 'eval'] as const)('a value that is not a finite number is refused (fix-u373), %s', (execution) => {
  afterAll(closeCalcs)
  const cell = (v: number): CellInput => ({ value: v })
  const at = (v: number): BookInput => ({ sheets: [{ name: 'S', cells: { A1: cell(v), A2: { formula: 'A1+1' } } }] })
  const BAD = [NaN, Infinity, -Infinity]
  const why = { path: 'book.sheets[0].cells["A1"].value', message: expect.stringMatching(/expected a finite number/) }

  it('compute and createSession, for each value', () => {
    for (const v of BAD) {
      expect(refusal(() => compute(at(v), { clock, execution })), String(v)).toEqual(why)
      expect(refusal(() => createSession(at(v), { clock, execution })), String(v)).toEqual(why)
    }
    expect(compute(at(-0), { clock, execution }).sheets[0].cells.A2.value).toBe(1)
  })

  it.each(ASYNC_CALCS)('compute and createSession %s', async (_, make) => {
    const calc = await make()
    for (const v of BAD) {
      await expect(calc.compute(at(v), { clock, execution })).rejects.toMatchObject(why)
      await expect(calc.createSession(at(v), { clock, execution })).rejects.toMatchObject(why)
    }
    await calc.close()
  })

  describe.each(KINDS)('the edits of %s', (kind) => {
    const open = openerFor(kind)
    it('setCells, replaceCells and applyChanges refuse it and change nothing', async () => {
      const s = await open(at(5), { clock, execution })
      const before = (await s.getResult()).hash
      for (const v of BAD) {
        const edit = [{ sheet: 'S', cell: 'A1', input: cell(v) }]
        await expect(s.setCells(edit), String(v)).rejects.toMatchObject({ path: 'edits[0].input.value' })
        await expect(s.replaceCells(edit), String(v)).rejects.toMatchObject({ path: 'edits[0].input.value' })
        await expect(s.applyChanges([{ op: 'setCells', edits: edit }]), String(v)).rejects.toMatchObject({ path: 'operations[0].edits[0].input.value' })
        await expect(s.applyChanges([{ op: 'replaceCells', edits: edit }]), String(v)).rejects.toMatchObject({ path: 'operations[0].edits[0].input.value' })
        await expect(s.validateEntry('S', 'A1', cell(v)), String(v)).rejects.toMatchObject({ path: 'input.value' })
      }
      expect((await s.getResult()).hash).toBe(before)
      await s.close()
    })
  })
})
