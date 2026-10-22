/**
 * Reproductions from the 21.10 review, refusals: a call that cannot be made fails with the library's own error class in
 * every kind of session. A closed session refuses with `CalcError` code `closed`, `setClock` on a live session with an
 * `InputError` on `clock`. `editStructure` refuses an operation named as a member of `Object.prototype` and `addTable` a
 * column without a name with `InputError` and its path. In a synchronous session, in the same thread and in a thread,
 * with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { CalcError, InputError, editStructure, type BookInput, type StructureOp } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

afterAll(closeCalcs)

const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { formula: 'A1*2' } } }] }

/** The error a call failed with. */
const failure = async (p: () => Promise<unknown>) => {
  try {
    await p()
  } catch (e) {
    return e
  }
  throw new Error('the call did not fail')
}

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)

    it("a closed session refuses its calls with CalcError code 'closed'", async () => {
      const s = await open(book, { clock: CLOCK, execution })
      await s.close()
      for (const call of [() => s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }]), () => s.getCell('S', 'A2')]) {
        const e = await failure(call)
        expect(e).toBeInstanceOf(CalcError)
        expect(e).toMatchObject({ code: 'closed', message: 'the session is closed' })
      }
    })

    it('setClock on a live session throws InputError on clock, and the session goes on', async () => {
      const s = await open(book, { live: true, execution })
      const e = await failure(() => s.setClock(CLOCK))
      expect(e).toBeInstanceOf(InputError)
      expect(e).toMatchObject({ path: 'clock', message: expect.stringContaining('a live session takes its clock from the system') })
      expect((await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])).changes.length).toBe(2)
      expect((await s.getCell('S', 'A2'))?.value).toBe(6)
      await s.close()
    })

    it.each(['constructor', 'toString', 'hasOwnProperty', '__proto__'])('editStructure refuses op %s with InputError on op.op', async (name) => {
      const op = { op: name } as unknown as StructureOp
      expect(() => editStructure(book, op)).toThrow(InputError)
      expect(() => editStructure(book, op)).toThrow(/^op\.op: expected one of insertRows/)
      const s = await open(book, { clock: CLOCK, execution })
      const e = await failure(() => s.editStructure(op))
      expect(e).toBeInstanceOf(InputError)
      expect(e).toMatchObject({ path: 'op.op' })
      await s.close()
    })

    it.each([[{}], [null], [5], [{ nam: 'b' }], [{ name: 5 }]])('addTable with a column %j over an empty header cell refuses with InputError', async (column) => {
      const op = { op: 'addTable', sheet: 'S', table: { name: 'T', range: 'C1:D2', columns: [{ name: 'a' }, column] } } as unknown as StructureOp
      expect(() => editStructure(book, op)).toThrow(/^op\.sheets\[0\]\.tables\[0\]\.columns\[1\]/)
      const s = await open(book, { clock: CLOCK, execution })
      const e = await failure(() => s.editStructure(op))
      expect(e).toBeInstanceOf(InputError)
      expect((e as InputError).path).toMatch(/^op\.sheets\[0\]\.tables\[0\]\.columns\[1\]/)
      expect(Object.keys((await s.getBook()).sheets[0].cells)).toEqual(['A1', 'A2'])
      await s.close()
    })

    it('addTable over a header cell with a formula writes the column name in its place', async () => {
      const s = await open({ sheets: [{ name: 'S', cells: { D1: { value: 1, formula: '1' } } }] }, { clock: CLOCK, execution })
      await s.editStructure({ op: 'addTable', sheet: 'S', table: { name: 'T', range: 'C1:D2', columns: [{ name: 'a' }, { name: 'b' }] } })
      expect((await s.getBook()).sheets[0].cells).toEqual({ C1: { value: 'a' }, D1: { value: 'b' } })
      await s.close()
    })
  })
})
