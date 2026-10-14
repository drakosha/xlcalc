/**
 * What the tests of structure operations in sessions share (`merge-cells.test.ts`, `fill-range.test.ts`): the cases of
 * struct sets with the book `editStructure` gives after their operations, the books' hashes, the path of an input's
 * `InputError`, and the check that the asynchronous kinds of session give the synchronous one's books and undo them.
 */
import { afterAll, expect, it } from 'vitest'
import { compute, createInlineCalc, editStructure, InputError, type BookInput, type Calc, type StructureOp } from '../src/index.js'
import { takenBook } from '../bench/coverage/stored.js'
import type { StructSet } from '../bench/coverage/struct/cases.js'
import { pool } from './fixtures/pool.js'

export const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** The input hash and the values' hash of a book. */
export const hashes = (book: BookInput) => {
  const r = compute(book, { clock })
  return [r.inputHash, r.hash]
}

/** A case of a struct set: its book, its operations in order (`before` first), the book `editStructure` gives after them (null when it refuses). */
export interface OpCase {
  id: string
  book: BookInput
  ops: StructureOp[]
  after: BookInput | null
}

/** The cases of the struct sets `names` (`bench/coverage/struct/<name>.ts`). */
export async function opCases(names: string[]): Promise<OpCase[]> {
  const sets = await Promise.all(names.map(async (n) => ((await import(`../bench/coverage/struct/${n}.ts`)) as StructSet).buildStructCases()))
  return sets.flatMap((set) =>
    set.cases.map((c) => {
      const book = takenBook(set.books[c.book])
      const ops = [...(c.before ? [c.before] : []), c.op]
      let after: BookInput | null = book
      for (const op of ops) {
        const r: ReturnType<typeof editStructure> | null = after && editStructure(after, op)
        after = r?.ok ? r.book : null
      }
      return { id: c.id, book, ops, after }
    }),
  )
}

/** The path of the `InputError` the operation throws on the book; another throw's text, or undefined when it throws none. */
export function inputErrorPath(book: BookInput, op: unknown): string | undefined {
  try {
    editStructure(book, op as StructureOp)
  } catch (x) {
    return x instanceof InputError ? x.path : `not an InputError: ${String(x)}`
  }
  return undefined
}

/** Tests that a session in the same thread and one in a thread give the books of `cases` and undo them back. */
export function inOtherSessions(cases: OpCase[]): void {
  const calcs: Calc[] = []
  afterAll(() => Promise.all(calcs.map((c) => c.close())))

  async function same(calc: Calc) {
    calcs.push(calc)
    for (const c of cases) {
      const s = await calc.createSession(c.book, { clock })
      for (const op of c.ops) await s.editStructure(op)
      expect(hashes(await s.getBook()), c.id).toEqual(hashes(c.after!))
      for (const _ of c.ops) await s.undo()
      expect(hashes(await s.getBook()), c.id).toEqual(hashes(c.book))
    }
  }

  it('in the same thread', () => same(createInlineCalc({ sliceMs: 1 })))
  it('in a thread', async () => same(await pool({ threads: 1 })), 60_000)
}
