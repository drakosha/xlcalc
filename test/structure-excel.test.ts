/**
 * Acceptance of `restructure` against Excel en-US (`docs/excel-struct/`, see `test/structure-excel.ts`):
 * each stored case is the generator's current case, and the engine's book after the operation shows what
 * Excel showed, or Excel refused and so does the engine. Known differences would be listed in
 * STRUCT_KNOWN with their lines; there are none.
 */
import { describe, expect, it } from 'vitest'
import type { StructCases } from '../bench/coverage/struct/cases.js'
import { restructure } from '../src/index.js'
import { differences, STRUCT_SETS } from './structure-excel.js'

/** `<set>:<case id>` → the difference lines pinned for that case, each with its reason in a comment. */
const STRUCT_KNOWN: Readonly<Record<string, string[]>> = {}

describe('structure operations against Excel en-US', () => {
  it('every set has its cases answered', () => {
    expect(STRUCT_SETS.map((s) => s.name)).toEqual(['checks', 'edge', 'move', 'moves', 'rows', 'sheets', 'styles'])
  })

  it('every STRUCT_KNOWN entry names an existing case', () => {
    const keys = new Set(STRUCT_SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect(Object.keys(STRUCT_KNOWN).filter((k) => !keys.has(k))).toEqual([])
  })

  it('the comparison sees the book before the operation differ from Excel', () => {
    const set = STRUCT_SETS.find((s) => s.name === 'rows')!
    const c = set.cases[0]
    expect(differences(set.books[c.book], set.answers[c.id], String(set.meta.read)).length).toBeGreaterThan(20)
  })
})

for (const set of STRUCT_SETS) {
  const built = ((await import(`../bench/coverage/struct/${set.name}.ts`)) as { buildStructCases: () => StructCases }).buildStructCases()

  describe(`structure set ${set.name}`, () => {
    it('the stored Excel answers are for the current cases and books', () => {
      expect(set.cases).toEqual(built.cases)
      expect(set.books).toEqual(built.books)
      expect(built.cases.filter((c) => !set.answers[c.id])).toEqual([])
    })

    const look = set.meta.look === undefined ? undefined : String(set.meta.look)
    for (const c of set.cases) {
      it(`${c.id}: as Excel`, () => {
        const a = set.answers[c.id]
        const r = restructure(set.books[c.book], c.op)
        if (a.error) {
          expect(r.ok, `Excel refused: ${a.error}`).toBe(false)
          // Excel left the book as it was
          expect(differences(set.books[c.book], a, String(set.meta.read), look)).toEqual([])
          return
        }
        expect(r.ok && differences(r.book, a, String(set.meta.read), look)).toEqual(STRUCT_KNOWN[`${set.name}:${c.id}`] ?? [])
      })
    }
  })
}
