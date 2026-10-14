/**
 * Acceptance of `editStructure` against Excel en-US (`test/excel/struct/`, see `test/structure-excel.ts`):
 * each stored case is the generator's current case, and the engine's book after the operation shows what
 * Excel showed, or Excel refused and so does the engine. Known differences are listed in STRUCT_KNOWN
 * with their lines.
 */
import { describe, expect, it } from 'vitest'
import type { StructCases } from '../bench/coverage/struct/cases.js'
import { takenBook } from '../bench/coverage/stored.js'
import { editStructure } from '../src/index.js'
import { differences, STRUCT_KNOWN, STRUCT_SETS } from './structure-excel.js'
import { currentStructSet } from './stored-cases.js'


describe('structure operations against Excel en-US', () => {
  it('every set has its cases answered', () => {
    expect(STRUCT_SETS.map((s) => s.name)).toEqual(['arrays', 'checks', 'cube', 'edge-layout', 'edge-wide', 'edge', 'hidden', 'history', 'move', 'moves', 'rows', 'sheets', 'spill-copy', 'spill-more', 'spill', 'styles-checks', 'styles', 'zero-held'])
  })

  it('every STRUCT_KNOWN entry names an existing case', () => {
    const keys = new Set(STRUCT_SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect(Object.keys(STRUCT_KNOWN).filter((k) => !keys.has(k))).toEqual([])
  })

  it('the comparison sees the book before the operation differ from Excel, its look too', () => {
    const set = STRUCT_SETS.find((s) => s.name === 'rows')!
    const c = set.cases[0]
    expect(differences(set.books[c.book], set.answers[c.id], String(set.meta.read)).length).toBeGreaterThan(20)
    const styled = STRUCT_SETS.find((s) => s.name === 'styles')!
    const i = styled.cases.find((x) => x.id === 'insert-rows-7')!
    const lines = differences(styled.books[i.book], styled.answers[i.id], String(styled.meta.read), String(styled.meta.look))
    expect(lines.filter((l) => / (look|height|merges)/.test(l)).length).toBeGreaterThan(10)
  })
})

for (const set of STRUCT_SETS) {
  const built = ((await import(`../bench/coverage/struct/${set.name}.ts`)) as { buildStructCases: () => StructCases }).buildStructCases()

  describe(`structure set ${set.name}`, () => {
    currentStructSet(set, built)

    const look = set.meta.look === undefined ? undefined : String(set.meta.look)
    for (const c of set.cases) {
      it(`${c.id}: as Excel`, () => {
        const a = set.answers[c.id]
        const r = editStructure(takenBook(set.books[c.book]), c.op)
        if (a.error) {
          expect(r.ok, `Excel refused: ${a.error}`).toBe(false)
          // Excel left the book as it was
          expect(differences(set.books[c.book], a, String(set.meta.read), look, set.widths !== false)).toEqual([])
          return
        }
        // Excel leaves the book as it was, without an error, where it does not change part of an old-kind array
        if (!r.ok && r.code === 'array') return void expect(differences(set.books[c.book], a, String(set.meta.read), look, set.widths !== false)).toEqual([])
        expect(r.ok ? differences(r.book, a, String(set.meta.read), look, set.widths !== false) : [`engine refused (${r.code}): ${r.message}`]).toEqual(STRUCT_KNOWN[`${set.name}:${c.id}`] ?? [])
      })
    }
  })
}
