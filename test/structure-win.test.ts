/**
 * Structure operations against Excel for Windows (`test/excel/win/struct/`, `bench/windows`): after each operation the
 * engine's book shows what Excel for Windows showed, or Excel refused and so does the engine, as
 * `structure-excel.test.ts` checks the Mac's. Known differences are pinned in `WIN_STRUCT` with their lines.
 */
import { describe, expect, it } from 'vitest'
import type { StoredStruct } from '../bench/coverage/struct.js'
import type { StructCases } from '../bench/coverage/struct/cases.js'
import { restructure, type BookInput } from '../src/index.js'
import { winFiles } from './excel-win-cases.js'
import { differences } from './structure-excel.js'

const SETS = winFiles<StoredStruct>('struct/')

/** `<set>:<case id>` → the difference lines pinned for that case. */
const WIN_STRUCT: Readonly<Record<string, string[]>> = {}

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/struct/${set.name}.ts`)) as { buildStructCases: () => StructCases }).buildStructCases()

  describe(`structure set ${set.name} against Excel for Windows`, () => {
    it('the stored answers are for the current cases and books', () => {
      expect(set.cases).toEqual(built.cases)
      expect(set.books).toEqual(built.books)
      expect(built.cases.filter((c) => !set.answers[c.id])).toEqual([])
    })

    const look = set.meta.look === undefined ? undefined : String(set.meta.look)
    for (const c of set.cases)
      it(`${c.id}: as Excel for Windows`, () => {
        const a = set.answers[c.id]
        const r = restructure(set.books[c.book], c.op)
        const same = (book: BookInput) => differences(book, a, String(set.meta.read), look, set.widths !== false)
        if (a.error) {
          expect(r.ok, `Excel refused: ${a.error}`).toBe(false)
          return void expect(same(set.books[c.book])).toEqual([])
        }
        if (!r.ok && r.code === 'array') return void expect(same(set.books[c.book])).toEqual([])
        expect(r.ok ? same(r.book) : [`engine refused (${r.code}): ${r.message}`]).toEqual(WIN_STRUCT[`${set.name}:${c.id}`] ?? [])
      })
  })
}

describe('known differences from Excel for Windows: structure operations', () => {
  it('every WIN_STRUCT entry names a stored case', () => {
    const keys = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect(Object.keys(WIN_STRUCT).filter((k) => !keys.has(k))).toEqual([])
  })
})
