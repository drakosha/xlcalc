/**
 * Structure operations against Excel for Windows (`test/excel/win/struct/`, `bench/windows`): after each operation the
 * engine's book shows what Excel for Windows showed, or Excel refused and so does the engine, as
 * `structure-excel.test.ts` checks the Mac's. Known differences are pinned in `WIN_STRUCT` with the reason.
 */
import { describe, expect, it } from 'vitest'
import type { StoredStruct } from '../bench/coverage/struct.js'
import type { StructCases } from '../bench/coverage/struct/cases.js'
import type { BookInput } from '../src/index.js'
import { winFiles } from './excel-win-cases.js'
import { differences, STRUCT_KNOWN, structureOf } from './structure-excel.js'
import { currentStructSet } from './stored-cases.js'

const SETS = winFiles<StoredStruct>('struct/')

/** The look at the sheet's edges (styles of rows and columns that reach row 1048576 or column XFD) as Excel for Windows reports it. */
const EDGE = 'Excel for Windows reports the look at the sheet\'s edges otherwise: the empty cells a span reaching the end takes after the operation read plain'

/**
 * Sizes in Excel for Windows's units, which the engine's look (in the Mac's) does not share: a row's height in points,
 * 3/4 of what Excel for Mac reports (30 for the Mac's 40, 20 for 26.75), and a column's width rounded to whole pixels
 * of a 7-pixel digit (13.29 for the Mac's 13.17). Such a line is the same size and does not count as a difference.
 */
function sameSize(line: string): boolean {
  const m = /^'[^']+' (height|width) \w+: Excel ([\d.]+), engine ([\d.]+)$/.exec(line)
  if (!m) return false
  const [win, engine] = [Number(m[2]), Number(m[3])]
  return m[1] === 'height' ? Math.abs((win * 4) / 3 - engine) <= 0.25 : Math.abs(win - engine) <= 0.2
}

/**
 * `<set>:<case id>` → why the case differs from Excel for Windows: a reason, and its difference lines must all be of the
 * look; or the lines themselves, pinned as `structure-excel.test.ts` pins the Mac's.
 */
const WIN_STRUCT: Readonly<Record<string, string | string[]>> = {
  ...Object.fromEntries(
    ['delete-rows-3-4', 'delete-row-1048572', 'delete-row-1048576', 'delete-rows-1048575-2', 'delete-rows-1048568-4', 'delete-columns-B-C', 'delete-column-XFB', 'delete-column-XFD', 'delete-columns-XEZ-3'].map((id) => [`edge-layout:${id}`, EDGE]),
  ),
  // the Mac's: whole columns copied onto part of a merged area, the fill handle over a filtered sheet
  ...STRUCT_KNOWN,
}

/** A difference line of the look: a cell's style, a row's height, a column's width. */
const LOOK = /^'[^']+' (look [A-Z]+\d+|height \d+|width [A-Z]+):/

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/struct/${set.name}.ts`)) as { buildStructCases: () => StructCases }).buildStructCases()

  describe(`structure set ${set.name} against Excel for Windows`, () => {
    currentStructSet(set, built)

    const look = set.meta.look === undefined ? undefined : String(set.meta.look)
    for (const c of set.cases)
      it(`${c.id}: as Excel for Windows`, () => {
        const a = set.answers[c.id]
        const r = structureOf(set.books[c.book], c)
        const same = (book: BookInput) => differences(book, a, String(set.meta.read), look, set.widths !== false).filter((l) => !sameSize(l))
        if (a.error) {
          expect(r.ok, `Excel refused: ${a.error}`).toBe(false)
          return void expect(same(set.books[c.book])).toEqual([])
        }
        if (!r.ok && r.code === 'array') return void expect(same(set.books[c.book])).toEqual([])
        const lines = r.ok ? same(r.book) : [`engine refused (${r.code}): ${r.message}`]
        const why = WIN_STRUCT[`${set.name}:${c.id}`] ?? []
        if (Array.isArray(why)) return void expect(lines).toEqual(why)
        expect(lines.length, why).toBeGreaterThan(0)
        expect(lines.filter((l) => !LOOK.test(l))).toEqual([])
      })
  })
}

describe('known differences from Excel for Windows: structure operations', () => {
  it('every WIN_STRUCT entry names a stored case', () => {
    const keys = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect(Object.keys(WIN_STRUCT).filter((k) => !keys.has(k))).toEqual([])
  })
})
