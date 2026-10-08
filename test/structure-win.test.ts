/**
 * Structure operations against Excel for Windows (`test/excel/win/struct/`, `bench/windows`): after each operation the
 * engine's book shows what Excel for Windows showed, or Excel refused and so does the engine, as
 * `structure-excel.test.ts` checks the Mac's. Known differences are pinned in `WIN_STRUCT` with the reason.
 */
import { describe, expect, it } from 'vitest'
import type { StoredStruct } from '../bench/coverage/struct.js'
import type { StructCases } from '../bench/coverage/struct/cases.js'
import { restructure, type BookInput } from '../src/index.js'
import { winFiles } from './excel-win-cases.js'
import { differences } from './structure-excel.js'

const SETS = winFiles<StoredStruct>('struct/')

/**
 * The look at the sheet's edges (styles of rows and columns that reach row 1048576 or column XFD), as Excel for Windows
 * reports it: the empty cells such a span takes after the operation read plain, rows read their height in points where
 * Excel for Mac reports 4/3 of it, and columns their width rounded to pixels (19.29 where the Mac has 19.17).
 */
const EDGE = 'Excel for Windows reports the look at the sheet\'s edges otherwise: empty cells of a span reaching the end read plain, row heights in its points (30, the Mac 40) and widths rounded to pixels'
const PIXEL_WIDTHS = 'Excel for Windows reports a column width rounded to whole pixels (9.29 where the Mac has 9.17)'

/**
 * `<set>:<case id>` → why the case differs from Excel for Windows: a reason, and its difference lines must all be of the
 * look (cells, heights, widths); or the lines themselves, pinned as `structure-excel.test.ts` pins the Mac's.
 */
const WIN_STRUCT: Readonly<Record<string, string | string[]>> = {
  ...Object.fromEntries(
    [
      ...['delete-rows-3-4', 'delete-row-1048572', 'delete-row-1048576', 'delete-rows-1048575-2', 'delete-rows-1048568-4', 'delete-columns-B-C', 'delete-column-XFB', 'delete-column-XFD', 'delete-columns-XEZ-3'].map((id) => `edge-layout:${id}`),
      ...['delete-column-A', 'delete-column-D', 'insert-column-B', 'delete-row-3'].map((id) => `edge-wide:${id}`),
    ].map((k) => [k, EDGE]),
  ),
  'hidden:columns-insert-inside': PIXEL_WIDTHS,
  'hidden:columns-insert-after': PIXEL_WIDTHS,
  // Excel's fill handle (Range.AutoFill) fills the hidden rows of a filtered sheet too, on Windows as on the Mac; the
  // engine's copyRange fills the visible rows only
  'hidden:filtered-fill-handle': ['D4', 'E4', 'D7', 'E7', 'D8', 'E8', 'D11', 'E11'].map((a) => `'S'!${a}: Excel ${JSON.stringify(a[0] === 'D' ? 'top' : `=A${a.slice(1)}+1000`)}, engine ""`),
}

/** A difference line of the look: a cell's style, a row's height, a column's width. */
const LOOK = /^'[^']+' (look [A-Z]+\d+|height \d+|width [A-Z]+):/

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
