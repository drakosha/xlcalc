/**
 * Book-level features against Excel for Windows (`test/excel/win/workbook/`, `bench/windows`): at each read the cells
 * of a case's rectangles hold what Excel for Windows showed, as `workbook-excel.test.ts` checks the Mac's (a cycle's
 * cells after the first read may differ where Excel's value moved, `goesOn`). Cells that knowingly differ are pinned in
 * `WIN_BOOK` by case and cell with the reason; they must still differ.
 */
import { describe, expect, it } from 'vitest'
import type { StoredWorkbook } from '../bench/coverage/workbook.js'
import type { WorkbookCase } from '../bench/coverage/workbook/cases.js'
import { winFiles } from './excel-win-cases.js'
import { differences, engineReads, goesOn } from './workbook-cases.js'

const SETS = winFiles<StoredWorkbook>('workbook/')

/** `<set>:<case>` → cell → why it differs from Excel for Windows (at any read). */
const WIN_BOOK: Readonly<Record<string, Readonly<Record<string, string>>>> = {}

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/workbook/${set.name}.ts`)) as { buildWorkbookCases: () => WorkbookCase[] }).buildWorkbookCases()

  describe(`workbook features against Excel for Windows: set ${set.name}`, () => {
    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    for (const c of set.cases)
      it(`${c.id}: ${c.about}`, () => {
        const reads = engineReads(c)
        const win = set.answers[c.id].reads
        expect(reads.length).toBe(win.length)
        const diffs = win.flatMap((a, i) => differences(c, a.values, a.shown, reads[i]).map((d) => ({ read: i, cell: d.split(/:? /)[0], text: `read ${i + 1}: ${d}` })))
        const known = WIN_BOOK[`${set.name}:${c.id}`] ?? {}
        expect(Object.keys(known).filter((k) => !diffs.some((d) => d.cell === k))).toEqual([])
        expect(diffs.filter((d) => !known[d.cell] && !goesOn(set.name, win, d.read, d.cell)).map((d) => d.text)).toEqual([])
      })
  })
}

describe('known differences from Excel for Windows: book-level features', () => {
  it('every WIN_BOOK entry names a stored case', () => {
    const keys = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect(Object.keys(WIN_BOOK).filter((k) => !keys.has(k))).toEqual([])
  })
})
