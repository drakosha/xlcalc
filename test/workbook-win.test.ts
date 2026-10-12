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
const PIXELS = 'a date that does not fit shows `#` across the cell: Excel for Windows measures the column in pixels, the engine in the Mac\'s points, so the count of `#` differs'
const OUTSIDE = {
  'F!A82': 'COUPDAYS is not in the whitelist (#NAME?), as on the Mac',
  'F!A92': 'XIRR stops within 1e-8 of the root in Excel, the engine finds the root (SOFT in the case sets), as on the Mac',
}
/** Excel starts a cycle from what the file stores, on Windows as on the Mac; the engine from 0. */
const START = 'Excel starts a cycle from what the file stores (50 in counter-cached, the fixed point 2 in cycles:iterate-file), empty text (#VALUE! in arithmetic) where it stores none; the engine starts from 0'
/** A cycle without iterative calculation, the set `cycles`. */
const OFF = 'without iterative calculation Excel keeps the cells of a cycle and the cells that read them as they are (the stored value, empty text, 0 for a formula typed); the engine gives #CYCLE!'
const UNLISTED = 'INDIRECT and OFFSET are not in the whitelist (#NAME?); Excel keeps the cycle through them as any other'
/** The cells of each case of `cycles` pinned with OFF. */
const CYCLES: Readonly<Record<string, readonly string[]>> = {
  typed: ['S!A1', 'S!B1', 'S!D1', 'S!E1'],
  stored: ['S!A1', 'S!B1', 'S!C1'],
  'stored-other': ['S!A1', 'S!B1', 'S!C1'],
  'stored-none': ['S!A1', 'S!B1', 'S!C1'],
  precedent: ['S!A1', 'S!B1', 'S!D1'],
  'typed-again': ['S!A1', 'S!B1', 'S!C1'],
  broken: ['S!A1', 'S!B1', 'S!C1'],
  self: ['S!A1', 'S!B1', 'S!C1'],
  name: ['S!A1', 'S!B1'],
  spill: ['S!A1', 'S!B5'],
  unrelated: ['S!A1', 'S!B1', 'S!C1'],
  above: ['S!A1', 'S!C3', 'S!D3'],
  sheet: ['S!A1', 'S!B1', 'S!C1', 'T!A1'],
  'other-precedent': ['S!A1', 'S!B1', 'S!C1'],
  'new-dependent': ['S!A1', 'S!B1', 'S!C1', 'S!E1'],
  'typed-chain': ['S!A1', 'S!B1', 'S!C1'],
  'after-iterate': ['S!A1', 'S!B1', 'S!C1'],
}
const WIN_BOOK: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  ...Object.fromEntries(Object.entries(CYCLES).map(([id, cells]) => [`cycles:${id}`, Object.fromEntries(cells.map((k) => [k, OFF]))])),
  'cycles:indirect': { 'S!A1': UNLISTED, 'S!B1': UNLISTED },
  'cycles:offset': { 'S!A1': UNLISTED, 'S!B1': UNLISTED },
  'cycles:iterate-file': { 'S!A1': START, 'S!B1': START, 'S!C1': START },
  'date1904-held:held': { 'S!A9': PIXELS },
  'date1904:d1900': { ...OUTSIDE, ...Object.fromEntries(['S!A4', 'S!A5', 'S!A7', 'S!A14', 'S!A17', 'S!A18', 'S!A20'].map((k) => [k, PIXELS])) },
  'date1904:d1904': { ...OUTSIDE, 'S!A11': PIXELS },
  'iterate:counter-blank': { 'S!A1': START, 'S!B1': START },
  'iterate:counter-cached': { 'S!A1': START, 'S!B1': START },
}

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
