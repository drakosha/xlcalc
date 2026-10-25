/**
 * Book-level features against Excel for Windows (`test/excel/win/workbook/`, `bench/windows`): at each read the cells
 * of a case's rectangles hold what Excel for Windows showed, as `workbook-excel.test.ts` checks the Mac's (a cycle's
 * cells after the first read may differ where Excel's value moved, `goesOn`). Cells that knowingly differ are pinned in
 * `WIN_BOOK` by case and cell with the reason, whole cases in `WIN_CASES`; they must still differ.
 */
import { describe, expect, it } from 'vitest'
import type { StoredWorkbook } from '../bench/coverage/workbook.js'
import type { WorkbookCase } from '../bench/coverage/workbook/cases.js'
import { storedSets } from './stored-sets.js'
import { differences, differsOrThrows, engineReads, goesOn, predictionCheck } from './workbook-cases.js'
import { currentCases, pinnedCases } from './stored-cases.js'

const SETS = storedSets<StoredWorkbook>('win/workbook')

/** `<set>:<case>` → cell → why it differs from Excel for Windows (at any read). */
const PIXELS = 'a date that does not fit shows `#` across the cell: Excel for Windows measures the column in pixels, the engine in the Mac\'s points, so the count of `#` differs'
const OUTSIDE = {
  'F!A82': 'COUPDAYS is not in the whitelist (#NAME?), as on the Mac',
  'F!A92': 'XIRR stops within 1e-8 of the root in Excel, the engine finds the root (SOFT in the case sets), as on the Mac',
}
/** Excel starts a cycle from what the file stores, on Windows as on the Mac; the engine from 0. */
const START = 'Excel starts a cycle from what the file stores (50 in counter-cached, the fixed point 2 in cycles:iterate-file), empty text (#VALUE! in arithmetic) where it stores none; the engine starts from 0'
const UNLISTED = 'INDIRECT and OFFSET are not in the whitelist (#NAME?); Excel keeps the cycle through them as any other'
/** The one miss of the predictions of `cycles-held`: the engine still answers as predicted there. */
const ARRAY = 'a spilled array a cycle comes to holds its first value here and spills no more (a held value is one value a cell); Excel keeps its whole spill as it was'
const WIN_BOOK: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  'cycles:indirect': { 'S!A1': UNLISTED, 'S!B1': UNLISTED },
  'cycles:offset': { 'S!A1': UNLISTED, 'S!B1': UNLISTED },
  'cycles:iterate-file': { 'S!A1': START, 'S!B1': START, 'S!C1': START },
  'cycles-held:array-enters': { 'S!A2': ARRAY, 'S!A3': ARRAY },
  'date1904-held:held': { 'S!A9': PIXELS },
  'date1904:d1900': { ...OUTSIDE, ...Object.fromEntries(['S!A4', 'S!A5', 'S!A7', 'S!A14', 'S!A17', 'S!A18', 'S!A20'].map((k) => [k, PIXELS])) },
  'date1904:d1904': { ...OUTSIDE, 'S!A11': PIXELS },
  'iterate:counter-blank': { 'S!A1': START, 'S!B1': START },
  'iterate:counter-cached': { 'S!A1': START, 'S!B1': START },
}

/**
 * Arrays in a cycle through their own spill or cells (`spill-cycles-held`, `-2` with iterative calculation, `-off-held`
 * without): cases where the engine differs from Excel for Windows as a whole, `<set>:<case>` → why. Each must still
 * differ from Excel or throw, and its predictions are not checked.
 */
const ORDER = 'Excel depends on the order the formulas were typed in (a full calculation and a rebuild keep it); the engine gives the answer of the book from its file'
const TWO_FILES = 'from the file Excel for Windows spills A2 and gives B1 #SPILL!, Excel for Mac the other way round; the engine follows the Mac'
const SPILLS = 'without iterative calculation Excel still computes an array a cycle runs through, as a spill or #SPILL!; the engine leaves the cells of the cycle empty'
const ARRAY_CELL = 'without iterative calculation Excel shows 0 in the other cell of an array of the old kind in a cycle; the engine leaves it empty'
const WIN_CASES: Readonly<Record<string, string>> = {
  ...Object.fromEntries(['sort-typed-B3-C2-A4', 'sort-typed-B3-A4-C2', 'sort-typed-C2-B3-A4', 'sheets-typed-Sheet2.E5-Sheet2.D7-Q_data.F5-Q_data.D4', 'sheets-typed-Q_data.F5-Q_data.D4-Sheet2.E5-Sheet2.D7'].map((id) => [`spill-cycles-held:${id}`, ORDER])),
  'spill-cycles-held:wait-file': TWO_FILES,
  'spill-cycles-held:wait-typed-B1-A2-C1-C2': ORDER,
  'spill-cycles-held:wait-typed-C2-C1-A2-B1': ORDER,
  ...Object.fromEntries(
    ['sort-typed-B3-C2-A4', 'munit-typed-D11-C14', 'sheets-typed-Sheet2.E5-Sheet2.D7-Q_data.F5-Q_data.D4', 'sheets-typed-Q_data.F5-Q_data.D4-Sheet2.E5-Sheet2.D7', 'sum-typed-array-last', 'unique-typed-A4-B2', 'unique-typed-B2-A4'].map((id) => [`spill-cycles-off-held:${id}`, ORDER]),
  ),
  ...Object.fromEntries(['wait-file', 'wait-typed-C2-C1-A2-B1', 'unique-file', 'unique-typed', 'unique-found-typed'].map((id) => [`spill-cycles-off-held:${id}`, SPILLS])),
  ...Object.fromEntries(['sum-file', 'sum-typed-value-last'].map((id) => [`spill-cycles-off-held:${id}`, ARRAY_CELL])),
}

/** Predictions of held-out cases that Excel for Windows answered otherwise: `<set>:<case>` → cell → why; the engine gives Excel's answer since. */
const RANKED = 'the engine left the cells of a cycle through a spill empty before an array closing such a cycle gave way to one that closes none (x-fuzz-modes-6)'
const WIN_MISSED: Readonly<Record<string, Readonly<Record<string, string>>>> = Object.fromEntries(
  Object.entries({
    'munit-file': ['S!D11'],
    'sheets-file-Sheet2-Q_data': ['Sheet2!E5', 'Q data!D4', 'Q data!E4', 'Q data!D5', 'Q data!E5', 'Q data!D6', 'Q data!E6'],
    'wait-typed-B1-A2-C1-C2': ['S!A2'],
  }).map(([id, cells]) => [`spill-cycles-off-held:${id}`, Object.fromEntries(cells.map((k) => [k, RANKED]))]),
)

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/workbook/${set.name}.ts`)) as { buildWorkbookCases: () => WorkbookCase[] }).buildWorkbookCases()

  describe(`workbook features against Excel for Windows: set ${set.name}`, () => {
    currentCases(set, built)

    for (const c of set.cases) {
      const whole = WIN_CASES[`${set.name}:${c.id}`]
      if (whole) {
        it(`${c.id}: ${c.about} (known)`, () => expect(differsOrThrows(c, set.answers[c.id].reads), whole).toBe(true))
        continue
      }
      it(`${c.id}: ${c.about}`, () => {
        const reads = engineReads(c)
        const win = set.answers[c.id].reads
        expect(reads.length).toBe(win.length)
        const diffs = win.flatMap((a, i) => differences(c, a.values, a.shown, reads[i]).map((d) => ({ read: i, cell: d.split(/:? /)[0], text: `read ${i + 1}: ${d}` })))
        const known = WIN_BOOK[`${set.name}:${c.id}`] ?? {}
        expect(Object.keys(known).filter((k) => !diffs.some((d) => d.cell === k))).toEqual([])
        expect(diffs.filter((d) => !known[d.cell] && !goesOn(set.name, win, d.read, d.cell)).map((d) => d.text)).toEqual([])
      })
    }
    // the predictions of a held-out case Excel for Windows alone runs (`cycle`; the Mac's test checks the others), written
    // before Excel ran: the engine gives them all, its misses pinned in WIN_BOOK and WIN_MISSED
    for (const c of set.cases.filter((x) => x.cycle && x.predicted && !WIN_CASES[`${set.name}:${x.id}`]))
      it(`${c.id}: the answers predicted before Excel ran`, () => {
        expect(predictionCheck(c, set.answers[c.id].reads, WIN_MISSED[`${set.name}:${c.id}`] ?? {})).toEqual([[], []])
      })
  })
}

describe('known differences from Excel for Windows: book-level features', () => {
  it('every WIN_BOOK entry names a stored case', () => {
    expect(pinnedCases(SETS, WIN_BOOK)).toEqual([])
    expect(pinnedCases(SETS, WIN_CASES)).toEqual([])
    expect(pinnedCases(SETS, WIN_MISSED)).toEqual([])
  })
})
