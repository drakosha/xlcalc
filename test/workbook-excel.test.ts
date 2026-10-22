/**
 * Book-level features against Excel en-US (`test/excel/workbook/<set>.json`, from `run.ts workbook <set>`): each
 * stored case is its generator's current case, and at each read the cells of its rectangles hold what Excel showed.
 * A held-out set also keeps the answers predicted before Excel ran, and the engine gives those too.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredWorkbook } from '../bench/coverage/workbook.js'
import type { WorkbookCase } from '../bench/coverage/workbook/cases.js'
import { differences, engineReads, goesOn, predictionCheck } from './workbook-cases.js'
import { currentCases } from './stored-cases.js'

const dir = new URL('./excel/workbook/', import.meta.url)
const SETS = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredWorkbook) }))
  : []

/** Cells whose answer differs from Excel's for a reason outside the case's feature: `<set>:<case>` → cell → why. */
const OUTSIDE = {
  'F!A82': 'COUPDAYS is not in the whitelist (#NAME?)',
  'F!A92': 'XIRR stops within 1e-8 of the root in Excel, the engine finds the root (SOFT in the case sets)',
}
/** Excel starts a cycle from the value the file stores for its cells; a book stores none and the engine starts from 0. */
const START = 'Excel starts a cycle from what the file stores: 50 here, empty text (#VALUE! in arithmetic) where it stores no value; the engine starts from 0'
const KNOWN: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  'date1904:d1900': OUTSIDE,
  'date1904:d1904': OUTSIDE,
  'iterate:counter-blank': { 'S!A1': START, 'S!B1': START },
  'iterate:counter-cached': { 'S!A1': START, 'S!B1': START },
  'spill-cycles-held:wait-file': { 'S!C1': START, 'S!C2': START },
}

/** The session kept the outcome of its edits, other than its compute(), before the engine gave a book one answer (x-fuzz-modes-6). */
const ONE = 'the session kept another answer than a full computation of its book before the engine gave a book one answer (x-fuzz-modes-6)'
/** The same cells missed for the same reason in each of the cases `ids` of `set`. */
const missed = (set: string, ids: string[], cells: string[], why: string) => Object.fromEntries(ids.map((id) => [`${set}:${id}`, Object.fromEntries(cells.map((k) => [k, why]))]))
/** Predictions of held-out cases that Excel answered otherwise: `<set>:<case>` → cells, each the reason of the miss. */
const MISSED: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  'date1904-held:held': { 'F!A24': 'EDATE of date text typed in the formula: Excel takes the 1904 serial as a 1900 one (the set date1904-text measures it)' },
  'iterate-held:text': Object.fromEntries(['S!A1', 'S!B1'].map((k) => [k, 'a change of text into other text does not count to go on iterating (the set iterate measures it)'])),
  ...missed('spill-cycles-held', ['wait-file', 'wait-typed-B1-A2-C1-C2'], ['S!B1', 'S!C1', 'S!A2', 'S!B2', 'S!C2'], 'the engine threw "spilled arrays did not settle" before it gave a book one answer (x-fuzz-modes-6)'),
  ...missed('spill-cycles-held-2', ['unique-typed', 'unique-typed-20', 'unique-found-typed'], ['S!B3', 'S!A5'], ONE),
  ...missed('spill-cycles-held-2', ['sum-typed-array-last', 'sum-typed-value-last'], ['S!C3', 'S!D3'], ONE),
}

/**
 * Arrays in a cycle through their own spill beside another array over the same cells (`spill-cycles-held`): Excel's answer
 * follows the order the formulas were typed in and stays after a full calculation and a rebuild of the dependencies, while
 * a book opened from its file gets one answer. The engine gives a book one answer, the file's, whatever order its cells
 * were entered in (`docs/limitations.md`): each case pinned here must still differ from Excel, and its predictions are not checked.
 */
const ORDER = 'Excel keeps the outcome the typing order gave it (a full calculation and a rebuild keep it too); the engine gives the answer of the book opened from its file'
const PENDING: Readonly<Record<string, string>> = Object.fromEntries(
  ['sort-typed-B3-C2-A4', 'sort-typed-B3-A4-C2', 'sort-typed-C2-B3-A4', 'sheets-typed-Sheet2.E5-Sheet2.D7-Q_data.F5-Q_data.D4', 'sheets-typed-Q_data.F5-Q_data.D4-Sheet2.E5-Sheet2.D7', 'wait-typed-C2-C1-A2-B1'].map((id) => [`spill-cycles-held:${id}`, ORDER]),
)

/** Whether the engine differs from Excel at some read. */
function differs(c: WorkbookCase, excel: StoredWorkbook['answers'][string]['reads']): boolean {
  const reads = engineReads(c)
  return excel.some((a, i) => differences(c, a.values, a.shown, reads[i]).length > 0)
}

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/workbook/${set.name}.ts`)) as { buildWorkbookCases: () => WorkbookCase[] }).buildWorkbookCases()

  describe(`workbook features against Excel en-US: set ${set.name}`, () => {
    currentCases(set, built)

    for (const c of set.cases) {
      const pending = PENDING[`${set.name}:${c.id}`]
      if (pending) {
        it(`${c.id}: ${c.about} (known difference)`, () => expect(differs(c, set.answers[c.id].reads), pending).toBe(true))
        continue
      }
      it(`${c.id}: ${c.about}`, () => {
        const reads = engineReads(c)
        const excel = set.answers[c.id].reads
        expect(reads.length).toBe(excel.length)
        const diffs = excel.flatMap((a, i) => differences(c, a.values, a.shown, reads[i]).map((d) => ({ read: i, cell: d.split(/:? /)[0], text: `read ${i + 1}: ${d}` })))
        // a known difference is pinned: its cell must still differ
        const known = KNOWN[`${set.name}:${c.id}`] ?? {}
        expect(Object.keys(known).filter((k) => !diffs.some((d) => d.cell === k))).toEqual([])
        expect(diffs.filter((d) => !known[d.cell] && !goesOn(set.name, excel, d.read, d.cell)).map((d) => d.text)).toEqual([])
      })
      if (c.predicted)
        it(`${c.id}: the answers predicted before Excel ran, but the missed ones`, () => {
          // a miss is pinned: Excel answered that cell otherwise than predicted, and it holds Excel's answer instead
          const excel = set.answers[c.id].reads
          const known = (read: number, cell: string) => !!KNOWN[`${set.name}:${c.id}`]?.[cell] || goesOn(set.name, excel, read, cell)
          expect(predictionCheck(c, excel, MISSED[`${set.name}:${c.id}`] ?? {}, known)).toEqual([[], []])
        })
    }
  })
}
