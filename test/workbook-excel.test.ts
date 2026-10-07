/**
 * Book-level features against Excel en-US (`test/excel/workbook/<set>.json`, from `run.ts workbook <set>`): each
 * stored case is its generator's current case, and at each read the cells of its rectangles hold what Excel showed.
 * A held-out set also keeps the answers predicted before Excel ran, and the engine gives those too.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredWorkbook } from '../bench/coverage/workbook.js'
import type { Val } from '../bench/coverage/formulajs.js'
import type { WorkbookCase } from '../bench/coverage/workbook/cases.js'
import { differences, engineReads, goesOn } from './workbook-cases.js'

/** A predicted read as Excel's answer: the values, and the shown text of `show:` keys. */
const predictedRead = (p: Record<string, Val | string>): [Record<string, Val>, Record<string, string>] => [
  Object.fromEntries(Object.entries(p).filter(([k]) => !k.startsWith('show:'))) as Record<string, Val>,
  Object.fromEntries(Object.entries(p).flatMap(([k, v]) => (k.startsWith('show:') ? [[k.slice(5), v as string]] : []))),
]

const dir = new URL('./excel/workbook/', import.meta.url)
const SETS = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredWorkbook) }))
  : []

/** Cells whose answer differs from Excel's for a reason outside the case's feature: `<set>:<case>` → cell → why. */
const OUTSIDE = {
  'F!A82': 'COUPDAYS is not in the whitelist (#NAME?)',
  'F!A86': 'ACCRINT with basis 1 over several coupon periods differs from Excel in both date systems',
  'F!A91': 'XNPV with a date of serial 0 is #NUM! in Excel in both date systems, computed here',
  'F!A92': 'XIRR stops within 1e-8 of the root in Excel, the engine finds the root (SOFT in the case sets)',
}
const HOLIDAY = { 'F!A10': 'WORKDAY reads date text in a one-cell reference among its holidays in Excel, the engine takes text in a reference there as #VALUE!' }
/** Excel starts a cycle from the value the file stores for its cells; a book stores none and the engine starts from 0. */
const START = 'Excel starts a cycle from what the file stores: 50 here, empty text (#VALUE! in arithmetic) where it stores no value; the engine starts from 0'
const KNOWN: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  'date1904:d1900': OUTSIDE,
  'date1904:d1904': OUTSIDE,
  'date1904-text:t1900': HOLIDAY,
  'date1904-text:t1904': HOLIDAY,
  'iterate:counter-blank': { 'S!A1': START, 'S!B1': START },
  'iterate:counter-cached': { 'S!A1': START, 'S!B1': START },
}

/** Predictions of held-out cases that Excel answered otherwise: `<set>:<case>` → cells, each the reason of the miss. */
const MISSED: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  'date1904-held:held': { 'F!A24': 'EDATE of date text typed in the formula: Excel takes the 1904 serial as a 1900 one (the set date1904-text measures it)' },
  'iterate-held:text': Object.fromEntries(['S!A1', 'S!B1'].map((k) => [k, 'a change of text into other text does not count to go on iterating (the set iterate measures it)'])),
}

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/workbook/${set.name}.ts`)) as { buildWorkbookCases: () => WorkbookCase[] }).buildWorkbookCases()

  describe(`workbook features against Excel en-US: set ${set.name}`, () => {
    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    for (const c of set.cases) {
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
          const missed = MISSED[`${set.name}:${c.id}`] ?? {}
          // a miss is pinned: Excel answered that cell otherwise than predicted
          const excel = set.answers[c.id].reads
          expect(Object.keys(missed).filter((k) => !c.predicted!.some((p, i) => JSON.stringify(p[k]) !== JSON.stringify(excel[i].values[k])))).toEqual([])
          const reads = engineReads(c)
          // the missed cells hold Excel's answer instead
          const kept = c.predicted!.map((p, i) => ({ ...p, ...Object.fromEntries(Object.keys(missed).map((k) => [k, excel[i].values[k]])) }))
          expect(kept.flatMap((p, i) => differences(c, ...predictedRead(p), reads[i]))).toEqual([])
        })
    }
  })
}
