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
import { differences, engineReads } from './workbook-cases.js'

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
const KNOWN: Readonly<Record<string, Readonly<Record<string, string>>>> = { 'date1904:d1900': OUTSIDE, 'date1904:d1904': OUTSIDE }

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
        const diffs = excel.flatMap((a, i) => differences(c, a.values, a.shown, reads[i]).map((d) => `read ${i + 1}: ${d}`))
        // a known difference is pinned: its cell must still differ
        const known = Object.keys(KNOWN[`${set.name}:${c.id}`] ?? {})
        const at = (d: string) => known.find((k) => d.includes(`: ${k}: `))
        expect(known.filter((k) => !diffs.some((d) => at(d) === k))).toEqual([])
        expect(diffs.filter((d) => !at(d))).toEqual([])
      })
      if (c.predicted)
        it(`${c.id}: the answer predicted before Excel ran`, () => {
          const reads = engineReads(c)
          expect(c.predicted!.flatMap((p, i) => differences(c, ...predictedRead(p), reads[i]))).toEqual([])
        })
    }
  })
}
