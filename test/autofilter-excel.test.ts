/**
 * The AutoFilter under what is entered on its sheet, against Excel (`test/excel/filter/` from Excel for Mac, its steps
 * a person's on the keyboard, `test/excel/win/filter/` from Excel for Windows through the object model; the step
 * `filter` of `bench/coverage` and `bench/windows`): after each step `getAutoFilterRange` is the range Excel reads
 * (`AutoFilter.Range`), and the AutoFilter's own range is the one the file Excel saved holds. The predictions committed
 * before Excel ran each set (`bench/coverage/filter/<set>.predicted.json`) are kept with the cases Excel did not confirm.
 * The filter-acts sets (`test/excel/filter-acts/`, `test/excel/win/filter-acts/`) ask what criteria, a sort and removing
 * a filter do to the rows, which the engine has no step for: they are read for the range alone, against `getAutoFilterRange`
 * over the range each file holds, and the rest against the predictions committed before each Excel ran them.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredFilter } from '../bench/coverage/filter.js'
import type { StoredActs } from '../bench/coverage/filter-acts.js'
import { actBook, asPredicted, missedBy, type ActCase } from '../bench/coverage/filter-acts/cases.js'
import { filterInEngine, type FilterAnswer } from '../bench/coverage/filter/cases.js'
import { createSession } from '../src/index.js'
import { parseRange } from '../src/structure/area.js'

const SETS = ['grow', 'grow-held', 'grow-held-2']

/**
 * The cases each prediction missed: `grow`'s engine grew the AutoFilter's own range over rows typed under it in its
 * columns, where Excel keeps that range and reads it grown over rows beside it too, and back when they are emptied;
 * `grow-held`'s took the rows beside it alone, where Excel reaches diagonally further (`far-3`), and left the range a
 * host filters as it was (`criteria-*`).
 */
const MISSED: Record<string, string[]> = {
  grow: ['under-1', 'under-2', 'under-3', 'under-4', 'under-5', 'under-6', 'under-7', 'gap-2', 'paste-1', 'paste-2', 'paste-4', 'copy-1', 'copy-2', 'copy-3', 'fill-1', 'fill-2', 'fill-3', 'entry-1', 'entry-2', 'criteria-1', 'criteria-2', 'criteria-3', 'header-1', 'clear-2', 'clear-3', 'undo-3', 'set-1', 'set-2'],
  'grow-held': ['far-3', 'criteria-1', 'criteria-2'],
  'grow-held-2': [],
}

const load = (file: string) => JSON.parse(readFileSync(new URL(file, import.meta.url), 'utf8')) as StoredFilter

/** An answer with its ranges as the cells they cover (`A:C` and `A1:C1048576` alike). */
const cells = (a: FilterAnswer) => ({ after: a.after.map((r) => (r ? parseRange(r) : null)), saved: a.saved ? parseRange(a.saved) : null })

describe.each([
  ['Mac', ''],
  ['Windows', 'win/'],
])('the AutoFilter under what is entered, Excel for %s', (_, dir) => {
  for (const set of SETS) {
    const file = `./excel/${dir}filter/${set}.json`
    it.skipIf(!existsSync(new URL(file, import.meta.url)))(`${set}: the range Excel reads after each step, and the one it saves`, () => {
      const stored = load(file)
      for (const c of stored.cases) expect(cells(filterInEngine(c)), c.id).toEqual(cells(stored.answers[c.id]))
    })
  }
})

describe('the predictions committed before Excel for Mac ran each set', () => {
  for (const set of SETS)
    it(`${set}: the cases Excel did not confirm`, () => {
      const stored = load(`./excel/filter/${set}.json`)
      const predicted = (JSON.parse(readFileSync(new URL(`../bench/coverage/filter/${set}.predicted.json`, import.meta.url), 'utf8')) as { cases: Record<string, FilterAnswer> }).cases
      const missed = stored.cases.filter((c) => JSON.stringify(cells(predicted[c.id])) !== JSON.stringify(cells(stored.answers[c.id]))).map((c) => c.id)
      expect(missed).toEqual(MISSED[set])
    })
})

/** The filter-acts sets, each listed in every comparison below. */
const ACT_SETS = ['totals', 'removal', 'moved', 'kept', 'totals-held', 'removal-held', 'totals-own', 'totals-functions']

/**
 * The cases each prediction missed, in each Excel: `totals` took a SUBTOTAL row in as any other, had criteria and a sort
 * save the range they read, and a table's filter leave its totals row and the filter mode out; `removal` had a row of
 * height 0 shown; `moved` had `Range.AutoFilter` move or remove the AutoFilter; `totals-held` left two rows of SUBTOTAL out, and
 * `Range.AutoFilter` on a cell over such a row save the range above it, as it does on the Mac alone; `removal-held` had a
 * hidden row under the range of a file's filter shown.
 */
const ACT_MISSED: Record<string, string[]> = {
  'Mac totals': ['sum-sort-model', 'sum-sort-button', 'sub109-criteria-model', 'sub109-criteria-button', 'sub109-sort-model', 'sub109-sort-button', 'sub9-criteria-model', 'sub9-criteria-button', 'sub9-sort-model', 'sub9-sort-button', 'label-sort-model', 'label-sort-button', 'table-criteria-model', 'table-criteria-button', 'table-sort-model', 'table-sort-button', 'pick-sum-range-button', 'pick-sub109-cell-model', 'pick-sub109-range-button'],
  'Windows totals': ['sum-sort-model', 'sum-sort-button', 'sub109-criteria-model', 'sub109-sort-model', 'sub109-sort-button', 'sub9-criteria-model', 'sub9-sort-model', 'sub9-sort-button', 'label-sort-model', 'label-sort-button', 'table-criteria-model', 'table-sort-model', 'table-sort-button', 'pick-sum-range-button', 'pick-sub109-range-button'],
  'Mac removal': ['b-fail-model', 'b-pass-model', 'b-fail-button', 'b-pass-button'],
  'Windows removal': ['b-fail-model', 'b-pass-model', 'b-fail-button', 'b-pass-button'],
  'Mac moved': ['a-field', 'a-plain', 'a-off-model', 'a-off-button', 'b-field', 'b-plain', 'c-field', 'c-plain', 'd-field', 'd-plain'],
  'Windows moved': ['a-field', 'a-off-button', 'b-field', 'c-field', 'd-field'],
  'Mac totals-held': ['sub-then-sub-criteria', 'sub-then-sub-sort'],
  'Windows totals-held': ['sub-then-sub-criteria', 'sub-then-sub-sort', 'order-2-model', 'order-4-model', 'order-5-model', 'order-7-model', 'one-book-button-model'],
  'Mac removal-held': ['file-sum-model', 'file-sum-button'],
  'Windows removal-held': ['file-sum-model', 'file-sum-button'],
}

/** The range `getAutoFilterRange` reads over a case's cells for an AutoFilter on `own`, empty without one. */
function readOver(c: ActCase, own: string | null): string {
  if (!own) return ''
  const s = createSession(actBook({ ...c, filter: own, keptOff: undefined }))
  const got = s.getAutoFilterRange('S')!
  s.close()
  return got
}

describe.each([
  ['Mac', ''],
  ['Windows', 'win/'],
])('what an AutoFilter does to its rows, Excel for %s', (name, dir) => {
  for (const set of ACT_SETS)
    describe(set, () => {
      const stored = JSON.parse(readFileSync(new URL(`./excel/${dir}filter-acts/${set}.json`, import.meta.url), 'utf8')) as StoredActs
      // a table's filter reads its totals row where it has one, and the engine has no such range
      const sheets = stored.cases.filter((c) => !c.table)
      it('every case is answered', () => expect(stored.cases.filter((c) => !stored.answers[c.id]).map((c) => c.id)).toEqual([]))
      it('getAutoFilterRange reads the range Excel reads as the book opens and at the end, over the range the file holds', () => {
        for (const c of sheets) {
          const a = stored.answers[c.id]
          if (!c.keptOff) expect(readOver(c, c.filter ?? null), `${c.id} opened`).toBe(a.opened!.range)
          const end = a.after?.at(-1)
          if (end) expect(readOver(c, a.saved!.ref), `${c.id} at the end`).toBe(end.range)
        }
      })
      it('the cases the prediction committed before the run missed', () => {
        const predicted = (JSON.parse(readFileSync(new URL(`../bench/coverage/filter-acts/${set}.predicted.json`, import.meta.url), 'utf8')) as { cases: Record<string, unknown> }).cases
        const missed = stored.cases.filter((c) => missedBy(predicted[c.id], asPredicted(stored.answers[c.id])).length).map((c) => c.id)
        expect(missed).toEqual(ACT_MISSED[`${name} ${set}`] ?? [])
      })
    })
})
