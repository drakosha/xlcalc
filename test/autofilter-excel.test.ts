/**
 * The AutoFilter under what is entered on its sheet, against Excel (`test/excel/filter/` from Excel for Mac, its steps
 * a person's on the keyboard, `test/excel/win/filter/` from Excel for Windows through the object model; the step
 * `filter` of `bench/coverage` and `bench/windows`): after each step `getAutoFilterRange` is the range Excel reads
 * (`AutoFilter.Range`), and the AutoFilter's own range is the one the file Excel saved holds. The predictions committed
 * before Excel ran each set (`bench/coverage/filter/<set>.predicted.json`) are kept with the cases Excel did not confirm.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredFilter } from '../bench/coverage/filter.js'
import { filterInEngine, type FilterAnswer } from '../bench/coverage/filter/cases.js'
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
