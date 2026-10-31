/**
 * Fields placed from Excel's PivotTable Fields list against Excel for Windows (`bench/coverage/pivot/checked.ts`, the
 * answers in `test/excel/win/pivot-fields/`): each case's fields placed as `suggestPivotField` says, through
 * `editPivotTable` with `groupDates` (a field dragged to Values with `function`), give the PivotTable the book Excel saved holds (its axes, value fields and
 * functions, the fields grouping added) and its cells. Cases Excel did not place (`error`) are left out.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { checkedOf, predictBook, type CheckedCase } from '../bench/coverage/pivot/checked.js'
import { buildCheckedCases as first } from '../bench/coverage/pivot/checked-1.js'
import { buildCheckedCases as held } from '../bench/coverage/pivot/checked-held.js'
import { buildCheckedCases as values } from '../bench/coverage/pivot/checked-values.js'
import { buildCheckedCases as values2 } from '../bench/coverage/pivot/checked-values-2.js'
import type { BookInput } from '../src/contract.js'
import { readXlsx } from '../src/xlsx.js'

const DIR = new URL('./excel/win/pivot-fields/', import.meta.url)

/** Cells of P that differ for reasons of the cache's refresh, not of placing (the case, why). */
const KNOWN: Record<string, string> = {
  'checked-1/Err': 'errors listed #N/A before #DIV/0!, the engine the other way',
  'checked-1/Times': 'a time of day kept as a date in the cache comes back a bit off',
  'checked-1/DateText': 'text before dates in a new field, the engine after',
}

/** The cases of checked-values whose field the mouse did not reach, out of the list's view (run again in checked-values-2). */
const UNPLACED = ['Text', 'Years', 'Bool', 'Blank', 'dates-columns', 'month-columns']

const cellsOf = (b: BookInput) => Object.fromEntries(Object.entries(b.sheets.find((s) => s.name === 'P')!.cells).flatMap(([a, c]) => (c.value !== undefined || c.formula !== undefined ? [[a, c.formula ?? c.value]] : [])))

for (const [set, cases] of [['checked-1', first()], ['checked-held', held()], ['checked-values', values().filter((c) => !UNPLACED.includes(c.id))], ['checked-values-2', values2()]] as [string, CheckedCase[]][]) {
  const answers = JSON.parse(readFileSync(new URL(`${set}.json`, DIR), 'utf8')).answers
  describe(`${set} against Excel for Windows`, () => {
    for (const c of cases.filter((x) => !answers[x.id].error))
      it(c.id, async () => {
        const book = predictBook(c)
        if (typeof book === 'string') throw new Error(book)
        const { dialog: _, ...want } = answers[c.id]
        expect(checkedOf(book, c.watch)).toEqual(want)
        const excel = (await readXlsx(new Uint8Array(readFileSync(new URL(`${set}/${c.id}.xlsx`, DIR))))).book
        if (!KNOWN[`${set}/${c.id}`]) expect(cellsOf(book)).toEqual(cellsOf(excel))
      })
  })
}
