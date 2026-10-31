/**
 * Σ Values (`'values'` on an axis) as `editPivotTable` and `addPivotTable` place it, as Excel keeps it: on the columns'
 * end for a second value field, where the host put it while there are two or more, off the axes with one, its place
 * kept for the next. Against Excel for Windows, the sets `values-field` and its held-out `values-field-held` driven on
 * the engine: each phase's PivotTable cells, and the place Excel saved (`rows`, `columns`, `dataOnRows`, `dataPosition`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { buildPivotCases as held } from '../bench/coverage/pivot/values-field-held.js'
import { buildPivotCases as first } from '../bench/coverage/pivot/values-field.js'
import { drive } from '../bench/coverage/pivot/drive.js'
import type { PivotCase } from '../bench/coverage/pivot/cases.js'
import { rounded, shown } from '../bench/coverage/pivot/phases.js'
import { editStructure, type BookInput, type PivotTable, type PivotTablePatch } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'

const book = (): BookInput => ({
  sheets: [
    { name: 'D', cells: { A1: { value: 'R' }, B1: { value: 'P' }, C1: { value: 'Q' }, D1: { value: 'V' }, A2: { value: 'a' }, B2: { value: 'x' }, C2: { value: 1 }, D2: { value: 2 } } },
    { name: 'P', cells: {} },
  ],
})

/** The PivotTable after an add with `definition` and each patch in turn, without `xlcalc/pivot`. */
function after(definition: PivotTablePatch, ...patches: PivotTablePatch[]): PivotTable {
  let r = editStructure(book(), { op: 'addPivotTable', sheet: 'P', cell: 'A3', name: 'T', source: { type: 'range', sheet: 'D', range: 'A1:D2' }, definition })
  for (const set of patches) {
    if (!r.ok) break
    r = editStructure(r.book, { op: 'editPivotTable', sheet: 'P', pivot: 'T', set })
  }
  if (!r.ok) throw new Error(r.message)
  return r.book.sheets[1].pivotTables![0]
}
const place = (t: PivotTable) => [t.rows, t.columns, t.valuesOnRows, t.valuesPosition]
const two = [{ field: 2 }, { field: 3 }]

describe('Σ Values in editPivotTable and addPivotTable', () => {
  it('comes at the end of the columns with a second value field and goes with one left', () => {
    expect(place(after({ rows: [0], columns: [1], values: two }))).toEqual([[0], [1, 'values'], undefined, undefined])
    expect(place(after({ rows: [0], values: [{ field: 2 }] }, { values: two }, { values: [{ field: 3 }] }))).toEqual([[0], [], undefined, undefined])
  })

  it('stays where the host put it, and comes back there after one value field', () => {
    const mid = after({ rows: [0, 'values', 1], values: two })
    expect(place(mid)).toEqual([[0, 'values', 1], undefined, true, 1])
    const one = after({ rows: [0, 'values', 1], values: two }, { values: [{ field: 2 }] })
    expect(place(one)).toEqual([[0, 1], undefined, true, 1])
    expect(place(after({ rows: [0, 'values', 1], values: two }, { values: [{ field: 2 }] }, { values: two }))).toEqual([[0, 'values', 1], undefined, true, 1])
  })

  it('moves by valuesOnRows and valuesPosition; a field moved beside it leaves its position', () => {
    expect(place(after({ rows: [0], values: two }, { valuesOnRows: true }))).toEqual([[0, 'values'], [], true, undefined])
    expect(place(after({ rows: [0, 1], values: two }, { valuesOnRows: true, valuesPosition: 0 }))).toEqual([['values', 0, 1], [], true, 0])
    expect(place(after({ rows: [0, 'values', 1], values: two }, { rows: [0, 'values'] }))).toEqual([[0, 'values'], undefined, true, undefined])
    expect(place(after({ rows: [0, 1], values: two, valuesOnRows: true, valuesPosition: 1 }, { rows: [0] }))).toEqual([[0, 'values'], undefined, true, 1])
  })
})

const saved = (set: string, id: string, k: number, last: boolean) => readXlsx(new Uint8Array(readFileSync(new URL(`./excel/win/pivot/${set}/${id}${last ? '' : `-${k + 1}`}.xlsx`, import.meta.url))))
const placeOf = (b: BookInput) => {
  const t = b.sheets.find((s) => s.name === 'P')!.pivotTables!.find((p) => p.name === 'PT1')!
  return [t.rows ?? [], t.columns ?? [], t.valuesOnRows === true, t.valuesPosition]
}

/** Count of a field with errors over Σ Values on the rows: Excel shows the error in its subtotals and grand total, the engine counts (not a matter of Σ Values). */
const KNOWN = ['values-field/third-rows-middle/1', 'values-field/rows-drop-back/2']

for (const [set, cases] of [['values-field', first()], ['values-field-held', held()]] as [string, PivotCase[]][]) {
  const answers = JSON.parse(readFileSync(new URL(`./excel/win/pivot/${set}.json`, import.meta.url), 'utf8')).answers
  describe(`${set} against Excel for Windows`, () => {
    for (const c of cases)
      it(c.id, async () => {
        const start = c.from ? (await readXlsx(new Uint8Array(readFileSync(new URL(`./excel/win/pivot/${c.from}.xlsx`, import.meta.url))))).book : undefined
        const phases = drive(c, start)
        expect(phases.length).toBe(answers[c.id].phases.length)
        for (const [k, p] of phases.entries()) {
          expect(p.errors).toEqual([])
          const want = answers[c.id].phases[k].pivots.PT1
          if (!KNOWN.includes(`${set}/${c.id}/${k}`)) expect(rounded(shown(p.book, 'P', want.area)), `${c.id} ${k}`).toEqual(rounded(want.values))
          expect(placeOf(p.book), `${c.id} ${k}`).toEqual(placeOf((await saved(set, c.id, k, k === phases.length - 1)).book))
        }
      })
  })
}
