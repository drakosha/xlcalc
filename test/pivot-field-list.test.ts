/**
 * What Excel's PivotTable Fields list does that a macro does not, as `editPivotTable` and `addPivotTable` take it:
 * `groupDates` groups a field of dates put on the rows or columns by periods (Excel 365's time grouping), in every
 * kind of session and both execution modes, one step of undo; refused without `xlcalc/pivot`, a wrong flag thrown.
 * Against Excel for Windows: dates placed through the object model are not grouped (the set `com-dates`), and a field
 * put on Values without a function takes the one `suggestPivotField` names (`com-values`).
 */
import { afterAll, describe, expect, it } from 'vitest'
import { editStructure, InputError, registerPivot, type BookInput, type StructureOp } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { setPivotCacheEdit, pivotCacheEdit } from '../src/pivot/hook.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { readFileSync } from 'node:fs'
import { buildPivotCases as comDates } from '../bench/coverage/pivot/com-dates.js'
import { buildPivotCases as comValues } from '../bench/coverage/pivot/com-values.js'
import { drive } from '../bench/coverage/pivot/drive.js'
import { rounded, shown } from '../bench/coverage/pivot/phases.js'

afterAll(closeCalcs)
registerPivot(pivotEngine)

const clock = { now: '2026-10-31T12:00:00Z', seed: 1 }
const day = (y: number, m: number, d: number) => Date.UTC(y, m - 1, d) / 86_400_000 + 25_569

const book = (): BookInput => ({
  sheets: [
    {
      name: 'Data',
      cells: {
        A1: { value: 'Date' }, B1: { value: 'Amount' },
        A2: { value: day(2024, 3, 5), numberFormat: 'yyyy-mm-dd' }, B2: { value: 10 },
        A3: { value: day(2024, 11, 20), numberFormat: 'yyyy-mm-dd' }, B3: { value: 20 },
        A4: { value: day(2025, 6, 1), numberFormat: 'yyyy-mm-dd' }, B4: { value: 5 },
      },
    },
    { name: 'P', cells: {} },
  ],
})

const add: StructureOp = { op: 'addPivotTable', sheet: 'P', cell: 'A3', source: { type: 'range', sheet: 'Data', range: 'A1:B4' }, definition: { values: [{ field: 1 }] } }
const byDates: StructureOp = { op: 'editPivotTable', sheet: 'P', pivot: 'PivotTable1', set: { rows: [0] }, groupDates: true }

describe('groupDates', () => {
  it('groups dates over years by years, quarters and months, the dates last, the periods collapsed', () => {
    const r = editStructure(book(), { ...add, definition: { rows: [0], values: [{ field: 1 }] }, groupDates: true } as StructureOp)
    if (!r.ok) throw new Error(r.message)
    const fields = r.book.pivotCaches![0].fields.map((f) => f.name)
    expect(r.book.sheets[1].pivotTables![0].rows!.map((f) => fields[f as number])).toEqual(['Years (Date)', 'Quarters (Date)', 'Months (Date)', 'Date'])
    expect(['A4', 'B4', 'A5', 'B5', 'A6', 'B6'].map((a) => r.book.sheets[1].cells[a]?.value)).toEqual(['2024', 30, '2025', 5, 'Grand Total', 35])
  })

  it('leaves a field grouped, dates it does not place and a macro-like edit without the flag alone', () => {
    const r = editStructure(book(), add)
    if (!r.ok) throw new Error(r.message)
    const plain = editStructure(r.book, { ...byDates, groupDates: undefined } as StructureOp)
    if (!plain.ok) throw new Error(plain.message)
    expect(plain.book.pivotCaches![0].fields.map((f) => f.name)).toEqual(['Date', 'Amount'])
    const again = editStructure(plain.book, byDates)
    if (!again.ok) throw new Error(again.message)
    expect(again.book.pivotCaches![0].fields).toHaveLength(2)
  })

  it('without xlcalc/pivot a field of dates to group is refused; a wrong flag throws', () => {
    const r = editStructure(book(), add)
    if (!r.ok) throw new Error(r.message)
    const edit = pivotCacheEdit()
    setPivotCacheEdit(null as never)
    try {
      expect(editStructure(r.book, byDates)).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('registerPivot') })
    } finally {
      setPivotCacheEdit(edit!)
    }
    expect(() => editStructure(r.book, { ...byDates, groupDates: 'yes' } as never)).toThrow(InputError)
  })
})

describe.each(KINDS)('groupDates in %s', (kind) => {
  it.each(['closures', 'eval'] as const)('one step of undo (execution %s)', async (execution) => {
    const s = await (await opener(kind))(book(), { clock, execution })
    await s.editStructure(add)
    expect(await s.editStructure(byDates)).toMatchObject({ ok: true })
    expect(await s.getValues('P', 'A4:B6')).toEqual([['2024', 30], ['2025', 5], ['Grand Total', 35]])
    await s.undo()
    expect((await s.getBook()).pivotCaches![0].fields).toHaveLength(2)
    expect(await s.getValues('P', 'A3:A4')).toEqual([['Sum of Amount'], [35]])
    await s.redo()
    expect(await s.getValues('P', 'A4:A5')).toEqual([['2024'], ['2025']])
    await s.close()
  })
})

for (const [set, cases] of [['com-dates', comDates()], ['com-values', comValues()]] as const) {
  const answers = JSON.parse(readFileSync(new URL(`./excel/win/pivot/${set}.json`, import.meta.url), 'utf8')).answers
  describe(`${set} against Excel for Windows`, () => {
    // a PivotTable of a range has no CubeField: Excel refused the step, which the driver plays as Orientation
    for (const c of cases.filter((x) => x.id !== 'cube-field-rows'))
      it(c.id, () => {
        const [p] = drive(c)
        const want = answers[c.id].phases[0].pivots.PT1
        expect(p.errors).toEqual([])
        expect(rounded(shown(p.book, 'P', want.area))).toEqual(rounded(want.values))
      })
  })
}
