/**
 * Outlines in the model: outline levels and collapsed marks of rows and columns, and where summaries are; checked as
 * input, in `inputHash` only when a book has them, moved by structure operations (the rules measured in
 * `test/excel/valid/`), a collapsed group's hidden rows counted out by SUBTOTAL 109, set through a session's layout, and
 * written and read in xlsx, in both execution modes.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, openSession, restructure, type BookInput, type Execution, type StructureOp } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

/** Rows 2:4 grouped and collapsed under row 5 (its total), row 3 grouped twice; columns B:C grouped, D their summary. */
const book = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: { A2: { value: 1 }, A3: { value: 2 }, A4: { value: 4 }, A5: { formula: 'SUBTOTAL(109,A2:A4)' }, A6: { formula: 'SUBTOTAL(9,A2:A4)' } },
      rows: { 2: { outlineLevel: 1, hidden: true }, 3: { outlineLevel: 2, hidden: true }, 4: { outlineLevel: 1, hidden: true }, 5: { collapsed: true } },
      columns: { 'B:C': { outlineLevel: 1 } },
    },
  ],
})

const rows = (b: BookInput) => b.sheets[0].rows
const after = (op: StructureOp, b = book()) => {
  const r = restructure(b, op)
  if (!r.ok) throw new Error(r.message)
  return r.book
}

describe('outlines as input', () => {
  it('levels from 1 to 7 and booleans, else InputError with the path', () => {
    const fails = (sheet: object, message: string) => expect(() => compute({ sheets: [{ name: 'S', cells: {}, ...sheet }] })).toThrow(message)
    fails({ rows: { 2: { outlineLevel: 8 } } }, 'rows["2"].outlineLevel: expected an integer from 1 to 7')
    fails({ rows: { 2: { outlineLevel: 0 } } }, 'rows["2"].outlineLevel')
    fails({ columns: { B: { collapsed: 1 } } }, 'columns["B"].collapsed: expected a boolean')
    fails({ outline: { summaryBelow: 'no' } }, 'outline.summaryBelow: expected a boolean')
    fails({ outline: { below: false } }, 'unknown field "below"')
    expect(() => compute({ sheets: [{ name: 'S', cells: {}, rows: { '2:9': { outlineLevel: 7 } }, outline: { summaryBelow: false } }] })).not.toThrow()
  })

  it('a collapsed group is hidden rows: SUBTOTAL 109 leaves them out; the values hash is that of the hidden rows alone', () => {
    const r = compute(book(), { clock })
    expect([r.sheets[0].cells.A5.value, r.sheets[0].cells.A6.value]).toEqual([0, 7])
    const plain = { ...book(), sheets: [{ ...book().sheets[0], rows: { '2:4': { hidden: true } } }] }
    expect(compute(plain, { clock }).hash).toBe(r.hash)
    expect(compute(plain, { clock }).inputHash).not.toBe(r.inputHash)
    expect(compute({ ...book(), sheets: [{ ...book().sheets[0], outline: { summaryBelow: false } }] }, { clock }).inputHash).not.toBe(r.inputHash)
  })
})

describe('outlines under structure operations', () => {
  it('inserted rows take the level of the row above, shown, and open the collapsed group they go into', () => {
    expect(rows(after({ op: 'insertRows', sheet: 'S', row: 3 }))).toEqual({ 2: { outlineLevel: 1, hidden: true }, 3: { outlineLevel: 1 }, 4: { outlineLevel: 2, hidden: true }, 5: { outlineLevel: 1, hidden: true } })
    // above the group, and at the summary row: the group stays collapsed
    expect(rows(after({ op: 'insertRows', sheet: 'S', row: 1 }))).toEqual({ 3: { outlineLevel: 1, hidden: true }, 4: { outlineLevel: 2, hidden: true }, 5: { outlineLevel: 1, hidden: true }, 6: { collapsed: true } })
    expect(rows(after({ op: 'insertRows', sheet: 'S', row: 6 }))).toEqual({ 2: { outlineLevel: 1, hidden: true }, 3: { outlineLevel: 2, hidden: true }, 4: { outlineLevel: 1, hidden: true }, 5: { collapsed: true } })
    // summaries above: the row before the group loses its mark
    const above = book()
    above.sheets[0].rows = { 1: { collapsed: true }, '2:4': { outlineLevel: 1, hidden: true } }
    above.sheets[0].outline = { summaryBelow: false }
    expect(rows(after({ op: 'insertRows', sheet: 'S', row: 4 }, above))).toEqual({ '2:3': { outlineLevel: 1, hidden: true }, 4: { outlineLevel: 1 }, 5: { outlineLevel: 1, hidden: true } })
    expect(after({ op: 'insertColumns', sheet: 'S', column: 'C' }).sheets[0].columns).toEqual({ 'B:D': { outlineLevel: 1 } })
  })

  it('deleted rows take their levels along; moved and copied whole rows carry them', () => {
    expect(rows(after({ op: 'deleteRows', sheet: 'S', row: 2, count: 3 }))).toEqual({ 2: { collapsed: true } })
    expect(rows(after({ op: 'copyRange', from: { sheet: 'S', range: '2:5' }, to: { sheet: 'S', range: '10:13' } }))).toMatchObject({ 10: { outlineLevel: 1, hidden: true }, 11: { outlineLevel: 2, hidden: true }, 13: { collapsed: true } })
  })
})

describe.each(['closures', 'eval'] as Execution[])('outlines in a session (%s)', (execution) => {
  it('expanding a group through setLayout shows its rows and recomputes SUBTOTAL 109', () => {
    const s = openSession(book(), { clock, execution })
    const { changes } = s.setLayout([
      { sheet: 'S', rows: '2:4', set: { hidden: null } },
      { sheet: 'S', rows: '5', set: { collapsed: null } },
    ])
    expect(changes.map((c) => [c.cell, c.value])).toEqual([['A5', 7]])
    expect(rows(s.book())).toEqual({ 2: { outlineLevel: 1 }, 3: { outlineLevel: 2 }, 4: { outlineLevel: 1 } })
    expect(s.setLayout([{ sheet: 'S', outline: { summaryBelow: false } }]).changes).toEqual([])
    expect(s.book().sheets[0].outline).toEqual({ summaryBelow: false })
    s.setLayout([{ sheet: 'S', outline: null }])
    expect(s.book().sheets[0].outline).toBeUndefined()
    expect(() => s.setLayout([{ sheet: 'S', rows: '2', set: { outlineLevel: 9 } }])).toThrow(InputError)
    expect(compute(s.book(), { clock, execution }).hash).toBe(s.hash())
  })
})

describe('outlines in xlsx', () => {
  it('levels, collapsed marks and summaries come back as written; the deepest levels are in sheetFormatPr', async () => {
    const b = book()
    b.sheets[0].outline = { summaryBelow: false, summaryRight: true }
    b.sheets.push({ name: 'T', cells: {}, columns: { 'C:E': { outlineLevel: 3, hidden: true, width: 5 }, F: { collapsed: true } } })
    const w = await writeXlsx({ book: b }, { values: 'omit' })
    expect(w.issues).toEqual([])
    const r = await readXlsx(w.data)
    expect(r.issues).toEqual([])
    expect(r.book.sheets.map((s) => [s.rows, s.columns, s.outline])).toEqual(b.sheets.map((s) => [s.rows, s.columns, s.outline]))
    const archive = openArchive(w.data)
    expect((await archive.text('xl/worksheets/sheet1.xml'))!).toContain('<sheetPr><outlinePr summaryBelow="0" summaryRight="1"/></sheetPr>')
    expect((await archive.text('xl/worksheets/sheet1.xml'))!).toContain('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="2" outlineLevelCol="1"/>')
    expect((await archive.text('xl/worksheets/sheet2.xml'))!).toContain('<col min="6" max="6" width="8.83203125" collapsed="1"/>')
  })

  it('a filtered sheet keeps its filter mark beside the outline settings', async () => {
    const b = book()
    b.sheets[0].filtered = true
    b.sheets[0].outline = { summaryRight: false }
    const w = await writeXlsx({ book: b }, { values: 'omit' })
    expect((await openArchive(w.data).text('xl/worksheets/sheet1.xml'))!).toContain('<sheetPr filterMode="1"><outlinePr summaryRight="0"/></sheetPr>')
    const r = await readXlsx(w.data)
    expect([r.book.sheets[0].filtered, r.book.sheets[0].outline]).toEqual([true, { summaryRight: false }])
  })
})
