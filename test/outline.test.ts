/**
 * Outlines in the model: outline levels and collapsed marks of rows and columns, and where summaries are; checked as
 * input, in `inputHash` only when a book has them, the groups Excel draws (`outlineGroups`), moved by structure
 * operations (the rules measured in `test/excel/valid/`), a collapsed group's hidden rows counted out by SUBTOTAL 109,
 * set through a session's layout, and written and read in xlsx with the marks Excel keeps, in both execution modes.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, createSession, editStructure, outlineGroups, type BookInput, type Execution, type StructureOp } from '../src/index.js'
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
  const r = editStructure(b, op)
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

describe('outlineGroups', () => {
  it("the groups Excel draws: the longest runs at each level, their summary lines, collapsed as Excel keeps the marks", () => {
    // the grid's example: 2:10 of level 1, 2:4 and the collapsed 6:9 of level 2
    const rows = { '2:4': { outlineLevel: 2 }, 5: { outlineLevel: 1 }, '6:9': { outlineLevel: 2, hidden: true }, 10: { outlineLevel: 1, collapsed: true } }
    expect(outlineGroups({ rows })).toEqual({
      rows: [
        { first: 2, last: 10, level: 1, summary: 11, collapsed: false },
        { first: 2, last: 4, level: 2, summary: 5, collapsed: false },
        { first: 6, last: 9, level: 2, summary: 10, collapsed: true },
      ],
      columns: [],
    })
    // summaries above: 6:9's summary row 5 holds no mark, but Excel marks it, its rows all hidden; 2:10 has none, past row 1
    expect(outlineGroups({ rows, outline: { summaryBelow: false } }).rows.map((g) => [g.first, g.summary, g.collapsed])).toEqual([[2, 1, false], [2, 1, false], [6, 5, true]])
    // rows hidden with no mark (setLayout), a mark over rows shown; columns at the sheet's right edge have no summary column
    expect(outlineGroups({ rows: { '3:4': { outlineLevel: 1, hidden: true }, 6: { outlineLevel: 1 }, 7: { collapsed: true } }, columns: { 'XFC:XFD': { outlineLevel: 1, width: 0 } } })).toEqual({
      rows: [
        { first: 3, last: 4, level: 1, summary: 5, collapsed: true },
        { first: 6, last: 6, level: 1, summary: 7, collapsed: false },
      ],
      columns: [{ first: 16_383, last: 16_384, level: 1, summary: null, collapsed: true }],
    })
    expect(() => outlineGroups({ rows: { '5:3': {} } })).toThrow(new InputError('sheet.rows["5:3"]', 'expected an object under a span of rows ("5", "5:9")'))
    expect(() => outlineGroups({ columns: { 'B:D': {}, C: {} } })).toThrow(new InputError('sheet.columns', 'two spans overlap'))
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
    const s = createSession(book(), { clock, execution })
    const { changes } = s.setLayout([
      { sheet: 'S', rows: '2:4', set: { hidden: null } },
      { sheet: 'S', rows: '5', set: { collapsed: null } },
    ])
    expect(changes.map((c) => [c.cell, c.value])).toEqual([['A5', 7]])
    expect(rows(s.getBook())).toEqual({ 2: { outlineLevel: 1 }, 3: { outlineLevel: 2 }, 4: { outlineLevel: 1 } })
    expect(s.setLayout([{ sheet: 'S', outline: { summaryBelow: false } }]).changes).toEqual([])
    expect(s.getBook().sheets[0].outline).toEqual({ summaryBelow: false })
    s.setLayout([{ sheet: 'S', outline: null }])
    expect(s.getBook().sheets[0].outline).toBeUndefined()
    expect(() => s.setLayout([{ sheet: 'S', rows: '2', set: { outlineLevel: 9 } }])).toThrow(InputError)
    expect(compute(s.getBook(), { clock, execution }).hash).toBe(s.getHash())
  })
})

describe('outlines in xlsx', () => {
  it('levels, collapsed marks and summaries come back as written; the deepest levels are in sheetFormatPr', async () => {
    const b = book()
    // summaries above: the mark of the group 2:4 is on row 1
    b.sheets[0].outline = { summaryBelow: false, summaryRight: true }
    b.sheets[0].rows = { 1: { collapsed: true }, 2: { outlineLevel: 1, hidden: true }, 3: { outlineLevel: 2, hidden: true }, 4: { outlineLevel: 1, hidden: true } }
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

  it('the collapsed marks are written as Excel keeps them: on a group whose lines are all hidden, not on one shown', async () => {
    const sheet = (rows: BookInput['sheets'][0]['rows']): BookInput => ({ sheets: [{ name: 'S', cells: {}, rows, columns: { 'B:C': { outlineLevel: 1, hidden: true } } }] })
    const written = async (b: BookInput) => {
      const r = await readXlsx((await writeXlsx({ book: b }, { values: 'omit' })).data)
      return [r.book.sheets[0].rows, r.book.sheets[0].columns]
    }
    // rows hidden through setLayout get the mark; a mark on rows shown is dropped; the mark of an inner group inside a hidden one stays
    expect(await written(sheet({ '2:4': { outlineLevel: 1, hidden: true }, 9: { collapsed: true }, '7:8': { outlineLevel: 1 } }))).toEqual([
      { '2:4': { outlineLevel: 1, hidden: true }, 5: { collapsed: true }, '7:8': { outlineLevel: 1 } },
      { 'B:C': { outlineLevel: 1, hidden: true }, D: { collapsed: true } },
    ])
    const inner = { '3:4': { outlineLevel: 1, hidden: true }, '5:7': { outlineLevel: 2, hidden: true }, '9:10': { outlineLevel: 1, hidden: true } }
    expect((await written(sheet({ ...inner, 8: { outlineLevel: 1, hidden: true, collapsed: true } })))[0]).toEqual({ ...inner, 8: { outlineLevel: 1, hidden: true, collapsed: true }, 11: { collapsed: true } })
    expect((await written(sheet({ ...inner, 8: { outlineLevel: 1, hidden: true } })))[0]).toEqual({ '3:4': inner['3:4'], '5:7': inner['5:7'], '8:10': { outlineLevel: 1, hidden: true }, 11: { collapsed: true } })
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
