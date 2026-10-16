/**
 * `outlineEdit` of `xlcalc/a1`: Excel's outline commands on a sheet alone give the rows and columns `editStructure` gives
 * in a book of that sheet (en-US), and its refusals, each of the three codes; the sheet and the operation checked.
 */
import { describe, expect, it } from 'vitest'
import { editStructure, InputError, type OutlineEditOp, type SheetInput, type StructureOp } from '../src/index.js'
import { outlineEdit } from '../src/a1.js'

type Sheet = Pick<SheetInput, 'rows' | 'columns' | 'outline'>

/** Each operation in turn on `sheet`: `outlineEdit`'s answer, and the sheet `editStructure` leaves for the next. */
function run(sheet: Sheet, ops: OutlineEditOp[]): unknown[] {
  return ops.map((op) => {
    const book = { sheets: [{ name: 'S', cells: {}, ...sheet }] }
    const r = editStructure(book, { ...op, sheet: 'S' } as StructureOp)
    const got = outlineEdit(sheet, op)
    if (!r.ok) {
      expect([op, got]).toEqual([op, { ok: false, code: r.code, message: r.message }])
      return got
    }
    const { rows, columns } = r.book.sheets[0]
    expect([op, got]).toEqual([op, { ok: true, ...(rows ? { rows } : {}), ...(columns ? { columns } : {}) }])
    sheet = { ...sheet, rows, columns }
    return got
  })
}

const codes = (answers: unknown[]) => answers.map((a) => (a as { ok: boolean; code?: string }).code ?? 'ok')

/** Rows 3:10 grouped (level 1, summary row 11) and 5:7 inside (level 2, summary row 8). */
const ROWS: Sheet = { rows: { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 }, '8:10': { outlineLevel: 1 } } }

describe('outlineEdit: the commands as editStructure does them', () => {
  it('Hide and Show Detail: the inner group stays collapsed when the outer one opens; refused when shown or hidden already', () => {
    const answers = run(ROWS, [
      { op: 'hideDetail', row: 8 },
      { op: 'hideDetail', row: 4 },
      { op: 'hideDetail', row: 11 },
      { op: 'showDetail', row: 11 },
      { op: 'showDetail', row: 11 },
      { op: 'showDetail', row: 6 },
      { op: 'hideDetail', row: 6 },
    ])
    expect(codes(answers)).toEqual(['ok', 'ok', 'detail', 'ok', 'detail', 'ok', 'ok'])
    expect(answers[3]).toEqual({ ok: true, rows: { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2, hidden: true }, 8: { outlineLevel: 1, collapsed: true }, '9:10': { outlineLevel: 1 } } })
    expect(answers[4]).toEqual({ ok: false, code: 'detail', message: 'the detail of row 11 is shown already' })
    expect(answers[2]).toEqual({ ok: false, code: 'detail', message: 'the detail of row 11 is hidden already' })
  })

  it('Group, Ungroup and the level buttons; a line in no group, a group with no summary line, the deepest level', () => {
    const answers = run(ROWS, [
      { op: 'group', rows: '11:12' },
      { op: 'showLevels', rows: 2 },
      { op: 'showLevels', rows: 1 },
      { op: 'showLevels', rows: 8 },
      { op: 'ungroup', rows: '3:12' },
      { op: 'ungroup', columns: 'A:C' },
      { op: 'hideDetail', row: 20 },
      { op: 'showDetail', column: 'C' },
      ...Array.from({ length: 6 }, (): OutlineEditOp => ({ op: 'group', rows: '5:7' })),
      { op: 'group', rows: '6' },
      { op: 'ungroup', rows: '1:4' },
    ])
    expect(codes(answers)).toEqual(['ok', 'ok', 'ok', 'ok', 'ok', 'not_grouped', 'not_grouped', 'not_grouped', 'ok', 'ok', 'ok', 'ok', 'ok', 'ok', 'outline_deep', 'not_grouped'])
    expect(answers.slice(5, 8).map((a) => (a as { message: string }).message)).toEqual(['the columns A:C are in no group', 'row 20 is in no group with a summary row', 'column C is in no group with a summary column'])
    expect(answers[14]).toEqual({ ok: false, code: 'outline_deep', message: 'the rows 6 are at the deepest outline level, 7, and are not grouped further' })
  })

  it('columns, summaries above and left, groups at the sheet\'s edge, both axes at once; an axis left with no lines is out', () => {
    const sheet: Sheet = { columns: { 'B:C': { outlineLevel: 1 }, 'XFC:XFD': { outlineLevel: 1 } }, rows: { '2:3': { outlineLevel: 1 } }, outline: { summaryBelow: false, summaryRight: false } }
    const answers = run(sheet, [
      { op: 'hideDetail', column: 'A' },
      { op: 'hideDetail', row: 1 },
      { op: 'showLevels', rows: 1, columns: 1 },
      { op: 'showLevels', rows: 2, columns: 2 },
      // summaries left: XFB is the summary column of XFC:XFD
      { op: 'hideDetail', column: 'XFD' },
      { op: 'group', columns: 'XFD' },
      { op: 'ungroup', rows: '2:3' },
    ])
    expect(codes(answers)).toEqual(['ok', 'ok', 'ok', 'ok', 'ok', 'ok', 'ok'])
    expect(answers[6]).toEqual({ ok: true, columns: { 'B:C': { outlineLevel: 1 }, XFB: { collapsed: true }, XFC: { outlineLevel: 1, hidden: true }, XFD: { outlineLevel: 2, hidden: true } } })
    expect(outlineEdit({}, { op: 'group', rows: '5' })).toEqual({ ok: true, rows: { 5: { outlineLevel: 1 } } })
  })

  it('checks the sheet as outlineGroups does and the operation as editStructure does', () => {
    expect(() => outlineEdit(null as never, { op: 'group', rows: '5' })).toThrow(new InputError('sheet', 'expected a sheet or its layout, got null'))
    expect(() => outlineEdit({ rows: { '5:3': {} } }, { op: 'group', rows: '5' })).toThrow('sheet.rows["5:3"]')
    expect(() => outlineEdit({}, { op: 'insertRows', row: 2 } as never)).toThrow('op.op: expected one of group, ungroup, showDetail, hideDetail, showLevels')
    expect(() => outlineEdit({}, { op: 'group', sheet: 'S', rows: '5' } as never)).toThrow('unknown field "sheet"')
    expect(() => outlineEdit({}, { op: 'showLevels', rows: 9 })).toThrow('op.rows: expected an integer from 1 to 8, got 9')
    expect(() => outlineEdit({}, { op: 'group', rows: '5', columns: 'C' } as never)).toThrow('op: expected one of rows and columns')
    expect(() => outlineEdit({}, { op: 'group', rows: 'C' } as never)).toThrow('op.rows: expected a row number such as "5" or a span such as "5:9" within the sheet')
    // the sheet given is not changed
    const rows = structuredClone(ROWS.rows)
    outlineEdit(ROWS, { op: 'hideDetail', row: 8 })
    expect(ROWS.rows).toEqual(rows)
  })
})
