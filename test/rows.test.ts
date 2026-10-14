/**
 * Row heights and wrapped entries through the API: `fitRows` over a computed book, the layout edit `{ rows, fit: true }`
 * in a session (both execution modes, in the same thread and in a thread), text with a line break typed, entered into a
 * selection and pasted, each in one step of the undo, the row's `bestFit` in xlsx, and the checks of their input. The
 * rules themselves are checked against Excel in `test/rows-excel.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, fitRows, InputError, type BookInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { closedAfter } from './fixtures/pool.js'
import { ASYNC_CALCS } from './session-kinds.js'

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }
const WRAP = { alignment: { wrap: true } }

const book = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 'Ag', style: 0 },
        A2: { value: 'a\nb\nc', style: 1 },
        B2: { value: 'Ag', style: 2 },
        A3: { value: 'a\nb', style: 1 },
        A5: { value: 'merged\ndown', style: 1 },
        A7: { value: 'The quick brown fox jumps over the lazy dog.', style: 1 },
      },
      rows: { 3: { height: 30 }, 4: { hidden: true, height: 40 } },
      merges: ['A5:A6'],
    },
  ],
  styles: [{ font: { size: 16 } }, WRAP, { font: { size: 24 } }],
})

/**
 * The heights AutoFit gives rows 1 to 7: a line of 16 points; three wrapped lines of Calibri 11 (44 points above the
 * line, 4 below) beside a line of 24 points (24 above, 7 below); two wrapped lines; the default for an empty row and a
 * cell merged down; five lines of a sentence in 46 points.
 */
const FITTED = { 1: 21, 2: 51, 3: 32, '4:6': null, 7: 80 }

describe('fitRows', () => {
  it('fits the rows of a computed book by their cells, fonts and wrapping, in spans of one height', () => {
    const b = book()
    expect(fitRows(b, compute(b, { clock }), 'S', '1:7')).toEqual(FITTED)
    expect(fitRows(b, compute(b, { clock }), 's', '4:9')).toEqual({ '4:6': null, 7: 80, '8:9': null })
  })
  it('stops at 409.5 points, the highest row', () => {
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: Array(40).fill('x').join('\n'), style: 0 } } }], styles: [WRAP] }
    expect(fitRows(b, compute(b, { clock }), 'S', '1')).toEqual({ 1: 409.5 })
  })
  it('refuses an unknown sheet, rows that are none and a result without the sheet', () => {
    const b = book()
    const r = compute(b, { clock })
    expect(() => fitRows(b, r, 'T', '1')).toThrow(InputError)
    expect(() => fitRows(b, r, 'S', 'A1')).toThrow(/rows such as/)
    expect(() => fitRows(b, { ...r, sheets: [] }, 'S', '1')).toThrow(/no sheet/)
  })
})

describe.each(['closures', 'eval'] as const)('rows in a session, execution %s', (execution) => {
  it('AutoFit sets the heights as Excel fitted them, shows a hidden row, drops a height set by hand, and undo takes it back', () => {
    const s = createSession(book(), { clock, execution })
    const r = s.setLayout([{ sheet: 'S', rows: '1:7', fit: true }])
    expect(r.layout).toEqual(['S'])
    expect(s.getLayout().sheets[0].rows).toEqual({ 1: { height: 21, bestFit: true }, 2: { height: 51, bestFit: true }, 3: { height: 32, bestFit: true }, 7: { height: 80, bestFit: true } })
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    s.undo()
    expect(s.getLayout().sheets[0].rows).toEqual(book().sheets[0].rows)
  })
  it('fits a row from the layout the edits before it leave', () => {
    const s = createSession(book(), { clock, execution })
    s.setLayout([
      { sheet: 'S', columns: 'A', set: { width: 40 } },
      { sheet: 'S', rows: '7', fit: true },
    ])
    expect(s.getLayout().sheets[0].rows?.[7]).toEqual({ height: 16, bestFit: true })
  })
  it('wraps a cell typed a line break into, by a style of the table, and fits its row, in one step', () => {
    const s = createSession(book(), { clock, execution })
    const r = s.setCellText('S', 'C1', 'x\ny')
    expect(r.ok && r.layout).toEqual(['S'])
    const b = s.getBook()
    // the table has the style that wraps
    expect(b.styles).toEqual(book().styles)
    expect(b.sheets[0].cells.C1).toEqual({ value: 'x\ny', style: 1 })
    // two wrapped lines of Calibri 11 (28 points above the line) beside a line of 16 points (5 below)
    expect(b.sheets[0].rows?.[1]).toEqual({ height: 33, bestFit: true })
    s.undo()
    expect(s.getBook().sheets[0].cells.C1).toBeUndefined()
    expect(s.getBook().sheets[0].rows?.[1]).toBeUndefined()
  })
  it('adds a style that wraps the cell’s own, keeps a height set by hand, and fits a fitted row again as text grows and shrinks', () => {
    const s = createSession(book(), { clock, execution })
    const grown = [...book().styles!, { font: { size: 16 }, alignment: { wrap: true } }]
    expect(s.setCellText('S', 'A1', 'a\nb')).toMatchObject({ ok: true, styles: grown })
    expect(s.getBook().styles).toEqual(grown)
    expect(s.getBook().sheets[0].cells.A1.style).toBe(3)
    expect(s.getLayout().sheets[0].rows?.[1]).toEqual({ height: 44, bestFit: true })
    s.setCellText('S', 'A1', 'a')
    expect(s.getLayout().sheets[0].rows?.[1]).toEqual({ height: 22, bestFit: true })
    s.setCellText('S', 'B3', 'p\nq\nr')
    expect(s.getLayout().sheets[0].rows?.[3]).toEqual({ height: 30 })
    s.undo()
    s.undo()
    s.undo()
    expect(s.getBook()).toEqual(createSession(book(), { clock, execution }).getBook())
  })
  it('wraps the cells of setCellTexts and of an entry into a selection, and their rows fit', () => {
    const s = createSession(book(), { clock, execution })
    s.setCellTexts([
      { sheet: 'S', cell: 'D8', text: 'a\nb' },
      { sheet: 'S', cell: 'D9', text: 'plain' },
    ])
    expect(s.getLayout().sheets[0].rows).toMatchObject({ 8: { height: 32, bestFit: true } })
    expect(s.getLayout().sheets[0].rows?.[9]).toBeUndefined()
    s.fillCellText('S', 'E10:F11', 'a\nb\nc')
    const b = s.getBook()
    expect([b.sheets[0].cells.E10, b.sheets[0].cells.F11].map((c) => b.styles![c.style!].alignment?.wrap)).toEqual([true, true])
    expect(s.getLayout().sheets[0].rows).toMatchObject({ 10: { height: 48, bestFit: true }, 11: { height: 48, bestFit: true } })
    expect(s.history.undo).toBe(2)
  })
})

describe('the input of AutoFit of rows', () => {
  it('takes bestFit and 409.5 points on a row, and refuses more', () => {
    const s = createSession(book(), { clock })
    s.setLayout([{ sheet: 'S', rows: '9', set: { height: 409.5, bestFit: true } }])
    expect(s.getLayout().sheets[0].rows?.[9]).toEqual({ height: 409.5, bestFit: true })
    expect(() => s.setLayout([{ sheet: 'S', rows: '9', set: { height: 410 } }])).toThrow(InputError)
    expect(() => s.setLayout([{ sheet: 'S', rows: 'A', fit: true } as never])).toThrow(/row number/)
  })
})

const { keep } = closedAfter()

describe.each(ASYNC_CALCS)('rows in an asynchronous session, %s', (_, make) => {
  it('AutoFit goes through setLayout and a typed line break wraps its cell and fits its row', async () => {
    const calc = await make()
    keep(calc)
    const s = await calc.createSession(book(), { clock })
    await s.setLayout([{ sheet: 'S', rows: '1:3', fit: true }])
    expect((await s.getLayout()).sheets[0].rows).toEqual({ 1: { height: 21, bestFit: true }, 2: { height: 51, bestFit: true }, 3: { height: 32, bestFit: true }, 4: { hidden: true, height: 40 } })
    await s.setCellText('S', 'C12', 'a\nb')
    expect((await s.getLayout()).sheets[0].rows?.[12]).toEqual({ height: 32, bestFit: true })
    await s.undo()
    expect((await s.getLayout()).sheets[0].rows?.[12]).toBeUndefined()
  })
})

describe('bestFit of a row in xlsx', () => {
  it('is written without customHeight and read back', async () => {
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } }, rows: { 1: { height: 32, bestFit: true }, 2: { height: 20 } } }] }
    const { data } = await writeXlsx({ book: b })
    expect((await readXlsx(data)).book.sheets[0].rows).toEqual({ 1: { height: 32, bestFit: true }, 2: { height: 20 } })
  })
})
