/**
 * `Session.getViewport`, a window for a grid: empty cells with a style of their row, their column or both; a row with a
 * style alone over a column with a format (the cell shows the column's format); built-in formats as the locale's codes;
 * spills that reach into the window, a large one's unstored rest too; merged areas on its edge; the tables that meet it
 * and the frozen panes; the widths and heights of its lines; formulas in the book's locale inside a table; `maxCells`; styles by number, with the objects of those the
 * window shows; the cells `Session.getCells` gives, each with the same value and look. In a synchronous session, in the same thread
 * and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, createSession, type BookInput, type CellView, type Execution, type ViewportCell } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { randomBook } from './seeded-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { spillBook } from './session-spill-edits.js'
import './locales.js'

afterAll(closeCalcs)

const BOLD = { font: { bold: true } }
const FILL = { fill: { color: '#FFFF00' } }
const ITALIC = { font: { italic: true } }

/** Rows and columns with styles and formats, sized and hidden lines, merged areas, a built-in format, a note and a link. */
const lines = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  styles: [BOLD, FILL, ITALIC],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 1 }, B1: { value: 2 }, E1: { value: 46096, numberFormat: 14 },
        D3: { value: 5 }, D4: { value: 5 }, B4: { value: 0.5, numberFormat: '0%', style: 2 },
        B6: { note: { text: 'see' } }, B7: { link: { url: 'https://example.com/' } },
      },
      rows: { 3: { style: 0 }, 4: { style: 0, numberFormat: '0.0' }, 6: { height: 30 }, 7: { hidden: true }, 8: { height: 0 } },
      columns: { C: { style: 1 }, D: { numberFormat: '0.000' }, E: { width: 20 }, F: { width: 5, hidden: true } },
      merges: ['C5:D6', 'F1:G1', 'H10:I11'],
    },
  ],
})

/** A spill from outside the window and a whole column's spill, whose cells past its data are not stored. */
const spills = (): BookInput => ({
  sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(5,2)' }, D1: { value: 1 }, D2: { value: 2 }, E1: { formula: 'D:D*2' } } }],
})

/** A table whose calculated column refers to its own row, and formulas outside it. */
const orders = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 'Item' }, B1: { value: 'Qty' }, C1: { value: 'Price' }, D1: { value: 'Total' },
        A2: { value: 'tea' }, B2: { value: 2 }, C2: { value: 10 }, D2: { formula: '[@Qty]*[@Price]' },
        A3: { value: 'coffee' }, B3: { value: 3 }, C3: { value: 20 }, D3: { formula: '[@Qty]*[@Price]' },
        F1: { formula: 'SUM(Orders[Total])' }, F2: { formula: 'ROUND(B2*1.5,0)' }, F3: { formula: 'SUM(' },
      },
      tables: [{ name: 'Orders', range: 'A1:D3', columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Price' }, { name: 'Total', formula: '[@Qty]*[@Price]' }] }],
    },
  ],
})

/** A window cell as `Session.getCells` gives it: the conditional formats' look, the cell's own format. */
function asView(w: ViewportCell, own: CellView['numberFormat']): CellView {
  const { effectiveStyle: _e, numberFormat: _n, numberFormatSource: _s, appearance, formulaLocal: _f, issue: _i, ...rest } = w
  const out: CellView = { ...rest }
  if (own !== undefined) out.numberFormat = own
  if (appearance) out.appearance = appearance
  return out
}

describe.each(['closures', 'eval'] as const)('a window for a grid, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it("empty cells with a style of their row, their column or both (the row's); a row's style over a column's format", async () => {
      const s = await (await opener(kind))(lines(), { clock: CLOCK, execution })
      const v = await s.getViewport('S', 'A1:G8')
      expect(v.cells.A3).toEqual({ value: null, text: '', effectiveStyle: 0 })
      expect(v.cells.C1).toEqual({ value: null, text: '', effectiveStyle: 1 })
      expect(v.cells.C3).toEqual({ value: null, text: '', effectiveStyle: 0 })
      expect(v.cells.C4).toEqual({ value: null, text: '', effectiveStyle: 0, numberFormat: '0.0', numberFormatSource: 'row' })
      // a row with a style alone: the cell takes its style, and the column's format
      expect(v.cells.D3).toEqual({ value: 5, text: '5.000', effectiveStyle: 0, numberFormat: '0.000', numberFormatSource: 'column' })
      expect(v.cells.D4).toEqual({ value: 5, text: '5.0', effectiveStyle: 0, numberFormat: '0.0', numberFormatSource: 'row' })
      expect(v.cells.B4).toEqual({ value: 0.5, text: '50%', style: 2, effectiveStyle: 2, numberFormat: '0%', numberFormatSource: 'cell' })
      expect(v.cells.E1).toEqual({ value: 46096, text: '3/15/2026', numberFormat: 'm/d/yyyy', numberFormatSource: 'cell' })
      expect(v.cells.B6).toEqual({ value: null, text: '', note: { text: 'see' } })
      expect(v.cells.B7).toEqual({ value: null, text: '', link: { url: 'https://example.com/' } })
      // nothing to show: no style, a column's format alone
      for (const a of ['B2', 'D5', 'E2', 'G2']) expect(v.cells[a], a).toBeUndefined()
      expect(Object.keys(v.cells).slice(0, 4)).toEqual(['A1', 'B1', 'C1', 'E1'])
      expect(v.columns).toEqual([{ width: 8.83203125 }, { width: 8.83203125 }, { width: 8.83203125 }, { width: 8.83203125 }, { width: 20 }, { width: 5, hidden: true }, { width: 8.83203125 }])
      expect(v.rows).toEqual([{ height: 15 }, { height: 15 }, { height: 15 }, { height: 15 }, { height: 15 }, { height: 30 }, { height: 15, hidden: true }, { height: 0, hidden: true }])
      expect(v.merges).toEqual(['C5:D6', 'F1:G1'])
      expect(v).toMatchObject({ range: 'A1:G8', truncated: false, usedRange: { rows: 11, columns: 9 } })
      // the styles the window shows by index, the answer's own copies; a window without a style has none of it
      expect(v.styles).toEqual({ 0: BOLD, 1: FILL, 2: ITALIC })
      v.styles[0].font!.bold = false
      expect((await s.getLayout()).styles![0]).toEqual(BOLD)
      expect((await s.getViewport('S', 'A1:C2')).styles).toEqual({ 1: FILL })
      expect((await s.getViewport('S', 'A1:B2')).styles).toEqual({})
      await s.close()
    })

    it("built-in formats as the codes of the book's locale; without the display, no text", async () => {
      const open = await opener(kind)
      const ru = await open(lines('ru-RU'), { clock: CLOCK, execution })
      expect((await ru.getViewport('S', 'E1')).cells.E1).toEqual({ value: 46096, text: '15.03.2026', numberFormat: 'dd.mm.yyyy', numberFormatSource: 'cell' })
      const bare = await open(lines(), { clock: CLOCK, execution, display: false })
      expect((await bare.getViewport('S', 'C3:D3')).cells).toEqual({ C3: { value: null, effectiveStyle: 0 }, D3: { value: 5, effectiveStyle: 0, numberFormat: '0.000', numberFormatSource: 'column' } })
      await ru.close()
      await bare.close()
    })

    it("spills in the window: from a formula outside it, and a whole column's unstored rest", async () => {
      const s = await (await opener(kind))(spills(), { clock: CLOCK, execution })
      const v = await s.getViewport('S', 'B3:E6')
      expect(v.cells.B3).toEqual({ value: 6, text: '6', spillOf: 'A1' })
      expect(v.cells.B5).toEqual({ value: 10, text: '10', spillOf: 'A1' })
      expect(v.cells.E6).toEqual({ value: 0, text: '0', spillOf: 'E1' })
      expect(Object.keys(v.cells)).toEqual(['B3', 'E3', 'B4', 'E4', 'B5', 'E5', 'E6'])
      expect((await s.getViewport('S', 'E1')).cells.E1).toEqual({ value: 2, text: '2', formula: 'D:D*2', spill: 'E1:E1048576', spillFill: 0 })
      expect((await s.getViewport('S', 'C1:C9')).usedRange).toEqual({ rows: 1_048_576, columns: 5 })
      await s.close()
    })

    it('merged areas that meet the window, on its edge too, whole', async () => {
      const s = await (await opener(kind))(lines(), { clock: CLOCK, execution })
      expect((await s.getViewport('S', 'A1:C5')).merges).toEqual(['C5:D6'])
      expect((await s.getViewport('S', 'D6:E9')).merges).toEqual(['C5:D6'])
      expect((await s.getViewport('S', 'G1:H9')).merges).toEqual(['F1:G1'])
      expect((await s.getViewport('S', 'J12:K20')).merges).toEqual([])
      await s.close()
    })

    it("the tables that meet the window, whole, in the sheet's order, and the sheet's frozen panes", async () => {
      const book: BookInput = {
        sheets: [
          {
            name: 'S',
            cells: { A1: { value: 'Item' }, B1: { value: 'Qty' }, A2: { value: 'tea' }, B2: { value: 2 }, A3: { value: 'Total' }, E5: { value: 1 }, E6: { value: 2 } },
            tables: [
              { name: 'T', range: 'A1:B3', totals: true, columns: [{ name: 'Item', totals: { label: 'Total' } }, { name: 'Qty' }] },
              { name: 'Bare', range: 'E5:E6', headers: false, columns: [{ name: 'Column1' }] },
            ],
            frozen: { rows: 1 },
          },
          { name: 'Other', cells: {} },
        ],
      }
      const s = await (await opener(kind))(book, { clock: CLOCK, execution })
      const [t, bare] = [{ name: 'T', range: 'A1:B3', headers: true, totals: true }, { name: 'Bare', range: 'E5:E6', headers: false, totals: false }]
      expect(await s.getViewport('S', 'B3:C4')).toMatchObject({ tables: [t], frozen: { rows: 1, columns: 0 } })
      expect((await s.getViewport('S', 'D6:F9')).tables).toEqual([bare])
      expect((await s.getViewport('S', 'B2:E5')).tables).toEqual([t, bare])
      expect((await s.getViewport('S', 'C1:D4')).tables).toEqual([])
      expect(await s.getViewport('Other', 'A1:E6')).toMatchObject({ tables: [], frozen: { rows: 0, columns: 0 } })
      // a change of the panes and a structure operation: the window as the layout has them now
      await s.setLayout([{ sheet: 'S', frozen: { rows: 2, columns: 1 } }])
      await s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
      const [v, layout] = [await s.getViewport('S', 'A1:F9'), (await s.getLayout()).sheets[0]]
      expect(v.tables).toEqual([{ ...t, range: 'A2:B4' }, { ...bare, range: 'E6:E7' }])
      expect(v.frozen).toEqual({ rows: layout.frozen?.rows ?? 0, columns: layout.frozen?.columns ?? 0 })
      await s.close()
    })

    it("formulas in the book's locale inside a table and outside it, only when asked", async () => {
      const open = await opener(kind)
      const ru = await open(orders('ru-RU'), { clock: CLOCK, execution })
      const v = await ru.getViewport('S', 'A1:F3', { formulaLocal: true })
      expect(v.cells.D2).toMatchObject({ formula: 'Orders[[#This Row],[Qty]]*Orders[[#This Row],[Price]]', formulaLocal: '[@Qty]*[@Price]' })
      expect(v.cells.F1.formulaLocal).toBe('СУММ(Orders[Total])')
      expect(v.cells.F2.formulaLocal).toBe('ОКРУГЛ(B2*1,5;0)')
      // a formula that does not read, as far as toLocal translates it
      expect(v.cells.F3).toMatchObject({ formula: 'SUM(', formulaLocal: 'СУММ(' })
      expect(Object.values((await ru.getViewport('S', 'A1:F3')).cells).some((c) => 'formulaLocal' in c)).toBe(false)
      expect((await ru.getViewport('S', 'D2')).cells.D2.formulaLocal).toBeUndefined()
      const en = await open(orders(), { clock: CLOCK, execution })
      const w = await en.getViewport('S', 'D3:F3', { formulaLocal: true })
      expect([w.cells.D3.formulaLocal, w.cells.F3.formulaLocal]).toEqual(['[@Qty]*[@Price]', 'SUM('])
      expect((await en.getViewport('S', 'F1', { formulaLocal: true })).cells.F1.formulaLocal).toBe('SUM(Orders[Total])')
      await ru.close()
      await en.close()
    })

    it('a window over maxCells gives its first rows that fit, and says so', async () => {
      const s = await (await opener(kind))(lines(), { clock: CLOCK, execution })
      const big = await s.getViewport('S', 'A1:J3000')
      expect([big.range, big.truncated, big.rows.length, big.columns.length]).toEqual(['A1:J2000', true, 2000, 10])
      const whole = await s.getViewport('S', 'C:C')
      expect([whole.range, whole.truncated, whole.rows.length, Object.keys(whole.cells).length]).toEqual(['C1:C20000', true, 20000, 20000])
      expect(await s.getViewport('S', 'A1:J9', { maxCells: 25 })).toMatchObject({ range: 'A1:J2', truncated: true })
      expect(await s.getViewport('S', 'A1:J9', { maxCells: 7 })).toMatchObject({ range: 'A1:G1', truncated: true })
      expect(await s.getViewport('S', 'A1:J2', { maxCells: 20 })).toMatchObject({ range: 'A1:J2', truncated: false })
      expect(await s.getViewport('S', 'B2')).toMatchObject({ range: 'B2', truncated: false, cells: {}, columns: [{ width: 8.83203125 }], rows: [{ height: 15 }] })
      await s.close()
    })

    it('the cells Session.getCells gives, with the same values and looks; the others with a style of their line or spilled', async () => {
      const open = await opener(kind)
      const cond = lines()
      cond.sheets[0].conditionalFormats = [
        { type: 'cellIs', ranges: ['A1:D9'], operator: 'greaterThan', formulas: ['1'], style: { font: { bold: true }, numberFormat: '0.0' } },
        { type: 'containsBlanks', ranges: ['H1:H3'], style: { fill: { color: '#FFC7CE' } } },
        { type: 'dataBar', ranges: ['A1:B1'], min: { type: 'min' }, max: { type: 'max' }, color: '#638EC6' },
      ]
      const books: [string, BookInput][] = [['lines', cond], ['spills', spills()], ['orders', orders('ru-RU')], ['the spill book', spillBook()], ['a seeded book', randomBook(5)]]
      for (const [name, book] of books) {
        const s = await open(book, { clock: CLOCK, execution })
        const b = await s.getBook()
        for (const sheet of b.sheets)
          for (const range of ['A1:J12', 'B2:D4', 'H1:H3', 'A5:Z40']) {
            const [cells, v] = [await s.getCells(sheet.name, range), await s.getViewport(sheet.name, range)]
            for (const [a, c] of Object.entries(cells)) expect(asView(v.cells[a], c.numberFormat), `${name}: ${sheet.name}!${a}`).toEqual(c)
            for (const [a, w] of Object.entries(v.cells)) if (!(a in cells)) expect(w.effectiveStyle !== undefined || w.spillOf !== undefined, `${name}: ${sheet.name}!${a}`).toBe(true)
          }
        await s.close()
      }
      const s = await open(cond, { clock: CLOCK, execution })
      const v = await s.getViewport('S', 'D3:H3')
      expect(v.cells.D3).toEqual({ value: 5, text: '5.0', effectiveStyle: 0, numberFormat: '0.000', numberFormatSource: 'column', appearance: { style: { font: { bold: true }, numberFormat: '0.0' } } })
      expect(v.cells.H3).toEqual({ value: null, text: '', effectiveStyle: 0, appearance: { style: { fill: { color: '#FFC7CE' } } } })
      await s.close()
    }, 60_000)

    it('errors with their reason and origin, only when asked', async () => {
      const s = await (await opener(kind))(orders(), { clock: CLOCK, execution })
      await s.setCells([{ sheet: 'S', cell: 'C2', input: { formula: '1/0' } }])
      const v = await s.getViewport('S', 'C1:F3', { diagnostics: true })
      expect(v.cells.C2.issue).toEqual({ error: 'DIV0', reason: 'division by zero', origin: { sheet: 'S', cell: 'C2' } })
      expect(v.cells.F1.issue).toEqual({ error: 'DIV0', reason: "division by zero, from 'S'!D2", origin: { sheet: 'S', cell: 'C2' } })
      expect(v.cells.F3.issue).toMatchObject({ error: 'ERROR', check: 'syntax', at: 4, origin: { sheet: 'S', cell: 'F3' } })
      expect(v.cells.C3.issue).toBeUndefined()
      expect((await s.getViewport('S', 'C2:F3')).cells.C2.issue).toBeUndefined()
      await s.close()
    })
  })
})

describe('a window for a grid', () => {
  it('the sheet, the range and the options are checked as input', () => {
    const s = createSession(lines(), { clock: CLOCK })
    expect(() => s.getViewport('Nope', 'A1')).toThrow(InputError)
    expect(() => s.getViewport('S', 'A0')).toThrow('range: expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got "A0"')
    expect(() => s.getViewport('S', 'A1', { maxCells: 0 })).toThrow('opts.maxCells')
    expect(() => s.getViewport('S', 'A1', { maxcells: 5 } as never)).toThrow(InputError)
    expect(() => s.getViewport('S', 'A1', { formulaLocal: 1 } as never)).toThrow('opts.formulaLocal')
    expect(s.getViewport('s', 'b2:a1'.toUpperCase()).range).toBe('A1:B2')
  })
})
