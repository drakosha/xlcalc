/**
 * Tables computed: the rectangle each structured reference takes, references to the formula's row, names over tables,
 * spills inside and into tables, the input checks; both execution modes. Excel's answers are in `test/excel/tables/`.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, type BookInput, type CellInput, type TableInput } from '../src/index.js'
import { cellIssues } from './diagnostics.js'

/** Sales A1:D5 (Region, Qty, Price, Total = Qty*Price as a calculated column) and the formulas given on S and Other. */
export function salesBook(formulas: Record<string, string>, other: Record<string, string> = {}, table: Partial<TableInput> = {}): BookInput {
  const cells: Record<string, CellInput> = {
    A1: { value: 'Region' }, B1: { value: 'Qty' }, C1: { value: 'Price' }, D1: { value: 'Total' },
    A2: { value: 'North' }, B2: { value: 2 }, C2: { value: 10 },
    A3: { value: 'South' }, B3: { value: 3 }, C3: { value: 20 },
    A4: { value: 'North' }, B4: { value: 5 }, C4: { value: 30 },
    A5: { value: 'East' }, B5: { value: 1 }, C5: { value: 40 },
  }
  for (let r = 2; r <= 5; r++) cells[`D${r}`] = { formula: '[@Qty]*[@Price]' }
  for (const [a, f] of Object.entries(formulas)) cells[a] = { formula: f }
  const columns = [{ name: 'Region' }, { name: 'Qty' }, { name: 'Price' }, { name: 'Total', formula: 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]' }]
  return {
    sheets: [
      { name: 'S', cells, tables: [{ name: 'Sales', range: 'A1:D5', columns, ...table }] },
      { name: 'Other', cells: Object.fromEntries(Object.entries(other).map(([a, f]) => [a, { formula: f }])) },
    ],
  }
}

const CLOCK = { now: '2026-01-01T00:00:00Z', seed: 1 }

/** The values of a book in both modes, which must agree. */
function values(book: BookInput) {
  const a = compute(book, { clock: CLOCK })
  const b = compute(book, { clock: CLOCK, execution: 'eval' })
  expect(b.hash).toBe(a.hash)
  return a
}

const v = (r: ReturnType<typeof compute>, sheet: number, cell: string) => r.sheets[sheet].cells[cell]?.value

describe('tables', () => {
  it('give each part of a table as Excel does (Excel 16.113 for Mac en-US)', () => {
    const r = values(
      salesBook({
        J1: 'SUM(Sales[Qty])', J2: 'Sales[@Qty]', J3: 'SUM(Sales[[Qty]:[Price]])', J4: 'ROWS(Sales[#All])', J5: 'ROWS(Sales[#Data])',
        J7: 'Sales[[#Headers],[Qty]]', J8: 'ROWS(Sales[[#Headers],[#Data]])', J9: 'Sales[#Totals]', J10: 'ROWS(Sales)', J11: 'COUNTA(Sales[[#Headers],[#Data],[Region]])',
        J12: 'SUM(Sales[[#Data],[#Totals],[Qty]])', J13: 'SUM(Sales[[#Totals],[Qty]])', J14: 'SUM(Sales[[#Headers],[Qty]:[Price]])', J15: 'Sales[@]', J16: 'ROWS(Sales[@])',
        J17: 'SUM((Sales[Qty],Sales[Price]))', J18: 'INDEX(Sales,2,2)', J19: 'ROW(Sales[#Headers])', J20: 'SUMIF(Sales[Region],"North",Sales[Total])',
        L1: 'Sales[#Headers]', L3: 'Sales[[#All],[Qty]]',
      }, { A1: 'SUM(Sales[Total])', B3: 'Sales[@Qty]', B9: 'Sales[@Qty]', A2: 'Sales[@]' }),
    )
    const s = (c: string) => v(r, 0, c)
    expect([s('J1'), s('J2'), s('J3'), s('J4'), s('J5'), s('J7'), s('J8'), s('J9'), s('J10'), s('J11')]).toEqual([11, 2, 111, 5, 4, 'Qty', 5, { error: 'REF' }, 4, 5])
    expect([s('J12'), s('J13'), s('J14'), s('J15'), s('J16'), s('J17'), s('J18'), s('J19'), s('J20')]).toEqual([11, { error: 'REF' }, 0, { error: 'VALUE' }, { error: 'VALUE' }, 111, 3, 1, 170])
    expect([s('L1'), s('M1'), s('O1'), s('L3'), s('L4'), s('L7')]).toEqual(['Region', 'Qty', 'Total', 'Qty', 2, 1])
    expect(r.sheets[0].cells.D3).toMatchObject({ value: 60, formula: 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]' })
    expect([v(r, 1, 'A1'), v(r, 1, 'B3'), v(r, 1, 'B9'), v(r, 1, 'A2'), v(r, 1, 'D2')]).toEqual([270, 3, { error: 'VALUE' }, 'North', 20])
  })

  it('a name over a table, its own row seen from the cell using it', () => {
    const book = salesBook({ J1: 'SUM(nm)', J3: 'nr', J9: 'nr', J10: 'COUNTA(nh)' }, { A4: 'nr' })
    book.names = [{ name: 'nm', ref: 'Sales[Qty]' }, { name: 'nr', ref: 'Sales[[#This Row],[Price]]' }, { name: 'nh', ref: 'Sales[#Headers]' }]
    const r = values(book)
    expect([v(r, 0, 'J1'), v(r, 0, 'J3'), v(r, 0, 'J9'), v(r, 0, 'J10'), v(r, 1, 'A4')]).toEqual([11, 20, { error: 'VALUE' }, 4, 30])
  })

  it('spill no array inside a table or into one', () => {
    const book = salesBook({ J1: 'SEQUENCE(3)', G1: 'SEQUENCE(3)' })
    book.sheets[0].cells.B2 = { formula: 'SEQUENCE(2)' }
    Object.assign(book.sheets[0].cells, { G3: { value: 'h' }, G4: { value: 1 } })
    book.sheets[0].tables!.push({ name: 'Small', range: 'G3:G4', columns: [{ name: 'h' }] })
    const r = values(book)
    expect([v(r, 0, 'J1'), v(r, 0, 'J3'), v(r, 0, 'B2'), v(r, 0, 'G1')]).toEqual([1, 3, { error: 'SPILL' }, { error: 'SPILL' }])
    expect(cellIssues(r).find((i) => i.cell === 'B2')?.reason).toMatch(/its cell is in the table Sales/)
    expect(cellIssues(r).find((i) => i.cell === 'G1')?.reason).toMatch(/it meets the table Small/)
  })

  it('without a header row or with a totals row', () => {
    const book = salesBook({ J1: 'Sales[[#Headers],[Qty]]', J2: 'ROWS(Sales[#All])', J3: 'SUM(Sales[Qty])' })
    const s = book.sheets[0]
    s.tables![0] = { ...s.tables![0], range: 'A2:D5', headers: false }
    expect([v(values(book), 0, 'J1'), v(values(book), 0, 'J2'), v(values(book), 0, 'J3')]).toEqual([{ error: 'REF' }, 4, 11])
    const t = salesBook({ J1: 'SUM(Sales[[#Totals],[Qty]])', J2: 'ROWS(Sales)', J3: 'ROWS(Sales[#All])', J4: 'SUM(Sales[[#Data],[#Totals],[Qty]])' }, {}, { range: 'A1:D6', totals: true })
    Object.assign(t.sheets[0].cells, { A6: { value: 'Total' }, B6: { formula: 'SUBTOTAL(109,Sales[Qty])' } })
    const r = values(t)
    expect([v(r, 0, 'J1'), v(r, 0, 'J2'), v(r, 0, 'J3'), v(r, 0, 'J4'), v(r, 0, 'B6')]).toEqual([11, 4, 6, 22, 11])
  })

  it('check their input', () => {
    const fails = (change: (b: BookInput) => void, message: RegExp) => {
      const b = salesBook({})
      change(b)
      expect(() => compute(b)).toThrow(InputError)
      expect(() => compute(b)).toThrow(message)
    }
    fails((b) => (b.sheets[0].tables![0].name = 'T2'), /must not look like a cell address/)
    fails((b) => (b.sheets[0].tables![0].name = 'Bad Name'), /has only letters/)
    fails((b) => (b.names = [{ name: 'sales', ref: "'S'!A1" }]), /also a defined name/)
    fails((b) => {
      b.sheets[1].cells.A1 = { value: 'x' }
      b.sheets[1].tables = [{ name: 'SALES', range: 'A1:A2', columns: [{ name: 'x' }] }]
    }, /also the name of the table Sales/)
    fails((b) => (b.sheets[0].cells.B1 = { value: 'Quantity' }), /header cell B1 holds "Quantity"/)
    fails((b) => (b.sheets[0].tables![0].columns[1].name = 'region'), /also the name of column 1/)
    fails((b) => (b.sheets[0].tables![0].range = 'A1:E5'), /has 5 columns, the table lists 4/)
    fails((b) => (b.sheets[0].tables![0].range = 'A1:D1'), /no row for data/)
    fails((b) => (b.sheets[0].merges = ['C3:D3']), /merged area C3:D3 meets the table Sales/)
    fails((b) => b.sheets[0].tables!.push({ name: 'More', range: 'D5:E6', headers: false, columns: [{ name: 'Total' }, { name: 'x' }] }), /overlaps the table Sales/)
  })
})
