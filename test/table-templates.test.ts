/**
 * The formula of a calculated column (`TableInput.columns[].formula`, the template a row inserted into the table gets)
 * under structure operations: it is rewritten as the formula of the column's first data row is, so it moves and copies
 * with its table and follows a renamed, moved or deleted sheet. Expectations come from that agreement, not from Excel.
 */
import { describe, expect, it } from 'vitest'
import { compute, copyFormula, createSession, editStructure, type BookInput, type CellInput, type StructureOp } from '../src/index.js'
import { formatAddress, parseAddress } from '../src/model.js'

const clock = { now: '2026-10-09T12:00:00Z', seed: 1 }
const EXECUTIONS = ['closures', 'eval'] as const

/** The book of the review: Sales at S!A1:B4 (Double = Qty*2 written for each row), T with 100 at A3. */
const reviewBook = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 'Qty' }, B1: { value: 'Double' }, A2: { value: 3 }, B2: { formula: 'A2*2' }, A3: { value: 7 }, B3: { formula: 'A3*2' }, A4: { value: 11 }, B4: { formula: 'A4*2' } },
      tables: [{ name: 'Sales', range: 'A1:B4', columns: [{ name: 'Qty' }, { name: 'Double', formula: 'A2*2' }] }],
    },
    { name: 'T', cells: { A3: { value: 100 } } },
  ],
})

/**
 * Sales at S!A1:B4 (Qty 3, 7, 11) whose Double column has the template `tpl` and its cells filled down from it, Rates
 * at S!D1:D2 (Rate 10), 1000 at S!C1, T with 100 at A3 and 5 at F3, and the sheets `more` after them.
 */
function salesBook(tpl: string, more: string[] = []): BookInput {
  const cells: Record<string, CellInput> = { A1: { value: 'Qty' }, B1: { value: 'Double' }, A2: { value: 3 }, A3: { value: 7 }, A4: { value: 11 }, C1: { value: 1000 }, D1: { value: 'Rate' }, D2: { value: 10 } }
  for (const r of [2, 3, 4]) cells[`B${r}`] = { formula: copyFormula(tpl, { sheet: 'S', cell: 'B2' }, { sheet: 'S', cell: `B${r}` }) }
  return {
    sheets: [
      {
        name: 'S',
        cells,
        tables: [
          { name: 'Sales', range: 'A1:B4', columns: [{ name: 'Qty' }, { name: 'Double', formula: tpl }] },
          { name: 'Rates', range: 'D1:D2', columns: [{ name: 'Rate' }] },
        ],
      },
      { name: 'T', cells: { A3: { value: 100 }, F3: { value: 5 } } },
      ...more.map((name) => ({ name, cells: { B1: { value: 1 } } })),
    ],
  }
}

function apply(book: BookInput, op: StructureOp): BookInput {
  const r = editStructure(book, op)
  if (!r.ok) throw new Error(`${op.op}: ${r.message}`)
  return r.book
}

/** Each calculated column of the book: its table's sheet, its first data cell, the template and that cell's formula. */
function calculated(book: BookInput) {
  return book.sheets.flatMap((s) =>
    (s.tables ?? []).flatMap((t) => {
      const { row, col } = parseAddress(t.range.split(':')[0])
      const first = row + (t.headers === false ? 0 : 1)
      return t.columns.flatMap((c, i) => (c.formula === undefined ? [] : [{ sheet: s.name, table: t.name, row: first, col: col + i, template: c.formula, cell: s.cells[formatAddress(first, col + i)]?.formula }]))
    }),
  )
}

const templateOf = (book: BookInput, table: string) => calculated(book).find((c) => c.table === table)?.template

/** The template of every calculated column equals the formula of its first data row. */
function expectAgreed(book: BookInput) {
  for (const c of calculated(book)) expect(c.template, `${c.table}: template against ${c.sheet}!${formatAddress(c.row, c.col)}`).toBe(c.cell)
}

/**
 * A row inserted into each table right after its first data row gets the first row's formulas copied down a row, and
 * after a value typed into it the session's values are those of the book computed afresh.
 */
function expectInsertedRows(book: BookInput, execution: (typeof EXECUTIONS)[number]) {
  for (const c of calculated(book)) {
    const s = createSession(book, { execution, clock })
    const r = s.editStructure({ op: 'insertRows', sheet: c.sheet, row: c.row + 2, count: 1 })
    if (!r.ok) throw new Error(r.message)
    const after = s.getBook()
    const sheet = after.sheets.find((x) => x.name === c.sheet)!
    const first = sheet.cells[formatAddress(c.row, c.col)].formula!
    const at = formatAddress(c.row + 1, c.col)
    expect(sheet.cells[at]?.formula, `${c.table}: the inserted ${c.sheet}!${at}`).toBe(copyFormula(first, { sheet: c.sheet, cell: formatAddress(c.row, c.col) }, { sheet: c.sheet, cell: at }, after))
    s.setCells([{ sheet: c.sheet, cell: formatAddress(c.row + 1, c.col - 1), input: { value: 5 } }])
    expect(s.getResult().hash).toBe(compute(s.getBook(), { execution, clock }).hash)
    s.close()
  }
}

describe.each(EXECUTIONS)('calculated column templates from the review, %s', (execution) => {
  it.each([
    ['copied', { op: 'copyRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'T', range: 'D5:E8' } }, 'Sales3'],
    ['moved', { op: 'moveRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'T', cell: 'D5' } }, 'Sales'],
  ] as [string, StructureOp, string][])('a table %s to another sheet gives an inserted row its own formula', (_, op, table) => {
    const book = apply(reviewBook(), op)
    expect(book.sheets[1].cells.E6.formula).toBe('D6*2')
    const s = createSession(book, { execution, clock })
    const r = s.editStructure({ op: 'insertRows', sheet: 'T', row: 7, count: 1 })
    if (!r.ok) throw new Error(r.message)
    s.setCells([{ sheet: 'T', cell: 'D7', input: { value: 5 } }])
    expect(s.getCells('T', 'E7').E7).toMatchObject({ formula: 'D7*2', value: 10 })
    expect(s.getResult().hash).toBe(compute(s.getBook(), { execution, clock }).hash)
    expect(templateOf(s.getBook(), table)).toBe('D6*2')
    s.close()
  })

  it('a renamed sheet: an inserted row refers to it by its new name', () => {
    const book = reviewBook()
    const S = book.sheets[0]
    for (const r of [2, 3, 4]) S.cells[`B${r}`] = { formula: `'S'!A${r}*2` }
    S.tables![0].columns[1].formula = "'S'!A2*2"
    const renamed = apply(book, { op: 'renameSheet', sheet: 'S', name: 'Renamed' })
    const s = createSession(renamed, { execution, clock })
    const r = s.editStructure({ op: 'insertRows', sheet: 'Renamed', row: 3, count: 1 })
    if (!r.ok) throw new Error(r.message)
    s.setCells([{ sheet: 'Renamed', cell: 'A3', input: { value: 5 } }])
    expect(s.getCells('Renamed', 'B3').B3).toMatchObject({ formula: "'Renamed'!A3*2", value: 10 })
    expect(s.getResult().hash).toBe(compute(s.getBook(), { execution, clock }).hash)
    expect(templateOf(s.getBook(), 'Sales')).toBe("'Renamed'!A2*2")
    s.close()
  })
})

const toT = (range = 'D5:E8'): StructureOp => ({ op: 'copyRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'T', range } })
const moveToT: StructureOp = { op: 'moveRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'T', cell: 'D5' } }
const copyHere: StructureOp = { op: 'copyRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'S', range: 'D10' } }
const moveHere: StructureOp = { op: 'moveRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'S', cell: 'D10' } }

/** [what, template of Sales, extra sheets, operation, the templates after it by table]. */
const CASES: [string, string, string[], StructureOp, Record<string, string>][] = [
  ['relative and absolute references, copied to another sheet', 'A2*2+$A$2+A$2+$A2', [], toT(), { Sales: 'A2*2+$A$2+A$2+$A2', Sales4: 'D6*2+$A$2+D$2+$A6' }],
  ['relative and absolute references, moved to another sheet', 'A2*2+$A$2+A$2+$A2', [], moveToT, { Sales: 'D6*2+$D$6+D$6+$D6' }],
  ['relative and absolute references, copied on the sheet', 'A2*2+$A$2+A$2+$A2', [], copyHere, { Sales: 'A2*2+$A$2+A$2+$A2', Sales4: 'D11*2+$A$2+D$2+$A11' }],
  ['relative and absolute references, moved on the sheet', 'A2*2+$A$2+A$2+$A2', [], moveHere, { Sales: 'D11*2+$D$11+D$11+$D11' }],
  ['a cell outside the range, copied to another sheet', 'A2*$C$1', [], toT(), { Sales4: 'D6*$C$1' }],
  ['a cell outside the range, moved to another sheet, gets its sheet', 'A2*$C$1', [], moveToT, { Sales: "D6*'S'!$C$1" }],
  ['a cell outside the range, moved on the sheet', 'A2*$C$1', [], moveHere, { Sales: 'D11*$C$1' }],
  ['a reference with its sheet, copied', "A2*'T'!$A$3", [], toT(), { Sales4: "D6*'T'!$A$3" }],
  ['a reference with its sheet, moved', "A2*'T'!$A$3", [], moveToT, { Sales: "D6*'T'!$A$3" }],
  ['tiled copies, each by its offset', 'A2*2', [], toT('D5:E12'), { Sales: 'A2*2', Sales4: 'D6*2', Sales5: 'D10*2' }],
  ['the row of its own table, copied', '[@Qty]*2', [], toT(), { Sales: '[@Qty]*2', Sales4: '[@Qty]*2' }],
  ['the row of its own table, moved', '[@Qty]*2', [], moveToT, { Sales: '[@Qty]*2' }],
  ['its table by name and another table, copied: the copy refers to itself', 'Sales[@Qty]*SUM(Rates[Rate])', [], toT(), { Sales: 'Sales[@Qty]*SUM(Rates[Rate])', Sales4: 'Sales4[[#This Row],[Qty]]*SUM(Rates[Rate])' }],
  ['its table by name and another table, moved', 'Sales[@Qty]*SUM(Rates[Rate])', [], moveToT, { Sales: 'Sales[@Qty]*SUM(Rates[Rate])' }],
  ['a renamed sheet it refers to', "A2*'T'!$A$3", [], { op: 'renameSheet', sheet: 'T', name: 'Data' }, { Sales: "A2*'Data'!$A$3" }],
  ['its own sheet renamed', "'S'!A2*2+A2", [], { op: 'renameSheet', sheet: 'S', name: 'Renamed' }, { Sales: "'Renamed'!A2*2+A2" }],
  ['a deleted sheet it refers to', "A2*'T'!$A$3", [], { op: 'deleteSheet', sheet: 'T' }, { Sales: 'A2*#REF!$A$3' }],
  ['a range it refers to moved, the table staying', 'A2*$C$1', [], { op: 'moveRange', from: { sheet: 'S', range: 'C1' }, to: { sheet: 'S', cell: 'F3' } }, { Sales: 'A2*$F$3' }],
  ['a range it refers to moved to another sheet', 'A2*$C$1', [], { op: 'moveRange', from: { sheet: 'S', range: 'C1' }, to: { sheet: 'T', cell: 'H2' } }, { Sales: "A2*'T'!$H$2" }],
  ['cells it refers to overwritten by a move', "A2*'T'!$F$3", [], { op: 'moveRange', from: { sheet: 'S', range: 'C1' }, to: { sheet: 'T', cell: 'F3' } }, { Sales: "A2*'T'!#REF!" }],
  ['a sheet moved out of its span across sheets', "A2+SUM('A:C'!B1)", ['A', 'B', 'C'], { op: 'moveSheet', sheet: 'A' }, { Sales: "A2+SUM('B:C'!B1)" }],
  ['an end of its span across sheets deleted', "A2+SUM('A:C'!B1)", ['A', 'B', 'C'], { op: 'deleteSheet', sheet: 'C' }, { Sales: "A2+SUM('A:B'!B1)" }],
]

describe.each(EXECUTIONS)('calculated column templates follow their cells, %s', (execution) => {
  it.each(CASES)('%s', (_, tpl, more, op, expected) => {
    const before = salesBook(tpl, more)
    expectAgreed(before)
    const after = apply(before, op)
    for (const [table, template] of Object.entries(expected)) expect(templateOf(after, table), table).toBe(template)
    expectAgreed(after)
    expectInsertedRows(after, execution)
  })

  it('references to a table of a deleted sheet are lost in a template too', () => {
    const book = salesBook('A2*2')
    const rates = book.sheets[0].tables!.pop()!
    book.sheets[1].tables = [{ ...rates, range: 'K1:K2' }]
    Object.assign(book.sheets[1].cells, { K1: { value: 'Rate' }, K2: { value: 10 } })
    for (const r of [2, 3, 4]) book.sheets[0].cells[`B${r}`] = { formula: `A${r}*SUM(Rates[Rate])` }
    book.sheets[0].tables![0].columns[1].formula = 'A2*SUM(Rates[Rate])'
    const after = apply(book, { op: 'deleteSheet', sheet: 'T' })
    expect(templateOf(after, 'Sales')).toBe('A2*SUM(#REF!)')
    expectAgreed(after)
    expectInsertedRows(after, execution)
  })
})
