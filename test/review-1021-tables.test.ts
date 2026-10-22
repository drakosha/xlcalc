/**
 * Tables under structure operations, from the review of 21.10: a copy of a whole table named as Excel would not name a
 * table, a fill continuing a series over header cells, rows inserted above a table at the sheet's edge. In each case the
 * book an operation gives, and a session's `getBook()`, computes, and the session's hash is a full computation's.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, registerLocale, type BookInput, type Execution, type StructureOp } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'
import { ruRU } from '../src/locales/ru-RU.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']

afterAll(closeCalcs)
registerLocale(ruRU)

/** The operation in a session of each kind: its answer, and a book the session holds that computes to the session's hash. */
function inSessions(execution: Execution, book: () => BookInput, op: StructureOp, check: (r: { ok: boolean; code?: string }, after: BookInput) => void) {
  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('answers as editStructure does and keeps a book that computes to its hash', async () => {
      const s = await open(book(), { execution, clock })
      const r = await s.editStructure(op)
      const after = await s.getBook()
      check(r, after)
      expect(compute(after, { execution, clock }).hash).toBe(await s.getHash())
      await s.close()
    })
  })
}

const tableBook = (name: string, names?: BookInput['names']): BookInput => ({
  ...(names ? { names } : {}),
  sheets: [{
    name: 'S',
    cells: { A1: { value: 'Qty' }, B1: { value: 'Price' }, A2: { value: 3 }, B2: { value: 1.5 }, A3: { value: 4 }, B3: { value: 2.5 }, D1: { formula: `SUM(${name}[Qty])` } },
    tables: [{ name, range: 'A1:B3', columns: [{ name: 'Qty' }, { name: 'Price' }] }],
  }],
})
const copyWhole: StructureOp = { op: 'copyRange', from: { sheet: 'S', range: 'A1:B3' }, to: { sheet: 'S', range: 'A10:B12' } }
const tableNames = (b: BookInput) => (b.sheets[0].tables ?? []).map((t) => t.name)

describe.each(EXECUTIONS)('a copy of a whole table (x-fix-structure-looks-1), %s', (execution) => {
  it('skips a numbered name a defined name has, at book scope or a sheet scope', () => {
    for (const names of [[{ name: 'Sales3', ref: '1' }], [{ name: 'sales3', ref: '1', scope: 'S' }, { name: 'Sales4', ref: '2' }]]) {
      const r = editStructure(tableBook('Sales', names), copyWhole)
      if (!r.ok) throw new Error(r.message)
      expect(tableNames(r.book)).toEqual(['Sales', names.length === 1 ? 'Sales4' : 'Sales5'])
      expect(compute(r.book, { execution, clock }).sheets[0].cells.D1.value).toBe(7)
    }
  })

  it('refuses a copy whose numbered name Excel would not take, as renameTable refuses it', () => {
    for (const name of ['Tab', 'Tbl', 'Rr']) {
      const r = editStructure(tableBook(name), copyWhole)
      expect(r.ok ? tableNames(r.book) : r.code, name).toBe('table_name')
    }
    // a copy of part of the table is not a table, whatever its name
    const part = editStructure(tableBook('Tab'), { op: 'copyRange', from: { sheet: 'S', range: 'A2:A3' }, to: { sheet: 'S', range: 'A10' } })
    expect(part.ok && tableNames(part.book)).toEqual(['Tab'])
  })

  inSessions(execution, () => tableBook('Tab'), copyWhole, (r, after) => {
    expect(r).toMatchObject({ ok: false, code: 'table_name' })
    expect(tableNames(after)).toEqual(['Tab'])
  })
  inSessions(execution, () => tableBook('Sales', [{ name: 'Sales3', ref: '1' }]), copyWhole, (r, after) => {
    expect(r.ok).toBe(true)
    expect(tableNames(after)).toEqual(['Sales', 'Sales4'])
  })
})

/** A table of one column at the bottom of the sheet: its header and one value above an empty last row. */
const edgeRows = (headers = true): BookInput => ({
  sheets: [{
    name: 'S',
    cells: headers ? { A1048574: { value: 'h' }, A1048575: { value: 1 } } : { A1048574: { value: 1 }, A1048575: { value: 2 } },
    tables: [{ name: 'T', range: 'A1048574:A1048576', columns: [{ name: 'h' }], ...(headers ? {} : { headers: false }) }],
  }],
})
/** A table without headers at the right edge of the sheet, its last column empty. */
const edgeColumns = (): BookInput => ({
  sheets: [{ name: 'S', cells: { XFC1: { value: 1 }, XFC2: { value: 2 } }, tables: [{ name: 'T', range: 'XFC1:XFD2', columns: [{ name: 'a' }, { name: 'b' }], headers: false }] }],
})

describe.each(EXECUTIONS)('lines inserted before a table at the sheet edge (fix-u297), %s', (execution) => {
  it.each([
    ['rows above the table', edgeRows(), { op: 'insertRows', sheet: 'S', row: 1 }, 'A1048576'],
    ['a row into the table at its last row', edgeRows(), { op: 'insertRows', sheet: 'S', row: 1048576 }, 'A1048576'],
    ['rows above a table without headers', edgeRows(false), { op: 'insertRows', sheet: 'S', row: 1 }, 'A1048576'],
    ['columns before a table without headers', edgeColumns(), { op: 'insertColumns', sheet: 'S', column: 'A' }, 'XFD2'],
  ] as [string, BookInput, StructureOp, string][])('refuses %s with off_sheet, naming the table\'s last cell', (_, book, op, cell) => {
    compute(book, { execution, clock })
    const r = editStructure(book, op)
    expect(r.ok ? r.book.sheets[0].tables : r.code).toBe('off_sheet')
    expect(!r.ok && r.message).toContain(`S!${cell}`)
  })

  it('takes rows inserted above a table that ends before the edge', () => {
    const book = edgeRows()
    book.sheets[0].tables![0].range = 'A1048574:A1048575'
    const r = editStructure(book, { op: 'insertRows', sheet: 'S', row: 1 })
    if (!r.ok) throw new Error(r.message)
    expect(r.book.sheets[0].tables![0].range).toBe('A1048575:A1048576')
    compute(r.book, { execution, clock })
  })

  inSessions(execution, edgeRows, { op: 'insertRows', sheet: 'S', row: 1 }, (r, after) => {
    expect(r).toMatchObject({ ok: false, code: 'off_sheet' })
    expect(after.sheets[0].tables![0].range).toBe('A1048574:A1048576')
  })
})

/** A table whose second header continues as a series: `Col1` over the column names `x` and `y`. */
const seriesHeaders = (): BookInput => ({
  sheets: [{
    name: 'S',
    cells: { A1: { value: 'Item' }, B1: { value: 'Col1' }, C1: { value: 'x' }, D1: { value: 'y' }, A2: { value: 'a' }, B2: { value: 1 }, C2: { value: 2 }, D2: { value: 3 }, E1: { formula: 'SUM(T[Col3])' } },
    tables: [{ name: 'T', range: 'A1:D3', columns: [{ name: 'Item' }, { name: 'Col1' }, { name: 'x' }, { name: 'y' }] }],
  }],
})
const fillHeaders: StructureOp = { op: 'fillRange', sheet: 'S', range: 'B1:B2', to: 'B1:D2' }
const headersOf = (b: BookInput) => {
  const t = b.sheets[0].tables![0]
  return { columns: t.columns.map((c) => c.name), cells: t.columns.map((_, i) => b.sheets[0].cells[`${String.fromCharCode(65 + 2 + i)}1`]?.value) }
}

describe.each(EXECUTIONS)('a fill continuing a series over header cells (x-fuzz-session-9), %s', (execution) => {
  it('names the columns by the text the series writes', () => {
    const r = editStructure(seriesHeaders(), fillHeaders)
    if (!r.ok) throw new Error(r.message)
    expect(r.book.sheets[0].tables![0].columns.map((c) => c.name)).toEqual(['Item', 'Col1', 'Col2', 'Col3'])
    expect([r.book.sheets[0].cells.C1.value, r.book.sheets[0].cells.D1.value, r.book.sheets[0].cells.C2.value]).toEqual(['Col2', 'Col3', 2])
    expect(compute(r.book, { execution, clock }).sheets[0].cells.E1.value).toBe(3)
    // a copy fill keeps the copy's names, as before
    const copied = editStructure(seriesHeaders(), { ...fillHeaders, type: 'copy' })
    expect(copied.ok && copied.book.sheets[0].tables![0].columns.map((c) => c.name)).toEqual(['Item', 'Col1', 'Col12', 'Col13'])
  })

  it('writes a number series into header cells as text', () => {
    const book: BookInput = {
      sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { value: 2 }, C1: { value: 'a' }, D1: { value: 'b' }, C2: { value: 5 } }, tables: [{ name: 'T', range: 'C1:D3', columns: [{ name: 'a' }, { name: 'b' }] }] }],
    }
    const r = editStructure(book, { op: 'fillRange', sheet: 'S', range: 'A1:B1', to: 'A1:D1' })
    if (!r.ok) throw new Error(r.message)
    expect(headersOf(r.book)).toEqual({ columns: ['3', '4'], cells: ['3', '4'] })
    compute(r.book, { execution, clock })
  })

  inSessions(execution, seriesHeaders, fillHeaders, (r, after) => {
    expect(r.ok).toBe(true)
    expect(after.sheets[0].tables![0].columns.map((c) => c.name)).toEqual(['Item', 'Col1', 'Col2', 'Col3'])
  })
})

/** Table T (Qty, Price) and formulas that name a column it lacks, beside one that names Price. */
const unknownColumn = (): BookInput => ({
  sheets: [{
    name: 'S',
    cells: {
      A1: { value: 'Qty' }, B1: { value: 'Price' }, A2: { value: 2 }, B2: { value: 3 }, A3: { value: 4 }, B3: { value: 6 },
      E1: { formula: 'SUM(T[Nope])' }, E2: { formula: 'SUM(T[Price])' }, E3: { formula: 'SUM(T[[Nope]:[Price]])' }, E4: { formula: 'SUM(T[[#Headers],[Nope]])' }, E5: { formula: 'T[@Nope]' },
    },
    tables: [{ name: 'T', range: 'A1:B3', columns: [{ name: 'Qty' }, { name: 'Price' }] }],
  }],
})
const convert: StructureOp = { op: 'convertTable', table: 'T' }

describe.each(EXECUTIONS)('a table converted to cells with a reference to a column it lacks (fix-u304), %s', (execution) => {
  it('turns that reference into #REF!, so the formula stays an error', () => {
    const before = compute(unknownColumn(), { execution, clock }).sheets[0].cells
    expect(['E1', 'E3', 'E4', 'E5'].map((c) => before[c].value)).toEqual(Array(4).fill({ error: 'NAME' }))
    const r = editStructure(unknownColumn(), convert)
    if (!r.ok) throw new Error(r.message)
    const cells = r.book.sheets[0].cells
    expect(['E1', 'E2', 'E3', 'E4', 'E5'].map((c) => cells[c].formula)).toEqual(['SUM(#REF!)', "SUM('S'!$B$2:$B$3)", 'SUM(#REF!)', 'SUM(#REF!)', '#REF!'])
    const after = compute(r.book, { execution, clock }).sheets[0].cells
    expect(['E1', 'E2', 'E3', 'E4', 'E5'].map((c) => after[c].value)).toEqual([{ error: 'REF' }, 9, { error: 'REF' }, { error: 'REF' }, { error: 'REF' }])
  })

  inSessions(execution, unknownColumn, convert, (r, after) => {
    expect(r.ok).toBe(true)
    expect(after.sheets[0].cells.E1.formula).toBe('SUM(#REF!)')
  })
})

/** Table Tab on S, and conditional formats and data validation on S and T whose formulas refer to it. */
const rulesOverTable = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 'Qty' }, B1: { value: 'Price' }, A2: { value: 3 }, B2: { value: 1.5 }, A3: { value: 4 }, B3: { value: 2.5 }, D1: { formula: 'SUM(Tab[Qty])' } },
      tables: [{ name: 'Tab', range: 'A1:B3', columns: [{ name: 'Qty' }, { name: 'Price' }] }],
      conditionalFormats: [{ type: 'expression', ranges: ['A2:A3'], formula: 'A2>AVERAGE(Tab[Qty])', style: { fill: { color: '#FF0000' } } }],
      validations: [{ range: 'E1', type: 'list', formula1: 'Tab[Qty]' }, { range: 'F1', type: 'custom', formula1: 'F1>SUM(Tab[Price])' }],
    },
    {
      name: 'T',
      cells: { A1: { value: 10 } },
      conditionalFormats: [{ type: 'cellIs', ranges: ['A1'], operator: 'greaterThan', formulas: ['SUM(Tab[Qty])'], style: { fill: { color: '#00FF00' } } }],
      validations: [{ range: 'B1', type: 'list', formula1: 'Tab[Qty]' }],
    },
  ],
})
/** The formulas of each sheet's rules: conditional formats, then validations. */
const ruleFormulas = (b: BookInput) => Object.fromEntries(b.sheets.map((s) => [s.name, [
  ...(s.conditionalFormats ?? []).map((x) => (x.type === 'cellIs' ? x.formulas[0] : 'formula' in x ? x.formula : '')),
  ...(s.validations ?? []).map((v) => v.formula1),
]]))

describe.each(EXECUTIONS)('table operations rewrite structured references in rules (x-fix-structure-looks-2), %s', (execution) => {
  it.each([
    [{ op: 'renameTable', table: 'Tab', name: 'Sales' }, { S: ['A2>AVERAGE(Sales[Qty])', 'Sales[Qty]', 'F1>SUM(Sales[Price])'], T: ['SUM(Sales[Qty])', 'Sales[Qty]'] }],
    [{ op: 'deleteTable', table: 'Tab' }, { S: ['A2>AVERAGE(#REF!)', '#REF!', 'F1>SUM(#REF!)'], T: ['SUM(#REF!)', '#REF!'] }],
    [{ op: 'convertTable', table: 'Tab' }, { S: ["A2>AVERAGE('S'!$A$2:$A$3)", "'S'!$A$2:$A$3", "F1>SUM('S'!$B$2:$B$3)"], T: ["SUM('S'!$A$2:$A$3)", "'S'!$A$2:$A$3"] }],
    [{ op: 'deleteColumns', sheet: 'S', column: 'A' }, { S: ['#REF!', 'E1>SUM(Tab[Price])'], T: ['SUM(#REF!)', '#REF!'] }],
    [{ op: 'deleteSheet', sheet: 'S' }, { T: ['SUM(#REF!)', '#REF!'] }],
  ] as [StructureOp, Record<string, string[]>][])('%j', (op, want) => {
    const r = editStructure(rulesOverTable(), op)
    if (!r.ok) throw new Error(r.message)
    expect(ruleFormulas(r.book)).toEqual(want)
    compute(r.book, { execution, clock })
  })

  it('keeps the rules working in a session after a rename', () => {
    const s = createSession(rulesOverTable(), { execution, clock })
    const looks = () => [!!s.getCells('S', 'A2:A3').A3.appearance, !!s.getCells('T', 'A1').A1.appearance]
    expect(looks()).toEqual([true, true])
    expect(s.editStructure({ op: 'renameTable', table: 'Tab', name: 'Sales' }).ok).toBe(true)
    expect(looks()).toEqual([true, true])
    expect(s.getValidation('S', 'E1')!.list).toMatchObject([{ value: 3 }, { value: 4 }])
    expect(s.validateEntry('S', 'F1', { value: 100 })).toEqual({ ok: true })
    expect(s.getHash()).toBe(compute(s.getBook(), { execution, clock }).hash)
  })

  inSessions(execution, rulesOverTable, { op: 'renameTable', table: 'Tab', name: 'Sales' }, (r, after) => {
    expect(r.ok).toBe(true)
    expect(ruleFormulas(after).T).toEqual(['SUM(Sales[Qty])', 'Sales[Qty]'])
  })
})

/** A legacy array formula over A1:B1 beside a table at D1:E3. */
const arrayBeside = (): BookInput => ({
  locale: 'en-US',
  sheets: [{
    name: 'S',
    cells: { A1: { formula: '{1,2}', array: 'A1:B1' }, A5: { formula: '{3}', array: 'A5' }, D1: { value: 'k' }, E1: { value: 'v' }, D2: { value: 1 }, E2: { value: 2 }, D3: { value: 3 }, E3: { value: 4 } },
    tables: [{ name: 'T', range: 'D1:E3', columns: [{ name: 'k' }, { name: 'v' }] }],
  }],
})
const arrayRefusal = 'the array formula over D2:E2 overlaps the table T (Excel does not allow array formulas of several cells in tables)'

describe.each(EXECUTIONS)('an array formula of several cells moved or copied onto a table (array-into-table), %s', (execution) => {
  const ops: StructureOp[] = [
    { op: 'moveRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', cell: 'D2' } },
    { op: 'copyRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', range: 'D2' } },
  ]
  it.each(ops)('refuses $op with code table and the message checkBook gives', (op) => {
    expect(editStructure(arrayBeside(), op)).toMatchObject({ ok: false, code: 'table', message: arrayRefusal })
    expect(editStructure({ ...arrayBeside(), locale: 'ru-RU' }, op)).toMatchObject({ ok: false, code: 'table', message: 'Формула массива на D2:E2 пересекается с таблицей «T»: в таблицах Excel не допускает формулы массива из нескольких ячеек' })
  })

  it('takes an array formula of one cell onto a table, and one of several cells beside it', () => {
    compute(edited(editStructure(arrayBeside(), { op: 'moveRange', from: { sheet: 'S', range: 'A5' }, to: { sheet: 'S', cell: 'D2' } })), { execution, clock })
    compute(edited(editStructure(arrayBeside(), { op: 'copyRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', range: 'G2' } })), { execution, clock })
  })

  for (const op of ops)
    inSessions(execution, arrayBeside, op, (r, after) => {
      expect(r).toMatchObject({ ok: false, code: 'table' })
      expect(after.sheets[0].cells.A1).toEqual({ formula: '{1,2}', array: 'A1:B1' })
    })
})

function edited(r: ReturnType<typeof editStructure>): BookInput {
  if (!r.ok) throw new Error(r.message)
  return r.book
}
