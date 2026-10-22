/**
 * `Session.undo` and `redo` on each kind of step, with closures and eval: edits of cells (a spill, several edits of one
 * call, a table's header that renamed its column), patches, names, the style table, the layout (heights, merged areas
 * over a spill, hidden rows under SUBTOTAL), structure operations (rows, a moved range, a sheet renamed, deleted, hidden),
 * what the answers carry (the changes with their text, the cells put back, `mapCell` each way); the history's counts, a
 * new change forgetting the redo, `setClock` no step, a live session's clock, refused calls, `depth` and `maxMb`. The
 * cells a state built again shares: none with a book a host was given, which stays as it was in every kind of session,
 * nor with a step of the history.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { CalcError, compute, InputError, createSession, type BookInput, type CellChange, type Execution, type Session, type StructureOp } from '../src/index.js'
import { run } from '../src/exec/budget.js'
import { openSteps } from '../src/session.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

const book = (): BookInput => ({
  styles: [{ font: { bold: true } }],
  names: [{ name: 'rate', ref: '2' }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3, numberFormat: '0.0' },
        B1: { formula: 'SUM(A1:A3)*rate' },
        C1: { formula: 'SEQUENCE(A2)' },
        D1: { formula: 'SUBTOTAL(109,A1:A3)' },
        F1: { value: 'Qty' }, F2: { value: 4 }, G1: { formula: 'SUM(T[Qty])' },
      },
      tables: [{ name: 'T', range: 'F1:F2', columns: [{ name: 'Qty' }] }],
    },
    { name: 'R', cells: { A1: { formula: 'S!B1+1' }, A2: { formula: "SUM('S'!A:A)" } } },
  ],
})

/** Changes as `sheet!cell=value text`, in the order of the sheets and addresses. */
const listed = (changes: CellChange[]) => changes.map((c) => `${c.sheet}!${c.cell}=${JSON.stringify(c.value)} ${c.text ?? ''}${c.spill ? ` ${c.spill}` : ''}`).sort()

/** Structure operations on `s`, each answered ok: the book after each, a copy. */
function applied(s: Session, ops: readonly StructureOp[]): BookInput[] {
  return ops.map((op) => {
    expect(s.editStructure(op).ok).toBe(true)
    return structuredClone(s.getBook())
  })
}

/** The structure operations that made `books` (from the first) undone one by one, then done again: each book in turn. */
function backAndAgain(s: Session, books: BookInput[], o: { execution: Execution }) {
  for (let k = books.length - 2; k >= 0; k--) {
    expect(s.undo()!.call).toBe('editStructure')
    same(s, books[k], o)
  }
  for (let k = 1; k < books.length; k++) {
    s.redo()
    same(s, books[k], o)
  }
}

/** The session's book and hash are those of `want`, and a full computation of it. */
function same(s: Session, want: BookInput, o: { execution: Execution }) {
  expect(s.getBook()).toEqual(want)
  expect(s.getHash()).toBe(compute(want, { clock, ...o }).hash)
}

describe.each(['closures', 'eval'] as const)('undo and redo, %s', (execution) => {
  const o = { clock, execution }

  it('edits of cells: one call is one step, its edits undone together; the changes with their text and the cells put back', () => {
    const s = createSession(book(), o)
    const before = structuredClone(s.getBook())
    s.setCells([
      { sheet: 'S', cell: 'A1', input: { value: 10 } },
      { sheet: 'S', cell: 'A2', input: { value: 3 } },
      { sheet: 'S', cell: 'E9', input: { value: 'new' } },
    ])
    const after = structuredClone(s.getBook())
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    const u = s.undo()!
    expect(u.call).toBe('setCells')
    expect(u.cells).toEqual([{ sheet: 'S', cell: 'A1' }, { sheet: 'S', cell: 'A2' }, { sheet: 'S', cell: 'E9' }])
    expect(listed(u.changes)).toEqual(['R!A1=13 13', 'R!A2=6 6', 'S!A1=1 1', 'S!A2=2 2', 'S!B1=12 12', 'S!C1=1 1 C1:C2', 'S!C3=null ', 'S!D1=6 6', 'S!E9=null '])
    same(s, before, o)
    expect(s.history).toEqual({ undo: 0, redo: 1 })
    expect(s.undo()).toBeNull()
    expect(listed(s.redo()!.changes)).toEqual(['R!A1=33 33', 'R!A2=16 16', 'S!A1=10 10', 'S!A2=3 3', 'S!B1=32 32', 'S!C1=1 1 C1:C3', 'S!C3=3 3', 'S!D1=16 16', 'S!E9="new" new'])
    same(s, after, o)
    expect(s.redo()).toBeNull()
    expect(s.history).toEqual({ undo: 1, redo: 0 })
  })

  it('clearing formats is undone and redone as clearCells, not as the patchCells it is made of', () => {
    const s = createSession(book(), o)
    s.patchCells([{ sheet: 'S', range: 'E9', set: { style: 0 } }])
    s.clearCells([{ sheet: 'S', range: 'E9' }], 'formats')
    expect(s.undo()!.call).toBe('clearCells')
    expect(s.redo()!.call).toBe('clearCells')
  })

  it('an old-kind array formula made, changed and cleared: each undone and redone, the whole area with it', () => {
    const s = createSession(book(), o)
    const books = [structuredClone(s.getBook())]
    const clear = ['E1', 'E2', 'E3'].map((cell) => ({ sheet: 'S', cell, input: null }))
    for (const edits of [[{ sheet: 'S', cell: 'E1', input: { formula: 'A1:A2*10', array: 'E1:E3' } }], [{ sheet: 'S', cell: 'E1', input: { formula: 'A2:A3*100', array: 'E1:E3' } }], clear]) {
      s.setCells(edits)
      books.push(structuredClone(s.getBook()))
    }
    s.patchCells([{ sheet: 'S', cell: 'E2', set: { style: 0 } }])
    expect(s.undo()!.cells).toEqual([{ sheet: 'S', cell: 'E2' }])
    for (let k = 2; k >= 0; k--) {
      s.undo()
      same(s, books[k], o)
    }
    s.redo()
    expect(s.getCells('S', 'E1:E3')).toMatchObject({ E1: { value: 10, array: 'E1:E3' }, E2: { value: 20 }, E3: { value: { error: 'NA' } } })
    expect(() => s.setCells([{ sheet: 'S', cell: 'E2', input: { value: 1 } }])).toThrow(InputError)
    expect(s.history).toEqual({ undo: 1, redo: 3 })
  })

  it("a table's header typed over renames its column: undo names it back, with the references", () => {
    const s = createSession(book(), o)
    const before = structuredClone(s.getBook())
    s.setCells([{ sheet: 'S', cell: 'F1', input: { value: 'Count' } }])
    expect(s.getBook().sheets[0].cells.G1.formula).toBe('SUM(T[Count])')
    // the column's name and the formula over it come back; the table is the sheet's layout
    expect(s.undo()).toEqual({
      call: 'setCells',
      revision: 2,
      changes: [{ sheet: 'S', cell: 'F1', value: 'Qty', text: 'Qty' }, { sheet: 'S', cell: 'G1', value: 4, text: '4', formula: 'SUM(T[Qty])' }],
      layout: ['S'],
    })
    same(s, before, o)
    s.redo()
    expect(s.getBook().sheets[0].tables![0].columns[0].name).toBe('Count')
  })

  it('patches: the style, format, note and link back, the shown text in the changes; a cell the patch made goes', () => {
    const s = createSession(book(), o)
    const before = structuredClone(s.getBook())
    const forward = [{ sheet: 'S', cell: 'A3', value: 3, text: '3', style: 0, numberFormat: null, note: { text: 'n' } }, { sheet: 'S', cell: 'H5', value: null, text: '', link: { url: 'https://example.com/' } }]
    expect(s.patchCells([{ sheet: 'S', cell: 'A3', set: { numberFormat: null, style: 0, note: { text: 'n' } } }, { sheet: 'S', cell: 'H5', set: { link: { url: 'https://example.com/' } } }]).changes).toEqual(forward)
    expect(s.undo()).toEqual({ call: 'patchCells', revision: 2, changes: [{ sheet: 'S', cell: 'A3', value: 3, text: '3.0', style: null, numberFormat: '0.0', numberFormatCode: '0.0', note: null }, { sheet: 'S', cell: 'H5', value: null, text: '', link: null }], cells: [{ sheet: 'S', cell: 'A3' }, { sheet: 'S', cell: 'H5' }] })
    same(s, before, o)
    expect(s.redo()!.changes).toEqual(forward)
    expect(s.getBook().sheets[0].cells.H5).toEqual({ link: { url: 'https://example.com/' } })
  })

  it('names and the style table', () => {
    const s = createSession(book(), o)
    const before = structuredClone(s.getBook())
    s.setNames([{ name: 'rate', ref: '3' }, { name: 'other', ref: 'S!$A$1' }])
    s.setStyles([{ font: { bold: true } }, { font: { italic: true } }])
    expect(s.undo()).toEqual({ call: 'setStyles', revision: 3, changes: [], styles: [{ font: { bold: true } }] })
    expect(s.getBook().styles).toEqual(before.styles)
    expect(s.undo()).toEqual({ call: 'setNames', revision: 4, changes: [{ sheet: 'S', cell: 'B1', value: 12, text: '12' }, { sheet: 'R', cell: 'A1', value: 13, text: '13' }] })
    same(s, before, o)
    s.redo()
    s.redo()
    expect(s.getBook().names).toEqual([{ name: 'rate', ref: '3' }, { name: 'other', ref: 'S!$A$1' }])
    expect(s.getBook().styles).toHaveLength(2)
  })

  it('the layout: heights and widths back, a merged area over a spill gone, hidden rows counted again by SUBTOTAL', () => {
    const s = createSession(book(), o)
    const before = structuredClone(s.getBook())
    const narrow = s.setLayout([{ sheet: 'S', rows: '1:3', set: { height: 30 } }, { sheet: 'S', columns: 'A', set: { width: 1 } }])
    s.setLayout([{ sheet: 'S', merges: ['C2:D3'] }])
    s.setLayout([{ sheet: 'S', rows: '2', set: { hidden: true } }])
    expect(s.getCells('S', 'C1:D1')).toMatchObject({ C1: { value: { error: 'SPILL' } }, D1: { value: 4 } })
    expect(s.undo()).toEqual({ call: 'setLayout', revision: 4, changes: [{ sheet: 'S', cell: 'D1', value: 6, text: '6' }], layout: ['S'] })
    expect(listed(s.undo()!.changes)).toEqual(['S!C1=1 1 C1:C2', 'S!C2=2 2'])
    // a width that shows numbers as ### shows them again
    expect(s.undo()!.changes.map((c) => `${c.cell} ${c.text}`)).toEqual(['A3 3.0'])
    same(s, before, o)
    expect(s.redo()!.changes).toEqual(narrow.changes)
    expect(narrow.changes.map((c) => c.cell)).toEqual(['A3'])
    expect(s.getBook().sheets[0].rows).toEqual({ '1:3': { height: 30 } })
    // a span cut in three by a change of its middle line, and whole again
    s.setLayout([{ sheet: 'S', rows: '2', set: { height: 40, numberFormat: '0' } }])
    expect(s.getBook().sheets[0].rows).toEqual({ 1: { height: 30 }, 2: { height: 40, numberFormat: '0' }, 3: { height: 30 } })
    expect(s.undo()!.changes.map((c) => `${c.cell} ${c.text}`)).toEqual(['A2 2'])
    expect(s.getBook().sheets[0].rows).toEqual({ '1:3': { height: 30 } })
  })

  it("cells() reads the rows' and columns' formats again after a change of the layout, a structure operation and an undo", () => {
    const s = createSession(book(), o)
    const text = () => Object.values(s.getCells('S', 'A1:A4')).map((c) => c.text)
    expect(text()).toEqual(['1', '2', '3.0'])
    s.setLayout([{ sheet: 'S', rows: '1:2', set: { numberFormat: '0.00' } }])
    expect(text()).toEqual(['1.00', '2.00', '3.0'])
    s.setLayout([{ sheet: 'S', columns: 'A', set: { numberFormat: '0%' } }, { sheet: 'S', rows: '2', set: { numberFormat: null } }])
    expect(text()).toEqual(['1.00', '200%', '3.0'])
    s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(Object.values(s.getCells('S', 'A2:A4')).map((c) => c.text)).toEqual(['1.00', '200%', '3.0'])
    s.undo()
    s.undo()
    expect(text()).toEqual(['1.00', '2.00', '3.0'])
    s.undo()
    expect(text()).toEqual(['1', '2', '3.0'])
  })

  it("AutoFit undone: the columns' widths, bestFit and hidden as they were", () => {
    const s = createSession(book(), o)
    s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 30 } }, { sheet: 'S', columns: 'B', set: { hidden: true } }])
    const before = structuredClone(s.getBook())
    s.setLayout([{ sheet: 'S', columns: 'A:B', fit: true }])
    expect(s.getBook().sheets[0].columns).toMatchObject({ A: { bestFit: true }, B: { bestFit: true } })
    expect(s.getBook().sheets[0].columns!.B.hidden).toBeUndefined()
    expect(s.undo()!.call).toBe('setLayout')
    same(s, before, o)
    s.redo()
    expect(s.getBook().sheets[0].columns).toMatchObject({ A: { bestFit: true }, B: { bestFit: true } })
  })

  it('rows inserted and deleted: the book before, mapCell from the book before the undo to the one after it', () => {
    const s = createSession(book(), o)
    const before = structuredClone(s.getBook())
    s.editStructure({ op: 'insertRows', sheet: 'S', row: 2, count: 2 })
    const inserted = structuredClone(s.getBook())
    const u = s.undo()!
    expect(u.call).toBe('editStructure')
    expect(u.mapCell!({ sheet: 'S', cell: 'A4' })).toEqual({ sheet: 'S', cell: 'A2' })
    expect(u.mapCell!({ sheet: 'S', cell: 'A3' })).toBeNull()
    expect(u.mapCell!({ sheet: 'r', cell: 'A3' })).toEqual({ sheet: 'R', cell: 'A3' })
    expect(u.mapCell!({ sheet: 'Nope', cell: 'A1' })).toBeNull()
    expect(() => u.mapCell!({ sheet: 'S', cell: 'A0' })).toThrow(InputError)
    // C2 of the inserted rows comes from no cell: the spill's value is listed, and the formulas over the rows with their text
    expect(u.changes).toEqual([
      { sheet: 'S', cell: 'B1', value: 12, text: '12', formula: 'SUM(A1:A3)*rate' },
      { sheet: 'S', cell: 'C1', value: 1, text: '1', formula: 'SEQUENCE(A2)', spill: 'C1:C2' },
      { sheet: 'S', cell: 'D1', value: 6, text: '6', formula: 'SUBTOTAL(109,A1:A3)' },
      { sheet: 'S', cell: 'C2', value: 2, text: '2', spillOf: 'C1' },
    ])
    expect(u.moved).toEqual([{ kind: 'lines', sheet: 'S', axis: 'rows', at: 2, count: -2 }])
    same(s, before, o)
    const r = s.redo()!
    expect(r.mapCell!({ sheet: 'S', cell: 'A2' })).toEqual({ sheet: 'S', cell: 'A4' })
    same(s, inserted, o)
    s.undo()
    s.editStructure({ op: 'deleteRows', sheet: 'S', row: 3 })
    expect(s.getBook().sheets[0].cells.B1.formula).toBe('SUM(A1:A2)*rate')
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    const d = s.undo()!
    expect([d.mapCell!({ sheet: 'S', cell: 'A2' }), d.mapCell!({ sheet: 'S', cell: 'A3' })]).toEqual([{ sheet: 'S', cell: 'A2' }, { sheet: 'S', cell: 'A4' }])
    same(s, before, o)
  })

  it('a range moved, a sheet renamed, deleted and hidden: each back', () => {
    const s = createSession(book(), o)
    const before = structuredClone(s.getBook())
    const ops = [
      { op: 'moveRange', from: { sheet: 'S', range: 'A1:A2' }, to: { sheet: 'R', cell: 'C5' } },
      { op: 'renameSheet', sheet: 's', name: 'Plan' },
      { op: 'insertSheet', name: 'N' },
      { op: 'hideSheet', sheet: 'R' },
      { op: 'deleteSheet', sheet: 'Plan' },
    ] as const
    const books = [before, ...applied(s, ops)]
    expect(s.getBook().sheets).toEqual([{ name: 'R', hidden: true, cells: { A1: { formula: '#REF!B1+1' }, A2: { formula: 'SUM(#REF!A:A)' }, C5: { value: 1 }, C6: { value: 2 } } }, { name: 'N', cells: {} }])
    backAndAgain(s, books, o)
  })

  it('a deleted sheet given back as the history keeps it, edited, deleted and given back again: each book as it was', () => {
    const s = createSession(book(), o)
    const before = structuredClone(s.getBook())
    s.editStructure({ op: 'deleteSheet', sheet: 'S' })
    const deleted = structuredClone(s.getBook())
    s.undo()
    same(s, before, o)
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
    const edited = structuredClone(s.getBook())
    s.editStructure({ op: 'deleteSheet', sheet: 'S' })
    same(s, deleted, o)
    s.undo()
    same(s, edited, o)
    s.undo()
    same(s, before, o)
    expect(s.undo()).toBeNull()
    s.redo()
    same(s, edited, o)
    s.redo()
    same(s, deleted, o)
    s.undo()
    same(s, edited, o)
  })

  it('structure operations over cells of a look alone under a spill and beside it, new ones from empty cells: each back, and again', () => {
    const cells = { A1: { formula: 'SEQUENCE(4)' }, A2: { style: 0 }, A3: { style: 0, numberFormat: '0.0' }, B5: { style: 0 }, C1: { value: 7 }, C3: { formula: 'A3*2' } }
    const s = createSession({ styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells, columns: { F: { style: 0 } } }] }, o)
    // the empty cells of a bold column moved out of it are bold cells; a table over empty cells writes its header
    const ops: StructureOp[] = [
      { op: 'moveRange', from: { sheet: 'S', range: 'F1:F3' }, to: { sheet: 'S', cell: 'B20' } },
      { op: 'insertRows', sheet: 'S', row: 3 },
      { op: 'deleteRows', sheet: 'S', row: 2 },
      { op: 'moveRange', from: { sheet: 'S', range: 'A2:B5' }, to: { sheet: 'S', cell: 'D10' } },
      { op: 'insertColumns', sheet: 'S', column: 'A' },
      { op: 'copyRange', from: { sheet: 'S', range: 'B1:D4' }, to: { sheet: 'S', range: 'C2' } },
      { op: 'deleteColumns', sheet: 'S', column: 'B' },
      { op: 'addTable', sheet: 'S', table: { name: 'Tab', range: 'H1:I3', columns: [{ name: 'x' }, { name: 'y' }] } },
    ]
    const books = [structuredClone(s.getBook()), ...applied(s, ops)]
    backAndAgain(s, books, o)
  })

  it('a column of more cells deleted than the way back looks up one by one: the sheet counted, its cells back', () => {
    const cells: BookInput['sheets'][0]['cells'] = { C1: { formula: 'SUM(A:A)' } }
    for (let r = 1; r <= 70_000; r++) cells[`A${r}`] = { value: r % 7 }
    const s = createSession({ sheets: [{ name: 'S', cells }] }, o)
    const before = structuredClone(s.getBook())
    expect(s.editStructure({ op: 'deleteColumns', sheet: 'S', column: 'A' }).ok).toBe(true)
    const after = structuredClone(s.getBook())
    expect(after.sheets[0].cells).toEqual({ B1: { formula: 'SUM(#REF!)' } })
    s.undo()
    same(s, before, o)
    s.redo()
    same(s, after, o)
  })

  it('a new change forgets the redo; setClock, refused calls and reads are no steps', () => {
    const s = createSession(book(), o)
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 6 } }])
    s.undo()
    expect(s.history).toEqual({ undo: 1, redo: 1 })
    s.setClock({ now: '2027-01-01T00:00:00Z', seed: 1 })
    s.getCells('S', 'A1:B2')
    s.validateEntry('S', 'A1', { value: 1 })
    expect(() => s.setCells([{ sheet: 'Nope', cell: 'A1', input: null }])).toThrow(InputError)
    expect(s.editStructure({ op: 'deleteSheet', sheet: 'S' }).ok).toBe(true)
    s.undo()
    expect(s.editStructure({ op: 'renameSheet', sheet: 'S', name: 'R' })).toMatchObject({ ok: false, code: 'sheet_name' })
    expect(() => s.setLayout([{ sheet: 'S', rows: '1', set: { style: 5 } }])).toThrow(InputError)
    expect(s.history).toEqual({ undo: 1, redo: 1 })
    s.setCells([])
    s.patchCells([])
    expect(s.history).toEqual({ undo: 1, redo: 1 })
    s.patchCells([{ sheet: 'S', cell: 'A1', set: { style: 0 } }])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
  })

  it('a call stopped by its limits is no step', () => {
    const s = createSession(book(), { ...o, limits: { cells: 1000 } })
    expect(() => s.setCells([{ sheet: 'S', cell: 'A9', input: { formula: 'SUM(SEQUENCE(2000))' } }])).toThrow(CalcError)
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    s.setCells([{ sheet: 'S', cell: 'A9', input: { formula: 'SUM(SEQUENCE(20))' } }])
    s.undo()
    expect(s.getBook().sheets[0].cells.A9).toBeUndefined()
  })

  it('a live session takes a new clock when an undo recomputes; volatile cells too', () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1+RAND()*0' }, C1: { formula: 'RAND()' } } }] }, { live: true, execution })
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }])
    const [at, rand] = [s.clock, s.getCells('S', 'C1').C1.value]
    const u = s.undo()!
    expect(s.clock).not.toEqual(at)
    expect(u.changes.find((c) => c.cell === 'C1')?.value).not.toEqual(rand)
    expect(s.getHash()).toBe(compute(s.getBook(), { clock: s.clock, execution }).hash)
  })

  it('depth: the oldest steps go; 0 keeps none', () => {
    const s = createSession(book(), { ...o, undo: { depth: 2 } })
    for (const v of [5, 6, 7]) s.setCells([{ sheet: 'S', cell: 'A1', input: { value: v } }])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
    s.undo()
    s.undo()
    expect(s.undo()).toBeNull()
    expect(s.getBook().sheets[0].cells.A1).toEqual({ value: 5 })
    const none = createSession(book(), { ...o, undo: { depth: 0 } })
    none.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
    none.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(none.history).toEqual({ undo: 0, redo: 0 })
    expect(none.undo()).toBeNull()
    expect(() => createSession(book(), { ...o, undo: { depth: -1 } })).toThrow('opts.undo.depth')
    expect(() => createSession(book(), { ...o, undo: { size: 1 } } as never)).toThrow('opts.undo: unknown field "size"')
  })

  it('maxMb: the oldest steps go past it; a step larger than it alone forgets the history, as Excel forgets what it cannot undo', () => {
    // about 3 KB a step of 20 cells of 60 characters
    const big = (v: string) => Array.from({ length: 20 }, (_, i) => ({ sheet: 'S', cell: `K${i + 1}`, input: { value: v.repeat(60) } }))
    const s = createSession(book(), { ...o, undo: { maxMb: 8 / 1024 } })
    for (const v of 'abcd') s.setCells(big(v))
    expect(s.history.undo).toBeGreaterThan(0)
    expect(s.history.undo).toBeLessThan(4)
    const kept = s.history.undo
    for (let i = 0; i < kept; i++) s.undo()
    expect(s.undo()).toBeNull()
    expect(s.getBook().sheets[0].cells.K1.value).toBe('abcd'[3 - kept].repeat(60))
    // the step keeps the inputs it replaces: 20 values of 600 characters, over 8 KB
    s.setCells(big('e'.repeat(10)))
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    s.setCells(big('f'))
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    expect(s.undo()).toBeNull()
  })

  it('a book never read between calls gives the answers and the book of one read after each: a state built again takes the cells no host was given as they are', () => {
    const calls: ((s: Session) => unknown)[] = [
      (s) => s.editStructure({ op: 'deleteSheet', sheet: 'S' }),
      (s) => s.undo(),
      (s) => s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 9 } }]),
      (s) => s.editStructure({ op: 'insertRows', sheet: 'R', row: 1 }),
      (s) => s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 8 } }]),
      (s) => s.undo(),
      (s) => s.undo(),
      (s) => s.redo(),
      (s) => s.setNames([{ name: 'rate', ref: '3' }]),
      (s) => s.setCells([{ sheet: 'R', cell: 'B1', input: { value: 7 } }]),
      (s) => s.editStructure({ op: 'deleteSheet', sheet: 'S' }),
      (s) => s.undo(),
      (s) => s.undo(),
      (s) => s.redo(),
    ]
    const [read, unread] = [createSession(book(), o), createSession(book(), o)]
    const books = []
    for (const [k, call] of calls.entries()) {
      expect(JSON.stringify(call(unread)), `call ${k}`).toBe(JSON.stringify(call(read)))
      books.push(structuredClone(read.getBook()))
    }
    expect(unread.getBook()).toEqual(books.at(-1))
    same(unread, books.at(-1)!, o)
    expect(unread.getBook().sheets[0].cells.A1).toEqual({ value: 9 })
  })

  it("an empty sheet whose cells an operation keeps (a rename), edited after: the operation's step holds its cells as they were, as a thread's snapshot reads them", () => {
    const s = run(openSteps({ sheets: [{ name: 'A', cells: { A1: { formula: 'SUM(E!A1:A3)' } } }, { name: 'E', cells: {} }] }, o))
    run(s.editStructure({ op: 'renameSheet', sheet: 'E', name: 'F' }))
    run(s.setCells([{ sheet: 'F', cell: 'A2', input: { value: 5 } }]))
    const [rename] = s.historyData().undo
    expect(rename.kind === 'book' && rename.diff.sheets[1]).toEqual({ rest: { name: 'E' }, put: {} })
    run(s.undo())
    run(s.undo())
    expect(s.getBook().sheets[1]).toEqual({ name: 'E', cells: {} })
  })
})

describe.each(['closures', 'eval'] as const)('a book a host was given, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('stays as it was through later operations, undos and edits; one read after them has them', async () => {
      const s = await (await opener(kind))(book(), { clock, execution })
      const given = await s.getBook()
      const was = structuredClone(given)
      await s.editStructure({ op: 'deleteSheet', sheet: 'S' })
      await s.undo()
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 9 } }, { sheet: 'R', cell: 'B1', input: { value: 8 } }])
      await s.editStructure({ op: 'insertRows', sheet: 'R', row: 1 })
      await s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 7 } }])
      expect(given).toEqual(was)
      const now = await s.getBook()
      expect([now.sheets[0].cells.A1, now.sheets[0].cells.A2, now.sheets[1].cells.B2]).toEqual([{ value: 9 }, { value: 7 }, { value: 8 }])
      await s.close()
    }, 60_000)
  })
})
