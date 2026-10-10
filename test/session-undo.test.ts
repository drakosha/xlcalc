/**
 * `Session.undo` and `redo` on each kind of step, with closures and eval: edits of cells (a spill, several edits of one
 * call, a table's header that renamed its column), patches, names, the style table, the layout (heights, merged areas
 * over a spill, hidden rows under SUBTOTAL), structure operations (rows, a moved range, a sheet renamed, deleted, hidden),
 * what the answers carry (the changes with their text, the cells put back, `mapCell` each way); the history's counts, a
 * new change forgetting the redo, `setClock` no step, a live session's clock, refused calls, `depth` and `maxMb`.
 */
import { describe, expect, it } from 'vitest'
import { CalcError, compute, InputError, openSession, type BookInput, type CellChange, type Execution, type Session } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

const book = (): BookInput => ({
  styles: [{ font: { bold: true } }],
  names: [{ name: 'rate', ref: '2' }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3, format: '0.0' },
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

/** The session's book and hash are those of `want`, and a full computation of it. */
function same(s: Session, want: BookInput, o: { execution: Execution }) {
  expect(s.book()).toEqual(want)
  expect(s.hash()).toBe(compute(want, { clock, ...o }).hash)
}

describe.each(['closures', 'eval'] as const)('undo and redo, %s', (execution) => {
  const o = { clock, execution }

  it('edits of cells: one call is one step, its edits undone together; the changes with their text and the cells put back', () => {
    const s = openSession(book(), o)
    const before = structuredClone(s.book())
    s.setCells([
      { sheet: 'S', cell: 'A1', input: { value: 10 } },
      { sheet: 'S', cell: 'A2', input: { value: 3 } },
      { sheet: 'S', cell: 'E9', input: { value: 'new' } },
    ])
    const after = structuredClone(s.book())
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

  it('an old-kind array formula made, changed and cleared: each undone and redone, the whole area with it', () => {
    const s = openSession(book(), o)
    const books = [structuredClone(s.book())]
    const clear = ['E1', 'E2', 'E3'].map((cell) => ({ sheet: 'S', cell, input: null }))
    for (const edits of [[{ sheet: 'S', cell: 'E1', input: { formula: 'A1:A2*10', array: 'E1:E3' } }], [{ sheet: 'S', cell: 'E1', input: { formula: 'A2:A3*100', array: 'E1:E3' } }], clear]) {
      s.setCells(edits)
      books.push(structuredClone(s.book()))
    }
    s.patchCells([{ sheet: 'S', cell: 'E2', set: { style: 0 } }])
    expect(s.undo()!.cells).toEqual([{ sheet: 'S', cell: 'E2' }])
    for (let k = 2; k >= 0; k--) {
      s.undo()
      same(s, books[k], o)
    }
    s.redo()
    expect(s.cells('S', 'E1:E3')).toMatchObject({ E1: { value: 10, array: 'E1:E3' }, E2: { value: 20 }, E3: { value: { error: 'NA' } } })
    expect(() => s.setCells([{ sheet: 'S', cell: 'E2', input: { value: 1 } }])).toThrow(InputError)
    expect(s.history).toEqual({ undo: 1, redo: 3 })
  })

  it("a table's header typed over renames its column: undo names it back, with the references", () => {
    const s = openSession(book(), o)
    const before = structuredClone(s.book())
    s.setCells([{ sheet: 'S', cell: 'F1', input: { value: 'Count' } }])
    expect(s.book().sheets[0].cells.G1.formula).toBe('SUM(T[Count])')
    expect(s.undo()).toEqual({ call: 'setCells', changes: [{ sheet: 'S', cell: 'F1', value: 'Qty', text: 'Qty' }] })
    same(s, before, o)
    s.redo()
    expect(s.book().sheets[0].tables![0].columns[0].name).toBe('Count')
  })

  it('patches: the style, format, note and link back, the shown text in the changes; a cell the patch made goes', () => {
    const s = openSession(book(), o)
    const before = structuredClone(s.book())
    expect(s.patchCells([{ sheet: 'S', cell: 'A3', set: { format: null, style: 0, note: { text: 'n' } } }, { sheet: 'S', cell: 'H5', set: { link: { url: 'https://example.com/' } } }])).toEqual([{ sheet: 'S', cell: 'A3', value: 3, text: '3' }])
    expect(s.undo()).toEqual({ call: 'patchCells', changes: [{ sheet: 'S', cell: 'A3', value: 3, text: '3.0' }], cells: [{ sheet: 'S', cell: 'A3' }, { sheet: 'S', cell: 'H5' }] })
    same(s, before, o)
    expect(s.redo()!.changes).toEqual([{ sheet: 'S', cell: 'A3', value: 3, text: '3' }])
    expect(s.book().sheets[0].cells.H5).toEqual({ link: { url: 'https://example.com/' } })
  })

  it('names and the style table', () => {
    const s = openSession(book(), o)
    const before = structuredClone(s.book())
    s.setNames([{ name: 'rate', ref: '3' }, { name: 'other', ref: 'S!$A$1' }])
    s.setStyles([{ font: { bold: true } }, { font: { italic: true } }])
    expect(s.undo()).toEqual({ call: 'setStyles', changes: [] })
    expect(s.book().styles).toEqual(before.styles)
    expect(s.undo()).toEqual({ call: 'setNames', changes: [{ sheet: 'S', cell: 'B1', value: 12, text: '12' }, { sheet: 'R', cell: 'A1', value: 13, text: '13' }] })
    same(s, before, o)
    s.redo()
    s.redo()
    expect(s.book().names).toEqual([{ name: 'rate', ref: '3' }, { name: 'other', ref: 'S!$A$1' }])
    expect(s.book().styles).toHaveLength(2)
  })

  it('the layout: heights and widths back, a merged area over a spill gone, hidden rows counted again by SUBTOTAL', () => {
    const s = openSession(book(), o)
    const before = structuredClone(s.book())
    const narrow = s.setLayout([{ sheet: 'S', rows: '1:3', set: { height: 30 } }, { sheet: 'S', columns: 'A', set: { width: 1 } }])
    s.setLayout([{ sheet: 'S', merges: ['C2:D3'] }])
    s.setLayout([{ sheet: 'S', rows: '2', set: { hidden: true } }])
    expect(s.cells('S', 'C1:D1')).toMatchObject({ C1: { value: { error: 'SPILL' } }, D1: { value: 4 } })
    expect(s.undo()).toEqual({ call: 'setLayout', changes: [{ sheet: 'S', cell: 'D1', value: 6, text: '6' }] })
    expect(listed(s.undo()!.changes)).toEqual(['S!C1=1 1 C1:C2', 'S!C2=2 2'])
    // a width that shows numbers as ### shows them again
    expect(s.undo()!.changes.map((c) => `${c.cell} ${c.text}`)).toEqual(['A3 3.0'])
    same(s, before, o)
    expect(s.redo()!.changes).toEqual(narrow)
    expect(narrow.map((c) => c.cell)).toEqual(['A3'])
    expect(s.book().sheets[0].rows).toEqual({ '1:3': { height: 30 } })
    // a span cut in three by a change of its middle line, and whole again
    s.setLayout([{ sheet: 'S', rows: '2', set: { height: 40, format: '0' } }])
    expect(s.book().sheets[0].rows).toEqual({ 1: { height: 30 }, 2: { height: 40, format: '0' }, 3: { height: 30 } })
    expect(s.undo()!.changes.map((c) => `${c.cell} ${c.text}`)).toEqual(['A2 2'])
    expect(s.book().sheets[0].rows).toEqual({ '1:3': { height: 30 } })
  })

  it("cells() reads the rows' and columns' formats again after a change of the layout, a structure operation and an undo", () => {
    const s = openSession(book(), o)
    const text = () => Object.values(s.cells('S', 'A1:A4')).map((c) => c.text)
    expect(text()).toEqual(['1', '2', '3.0'])
    s.setLayout([{ sheet: 'S', rows: '1:2', set: { format: '0.00' } }])
    expect(text()).toEqual(['1.00', '2.00', '3.0'])
    s.setLayout([{ sheet: 'S', columns: 'A', set: { format: '0%' } }, { sheet: 'S', rows: '2', set: { format: null } }])
    expect(text()).toEqual(['1.00', '200%', '3.0'])
    s.restructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(Object.values(s.cells('S', 'A2:A4')).map((c) => c.text)).toEqual(['1.00', '200%', '3.0'])
    s.undo()
    s.undo()
    expect(text()).toEqual(['1.00', '2.00', '3.0'])
    s.undo()
    expect(text()).toEqual(['1', '2', '3.0'])
  })

  it('rows inserted and deleted: the book before, mapCell from the book before the undo to the one after it', () => {
    const s = openSession(book(), o)
    const before = structuredClone(s.book())
    s.restructure({ op: 'insertRows', sheet: 'S', row: 2, count: 2 })
    const inserted = structuredClone(s.book())
    const u = s.undo()!
    expect(u.call).toBe('restructure')
    expect(u.mapCell!({ sheet: 'S', cell: 'A4' })).toEqual({ sheet: 'S', cell: 'A2' })
    expect(u.mapCell!({ sheet: 'S', cell: 'A3' })).toBeNull()
    expect(u.mapCell!({ sheet: 'r', cell: 'A3' })).toEqual({ sheet: 'R', cell: 'A3' })
    expect(() => u.mapCell!({ sheet: 'Nope', cell: 'A1' })).toThrow(InputError)
    // C2 of the inserted rows comes from no cell: the spill's value is listed
    expect(u.changes).toEqual([{ sheet: 'S', cell: 'C2', value: 2, text: '2', spillOf: 'C1' }])
    same(s, before, o)
    const r = s.redo()!
    expect(r.mapCell!({ sheet: 'S', cell: 'A2' })).toEqual({ sheet: 'S', cell: 'A4' })
    same(s, inserted, o)
    s.undo()
    s.restructure({ op: 'deleteRows', sheet: 'S', row: 3 })
    expect(s.book().sheets[0].cells.B1.formula).toBe('SUM(A1:A2)*rate')
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    const d = s.undo()!
    expect([d.mapCell!({ sheet: 'S', cell: 'A2' }), d.mapCell!({ sheet: 'S', cell: 'A3' })]).toEqual([{ sheet: 'S', cell: 'A2' }, { sheet: 'S', cell: 'A4' }])
    same(s, before, o)
  })

  it('a range moved, a sheet renamed, deleted and hidden: each back', () => {
    const s = openSession(book(), o)
    const before = structuredClone(s.book())
    const ops = [
      { op: 'moveRange', from: { sheet: 'S', range: 'A1:A2' }, to: { sheet: 'R', cell: 'C5' } },
      { op: 'renameSheet', sheet: 's', name: 'Plan' },
      { op: 'insertSheet', name: 'N' },
      { op: 'hideSheet', sheet: 'R' },
      { op: 'deleteSheet', sheet: 'Plan' },
    ] as const
    const books = [before]
    for (const op of ops) {
      expect(s.restructure(op).ok).toBe(true)
      books.push(structuredClone(s.book()))
    }
    expect(s.book().sheets).toEqual([{ name: 'R', hidden: true, cells: { A1: { formula: '#REF!B1+1' }, A2: { formula: 'SUM(#REF!A:A)' }, C5: { value: 1 }, C6: { value: 2 } } }, { name: 'N', cells: {} }])
    for (let k = ops.length - 1; k >= 0; k--) {
      const u = s.undo()!
      expect(u.call).toBe('restructure')
      same(s, books[k], o)
    }
    for (let k = 1; k <= ops.length; k++) {
      s.redo()
      same(s, books[k], o)
    }
  })

  it('a new change forgets the redo; setClock, refused calls and reads are no steps', () => {
    const s = openSession(book(), o)
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 6 } }])
    s.undo()
    expect(s.history).toEqual({ undo: 1, redo: 1 })
    s.setClock({ now: '2027-01-01T00:00:00Z', seed: 1 })
    s.cells('S', 'A1:B2')
    s.validateEntry('S', 'A1', { value: 1 })
    expect(() => s.setCells([{ sheet: 'Nope', cell: 'A1', input: null }])).toThrow(InputError)
    expect(s.restructure({ op: 'deleteSheet', sheet: 'S' }).ok).toBe(true)
    s.undo()
    expect(s.restructure({ op: 'renameSheet', sheet: 'S', name: 'R' })).toMatchObject({ ok: false, code: 'sheet_name' })
    expect(() => s.setLayout([{ sheet: 'S', rows: '1', set: { style: 5 } }])).toThrow(InputError)
    expect(s.history).toEqual({ undo: 1, redo: 1 })
    s.setCells([])
    s.patchCells([])
    expect(s.history).toEqual({ undo: 1, redo: 1 })
    s.patchCells([{ sheet: 'S', cell: 'A1', set: { style: 0 } }])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
  })

  it('a call stopped by its limits is no step', () => {
    const s = openSession(book(), { ...o, limits: { cells: 1000 } })
    expect(() => s.setCells([{ sheet: 'S', cell: 'A9', input: { formula: 'SUM(SEQUENCE(2000))' } }])).toThrow(CalcError)
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    s.setCells([{ sheet: 'S', cell: 'A9', input: { formula: 'SUM(SEQUENCE(20))' } }])
    s.undo()
    expect(s.book().sheets[0].cells.A9).toBeUndefined()
  })

  it('a live session takes a new clock when an undo recomputes; volatile cells too', () => {
    const s = openSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1+RAND()*0' }, C1: { formula: 'RAND()' } } }] }, { live: true, execution })
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }])
    const [at, rand] = [s.clock, s.cells('S', 'C1').C1.value]
    const u = s.undo()!
    expect(s.clock).not.toEqual(at)
    expect(u.changes.find((c) => c.cell === 'C1')?.value).not.toEqual(rand)
    expect(s.hash()).toBe(compute(s.book(), { clock: s.clock, execution }).hash)
  })

  it('depth: the oldest steps go; 0 keeps none', () => {
    const s = openSession(book(), { ...o, undo: { depth: 2 } })
    for (const v of [5, 6, 7]) s.setCells([{ sheet: 'S', cell: 'A1', input: { value: v } }])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
    s.undo()
    s.undo()
    expect(s.undo()).toBeNull()
    expect(s.book().sheets[0].cells.A1).toEqual({ value: 5 })
    const none = openSession(book(), { ...o, undo: { depth: 0 } })
    none.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
    none.restructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(none.history).toEqual({ undo: 0, redo: 0 })
    expect(none.undo()).toBeNull()
    expect(() => openSession(book(), { ...o, undo: { depth: -1 } })).toThrow('opts.undo.depth')
    expect(() => openSession(book(), { ...o, undo: { size: 1 } } as never)).toThrow('opts.undo: unknown field "size"')
  })

  it('maxMb: the oldest steps go past it; a step larger than it alone forgets the history, as Excel forgets what it cannot undo', () => {
    // about 3 KB a step of 20 cells of 60 characters
    const big = (v: string) => Array.from({ length: 20 }, (_, i) => ({ sheet: 'S', cell: `K${i + 1}`, input: { value: v.repeat(60) } }))
    const s = openSession(book(), { ...o, undo: { maxMb: 8 / 1024 } })
    for (const v of 'abcd') s.setCells(big(v))
    expect(s.history.undo).toBeGreaterThan(0)
    expect(s.history.undo).toBeLessThan(4)
    const kept = s.history.undo
    for (let i = 0; i < kept; i++) s.undo()
    expect(s.undo()).toBeNull()
    expect(s.book().sheets[0].cells.K1.value).toBe('abcd'[3 - kept].repeat(60))
    // the step keeps the inputs it replaces: 20 values of 600 characters, over 8 KB
    s.setCells(big('e'.repeat(10)))
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    s.setCells(big('f'))
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    expect(s.undo()).toBeNull()
  })
})
