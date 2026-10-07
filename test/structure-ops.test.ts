/** `restructure` beyond Excel's answers: the API, refusals, mapCell, cells' format and meta, formula text. */
import { describe, expect, it } from 'vitest'
import { compute, InputError, restructure, type BookInput, type StructureOp } from '../src/index.js'

const book = (): BookInput => ({
  sheets: [
    { name: 'S', cells: { A1: { value: 1 }, A5: { value: 5, format: '0.00', meta: { id: 7 } }, B1: { formula: 'sum( a5 , $A$1 ) + Нет!A5' }, B2: { formula: 'SUM(A5' }, B3: { formula: 'A5*2' } } },
    { name: 'Лист 2', cells: { A1: { formula: 'S!A5+A5' }, A5: { value: 50 } } },
  ],
  names: [
    { name: 'five', ref: "'S'!$A$5" },
    { name: 'loc', ref: "'S'!$A$1:$A$5", scope: 'S' },
  ],
})

function ok(b: BookInput, op: StructureOp) {
  const r = restructure(b, op)
  if (!r.ok) throw new Error(r.message)
  return r
}

const cells = (b: BookInput, sheet = 'S') => b.sheets.find((s) => s.name === sheet)!.cells

describe('restructure', () => {
  it('rewrites only the references that change, in place; formulas that do not parse and unknown sheets stay as written', () => {
    const r = ok(book(), { op: 'insertRows', sheet: 'S', row: 3, count: 2 })
    expect(cells(r.book).B1.formula).toBe('sum( A7 , $A$1 ) + Нет!A5')
    expect(cells(r.book).B2).toEqual({ formula: 'SUM(A5' })
    expect(cells(r.book).B5.formula).toBe('A7*2')
    expect(cells(r.book, 'Лист 2').A1.formula).toBe('S!A7+A5')
    expect(r.book.names).toEqual([
      { name: 'five', ref: "'S'!$A$7" },
      { name: 'loc', ref: "'S'!$A$1:$A$7", scope: 'S' },
    ])
  })

  it('cells keep their format and meta when they move; copies have no meta unless withMeta', () => {
    const r = ok(book(), { op: 'insertRows', sheet: 'S', row: 2 })
    expect(cells(r.book).A6).toEqual({ value: 5, format: '0.00', meta: { id: 7 } })
    expect(Object.keys(cells(r.book)).sort()).toEqual(['A1', 'A6', 'B1', 'B3', 'B4'])
    const m = ok(book(), { op: 'moveRange', from: { sheet: 'S', range: 'A5' }, to: { sheet: 'Лист 2', cell: 'C1' } })
    expect(cells(m.book, 'Лист 2').C1).toEqual(cells(book()).A5)
    const c = ok(book(), { op: 'copyRange', from: { sheet: 'S', range: 'A5' }, to: { sheet: 'Лист 2', range: 'C1:C2' } })
    expect([cells(c.book, 'Лист 2').C1, cells(c.book, 'Лист 2').C2]).toEqual([{ value: 5, format: '0.00' }, { value: 5, format: '0.00' }])
    const w = ok(book(), { op: 'copyRange', from: { sheet: 'S', range: 'A5' }, to: { sheet: 'Лист 2', range: 'C1' }, withMeta: true })
    expect(cells(w.book, 'Лист 2').C1).toEqual(cells(book()).A5)
  })

  it('does not change the input book and shares what it does not change', () => {
    const b = book()
    const frozen = JSON.stringify(b)
    const r = ok(b, { op: 'deleteRows', sheet: 'Лист 2', row: 6, count: 2 })
    expect(JSON.stringify(b)).toBe(frozen)
    expect(r.book.sheets[0]).toBe(b.sheets[0])
    expect(r.book.sheets[1].cells.A1).toBe(b.sheets[1].cells.A1)
    expect(r.book.names).toBe(b.names)
  })

  it('a deleted reference computes to #REF!; a formula on a deleted sheet is #REF! followed by the reference', () => {
    const r = ok(book(), { op: 'deleteRows', sheet: 'S', row: 5 })
    expect(cells(r.book).B3.formula).toBe('#REF!*2')
    expect(cells(r.book, 'Лист 2').A1.formula).toBe('S!#REF!+A5')
    expect(r.book.names![0]).toEqual({ name: 'five', ref: "'S'!#REF!" })
    expect(compute(r.book).sheets[0].cells.B3.value).toEqual({ error: 'REF' })
    const d = ok(book(), { op: 'deleteSheet', sheet: 's' })
    expect(d.book.sheets.map((s) => s.name)).toEqual(['Лист 2'])
    expect(cells(d.book, 'Лист 2').A1.formula).toBe('#REF!A5+A5')
    expect(d.book.names).toEqual([{ name: 'five', ref: '#REF!$A$5' }])
    expect(compute(d.book).sheets[0].cells.A1.value).toEqual({ error: 'REF' })
  })

  it('renames a sheet in references, names and scopes; a moved formula names the sheet it reads', () => {
    const r = ok(book(), { op: 'renameSheet', sheet: 'S', name: "It's" })
    expect(cells(r.book, 'Лист 2').A1.formula).toBe("'It''s'!A5+A5")
    expect(r.book.names).toEqual([
      { name: 'five', ref: "'It''s'!$A$5" },
      { name: 'loc', ref: "'It''s'!$A$1:$A$5", scope: "It's" },
    ])
    const m = ok(book(), { op: 'moveRange', from: { sheet: 'S', range: 'B3' }, to: { sheet: 'Лист 2', cell: 'B3' } })
    expect(cells(m.book, 'Лист 2').B3.formula).toBe("'S'!A5*2")
  })

  it('mapCell: where each old cell is now', () => {
    const ins = ok(book(), { op: 'insertColumns', sheet: 'S', column: 'A', count: 2 })
    expect([ins.mapCell({ sheet: 's', cell: 'A5' }), ins.mapCell({ sheet: 'Лист 2', cell: 'A5' })]).toEqual([
      { sheet: 'S', cell: 'C5' },
      { sheet: 'Лист 2', cell: 'A5' },
    ])
    const del = ok(book(), { op: 'deleteRows', sheet: 'S', row: 2, count: 2 })
    expect([del.mapCell({ sheet: 'S', cell: 'A2' }), del.mapCell({ sheet: 'S', cell: 'A5' })]).toEqual([null, { sheet: 'S', cell: 'A3' }])
    const mv = ok(book(), { op: 'moveRange', from: { sheet: 'S', range: 'A1:A2' }, to: { sheet: 'S', cell: 'A5' } })
    expect([mv.mapCell({ sheet: 'S', cell: 'A1' }), mv.mapCell({ sheet: 'S', cell: 'A6' }), mv.mapCell({ sheet: 'S', cell: 'A7' })]).toEqual([{ sheet: 'S', cell: 'A5' }, null, { sheet: 'S', cell: 'A7' }])
    const cp = ok(book(), { op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'C1:C3' } })
    expect([cp.mapCell({ sheet: 'S', cell: 'A1' }), cp.mapCell({ sheet: 'S', cell: 'C2' })]).toEqual([{ sheet: 'S', cell: 'A1' }, null])
    expect(ok(book(), { op: 'renameSheet', sheet: 'S', name: 'T' }).mapCell({ sheet: 'S', cell: 'A1' })).toEqual({ sheet: 'T', cell: 'A1' })
    expect(ok(book(), { op: 'deleteSheet', sheet: 'S' }).mapCell({ sheet: 'S', cell: 'A1' })).toBe(null)
  })

  it('tiles a copy over a target that is a whole multiple of the source, else pastes it once', () => {
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1*2' } } }] }
    const tiled = ok(b, { op: 'copyRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', range: 'D1:G3' } })
    expect(Object.keys(cells(tiled.book)).sort()).toEqual(['A1', 'B1', 'D1', 'D2', 'D3', 'E1', 'E2', 'E3', 'F1', 'F2', 'F3', 'G1', 'G2', 'G3'])
    expect([cells(tiled.book).G3.formula, cells(tiled.book).E2.formula]).toEqual(['F3*2', 'D2*2'])
    const once = ok(b, { op: 'copyRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', range: 'D1:F3' } })
    expect(Object.keys(cells(once.book)).sort()).toEqual(['A1', 'B1', 'D1', 'E1'])
  })

  it('refuses what Excel refuses', () => {
    const edge: BookInput = { sheets: [{ name: 'S', cells: { A1048576: { format: '0' }, XFD1: { meta: 1 } } }] }
    expect(restructure(edge, { op: 'insertRows', sheet: 'S', row: 1 })).toMatchObject({ ok: false, code: 'off_sheet' })
    expect(restructure(edge, { op: 'insertColumns', sheet: 'S', column: 'B' })).toMatchObject({ ok: false, code: 'off_sheet' })
    expect(restructure(edge, { op: 'deleteSheet', sheet: 'S' })).toMatchObject({ ok: false, code: 'last_sheet' })
    expect(restructure(book(), { op: 'renameSheet', sheet: 'S', name: 'лист 2' })).toMatchObject({ ok: false, code: 'sheet_name' })
    expect(restructure(book(), { op: 'renameSheet', sheet: 'S', name: '' })).toMatchObject({ ok: false, code: 'sheet_name' })
    expect(restructure(book(), { op: 'renameSheet', sheet: 'S', name: 's' }).ok).toBe(true)
    expect(restructure(book(), { op: 'moveRange', from: { sheet: 'S', range: 'A1:A5' }, to: { sheet: 'S', cell: 'A1048575' } })).toMatchObject({ ok: false, code: 'off_sheet' })
    expect(restructure(book(), { op: 'copyRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', range: 'XFD1' } })).toMatchObject({ ok: false, code: 'off_sheet' })
  })

  it('throws InputError on input of the wrong shape', () => {
    const fails = (op: unknown, message: string) => expect(() => restructure(book(), op as StructureOp)).toThrow(new InputError(message.split(': ')[0], message.slice(message.indexOf(': ') + 2)))
    fails({ op: 'insertRow', sheet: 'S', row: 1 }, 'op.op: expected one of insertRows, deleteRows, insertColumns, deleteColumns, moveRange, copyRange, renameSheet, deleteSheet, insertSheet, moveSheet, got "insertRow"')
    fails({ op: 'insertRows', sheet: 'Нет', row: 1 }, 'op.sheet: unknown sheet "Нет"')
    fails({ op: 'insertRows', sheet: 'S', row: 0 }, 'op.row: expected an integer from 1 to 1048576, got 0')
    fails({ op: 'deleteRows', sheet: 'S', row: 1048576, count: 2 }, 'op.count: expected an integer from 1 to 1, got 2')
    fails({ op: 'insertColumns', sheet: 'S', column: 'xfe' }, 'op.column: expected column letters from A to XFD, got "xfe"')
    fails({ op: 'deleteSheet', sheet: 'S', count: 1 }, 'op: unknown field "count" (expected op, sheet)')
    fails({ op: 'moveRange', from: { sheet: 'S', range: 'A1:' }, to: { sheet: 'S', cell: 'B1' } }, 'op.from.range: expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got "A1:"')
    fails({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', cell: 'B1' } }, 'op.to: unknown field "cell" (expected sheet, range)')
    fails({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B1' }, withMeta: 1 }, 'op.withMeta: expected a boolean, got 1')
    fails({ op: 'moveRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', cell: 'B1' }, withMeta: true }, 'op: unknown field "withMeta" (expected op, from, to)')
    const r = ok(book(), { op: 'insertRows', sheet: 'S', row: 1 })
    expect(() => r.mapCell({ sheet: 'S', cell: 'a1' })).toThrow(InputError)
    expect(() => r.mapCell({ sheet: 'Нет', cell: 'A1' })).toThrow('ref.sheet: unknown sheet "Нет"')
  })
})
