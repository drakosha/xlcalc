/**
 * The geometry of structure operations (`structure/geometry.ts`) against where the operations put the cells: every cell
 * of the old book the operation keeps as it was is in the new book where `mapCell` says, and no other; the inverse takes
 * back every place the geometry brings a cell to and nothing else; how addresses moved (`movesOf`) for each kind of
 * move. And the cells of a sheet read once for an operation (`structure/grid.ts`).
 */
import { describe, expect, it } from 'vitest'
import { restructure, type BookInput, type CellInput, type StructureOp } from '../src/index.js'
import { formatAddress, readAddress } from '../src/model.js'
import { inverse, movesOf, placer, type Place } from '../src/structure/geometry.js'
import { cellsIn, gridOf, withGrids } from '../src/structure/grid.js'
import { applyOp, namesOf } from '../src/structure/index.js'

/** Values over A1:F30 of S and A1:C10 of D, a table on T emptied by deleting its data, rows of S hidden by a filter in some books. */
function book(filtered = false): BookInput {
  const cells = (rows: number, cols: number, base: number) => {
    const out: Record<string, CellInput> = {}
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) out[formatAddress(r, c)] = { value: base + r * 10 + c }
    return out
  }
  return {
    sheets: [
      { name: 'S', cells: cells(30, 6, 0), ...(filtered ? { filtered: true, rows: { '3': { hidden: true }, '5:6': { hidden: true } } } : {}) },
      { name: 'D', cells: cells(10, 3, 1000) },
      {
        name: 'T',
        cells: { A1: { value: 'Qty' }, B1: { value: 'Price' }, A2: { value: 3 }, B2: { value: 1.5 }, A3: { value: 4 }, B3: { value: 2.5 }, D5: { value: 9 } },
        tables: [{ name: 'Tab', range: 'A1:B3', columns: [{ name: 'Qty' }, { name: 'Price' }] }],
      },
    ],
  }
}

const OPS: [string, StructureOp, boolean?][] = [
  ['rows inserted', { op: 'insertRows', sheet: 'S', row: 4, count: 3 }],
  ['rows deleted', { op: 'deleteRows', sheet: 'S', row: 2, count: 4 }],
  ['columns inserted', { op: 'insertColumns', sheet: 'D', column: 'B', count: 2 }],
  ['columns deleted', { op: 'deleteColumns', sheet: 'S', column: 'C', count: 2 }],
  ['a range moved down over itself', { op: 'moveRange', from: { sheet: 'S', range: 'B3:D8' }, to: { sheet: 'S', cell: 'C6' } }],
  ['a range moved to another sheet', { op: 'moveRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'D', cell: 'B2' } }],
  ['a row filled down', { op: 'copyRange', from: { sheet: 'S', range: 'A2:F2' }, to: { sheet: 'S', range: 'A3:F12' } }],
  ['a copy pasted once', { op: 'copyRange', from: { sheet: 'S', range: 'A1:B3' }, to: { sheet: 'D', range: 'A2:A3' } }],
  ['a sheet renamed', { op: 'renameSheet', sheet: 'D', name: 'Data' }],
  ['a sheet inserted', { op: 'insertSheet', name: 'N', before: 'D' }],
  ['a sheet deleted', { op: 'deleteSheet', sheet: 'D' }],
  ['a sheet moved', { op: 'moveSheet', sheet: 'S', before: 'T' }],
  ['a table deleted with its cells', { op: 'deleteTable', table: 'Tab' }],
  ["every data row of a table deleted: its first one kept, cleared", { op: 'deleteRows', sheet: 'T', row: 2, count: 2 }],
  ['visible rows deleted on a filtered sheet', { op: 'deleteRows', sheet: 'S', row: 2, count: 6 }, true],
  ['rows inserted among hidden ones', { op: 'insertRows', sheet: 'S', row: 3, count: 4 }, true],
  ['a copy over hidden rows', { op: 'copyRange', from: { sheet: 'D', range: 'A1:C1' }, to: { sheet: 'S', range: 'A2:C8' } }, true],
]

/** Each cell's input object of a book, by sheet and address. */
const inputs = (b: BookInput) => b.sheets.flatMap((s) => Object.entries(s.cells).map(([cell, input]) => ({ sheet: s.name, cell, input })))

describe('the geometry of each operation is where it puts the cells', () => {
  it.each(OPS)('%s', (_, op, filtered) => {
    const before = book(filtered)
    const r = restructure(before, op)
    if (!r.ok) throw new Error(r.message)
    const [was, now] = [inputs(before), inputs(r.book)]
    const old = new Set(was.map((x) => x.input))
    // a copy pastes its source's own inputs again; another operation keeps each cell it keeps once
    const copy = op.op === 'copyRange'
    for (const { sheet, cell, input } of was) {
      const to = r.mapCell({ sheet, cell })
      const there = now.filter((x) => x.input === input)
      if (to) expect(there.map(({ sheet, cell }) => ({ sheet, cell })), `${sheet}!${cell}`).toContainEqual(to)
      if (!copy) expect(there.map(({ sheet, cell }) => ({ sheet, cell })), `${sheet}!${cell}`).toEqual(to ? [to] : [])
    }
    // a cell of the new book kept from the old one came to where it is by mapCell
    if (copy) return
    for (const x of now) {
      if (!old.has(x.input)) continue
      const from = was.filter((y) => y.input === x.input && r.mapCell(y)?.sheet === x.sheet && r.mapCell(y)?.cell === x.cell)
      expect(from.length, `${x.sheet}!${x.cell}`).toBe(1)
    }
  })

  it.each(OPS)('%s: the inverse takes back every place the geometry brings a cell to, and no other', (_, op, filtered) => {
    const b = book(filtered)
    const r = applyOp(b, op, () => [])
    if (!r.ok) throw new Error(r.message)
    const [go, back] = [placer(r.geometry), placer(inverse(r.geometry))]
    const at = (s: number, row: number, col: number): Place => ({ s, r: row, c: col })
    const key = (p: Place) => `${p.s}:${p.r}:${p.c}`
    const reached = new Set<string>()
    for (let s = 0; s < b.sheets.length; s++)
      for (let row = 0; row < 40; row++)
        for (let col = 0; col < 10; col++) {
          const p = at(s, row, col)
          if (!go(p)) continue
          reached.add(key(p))
          const q = { ...p }
          expect(back(q) && key(q), `${s}:${row}:${col}`).toBe(`${s}:${row}:${col}`)
        }
    for (let s = 0; s < r.book.sheets.length; s++)
      for (let row = 0; row < 40; row++)
        for (let col = 0; col < 10; col++) {
          const q = at(s, row, col)
          const p = { ...q }
          if (!back(p) || p.r >= 40 || p.c >= 10) continue
          expect(reached.has(key(q)), key(q)).toBe(true)
        }
  })
})

describe('how addresses moved', () => {
  const moves = (op: StructureOp, filtered = false) => {
    const b = book(filtered)
    const r = applyOp(b, op, () => [])
    if (!r.ok) throw new Error(r.message)
    const [was, now] = [namesOf(b), namesOf(r.book)]
    return [movesOf(r.geometry, was, now), movesOf(inverse(r.geometry), now, was)]
  }

  it('lines deleted on a filtered sheet: each run of visible rows from the bottom up, and back in the other order', () => {
    // rows 2 to 7 hold the visible rows 2, 4 and 7
    expect(moves({ op: 'deleteRows', sheet: 'S', row: 2, count: 6 }, true)).toEqual([
      [7, 4, 2].map((at) => ({ kind: 'lines', sheet: 'S', axis: 'rows', at, count: -1 })),
      [2, 4, 7].map((at) => ({ kind: 'lines', sheet: 'S', axis: 'rows', at, count: 1 })),
    ])
  })

  it('rows inserted among hidden rows of a filtered sheet: as many as are visible', () => {
    // rows 3 to 6 hold the visible row 4
    expect(moves({ op: 'insertRows', sheet: 'S', row: 3, count: 4 }, true)[0]).toEqual([{ kind: 'lines', sheet: 'S', axis: 'rows', at: 3, count: 1 }])
  })

  it('a range moved, and back', () => {
    expect(moves({ op: 'moveRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'D', cell: 'B2' } })).toEqual([
      [{ kind: 'range', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'D', cell: 'B2' } }],
      [{ kind: 'range', from: { sheet: 'D', range: 'B2:C5' }, to: { sheet: 'S', cell: 'A1' } }],
    ])
  })

  it('a table emptied: the rows after its first data row, which stays', () => {
    expect(moves({ op: 'deleteRows', sheet: 'T', row: 2, count: 2 })[0]).toEqual([{ kind: 'lines', sheet: 'T', axis: 'rows', at: 3, count: -1 }])
  })

  it('sheets inserted, deleted, moved and renamed, and back', () => {
    expect(moves({ op: 'insertSheet', name: 'N', before: 'D' })).toEqual([[{ kind: 'insertSheet', sheet: 'N', index: 1 }], [{ kind: 'deleteSheet', sheet: 'N' }]])
    expect(moves({ op: 'deleteSheet', sheet: 'D' })).toEqual([[{ kind: 'deleteSheet', sheet: 'D' }], [{ kind: 'insertSheet', sheet: 'D', index: 1 }]])
    expect(moves({ op: 'moveSheet', sheet: 'S' })).toEqual([[{ kind: 'moveSheet', sheet: 'S', index: 2 }], [{ kind: 'moveSheet', sheet: 'S', index: 0 }]])
    expect(moves({ op: 'renameSheet', sheet: 'd', name: 'Data' })).toEqual([[{ kind: 'renameSheet', sheet: 'D', name: 'Data' }], [{ kind: 'renameSheet', sheet: 'Data', name: 'D' }]])
  })

  it('a copy and a table deleted move no address', () => {
    expect(moves({ op: 'copyRange', from: { sheet: 'S', range: 'A2:F2' }, to: { sheet: 'S', range: 'A3:F12' } })).toEqual([[], []])
    expect(moves({ op: 'deleteTable', table: 'Tab' })).toEqual([[], []])
  })
})

describe('mapCell', () => {
  it('a cell pushed past the last row by inserted rows is nowhere', () => {
    const r = restructure({ sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }, { op: 'insertRows', sheet: 'S', row: 1, count: 2 })
    if (!r.ok) throw new Error(r.message)
    expect([r.mapCell({ sheet: 's', cell: 'A1' }), r.mapCell({ sheet: 'S', cell: 'A1048575' }), r.mapCell({ sheet: 'S', cell: 'A1048574' })]).toEqual([{ sheet: 'S', cell: 'A3' }, null, { sheet: 'S', cell: 'A1048576' }])
  })
})

describe('the cells of a sheet read once for an operation', () => {
  const sheet = book().sheets[0]

  it('each cell with its row and column, in the sheet order; those of a rectangle', () => {
    const g = gridOf(sheet)
    expect(g.addr).toEqual(Object.keys(sheet.cells))
    const p = { r: 0, c: 0 }
    g.addr.forEach((a, i) => {
      readAddress(a, p)
      expect([g.row[i], g.col[i], g.cell[i]]).toEqual([p.r, p.c, sheet.cells[a]])
    })
    expect(cellsIn(g, { r1: 28, r2: 40, c1: 4, c2: 9 }).map((i) => g.addr[i])).toEqual(['E29', 'F29', 'E30', 'F30'])
  })

  it('shared by the steps of one operation, read again by the next', () => {
    const [a, b] = withGrids(() => [gridOf(sheet), gridOf(sheet)])
    expect(a).toBe(b)
    expect(withGrids(() => gridOf(sheet))).not.toBe(a)
    expect(gridOf(sheet)).not.toBe(gridOf(sheet))
  })

  it('a valid address read and written back', () => {
    const p = { r: 0, c: 0 }
    for (const a of ['A1', 'Z99', 'AA1', 'XFD1048576', 'ABC123']) {
      readAddress(a, p)
      expect(formatAddress(p.r, p.c)).toBe(a)
    }
  })
})
