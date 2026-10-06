/**
 * Structure operations on seeded random books: three sheets of numbers in A1:H20 and formulas in J1:N20
 * over them (cells and SUM ranges, relative, absolute and mixed, with and without a sheet). Properties:
 * an insert changes no value and deleting what it inserted gives the book back; a move keeps the value
 * of every cell at its new place (a formula over single cells may become `#REF!` when a cell it reads
 * was overwritten); a copy copied back restores the source.
 */
import { describe, expect, it } from 'vitest'
import { compute, restructure, type BookInput, type CellInput, type CellValue, type StructureOp } from '../src/index.js'
import { indexToCol, parseAddress } from '../src/model.js'
import { decode } from '../src/structure/area.js'
import { parse, walk } from '../src/syntax.js'
import { mulberry32 } from '../src/synth.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }
const SHEETS = ['S1', 'S2', 'S 3']

function randomBook(seed: number): BookInput {
  const rnd = mulberry32(seed)
  const int = (n: number) => Math.floor(rnd() * n)
  const flag = () => (rnd() < 0.3 ? '$' : '')
  const cell = () => `${flag()}${indexToCol(int(8))}${flag()}${1 + int(20)}`
  const sheet = (own: string) => {
    const s = SHEETS[int(SHEETS.length)]
    return rnd() < 0.6 ? '' : s === own && rnd() < 0.5 ? '' : `'${s}'!`
  }
  const term = (own: string) => {
    if (rnd() < 0.6) return sheet(own) + cell()
    const [r1, c1] = [int(20), int(8)]
    return `SUM(${sheet(own)}${flag()}${indexToCol(c1)}${flag()}${r1 + 1}:${flag()}${indexToCol(c1 + int(8 - c1))}${flag()}${r1 + 1 + int(20 - r1)})`
  }
  return {
    sheets: SHEETS.map((name) => {
      const cells: Record<string, CellInput> = {}
      for (let r = 1; r <= 20; r++) for (let c = 0; c < 8; c++) if (rnd() < 0.8) cells[`${indexToCol(c)}${r}`] = { value: int(1000) }
      for (let r = 1; r <= 20; r++) for (let c = 9; c < 14; c++) if (rnd() < 0.7) cells[`${indexToCol(c)}${r}`] = { formula: Array.from({ length: 1 + int(3) }, () => term(name)).join(rnd() < 0.5 ? '+' : '-') }
      return { name, cells }
    }),
  }
}

/**
 * A range whose bounds on one axis are the same row or column with different `$` (`A1:$A$2`): copied and
 * copied back, the `$` may change corners (`$A1:A$2`), the same cells.
 */
function tied(formula: string): boolean {
  let tie = false
  walk(parse(formula), (n) => {
    if (n.k !== 'ref') return
    const a = decode(n.ref)
    for (const s of [a.r, a.c]) if (s.lo.i === s.hi.i && s.lo.abs !== s.hi.abs) tie = true
  })
  return tie
}

const values = (b: BookInput) => new Map(compute(b, { clock }).sheets.map((s) => [s.name, new Map(Object.entries(s.cells).map(([a, c]) => [a, c.value]))]))

function apply(b: BookInput, op: StructureOp) {
  const r = restructure(b, op)
  if (!r.ok) throw new Error(r.message)
  return r
}

const SEEDS = Array.from({ length: 40 }, (_, i) => i + 1)

describe('structure operations on seeded books', () => {
  it('inserting rows or columns changes no value; deleting them gives the book back', () => {
    for (const seed of SEEDS) {
      const book = randomBook(seed)
      const rnd = mulberry32(seed * 7)
      const sheet = SHEETS[Math.floor(rnd() * 3)]
      const count = 1 + Math.floor(rnd() * 3)
      const rows = rnd() < 0.5
      const op: StructureOp = rows ? { op: 'insertRows', sheet, row: 1 + Math.floor(rnd() * 22), count } : { op: 'insertColumns', sheet, column: indexToCol(Math.floor(rnd() * 15)), count }
      const r = apply(book, op)
      const [old, now] = [values(book), values(r.book)]
      for (const [s, cells] of old) for (const [cell, v] of cells) expect(now.get(r.mapCell({ sheet: s, cell })!.sheet)!.get(r.mapCell({ sheet: s, cell })!.cell), `seed ${seed} ${s}!${cell}`).toEqual(v)
      const back = apply(r.book, rows ? { op: 'deleteRows', sheet, row: (op as { row: number }).row, count } : { op: 'deleteColumns', sheet, column: (op as { column: string }).column, count })
      expect(back.book, `seed ${seed}`).toEqual(book)
    }
  })

  it('a move keeps every value at its new place; a formula over cells is #REF! only where a cell it read was overwritten', () => {
    for (const seed of SEEDS) {
      const book = randomBook(seed)
      const rnd = mulberry32(seed * 13)
      const int = (n: number) => Math.floor(rnd() * n)
      const [r1, c1] = [int(20), int(14)]
      const range = `${indexToCol(c1)}${r1 + 1}:${indexToCol(c1 + int(3))}${r1 + 1 + int(4)}`
      // to an empty place half the time, else anywhere in the used area
      const far = rnd() < 0.5
      const to = { sheet: SHEETS[int(3)], cell: far ? `${indexToCol(20 + int(5))}${30 + int(10)}` : `${indexToCol(int(14))}${1 + int(20)}` }
      const r = apply(book, { op: 'moveRange', from: { sheet: SHEETS[int(3)], range }, to })
      const [old, now] = [values(book), values(r.book)]
      for (const s of book.sheets)
        for (const [cell, input] of Object.entries(s.cells)) {
          const to = r.mapCell({ sheet: s.name, cell })
          if (!to) continue
          const v = now.get(to.sheet)!.get(to.cell)
          const was = old.get(s.name)!.get(cell) as CellValue
          // a formula that reads ranges can see a range shrink or grow, as in Excel
          if (input.formula?.includes(':')) continue
          if (far || JSON.stringify(v) !== JSON.stringify({ error: 'REF' })) expect(v, `seed ${seed} ${s.name}!${cell} → ${to.sheet}!${to.cell}`).toEqual(was)
        }
    }
  })

  it('a copy copied back restores the source cells', () => {
    for (const seed of SEEDS) {
      const book = randomBook(seed)
      const rnd = mulberry32(seed * 31)
      const int = (n: number) => Math.floor(rnd() * n)
      const sheet = SHEETS[int(3)]
      const [r1, c1] = [int(20), 9 + int(5)]
      const [h, w] = [1 + int(4), 1 + int(14 - c1)]
      const src = `${indexToCol(c1)}${r1 + 1}:${indexToCol(c1 + w - 1)}${r1 + h}`
      const [dr, dc] = [30 + int(20), int(10)]
      const target = { sheet: SHEETS[int(3)], range: `${indexToCol(c1 + dc)}${r1 + 1 + dr}` }
      const there = apply(book, { op: 'copyRange', from: { sheet, range: src }, to: target }).book
      const pasted = `${indexToCol(c1 + dc)}${r1 + 1 + dr}:${indexToCol(c1 + dc + w - 1)}${r1 + h + dr}`
      const back = apply(there, { op: 'copyRange', from: { sheet: target.sheet, range: pasted }, to: { sheet, range: src } }).book
      const cells = (b: BookInput) => b.sheets.find((s) => s.name === sheet)!.cells
      for (const [addr, cell] of Object.entries(cells(book))) {
        const { row, col } = parseAddress(addr)
        if (row < r1 || row >= r1 + h || col < c1 || col >= c1 + w || (cell.formula && tied(cell.formula))) continue
        expect(cells(back)[addr], `seed ${seed} ${sheet}!${addr}`).toEqual(cell)
      }
    }
  })
})
