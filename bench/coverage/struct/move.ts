/**
 * Moving (cut and paste) and copying ranges of the base book and of a small fill book: references to the
 * moved cells, to the overwritten ones and inside the moved formulas; relative, absolute and mixed parts
 * of copied formulas, copies past the sheet's edge, tiling.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { RangeRef } from '../../../src/structure/types.js'
import { baseBook, type StructCase, type StructCases } from './cases.js'

const f = (formula: string): CellInput => ({ formula })

/** Data in A1:B10, a row of formulas in D1:J1 to fill down and right. */
function fillBook(): BookInput {
  const s: Record<string, CellInput> = {}
  for (let r = 1; r <= 10; r++) {
    s[`A${r}`] = { value: r }
    s[`B${r}`] = { value: r * 2 }
  }
  Object.assign(s, {
    D1: f('A1*2'),
    E1: f('$A1+B$1'),
    F1: f('SUM(A$1:A1)'),
    G1: f('S!A1'),
    H1: f('SUM(B:B)'),
    I1: f('x*A1'),
    J1: f('SUM(A1:B1)'),
  })
  return { sheets: [{ name: 'S', cells: s }, { name: 'T', cells: { A1: f('SUM(S!D1:D10)') } }], names: [{ name: 'x', ref: "'S'!$B$1" }] }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const at = (ref: string): RangeRef => {
    const i = ref.lastIndexOf('!')
    return { sheet: ref.slice(0, i), range: ref.slice(i + 1) }
  }
  const move = (id: string, from: string, to: string, book = 'base') => {
    const t = at(to)
    cases.push({ id, book, op: { op: 'moveRange', from: at(from), to: { sheet: t.sheet, cell: t.range } } })
  }
  const copy = (id: string, from: string, to: string, book = 'base') => cases.push({ id, book, op: { op: 'copyRange', from: at(from), to: at(to) } })

  move('move-A5-K5', 'S!A5', 'S!K5')
  move('move-A3:A8-K3', 'S!A3:A8', 'S!K3')
  move('move-A3:A8-T', 'S!A3:A8', 'T!K3')
  move('move-A5:A6-K20', 'S!A5:A6', 'S!K20')
  move('move-A4:B4-K4', 'S!A4:B4', 'S!K4')
  move('move-F1:G12-K1', 'S!F1:G12', 'S!K1')
  move('move-F1:H12-T', 'S!F1:H12', 'T!K1')
  move('move-A1-A5', 'S!A1', 'S!A5')
  move('move-B3:B8-A3', 'S!B3:B8', 'S!A3')
  move('move-A3:A8-A5', 'S!A3:A8', 'S!A5')
  move('move-cols-B-J', 'S!B:B', 'S!J1')
  move('move-rows-30:31-40', 'S!30:31', 'S!A40')
  move('move-MyData-S', 'My Data!A1:A3', 'S!K1')
  move('move-T-S', 'T!A1:A12', 'S!M1')

  copy('copy-F1:G12-K1', 'S!F1:G12', 'S!K1')
  copy('copy-F1:F12-F21', 'S!F1:F12', 'S!F21')
  copy('copy-G6-G2', 'S!G6', 'S!G2')
  copy('copy-F1:F4-A21', 'S!F1:F4', 'S!A21')
  copy('copy-F1:H12-T', 'S!F1:H12', 'T!K1')
  copy('copy-T-A1:A12-C1', 'T!A1:A12', 'T!C1')
  copy('fill-down', 'S!D1:J1', 'S!D2:J10', 'fill')
  copy('fill-right', 'S!D1', 'S!D1:H1', 'fill')
  copy('tile-2x2', 'S!D1:E2', 'S!D11:G14', 'fill')
  copy('copy-empty', 'S!M1:N2', 'S!A1', 'fill')
  copy('copy-cols-whole', 'S!H1', 'S!L5', 'fill')
  copy('copy-fill-T', 'S!D1:J1', 'T!D3', 'fill')
  copy('copy-not-multiple', 'S!D1:E2', 'S!D11:F13', 'fill')
  return { books: { base: baseBook(), fill: fillBook() }, cases }
}
