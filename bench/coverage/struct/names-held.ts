/**
 * Held out: names under structure operations, predicted by the engine before Excel took them (`names-held.predicted.json`,
 * from `predict.ts`). A reference to a sheet's own name (`S!loc`) in cells and in a global name when the sheet is renamed
 * or deleted; names whose references mix `$` and offsets (`$A1`, `A$12`, `A$1`, `$B$2:$B1`, `$A$2:$A3`, `$A1048574:$A$12`)
 * when rows and columns are inserted and deleted, each used in a cell beside a cell formula that reads the same cells.
 * The values the cells show after the same operations are measured by the tables set `name-values-held`.
 */
import type { BookInput, CellInput, StructureOp } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'
import { formulaCell as f } from '../stored.js'

/** S: 42 in A1 and a sheet-scoped name `loc` over it; T: `S!loc` in cells and in the global name `g`. */
export function qualifiedBook(): BookInput {
  return {
    names: [
      { name: 'loc', ref: "'S'!$A$1", scope: 'S' },
      { name: 'g', ref: 'S!loc+1' },
    ],
    sheets: [
      { name: 'S', cells: { A1: { value: 42 }, B1: f('loc*3') } },
      { name: 'T', cells: { A1: f('S!loc*2'), A2: f("SUM('S'!loc)"), A3: f('g'), A4: f('S!A1') } },
    ],
  }
}

/** Mixed names over values in A1:A10, row 12 and row 1: each used in column C, the same reference in a cell beside it. */
export function mixedBook(): BookInput {
  const cells: Record<string, CellInput> = { C5: f('mix'), E5: f('$A5'), C14: f('mixc'), E14: f('C$12'), C16: f('mrow'), E16: f('C$1'), C7: f('rel') }
  for (let r = 1; r <= 10; r++) cells[`A${r}`] = { value: r * 100 }
  'ABCDEF'.split('').forEach((c, i) => (cells[`${c}12`] = { value: (i + 1) * 7 }))
  'BCDEF'.split('').forEach((c, i) => (cells[`${c}1`] = { value: (i + 1) * 11 }))
  return {
    names: [
      { name: 'mix', ref: "'S'!$A1" },
      { name: 'mixc', ref: "'S'!A$12" },
      { name: 'mrow', ref: "'S'!A$1" },
      { name: 'rel', ref: "'S'!XFD1" },
    ],
    sheets: [{ name: 'S', cells }],
  }
}

/** A running total by a name, `run` = `$B$2:$B1` in C2:C6, the same sum written in D2:D6. */
export function runBook(): BookInput {
  const cells: Record<string, CellInput> = { B1: { value: 'amt' } }
  for (let r = 2; r <= 6; r++) Object.assign(cells, { [`B${r}`]: { value: r * 10 }, [`C${r}`]: f('SUM(run)'), [`D${r}`]: f(`SUM($B$2:$B${r})`) })
  return { names: [{ name: 'run', ref: "'S'!$B$2:$B1" }], sheets: [{ name: 'S', cells }] }
}

/** Areas with a fixed and an offset bound on rows: `part` (rows 2 to two below) and `top` (three above to row 12), used at row 10. */
export function areaBook(): BookInput {
  const cells: Record<string, CellInput> = { C10: f('SUM(part)'), D10: f('SUM($A$2:$A12)'), F10: f('SUM(top)'), G10: f('SUM($A7:$A$12)') }
  for (let r = 1; r <= 20; r++) cells[`A${r}`] = { value: r }
  return {
    names: [
      { name: 'part', ref: "'S'!$A$2:$A3" },
      { name: 'top', ref: "'S'!$A1048574:$A$12" },
    ],
    sheets: [{ name: 'S', cells }],
  }
}

const rows = (op: 'insertRows' | 'deleteRows', row: number, count?: number): StructureOp => ({ op, sheet: 'S', row, ...(count ? { count } : {}) })
const columns = (op: 'insertColumns' | 'deleteColumns', column: string): StructureOp => ({ op, sheet: 'S', column })

/** The operations by book: `[id, op]`; the tables set takes the same. */
export const OPS: Record<string, [string, StructureOp][]> = {
  qualified: [
    ['rename-U', { op: 'renameSheet', sheet: 'S', name: 'U' }],
    ['rename-quoted', { op: 'renameSheet', sheet: 'S', name: 'My Sheet' }],
    ['delete-S', { op: 'deleteSheet', sheet: 'S' }],
  ],
  mixed: [
    ['insert-row-1', rows('insertRows', 1)],
    ['delete-row-1', rows('deleteRows', 1)],
    ['insert-column-A', columns('insertColumns', 'A')],
    ['delete-column-A', columns('deleteColumns', 'A')],
    ['insert-row-3', rows('insertRows', 3)],
  ],
  run: [
    ['delete-row-2', rows('deleteRows', 2)],
    ['delete-row-4', rows('deleteRows', 4)],
    ['insert-row-2', rows('insertRows', 2)],
  ],
  area: [
    ['delete-row-2', rows('deleteRows', 2)],
    ['delete-rows-2-3', rows('deleteRows', 2, 2)],
    ['delete-row-12', rows('deleteRows', 12)],
    ['insert-row-2', rows('insertRows', 2)],
    ['insert-row-1', rows('insertRows', 1)],
    ['delete-rows-1-12', rows('deleteRows', 1, 12)],
  ],
}

export const BOOKS: Record<string, () => BookInput> = { qualified: qualifiedBook, mixed: mixedBook, run: runBook, area: areaBook }

export function buildStructCases(): StructCases {
  const cases: StructCase[] = Object.entries(OPS).flatMap(([book, ops]) => ops.map(([id, op]) => ({ id: `${book}-${id}`, book, op })))
  return { books: Object.fromEntries(Object.entries(BOOKS).map(([k, b]) => [k, b()])), cases }
}
