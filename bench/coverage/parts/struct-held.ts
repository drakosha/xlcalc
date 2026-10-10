/**
 * Held out: notes, hyperlinks and frozen panes under structure operations, predicted by the engine before Excel ran them
 * (`struct-held.predicted.json`, from `predict.ts`). The engine moves a note and a link with their cell, drops them with
 * a deleted or overwritten cell (a pasted cell without one too), copies them with a copy, gives none to inserted lines,
 * keeps a link's location as written whatever happens to its sheet, and moves the split of frozen panes with the lines
 * above it.
 */
import type { BookInput, CellInput, StructureOp } from '../../../src/contract.js'
import type { PartsCase, PartsCases } from './cases.js'

/**
 * S: values in A1:A10, notes at B2, B5 and C8, links at B3 (outside the book), C5 (to S!A1) and D6 (to T), rows 1:2 and
 * column A frozen. T: a link at A2 to S!B5, its top row frozen.
 */
function partsBook(): BookInput {
  const cells: Record<string, CellInput> = {
    B2: { value: 'b2', note: { author: 'Ann', text: 'note B2' } },
    B3: { value: 'site', link: { url: 'https://example.com/', tooltip: 'B3' } },
    B5: { value: 'b5', note: { author: 'Ann', text: 'note B5' } },
    C5: { value: 'to A1', link: { location: 'S!A1' } },
    D6: { value: 'to T', link: { location: "'T'!A1" } },
    C8: { note: { author: 'Ann', text: 'note C8' } },
  }
  for (let r = 1; r <= 10; r++) cells[`A${r}`] = { value: r }
  return {
    sheets: [
      { name: 'S', cells, frozen: { rows: 2, columns: 1 } },
      { name: 'T', cells: { A1: { value: 't' }, A2: { value: 'to B5', link: { location: 'S!B5' } } }, frozen: { rows: 1 } },
    ],
  }
}

export function buildPartsCases(): PartsCases {
  const cases: PartsCase[] = []
  const add = (id: string, op: StructureOp) => cases.push({ id, book: 'parts', op })
  const S = 'S'
  add('insert-row-1', { op: 'insertRows', sheet: S, row: 1 })
  add('insert-row-2', { op: 'insertRows', sheet: S, row: 2 })
  add('insert-row-3', { op: 'insertRows', sheet: S, row: 3 })
  add('insert-row-4', { op: 'insertRows', sheet: S, row: 4 })
  add('insert-rows-5x2', { op: 'insertRows', sheet: S, row: 5, count: 2 })
  add('delete-row-1', { op: 'deleteRows', sheet: S, row: 1 })
  add('delete-rows-1-2', { op: 'deleteRows', sheet: S, row: 1, count: 2 })
  add('delete-rows-2-3', { op: 'deleteRows', sheet: S, row: 2, count: 2 })
  add('delete-row-5', { op: 'deleteRows', sheet: S, row: 5 })
  add('insert-col-a', { op: 'insertColumns', sheet: S, column: 'A' })
  add('insert-col-b', { op: 'insertColumns', sheet: S, column: 'B' })
  add('insert-col-c', { op: 'insertColumns', sheet: S, column: 'C' })
  add('delete-col-a', { op: 'deleteColumns', sheet: S, column: 'A' })
  add('delete-col-b', { op: 'deleteColumns', sheet: S, column: 'B' })
  add('move-b2-b5', { op: 'moveRange', from: { sheet: S, range: 'B2:B5' }, to: { sheet: S, cell: 'F2' } })
  add('move-onto-note', { op: 'moveRange', from: { sheet: S, range: 'B2' }, to: { sheet: S, cell: 'C8' } })
  add('move-plain-onto-note', { op: 'moveRange', from: { sheet: S, range: 'A1' }, to: { sheet: S, cell: 'B5' } })
  add('move-to-t', { op: 'moveRange', from: { sheet: S, range: 'B2:C5' }, to: { sheet: 'T', cell: 'C1' } })
  add('copy-b2-d6', { op: 'copyRange', from: { sheet: S, range: 'B2:D6' }, to: { sheet: S, range: 'F10' } })
  add('copy-plain-onto-note', { op: 'copyRange', from: { sheet: S, range: 'A1' }, to: { sheet: S, range: 'B2' } })
  add('copy-plain-onto-link', { op: 'copyRange', from: { sheet: S, range: 'A1' }, to: { sheet: S, range: 'B3' } })
  add('copy-fill', { op: 'copyRange', from: { sheet: S, range: 'B2:B3' }, to: { sheet: S, range: 'F1:F6' } })
  add('copy-to-t', { op: 'copyRange', from: { sheet: S, range: 'C5:D6' }, to: { sheet: 'T', range: 'D4' } })
  add('rename-s', { op: 'renameSheet', sheet: S, name: 'Data' })
  add('rename-t', { op: 'renameSheet', sheet: 'T', name: 'T 2' })
  add('delete-t', { op: 'deleteSheet', sheet: 'T' })
  return { books: { parts: partsBook() }, cases }
}
