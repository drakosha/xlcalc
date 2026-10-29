/**
 * The rows a paste fits to the fonts it brings (Excel for Mac fits a row without a height of its own after Paste
 * Special's Formats): Formats, All, a plain copy, values and formulas of a large font onto a row of the default height;
 * looks of the default font and of a small one onto a row fitted to a large one and onto a default row; onto a row of its
 * own height, an empty cell, tiles of rows, a row whose other cell is larger, whole columns. The books are made as a
 * host makes them, their fonts patched in a session, so their rows have the heights the engine fits. Predicted by the
 * engine before Excel ran them (`fit.predicted.json`).
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

const STYLES = [{}, { font: { size: 20, bold: true } }, { font: { size: 8 } }, { font: { size: 16 } }]

/** A1 large, A2 small, A3 plain, A4 middle; targets C5, C7 (large), C9 (own height), C13:C15, C17 with D17 large; rows fitted as a session fits them. */
export function fitBook(): BookInput {
  const cells: Record<string, CellInput> = { A1: { value: 'big' }, A2: { value: 'small' }, A3: { value: 'plain' }, A4: { value: 'mid' }, C5: { value: 'x' }, C7: { value: 'y' }, C9: { value: 'z' }, C13: { value: 1 }, C14: { value: 2 }, C15: { value: 3 }, C17: { value: 'v' }, D17: { value: 'w' } }
  const s = createSession({ styles: STYLES, sheets: [{ name: 'S', cells, rows: { 9: { height: 30 } } }] }, { clock: { now: '2026-10-27T12:00:00Z', seed: 1 } })
  const at = (cell: string, style: number) => ({ sheet: 'S', range: cell, set: { style } })
  s.patchCells([at('A1', 1), at('A2', 2), at('A4', 3), at('C7', 1), at('D17', 1)])
  return s.getBook()
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteWhat) => cases.push({ id, book: 'fit', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what })
  for (const what of ['formats', 'all', 'copy', 'values', 'formulas'] as const) add(`${what}-up`, 'A1', 'C5', what)
  add('formats-down', 'A3', 'C7', 'formats')
  add('formats-small', 'A2', 'C5', 'formats')
  add('formats-small-down', 'A2', 'C7', 'formats')
  add('formats-own', 'A1', 'C9', 'formats')
  add('formats-empty', 'A1', 'E11', 'formats')
  add('formats-tiled', 'A1', 'C13:C15', 'formats')
  add('formats-under-larger', 'A4', 'C17', 'formats')
  add('formats-column', 'A:A', 'G:G', 'formats')
  add('copy-down', 'A3', 'C7', 'copy')
  return { books: { fit: fitBook() }, cases }
}
