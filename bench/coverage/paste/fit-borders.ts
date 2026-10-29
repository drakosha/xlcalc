/**
 * The rows All Except Borders fits to the fonts it brings, beside All: a large font onto a row of the default height, the
 * default font onto a row fitted to a large one, a small font onto a default row, onto a row of its own height, tiles of
 * rows, a row whose other cell is larger, whole columns; and a large font with a border onto a cell with one of its own.
 * The book of the set `fit`, its rows as a session fits them. Predicted by the engine before Excel ran them
 * (`fit-borders.predicted.json`): it fits them as for All.
 */
import type { BookInput, CellStyle } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'
import { fitBook } from './fit.js'

/** The book of `fit` with A6 large and a thin bottom border, C19 with a thick top border, patched in a session as a host patches them, which fits row 6. */
function bordersBook(): BookInput {
  const book = fitBook()
  const styles: CellStyle[] = [...(book.styles ?? []), { font: { size: 20 }, border: { bottom: { style: 'thin' } } }, { border: { top: { style: 'thick' } } }]
  const sheet = book.sheets[0]
  const s = createSession({ ...book, styles, sheets: [{ ...sheet, cells: { ...sheet.cells, A6: { value: 'edge' }, C19: { value: 'u' } } }] }, { clock: { now: '2026-10-27T12:00:00Z', seed: 1 } })
  s.patchCells([{ sheet: 'S', range: 'A6', set: { style: styles.length - 2 } }, { sheet: 'S', range: 'C19', set: { style: styles.length - 1 } }])
  return s.getBook()
}

export function buildPasteCases(): PasteCases {
  const fit = bordersBook()
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string) => {
    for (const what of ['allExceptBorders', 'all'] as PasteWhat[]) cases.push({ id: what === 'all' ? `all-${id}` : id, book: 'fit', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what })
  }
  add('up', 'A1', 'C5')
  add('down', 'A3', 'C7')
  add('small', 'A2', 'C5')
  add('own', 'A1', 'C9')
  add('tiled', 'A1', 'C13:C15')
  add('under-larger', 'A4', 'C17')
  add('column', 'A:A', 'G:G')
  add('bordered', 'A6', 'C19')
  return { books: { fit }, cases }
}
