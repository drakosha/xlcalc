/**
 * Which merged areas of the target a paste of looks takes away, after the set `held` (Formats over F1:K1 took F1:G1 away
 * and left J1:K1): a merged area at the paste's top-left cell or inside it elsewhere, one the paste covers part of, and
 * Paste Special's `all` over the same. Predicted by the engine before Excel ran them (`merges.predicted.json`).
 */
import type { BookInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

/** A source A3:F3 of looks and values; the target row 1 with the merged areas `merges`. */
function book(merges: string[]): BookInput {
  return {
    styles: [{}, { font: { bold: true } }, { fill: { color: '#FFFF00' } }],
    sheets: [{ name: 'S', cells: { A3: { value: 'p', style: 2 }, B3: { value: 2, style: 1 }, E3: { value: 5, style: 2 }, F1: { value: 'x' }, J1: { value: 'y' } }, merges }],
  }
}

export function buildPasteCases(): PasteCases {
  const books: Record<string, BookInput> = {
    right: book(['J1:K1']),
    corner: book(['F1:G1']),
    both: book(['F1:G1', 'J1:K1']),
    down: book(['J1:J2']),
  }
  const cases: PasteCase[] = []
  const add = (id: string, b: string, to: string, what: PasteWhat) => cases.push({ id, book: b, from: { sheet: 'S', range: 'A3:F3' }, to: { sheet: 'S', range: to }, what })
  add('formats-inside', 'right', 'F1', 'formats')
  add('formats-corner', 'corner', 'F1', 'formats')
  add('formats-before', 'both', 'E1', 'formats')
  add('formats-part', 'down', 'F1', 'formats')
  add('all-both', 'both', 'F1', 'all')
  add('all-inside', 'right', 'F1', 'all')
  add('all-part', 'down', 'F1', 'all')
  return { books, cases }
}
