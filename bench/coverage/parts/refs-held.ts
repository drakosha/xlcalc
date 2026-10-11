/**
 * Held out: the cell of a note as a file other than Excel's may spell it (`<comment ref>`): a row with a leading zero
 * (`A01`), a range of one cell (`A1:A1`), lower case, `$`, a range of several cells. `writeXlsx` writes the book (notes on
 * A1 and D4) and the spelling of A1's note is put in by hand (`raw`); its VML shape stays on A1. Predicted before Excel
 * ran it (`refs-held.predicted.json`).
 */
import type { BookInput } from '../../../src/contract.js'
import type { PartsCase, PartsCases } from './cases.js'

const book: BookInput = {
  sheets: [{ name: 'S', cells: { A1: { value: 1, note: { author: 'Ann', text: 'on A1' } }, B2: { value: 2 }, D4: { value: 4, note: { author: 'Ann', text: 'on D4' } } } }],
}

/** A case whose note on A1 is on `ref` in the file. */
const spelled = (id: string, ref: string): PartsCase => ({ id, book: 'notes', raw: [{ part: 'xl/comments1.xml', from: '<comment ref="A1" ', to: `<comment ref="${ref}" ` }] })

export function buildPartsCases(): PartsCases {
  return {
    books: { notes: book },
    cases: [spelled('zero', 'A01'), spelled('one-cell-range', 'A1:A1'), spelled('lower', 'a1'), spelled('dollar', '$A$1'), spelled('range', 'A1:B2')],
  }
}
