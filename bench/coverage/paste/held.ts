/**
 * Held out after the set `probe` gave the engine Paste Special: cases it did not have, predicted by the engine before
 * Excel ran them (`held.predicted.json`). Transposed references outside the copy, absolute, mixed, to another sheet, to
 * whole columns and rows, through a name, and one turned off the sheet; a transposed copy tiled; a paste onto part of a
 * merged area; contents onto merged areas the copy brings and does not; looks over a note and a hyperlink; whole
 * columns; looks onto part of an array formula; data validation tiled; a table's cells pasted as values and formulas.
 */
import type { BookInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

/** References a transposed copy turns: S!A1:D3 with references out of it. */
function refsBook(): BookInput {
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1 },
          B1: { formula: '$Z$9+A$9+$A9' },
          C1: { formula: 'T!A1+T!$B$2' },
          D1: { formula: 'SUM(A:A)' },
          A2: { formula: 'nm*2' },
          B2: { formula: 'SUM(A1:A2)' },
          C2: { formula: 'Z9' },
          D2: { formula: 'SUM(5:5)' },
          A3: { formula: 'A20000*1' },
          B3: { value: 'b' },
          Z9: { value: 5 },
        },
      },
      { name: 'T', cells: { A1: { value: 2 }, B2: { value: 3 } } },
    ],
    names: [{ name: 'nm', ref: 'S!$A$1' }],
  }
}

/** Merged areas: A1:B1 in the source, F1:G1 and I1:J1 on the target, a source of three cells, notes and a hyperlink on the target, a legacy array formula. */
function mergedBook(): BookInput {
  return {
    styles: [{}, { font: { bold: true } }, { fill: { color: '#FFFF00' } }],
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1, style: 1 },
          C1: { value: 3 },
          A3: { value: 'p', style: 2 },
          F1: { value: 'x' },
          J1: { value: 'y' },
          F3: { value: 'n', note: { text: 'own' }, link: { url: 'https://example.org/' } },
          F5: { formula: 'A1:A2*1', array: 'F5:F6' },
        },
        merges: ['A1:B1', 'F1:G1', 'J1:K1'],
      },
    ],
  }
}

/** A column with its own width and look, a value in it; a table T over A10:B12 with a calculated column. */
function linesBook(): BookInput {
  return {
    styles: [{}, { font: { bold: true } }],
    sheets: [
      {
        name: 'S',
        cells: { A1: { value: 4 }, A2: { value: 'q', style: 0 }, A10: { value: 'x' }, B10: { value: 'y' }, A11: { value: 1 }, B11: { formula: 'T[@x]*2' }, A12: { value: 2 }, B12: { formula: 'T[@x]*2' } },
        columns: { A: { width: 20, style: 1 } },
        tables: [{ name: 'T', range: 'A10:B12', columns: [{ name: 'x' }, { name: 'y', formula: 'T[@x]*2' }] }],
        validations: [{ range: 'C1', type: 'whole', operator: 'greaterThan', formula1: '0' }],
      },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, book: string, from: string, to: string, what: PasteWhat, transpose = false) =>
    cases.push({ id, book, from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what, ...(transpose ? { transpose } : {}) })
  add('refs-out-transposed', 'refs', 'A1:D3', 'F1', 'formulas', true)
  add('refs-out-all-transposed', 'refs', 'A1:D3', 'F1', 'all', true)
  add('tiled-transposed', 'refs', 'A1:B1', 'F10:F13', 'all', true)
  add('merged-part-all', 'merged', 'C1', 'G1', 'all')
  add('merged-part-wide-all', 'merged', 'A3:C3', 'G1', 'all')
  add('merged-same-values', 'merged', 'A1:B1', 'F1', 'values')
  add('merged-inside-values', 'merged', 'A3:F3', 'F1', 'values')
  add('merged-inside-formats', 'merged', 'A3:F3', 'F1', 'formats')
  add('formats-over-note', 'merged', 'A3', 'F3', 'formats')
  add('values-over-note', 'merged', 'C1', 'F3', 'values')
  add('formats-on-array-part', 'merged', 'A3', 'F6', 'formats')
  add('values-on-array-part', 'merged', 'C1', 'F6', 'values')
  add('column-formats', 'lines', 'A:A', 'D:D', 'formats')
  add('column-values', 'lines', 'A:A', 'D:D', 'values')
  add('validation-tiled', 'lines', 'C1', 'E1:E3', 'validation')
  add('table-values', 'lines', 'A10:B12', 'D10', 'values')
  add('table-formulas', 'lines', 'B11:B12', 'H11', 'formulas')
  add('table-formats', 'lines', 'A10:B12', 'D10', 'formats')
  return { books: { refs: refsBook(), merged: mergedBook(), lines: linesBook() }, cases }
}
