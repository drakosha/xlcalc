/**
 * The first look at Paste Special's kinds the engine does not have, before it has them: values and number formats,
 * formulas and number formats, column widths, all except borders and comments (notes), of a source with looks, borders,
 * notes, widths, a rule and a validation onto a target with its own; onto another sheet, tiled, transposed, from whole
 * columns and onto merged areas. The engine's prediction (`kinds.predicted.json`) has only the kinds it had then.
 */
import type { BookInput, CellInput, CellStyle } from '../../../src/contract.js'
import { RED } from '../cond/cases.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

const THIN = { style: 'thin' } as const
const STYLES: CellStyle[] = [
  {},
  { font: { bold: true }, fill: { color: '#FFFF00' }, border: { bottom: THIN } },
  { font: { italic: true }, border: { top: THIN, bottom: THIN, left: THIN, right: THIN } },
  { fill: { color: '#9BC2E6' }, border: { top: { style: 'thick' } } },
  { border: { left: THIN } },
]

/** S: the source A1:C3 (columns A and B of their own widths) and a target F1:H3 with its own (column G of its own width); T: a target with its own. */
function kindsBook(): BookInput {
  const cells: Record<string, CellInput> = {
    A1: { value: 1, style: 1, numberFormat: '0.00', note: { text: 'na' } },
    B1: { formula: 'A1*2', style: 2, numberFormat: '0%' },
    C1: { value: 'link', link: { url: 'https://example.com/' } },
    A2: { value: 3 },
    B2: { style: 4 },
    C2: { value: 'm', note: { text: 'nc' } },
    A3: { value: 46037, numberFormat: 'm/d/yyyy' },
    B3: { formula: 'A3+1' },
    C3: { formula: '1/0' },
    A5: { value: 'w', style: 2 },
    F1: { value: 'old', style: 3, numberFormat: '0.000' },
    G1: { value: 9, note: { text: 'own' } },
    H1: { value: 'x', link: { url: 'https://example.org/' } },
    G2: { value: 5, style: 3 },
    H3: { formula: 'G2+1', numberFormat: '#,##0' },
  }
  return {
    styles: STYLES,
    sheets: [
      {
        name: 'S',
        cells,
        merges: ['A5:B5'],
        columns: { A: { width: 20 }, B: { width: 5 }, G: { width: 15 } },
        conditionalFormats: [
          { type: 'cellIs', ranges: ['A1:C3'], operator: 'greaterThan', formulas: ['1'], style: RED },
          { type: 'expression', ranges: ['F1:H3'], formula: 'F1=9', style: { font: { italic: true } } },
        ],
        validations: [
          { range: 'A1:A3', type: 'whole', operator: 'between', formula1: '0', formula2: '10' },
          { range: 'F1:F3', type: 'list', formula1: '"a,b"' },
        ],
      },
      { name: 'T', cells: { B2: { value: 'own', style: 3, numberFormat: '0.0' }, C3: { value: 5, note: { text: 't' } } }, columns: { C: { width: 30 } } },
    ],
  }
}

/** A target merged over F1:G1 with a value, and a source A1:B1 with notes and looks. */
function mergedBook(): BookInput {
  return {
    styles: STYLES,
    sheets: [{ name: 'S', cells: { A1: { value: 1, style: 1, note: { text: 'a' } }, B1: { value: 2, numberFormat: '0%', note: { text: 'b' } }, F1: { value: 'x', style: 3 } }, merges: ['F1:G1'], columns: { A: { width: 20 } } }],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, book: string, from: string, to: string, what: PasteWhat, transpose = false, toSheet = 'S') =>
    cases.push({ id, book, from: { sheet: 'S', range: from }, to: { sheet: toSheet, range: to }, what, ...(transpose ? { transpose } : {}) })
  const KINDS = ['valuesAndNumberFormats', 'formulasAndNumberFormats', 'columnWidths', 'allExceptBorders', 'notes'] as const
  for (const what of KINDS) add(`${what}-F1`, 'kinds', 'A1:C3', 'F1', what)
  for (const what of KINDS) add(`${what}-T`, 'kinds', 'A1:C3', 'B2', what, false, 'T')
  for (const what of KINDS) add(`${what}-transposed`, 'kinds', 'A1:C3', 'K1', what, true)
  add('notes-tiled', 'kinds', 'A1', 'F10:G11', 'notes')
  add('notes-over-value', 'kinds', 'C2', 'H1', 'notes')
  add('notes-none-over-note', 'kinds', 'A2', 'G1', 'notes')
  add('columnWidths-cell', 'kinds', 'B1', 'H5', 'columnWidths')
  add('columnWidths-default-over-own', 'kinds', 'C1', 'G7', 'columnWidths')
  add('columnWidths-column', 'kinds', 'A:A', 'D:D', 'columnWidths')
  add('columnWidths-tiled', 'kinds', 'A1:B1', 'J20:M20', 'columnWidths')
  add('valuesAndNumberFormats-blank-styled', 'kinds', 'B2', 'G2', 'valuesAndNumberFormats')
  add('formulasAndNumberFormats-tiled', 'kinds', 'A1:B1', 'F10:I11', 'formulasAndNumberFormats')
  add('allExceptBorders-merged-source', 'kinds', 'A5:B5', 'F5', 'allExceptBorders')
  add('allExceptBorders-onto-border', 'kinds', 'A2', 'F1', 'allExceptBorders')
  add('allExceptBorders-column', 'kinds', 'A:A', 'D:D', 'allExceptBorders')
  add('notes-merged', 'merged', 'A1:B1', 'F1', 'notes')
  add('notes-merged-part', 'merged', 'A1', 'G1', 'notes')
  add('valuesAndNumberFormats-merged', 'merged', 'A1:B1', 'F1', 'valuesAndNumberFormats')
  add('valuesAndNumberFormats-merged-part', 'merged', 'A1', 'G1', 'valuesAndNumberFormats')
  add('columnWidths-merged', 'merged', 'A1:B1', 'F1', 'columnWidths')
  add('allExceptBorders-merged', 'merged', 'A1:B1', 'F1', 'allExceptBorders')
  add('allExceptBorders-merged-part', 'merged', 'A1', 'G1', 'allExceptBorders')
  return { books: { kinds: kindsBook(), merged: mergedBook() }, cases }
}
