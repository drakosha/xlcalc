/**
 * The first look at Paste Special's All Merging Conditional Formats (Excel for Windows alone: the Mac's AppleScript has
 * no such kind) and All Using Source Theme, beside All: a source with rules of its own onto a target with a rule over it
 * and beyond it, with the same rule over the same cells and over more of them, with the same rule of a relative formula,
 * a source without rules over a target's rule, onto another sheet's rule, onto cells without rules; transposed, with
 * Skip Blanks, an operation, tiled. The engine's prediction (`merging.predicted.json`) has only All's cases.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { RED } from '../cond/cases.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

const GREEN = { fill: { color: '#C6EFCE' } }
const ITALIC = { font: { italic: true } }

/**
 * S: the source A1:B3 with the rules A1:A3 (over 1) and B1:B3 (`B1<0`), P1:P2 without rules; targets D1:E3 under a rule
 * over D1:E5, G1:G3 under A's rule, L1:L5 under A's rule, N1:N3 under B's, J1 without; T: a rule over A1:B3.
 */
function mergingBook(): BookInput {
  const cells: Record<string, CellInput> = {
    A1: { value: 5, style: 1, numberFormat: '0.00' },
    A2: { value: 2 },
    A3: { value: 'a', note: { text: 'n' } },
    B1: { formula: 'A1*2' },
    B2: { value: -1, style: 2 },
    P1: { value: 7 },
    P2: { value: 8, style: 2 },
  }
  for (const a of ['D1', 'E1', 'D2', 'E2', 'D3', 'E3', 'D4', 'G1', 'G2', 'L1', 'L5', 'N1']) cells[a] = { value: 3 }
  return {
    styles: [{}, { font: { bold: true } }, { fill: { color: '#9BC2E6' } }],
    sheets: [
      {
        name: 'S',
        cells,
        conditionalFormats: [
          { type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['1'], style: RED },
          { type: 'expression', ranges: ['B1:B3'], formula: 'B1<0', style: ITALIC },
          { type: 'cellIs', ranges: ['D1:E5'], operator: 'lessThan', formulas: ['10'], style: GREEN },
          { type: 'cellIs', ranges: ['G1:G3'], operator: 'greaterThan', formulas: ['1'], style: RED },
          { type: 'cellIs', ranges: ['L1:L5'], operator: 'greaterThan', formulas: ['1'], style: RED },
          { type: 'expression', ranges: ['N1:N3'], formula: 'N1<0', style: ITALIC },
        ],
        validations: [{ range: 'A1:A3', type: 'whole', operator: 'between', formula1: '0', formula2: '10' }],
      },
      {
        name: 'T',
        cells: { A1: { value: 1 }, B2: { value: 2 } },
        conditionalFormats: [{ type: 'cellIs', ranges: ['A1:B3'], operator: 'equal', formulas: ['2'], style: GREEN }],
      },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteWhat, o: Partial<Pick<PasteCase, 'transpose' | 'operation' | 'skipBlanks'>> = {}, toSheet = 'S') =>
    cases.push({ id, book: 'merging', from: { sheet: 'S', range: from }, to: { sheet: toSheet, range: to }, what, ...o })
  for (const what of ['all', 'allMergingConditionalFormats', 'allUsingSourceTheme'] as const) {
    add(`${what}-over`, 'A1:B3', 'D1', what)
    add(`${what}-same`, 'A1:A3', 'G1', what)
    add(`${what}-inside`, 'A1:A3', 'L2', what)
    add(`${what}-relative`, 'B1:B3', 'N1', what)
    add(`${what}-plain`, 'P1:P2', 'D1', what)
    add(`${what}-other-sheet`, 'A1:B3', 'A1', what, {}, 'T')
    add(`${what}-empty`, 'A1:B3', 'J1', what)
  }
  for (const what of ['allMergingConditionalFormats', 'allUsingSourceTheme'] as const) {
    add(`${what}-transposed`, 'A1:B3', 'D1', what, { transpose: true })
    add(`${what}-skip`, 'A1:B3', 'D1', what, { skipBlanks: true })
    add(`${what}-add`, 'A1:B3', 'D1', what, { operation: 'add' })
  }
  add('allMergingConditionalFormats-tiled', 'A1:A3', 'L1:L6', 'allMergingConditionalFormats')
  return { books: { merging: mergingBook() }, cases }
}
