/**
 * Held out after the set `merging` gave the engine All Merging Conditional Formats (Excel for Windows alone): predicted
 * by the engine before Excel ran them (`merging-held.predicted.json`). A copy just below its own rule, where All grows
 * the rule; a copy over part of its own rule and over its own cells; a rule that lies inside the destination; two rules
 * of the destination; whole columns; one cell into a merged area; transposed over part of a rule; tiled; with an
 * operation, and Skip Blanks with one; onto another sheet with Skip Blanks; and All beside the cases where it differs.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { RED } from '../cond/cases.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

const GREEN = { fill: { color: '#C6EFCE' } }
const BOLD = { font: { bold: true } }

/**
 * S: the source A1:A3 under its rule (over 1), a rule over L1:L5, a destination D1:E4 under a rule over E2:E3 and one
 * over D1:D9, a rule over column H, a merged area J1:K2 under a rule; T: a rule over B2:C3.
 */
function heldBook(): BookInput {
  const cells: Record<string, CellInput> = { A1: { value: 5, style: 1 }, A2: { value: 2 }, B1: { value: -1 }, H1: { value: 4 }, J1: { value: 9 } }
  for (const a of ['D1', 'D2', 'E2', 'E3', 'L1', 'L2', 'L3']) cells[a] = { value: 3 }
  return {
    styles: [{}, BOLD],
    sheets: [
      {
        name: 'S',
        cells,
        merges: ['J1:K2'],
        conditionalFormats: [
          { type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['1'], style: RED },
          { type: 'cellIs', ranges: ['L1:L5'], operator: 'lessThan', formulas: ['10'], style: GREEN },
          { type: 'expression', ranges: ['E2:E3'], formula: 'E2>2', style: BOLD },
          { type: 'cellIs', ranges: ['D1:D9'], operator: 'equal', formulas: ['3'], style: GREEN },
          { type: 'cellIs', ranges: ['H:H'], operator: 'greaterThan', formulas: ['0'], style: BOLD },
          { type: 'cellIs', ranges: ['J1:K2'], operator: 'greaterThan', formulas: ['5'], style: GREEN },
        ],
      },
      {
        name: 'T',
        cells: { B2: { value: 1 } },
        conditionalFormats: [{ type: 'cellIs', ranges: ['B2:C3'], operator: 'equal', formulas: ['1'], style: GREEN }],
      },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, o: Partial<Pick<PasteCase, 'transpose' | 'operation' | 'skipBlanks'>> = {}, toSheet = 'S', kinds: PasteWhat[] = ['allMergingConditionalFormats']) => {
    for (const what of kinds) cases.push({ id: what === 'all' ? `all-${id}` : id, book: 'held', from: { sheet: 'S', range: from }, to: { sheet: toSheet, range: to }, what, ...o })
  }
  const both: PasteWhat[] = ['allMergingConditionalFormats', 'all']
  add('below-own', 'A1:A3', 'A4', {}, 'S', both)
  add('over-own-rule', 'L1:L2', 'L4', {}, 'S', both)
  add('over-own-cells', 'A1:A3', 'A2', {}, 'S', both)
  add('rule-inside', 'A1:B3', 'D1:E3', {}, 'S', both)
  add('columns', 'A:A', 'H:H', {}, 'S', both)
  add('into-merged', 'A1', 'J1', {}, 'S', both)
  add('transposed', 'A1:A3', 'D2', { transpose: true })
  add('tiled', 'A1:A2', 'D1:D4')
  add('multiply', 'A1:A3', 'D1', { operation: 'multiply' })
  add('skip-add', 'A1:A3', 'D1', { skipBlanks: true, operation: 'add' })
  add('other-sheet-skip', 'A1:A3', 'B2', { skipBlanks: true }, 'T')
  add('other-sheet', 'A1:B2', 'B2', {}, 'T')
  return { books: { held: heldBook() }, cases }
}
