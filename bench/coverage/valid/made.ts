/**
 * Data validation and outlines Excel makes itself, through its object model on a book of plain values: a rule of every
 * type and setting (`Range.Validation.Add` with the range's first cell active, then its properties), a rule over two
 * ranges with relative references, groups of rows and columns nested and collapsed (`Group`, `ShowDetail`), summaries
 * above and left. The engine holds what Excel made as `made` (in `cases.ts`) describes it; the reader must read the file
 * Excel saved as Excel read it. Predicted before Excel ran it (`made.predicted.json`).
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { ValidCases, ValidMake } from './cases.js'

/** S: values in A1:H12 (numbers, text, dates, times); T: a second sheet for a list over another sheet. */
function plainBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 12; r++) {
    cells[`A${r}`] = { value: r }
    cells[`B${r}`] = { value: r % 3 ? 'Yes' : 'no' }
    cells[`C${r}`] = { value: r * 1.5 }
    cells[`D${r}`] = { value: 45650 + r * 3 }
    cells[`E${r}`] = { value: r / 24 }
    cells[`F${r}`] = { value: 'x'.repeat(r) }
  }
  return { sheets: [{ name: 'S', cells }, { name: 'T', cells: { A1: { value: 'Yes' }, A2: { value: 'Maybe' } } }] }
}

const S = 'S'

/** A rule of each type with its settings. */
const rules: ValidMake[] = [
  { sheet: S, range: 'B1:B6', rule: { type: 'list', formula1: '"Yes,Maybe"', promptTitle: 'Pick', prompt: 'Yes or Maybe' } },
  { sheet: S, range: 'B7:B12', rule: { type: 'list', formula1: "'T'!$A$1:$A$2", dropdown: false, errorStyle: 'warning', errorTitle: 'Not listed', error: 'Pick from T' } },
  { sheet: S, range: 'A1:A12', rule: { type: 'whole', operator: 'lessThanOrEqual', formula1: '10' } },
  { sheet: S, range: 'C1:C12', rule: { type: 'decimal', operator: 'notBetween', formula1: '$A$3', formula2: '$A$9', allowBlank: false } },
  { sheet: S, range: 'D1:D12', rule: { type: 'date', operator: 'greaterThan', formula1: 'DATE(2025,1,1)', errorStyle: 'information' } },
  { sheet: S, range: 'E1:E12', rule: { type: 'time', formula1: '0.125', formula2: '0.375', showPrompt: false, prompt: 'hidden prompt' } },
  { sheet: S, range: 'F1:F12', rule: { type: 'textLength', operator: 'greaterThan', formula1: '5', showError: false } },
  { sheet: S, range: 'G1:G3 I5:I6', rule: { type: 'custom', formula1: 'COUNTIF($A$1:$A$12,A1)>0' } },
  { sheet: S, range: 'H1:H12', rule: { type: 'any', promptTitle: 'Free', prompt: 'Anything' } },
]

export function buildValidCases(): ValidCases {
  return {
    books: { plain: plainBook() },
    cases: [
      { id: 'rules', book: 'plain', make: rules },
      {
        id: 'groups',
        book: 'plain',
        make: [
          { sheet: S, group: '3:6' },
          { sheet: S, group: '4:5' },
          { sheet: S, group: '9:10' },
          { sheet: S, collapse: '11' },
          { sheet: S, group: 'B:C' },
          { sheet: S, collapse: 'D' },
        ],
      },
      {
        id: 'groups-above',
        book: 'plain',
        make: [
          { sheet: S, summary: { summaryBelow: false, summaryRight: false } },
          { sheet: S, group: '3:5' },
          { sheet: S, collapse: '2' },
          { sheet: S, group: 'C:D' },
          { sheet: S, collapse: 'B' },
        ],
      },
      {
        id: 'rules-and-groups',
        book: 'plain',
        make: [rules[0], rules[2], { sheet: S, group: '2:4' }, { sheet: S, collapse: '5' }],
        op: { op: 'insertRows', sheet: S, row: 3 },
      },
    ],
    window: 'A1:J14',
  }
}
