/**
 * All Merging Conditional Formats beside All with rules other than cell values and formulas, in both Excels (the Mac's
 * through its Paste Special window): a source under each of a color scale, a data bar, an icon set, a top rule, an
 * average rule and a duplicates rule pasted over a destination's rule of cell values; a source without rules over a
 * destination's color scale, top rule and data bar; a source under the same color scale and the same top rule as the
 * destination's; each of the color scale, the top rule and the duplicates rule copied just below its own cells. Predicted
 * by the engine before Excel ran them (`merging-rules.predicted.json`).
 */
import type { BookInput, CellInput, ConditionalFormat } from '../../../src/contract.js'
import { RED } from '../cond/cases.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

const GREEN = { fill: { color: '#C6EFCE' } }
const SCALE = { type: 'colorScale', points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] } as const
const TOP = { type: 'top10', rank: 1, style: RED } as const

/** By column of the source, its rule over rows 1 to 4 and its values. */
const SOURCES: Record<string, { rule: Record<string, unknown>; values: number[] }> = {
  A: { rule: SCALE, values: [1, 2, 3, 4] },
  B: { rule: { type: 'dataBar', min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', border: '#638EC6' }, values: [1, 5, 9, 3] },
  C: { rule: { type: 'iconSet', icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }] }, values: [1, 2, 3, 4] },
  D: { rule: TOP, values: [10, 20, 30, 40] },
  E: { rule: { type: 'aboveAverage', style: RED }, values: [1, 2, 3, 4] },
  F: { rule: { type: 'duplicateValues', style: RED }, values: [1, 1, 2, 3] },
}

/** S: the sources A1:F4 under their rules, P1:P4 without; destinations H1:H6 under a rule of cell values, J1:J6 under A's color scale, L1:L6 under D's top rule, N1:N6 under a data bar. */
function rulesBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (const [col, s] of Object.entries(SOURCES)) s.values.forEach((v, i) => (cells[`${col}${i + 1}`] = { value: v }))
  ;[7, 8, 9, 10].forEach((v, i) => (cells[`P${i + 1}`] = { value: v }))
  for (const col of ['H', 'J', 'L', 'N']) for (let r = 1; r <= 6; r++) cells[`${col}${r}`] = { value: r * 3 }
  const rules = [
    ...Object.entries(SOURCES).map(([col, s]) => ({ ...s.rule, ranges: [`${col}1:${col}4`] })),
    { type: 'cellIs', ranges: ['H1:H6'], operator: 'greaterThan', formulas: ['1'], style: GREEN },
    { ...SCALE, ranges: ['J1:J6'] },
    { ...TOP, ranges: ['L1:L6'] },
    { type: 'dataBar', ranges: ['N1:N6'], min: { type: 'min' }, max: { type: 'max' }, color: '#FFB628' },
  ] as ConditionalFormat[]
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }] }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string) => {
    for (const what of ['allMergingConditionalFormats', 'all'] as PasteWhat[]) cases.push({ id: what === 'all' ? `all-${id}` : id, book: 'rules', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what })
  }
  const names: Record<string, string> = { A: 'scale', B: 'bar', C: 'icons', D: 'top', E: 'average', F: 'duplicates' }
  for (const [col, name] of Object.entries(names)) add(`${name}-over`, `${col}1:${col}4`, 'H2')
  add('plain-on-scale', 'P1:P4', 'J2')
  add('plain-on-top', 'P1:P4', 'L2')
  add('plain-on-bar', 'P1:P4', 'N2')
  add('scale-same', 'A1:A4', 'J2')
  add('top-same', 'D1:D4', 'L2')
  for (const col of ['A', 'D', 'F']) add(`${names[col]}-below-own`, `${col}1:${col}4`, `${col}5`)
  return { books: { rules: rulesBook() }, cases }
}
