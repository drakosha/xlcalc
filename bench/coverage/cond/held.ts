/**
 * Held out after the sets looks, files and struct-held changed the engine: the rules they gave, on cases they did not
 * have, predicted before Excel ran them (`held.predicted.json`). Looks: a formula answering a range whose first value
 * holds, duplicates of TRUE and numeric text, a three-colour scale over equal numbers, a top rule over text alone.
 * Structure: a column inserted right of a rule, rows and a column deleted from an icon set, a top rule and a two-column
 * colour scale, a fill to the right, a copy into the rule's own cells, a move of the referenced cells to another sheet,
 * a move onto the rule's own cells.
 */
import type { BookInput, CellInput, ConditionalFormat, StructureOp } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'

const v = (value: CellInput['value']): CellInput => ({ value })
const RED = { fill: { color: '#FFC7CE' } }

function looksBook(): BookInput {
  const cells: Record<string, CellInput> = { A1: v(5), A2: v(1), A3: v(2) }
  for (let r = 1; r <= 3; r++) cells[`B${r}`] = v(`b${r}`)
  const dups = [v(true), v('TRUE'), v(2), v('2.0'), v(' 2'), v('x'), v('X '), v(0), v(false)]
  dups.forEach((x, i) => (cells[`C${i + 1}`] = x))
  ;[v(4), v(4), v(4)].forEach((x, i) => (cells[`D${i + 1}`] = x))
  ;[v('a'), v('b'), v('10')].forEach((x, i) => (cells[`E${i + 1}`] = x))
  const rules: ConditionalFormat[] = [
    { type: 'expression', ranges: ['B1:B3'], formula: '$A$1:$A$3>3', style: RED },
    { type: 'duplicateValues', ranges: ['C1:C9'], style: RED },
    { type: 'colorScale', ranges: ['D1:D3'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'percentile', value: '50' }, color: '#FFEB84' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'top10', ranges: ['E1:E3'], rank: 1, style: RED },
  ]
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }] }
}

/** S: numbers 0 to 9 in A1:A10 and 1 to 10 in B:H; a rule on each column; T empty. */
function structBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 10; r++) {
    cells[`A${r}`] = v(r - 1)
    for (const c of ['B', 'C', 'D', 'E', 'F', 'G']) cells[`${c}${r}`] = v(r)
  }
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['B2:B10'], operator: 'greaterThan', formulas: ['5'], style: RED },
    { type: 'colorScale', ranges: ['C1:D10'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'iconSet', ranges: ['E2:E10'], icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }] },
    { type: 'top10', ranges: ['F2:F10'], rank: 3, style: { font: { bold: true } } },
    { type: 'expression', ranges: ['G2:G10'], formula: '$A2>4', style: { font: { italic: true } } },
  ]
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }, { name: 'T', cells: {} }] }
}

export function buildCondCases(): CondCases {
  const cases: CondCase[] = [{ id: 'looks', book: 'looks', read: [['S', 'B1:E9']] }]
  const read: CondCase['read'] = [['S', 'B1:H12']]
  const add = (id: string, op: StructureOp, more: CondCase['read'] = []) => cases.push({ id, book: 'struct', op, read: [...read, ...more] })
  const S = 'S'
  add('insert-col-right', { op: 'insertColumns', sheet: S, column: 'H' })
  add('delete-rows-iconset', { op: 'deleteRows', sheet: S, row: 4, count: 2 })
  add('delete-col-scale', { op: 'deleteColumns', sheet: S, column: 'D' })
  add('fill-right', { op: 'copyRange', from: { sheet: S, range: 'B2:B10' }, to: { sheet: S, range: 'H2:H10' } })
  add('fill-right-adjacent', { op: 'copyRange', from: { sheet: S, range: 'G2:G10' }, to: { sheet: S, range: 'H2:H10' } })
  add('copy-into-own', { op: 'copyRange', from: { sheet: S, range: 'B2:B3' }, to: { sheet: S, range: 'B5' } })
  add('move-referenced-to-t', { op: 'moveRange', from: { sheet: S, range: 'A1:A10' }, to: { sheet: 'T', cell: 'A1' } })
  add('move-onto-own', { op: 'moveRange', from: { sheet: S, range: 'B2:B5' }, to: { sheet: S, cell: 'B4' } })
  return { books: { looks: looksBook(), struct: structBook() }, cases }
}
