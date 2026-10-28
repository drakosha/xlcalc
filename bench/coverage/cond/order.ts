/**
 * Which rules Excel puts after the others when lines are inserted or deleted, by what the operation does to a rule's
 * cells (untouched, moved, cut) and to the cells its formulas refer to (moved, lost, on another sheet), for each kind of
 * rule: Excel for Windows put a colour scale whose threshold a column deletion moved after the others, though it kept its
 * cells (`test/xlsx-keep-rules-excel.test.ts`). Each book's last rule, untouched by the operations, shows which rules
 * went after it. Predicted before Excel ran it (`order.predicted.json`).
 */
import type { BookInput, CellInput, ConditionalFormat, ConditionalValue, StructureOp } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'
import { valueCell as v } from '../stored.js'
import { RED } from './cases.js'

const num = (value: string): ConditionalValue => ({ type: 'num', value })

/** One rule of each kind over ten rows from `top` of the columns `cols`, from the first, referring to the cells `a` and `b`; a rule over `last` after them. */
function rules(cols: string[], top: number, a: string, b: string, other: string, last: string): ConditionalFormat[] {
  const over = (c: string) => `${c}${top}:${c}${top + 9}`
  const [scale, bar, icons, value, formula, ranked, across] = cols.map(over)
  const first = (r: string) => r.split(':')[0]
  return [
    { type: 'colorScale', ranges: [scale], points: [{ value: num(a), color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'dataBar', ranges: [bar], min: num(a), max: num(b), color: '#638EC6' },
    { type: 'iconSet', ranges: [icons], icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' } }, { value: num(a) }, { value: num(b) }] },
    { type: 'cellIs', ranges: [value], operator: 'greaterThan', formulas: [b], style: RED },
    { type: 'expression', ranges: [formula], formula: `${first(formula)}>${a}`, style: { font: { bold: true } } },
    { type: 'top10', ranges: [ranked], rank: 3, style: { font: { italic: true } } },
    { type: 'expression', ranges: [across], formula: `${first(across)}>${other}`, style: { fill: { color: '#C6EFCE' } } },
    { type: 'cellIs', ranges: [over(last)], operator: 'greaterThan', formulas: ['5'], style: { font: { color: '#0070C0' } } },
  ]
}

/** Numbers 1 to 10 down the rows `from` to `from + 9` of the columns. */
function numbers(cols: string[], from: number): Record<string, CellInput> {
  const cells: Record<string, CellInput> = {}
  for (const c of cols) for (let r = 0; r < 10; r++) cells[`${c}${from + r}`] = v(r + 1)
  return cells
}

const T = { name: 'T', cells: { A5: v(6) } }

/**
 * `refs`: rules over B to H, rows 1 to 10, referring to K15 and K16 below and right of them (and H to T's A5), the last
 * over A; lines deleted or inserted between the rules and those cells move what they refer to and leave their cells.
 */
function refsBook(): BookInput {
  const cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
  const cells = { ...numbers(['A', ...cols], 1), K15: v(4), K16: v(7) }
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules(cols, 1, '$K$15', '$K$16', "'T'!$A$5", 'A') }, T] }
}

/**
 * `left`: rules over D to J, rows 5 to 14, referring to A1 and A2 above and left of them, J's to L1 right of them, the
 * last over B; a column deleted or inserted at C moves the rules' cells and leaves A1 and A2, and moves L1 with J.
 */
function leftBook(): BookInput {
  const cols = ['D', 'E', 'F', 'G', 'H', 'I', 'J']
  const cells = { ...numbers(['B', ...cols], 5), A1: v(4), A2: v(7), L1: v(5) }
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules(cols, 5, '$A$1', '$A$2', '$L$1', 'B') }, T] }
}

export function buildCondCases(): CondCases {
  const cases: CondCase[] = []
  const add = (id: string, book: string, op: StructureOp, read: string) => cases.push({ id, book, op, read: [['S', read]] })
  const S = 'S'
  const refs = (id: string, op: StructureOp) => add(id, 'refs', op, 'A1:H10')
  refs('delete-cols-i', { op: 'deleteColumns', sheet: S, column: 'I', count: 2 })
  refs('delete-rows-12', { op: 'deleteRows', sheet: S, row: 12, count: 2 })
  refs('insert-cols-i', { op: 'insertColumns', sheet: S, column: 'I', count: 2 })
  refs('insert-rows-12', { op: 'insertRows', sheet: S, row: 12, count: 2 })
  refs('delete-col-k', { op: 'deleteColumns', sheet: S, column: 'K' })
  refs('delete-row-15', { op: 'deleteRows', sheet: S, row: 15 })
  refs('delete-row-5', { op: 'deleteRows', sheet: S, row: 5 })
  refs('delete-rows-t', { op: 'deleteRows', sheet: 'T', row: 2 })
  refs('insert-rows-t', { op: 'insertRows', sheet: 'T', row: 2 })
  refs('rename-t', { op: 'renameSheet', sheet: 'T', name: 'U' })
  refs('delete-sheet-t', { op: 'deleteSheet', sheet: 'T' })
  const left = (id: string, op: StructureOp) => add(id, 'left', op, 'B5:J14')
  left('left-delete-col-c', { op: 'deleteColumns', sheet: S, column: 'C' })
  left('left-delete-row-3', { op: 'deleteRows', sheet: S, row: 3 })
  left('left-insert-col-c', { op: 'insertColumns', sheet: S, column: 'C' })
  return { books: { refs: refsBook(), left: leftBook() }, cases }
}
