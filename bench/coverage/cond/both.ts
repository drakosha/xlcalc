/**
 * A sheet with both conditional formats and data validation (and grouped rows): Excel 2010's rules of both in the sheet's
 * extension list beside each other, as written by `writeXlsx` and after rows inserted in the engine and in Excel. Excel
 * must load each book, show what the engine shows, and save both kinds of rules as the book holds them. Predicted
 * before Excel ran it (`both.predicted.json`).
 */
import type { BookInput, CellInput, ConditionalFormat, DataValidation, StructureOp } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'

const v = (value: CellInput['value']): CellInput => ({ value })

function bothBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 10; r++) for (const c of ['A', 'B', 'C', 'D']) cells[`${c}${r}`] = v(c === 'D' ? ['x', 'y', 'z'][r % 3] : r * (c === 'A' ? 1 : 10))
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['A1:A10'], operator: 'greaterThan', formulas: ['5'], style: { fill: { color: '#FFC7CE' } } },
    { type: 'colorScale', ranges: ['B1:B10'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'dataBar', ranges: ['C1:C10'], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', border: '#638EC6' },
    { type: 'iconSet', ranges: ['C1:C10'], icons: '3Stars', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }] },
    { type: 'expression', ranges: ['D1:D10'], formula: "D1='L'!$A$1", style: { font: { bold: true } } },
  ]
  const validations: DataValidation[] = [
    { range: 'D1:D10', type: 'list', formula1: "'L'!$A$1:$A$3" },
    { range: 'A1:A10', type: 'whole', operator: 'between', formula1: '1', formula2: '100' },
    { range: 'B1:B10', type: 'custom', formula1: 'B1>A1' },
  ]
  return {
    sheets: [
      { name: 'S', cells, conditionalFormats: rules, validations, rows: { '3:4': { outlineLevel: 1 } } },
      { name: 'L', cells: { A1: v('x'), A2: v('y'), A3: v('z') } },
    ],
  }
}

export function buildCondCases(): CondCases {
  const add = (id: string, op?: StructureOp): CondCase => ({ id, book: 'both', ...(op ? { op } : {}), read: [['S', 'A1:D11']] })
  return { books: { both: bothBook() }, cases: [add('written'), add('insert-row', { op: 'insertRows', sheet: 'S', row: 5 }), add('delete-row', { op: 'deleteRows', sheet: 'S', row: 1 })] }
}
