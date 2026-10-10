/**
 * Conditional formats as files hold them, both ways: books written by `writeXlsx` with every kind of rule and option
 * (Excel 2010's data bars and icon sets, custom icons, formulas over other sheets, differential formats with borders,
 * number formats, underlines, a bold taken away), whole columns and rows, several ranges, many rules on one sheet, and
 * books written after structure operations in the engine. Excel must load each book, show what the engine shows and
 * save rules the reader reads back as the book holds them. Predicted before Excel ran it (`files.predicted.json`).
 */
import type { BookInput, CellInput, ConditionalFormat } from '../../../src/contract.js'
import { restructure } from '../../../src/index.js'
import type { CondCase, CondCases } from './cases.js'

const v = (value: CellInput['value']): CellInput => ({ value })

/** Numbers 1 to 10 down each of `cols`, and -2 below them. */
function numbers(cols: string[]): Record<string, CellInput> {
  const out: Record<string, CellInput> = {}
  for (const c of cols) {
    for (let r = 1; r <= 10; r++) out[`${c}${r}`] = v(r)
    out[`${c}11`] = v(-2)
  }
  return out
}

function kindsBook(): BookInput {
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['A1:A11'], operator: 'between', formulas: ['3', '$Z$1'], stopIfTrue: true, style: { fill: { color: '#FFC7CE' }, font: { color: '#9C0006', bold: false } } },
    { type: 'expression', ranges: ['B1:B11'], formula: 'AND($B1>2,ISODD($B1))', style: { font: { bold: true, underline: 'double' }, border: { bottom: { style: 'thin', color: '#FF0000' }, top: { style: 'dashed' } } } },
    { type: 'containsText', ranges: ['C1:C11'], text: '1', style: { font: { strike: true, italic: true } } },
    { type: 'beginsWith', ranges: ['C1:C11'], text: 'a"b' },
    { type: 'endsWith', ranges: ['C1:C11'], text: 'z', style: { fill: { color: '#00B050' } } },
    { type: 'notContainsText', ranges: ['C1:C11'], text: 'q', style: { format: '0.0' } },
    { type: 'timePeriod', ranges: ['D1:D11'], period: 'lastMonth', style: { fill: { color: '#FFEB9C' } } },
    { type: 'containsBlanks', ranges: ['D1:D11'], style: { fill: { color: '#D9D9D9' } } },
    { type: 'notContainsErrors', ranges: ['D1:D11'], style: { format: 10 } },
    { type: 'top10', ranges: ['E1:E11'], rank: 30, percent: true, bottom: true, style: { fill: { color: '#C6EFCE' } } },
    { type: 'aboveAverage', ranges: ['E1:E11'], below: true, stdDev: 2, style: { font: { color: '#0070C0' } } },
    { type: 'uniqueValues', ranges: ['F1:F11'], style: { fill: { color: '#C6EFCE' } } },
    { type: 'duplicateValues', ranges: ['F1:F11'], style: { format: '#,##0.00 "kg"' } },
    { type: 'colorScale', ranges: ['G1:G11'], points: [{ value: { type: 'num', value: '0' }, color: '#5A8AC6' }, { value: { type: 'percentile', value: '40' }, color: '#FCFCFF' }, { value: { type: 'formula', value: '$Z$1*2' }, color: '#F8696B' }] },
    { type: 'dataBar', ranges: ['H1:H11'], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', border: '#638EC6', negativeColor: '#FF0000', negativeBorder: '#FF0000' },
    { type: 'dataBar', ranges: ['I1:I11'], min: { type: 'num', value: '-5' }, max: { type: 'percentile', value: '90' }, color: '#FFB628', gradient: false, axis: 'middle', axisColor: '#7F7F7F', direction: 'rightToLeft', minLength: 5, maxLength: 95, showValue: false },
    { type: 'dataBar', ranges: ['J1:J11'], min: { type: 'min' }, max: { type: 'max' }, color: '#63C384', excel2007: true },
    { type: 'iconSet', ranges: ['K1:K11'], icons: '5Quarters', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '20' } }, { value: { type: 'num', value: '5' }, gte: false }, { value: { type: 'percentile', value: '60' } }, { value: { type: 'formula', value: '$Z$1' } }], reverse: true },
    { type: 'iconSet', ranges: ['L1:L11'], icons: '3Stars', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }], showValue: false },
    { type: 'iconSet', ranges: ['M1:M11'], icons: '3TrafficLights1', points: [{ value: { type: 'percent', value: '0' }, icon: { set: '3Flags', index: 0 } }, { value: { type: 'num', value: '4' }, icon: null }, { value: { type: 'num', value: '8' }, icon: { set: '5Boxes', index: 4 } }] },
    { type: 'expression', ranges: ['N1:N11'], formula: "$N1>'Other sheet'!$A$1", style: { fill: { color: '#FFC000' } } },
    { type: 'cellIs', ranges: ['O1:O11'], operator: 'greaterThan', formulas: ["'Other sheet'!$A$2"], style: { font: { color: '#FF0000' } } },
  ]
  const cells: Record<string, CellInput> = { ...numbers(['A', 'B', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']), Z1: v(8) }
  const words = ['a1', 'a"bc', 'xyz', 'q', 'b12', '', 'ab', 'z', 'Q1', 'x', 'a"b']
  words.forEach((w, i) => (cells[`C${i + 1}`] = v(w)))
  for (let r = 1; r <= 11; r++) cells[`D${r}`] = r === 5 ? { formula: 'NA()' } : r === 6 ? v('') : { formula: `EDATE(TODAY(),-1)+${r - 5}` }
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }, { name: 'Other sheet', cells: { A1: v(4), A2: v(6) } }] }
}

function wideBook(): BookInput {
  const rules: ConditionalFormat[] = [
    { type: 'expression', ranges: ['A:A'], formula: 'ISNUMBER(A1)', style: { fill: { color: '#DDEBF7' } } },
    { type: 'expression', ranges: ['3:4'], formula: 'MOD(COLUMN(),2)=0', style: { font: { bold: true } } },
    { type: 'cellIs', ranges: ['C2:C4', 'E6', 'G1:H2'], operator: 'equal', formulas: ['C2'], style: { fill: { color: '#FFC7CE' } } },
  ]
  // many rules on one sheet, one per row of column J
  for (let r = 1; r <= 60; r++) rules.push({ type: 'cellIs', ranges: [`J${r}`], operator: 'equal', formulas: [String(r)], style: { font: { color: '#C00000' } } })
  const cells: Record<string, CellInput> = { A1: v(1), A2: v('a'), B3: v(2), C2: v(1), C3: v(2), C4: v(3), E6: v(5), G1: v(1), H2: v(2) }
  for (let r = 1; r <= 60; r++) cells[`J${r}`] = v(r % 7 ? r : 0)
  return { sheets: [{ name: 'S', cells, conditionalFormats: rules }] }
}

/** The kinds book after a structure operation in the engine: the file is the engine's book. */
function after(op: Parameters<typeof restructure>[1]): BookInput {
  const r = restructure(kindsBook(), op)
  if (!r.ok) throw new Error(`files: ${op.op} refused`)
  return r.book
}

export function buildCondCases(): CondCases {
  const books: Record<string, BookInput> = {
    kinds: kindsBook(),
    wide: wideBook(),
    'kinds-insert-row': after({ op: 'insertRows', sheet: 'S', row: 3, count: 2 }),
    'kinds-delete-col': after({ op: 'deleteColumns', sheet: 'S', column: 'B' }),
    'kinds-move': after({ op: 'moveRange', from: { sheet: 'S', range: 'A1:C5' }, to: { sheet: 'S', cell: 'Q1' } }),
    'kinds-rename': after({ op: 'renameSheet', sheet: 'Other sheet', name: 'Другой' }),
  }
  const cases: CondCase[] = [
    { id: 'kinds', book: 'kinds', read: [['S', 'A1:G11']] },
    { id: 'wide', book: 'wide', read: [['S', 'A1:H6'], ['S', 'J1:J14']] },
    { id: 'kinds-insert-row', book: 'kinds-insert-row', read: [['S', 'A1:C13']] },
    { id: 'kinds-delete-col', book: 'kinds-delete-col', read: [['S', 'A1:C11']] },
    { id: 'kinds-move', book: 'kinds-move', read: [['S', 'A1:C11'], ['S', 'Q1:S5']] },
    { id: 'kinds-rename', book: 'kinds-rename', read: [['S', 'N1:O11']] },
  ]
  return { books, cases }
}
