/**
 * Rules Excel makes itself, through its object model (Excel for Windows: `FormatConditions.Add`, `AddTop10`,
 * `AddAboveAverage`, `AddUniqueValues`, `AddColorScale`, `AddDatabar`, `AddIconSetCondition`), over a book of plain values.
 * The reader must read the rules from the file Excel saved, and the engine show what Excel shows by them. The rules Excel
 * was expected to make (`madeRules`: its defaults, as the engine would hold them) gave the predictions
 * (`made.predicted.json`) before Excel ran it.
 */
import type { BookInput, CellInput, ConditionalFormat, ConditionalStyle } from '../../../src/contract.js'
import type { CondCase, CondCases, CondMake } from './cases.js'

const v = (value: CellInput['value']): CellInput => ({ value })
const f = (formula: string): CellInput => ({ formula })

/** The style the harness gives a rule Excel makes: a red fill, a dark red bold font. */
export const MADE_STYLE: ConditionalStyle = { fill: { color: '#FFC7CE' }, font: { color: '#9C0006', bold: true } }

const DATA = [v(1), v(5), v(10), v('abc'), undefined, v(true), f('NA()'), v(0), v(-3), v(7), v(5), v('xa')]
const COLS = ['B', 'C', 'D', 'E', 'F', 'G', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'R']

function plainBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (const c of ['A', ...COLS]) DATA.forEach((x, i) => x && (cells[`${c}${i + 1}`] = x))
  for (let r = 1; r <= 12; r++) cells[`H${r}`] = f(`TODAY()-${r - 3}`)
  return { sheets: [{ name: 'S', cells }] }
}

const MAKE: CondMake[] = [
  { sheet: 'S', range: 'B1:B12', add: 'cellIs', operator: 5, f1: '=5' },
  { sheet: 'S', range: 'C1:C12', add: 'cellIs', operator: 1, f1: '=2', f2: '=8' },
  { sheet: 'S', range: 'D1:D12', add: 'expression', f1: '=$A1>4' },
  { sheet: 'S', range: 'E1:E12', add: 'text', operator: 0, text: 'a' },
  { sheet: 'S', range: 'F1:F12', add: 'blanks' },
  { sheet: 'S', range: 'G1:G12', add: 'errors' },
  { sheet: 'S', range: 'H1:H12', add: 'period', n: 1 },
  { sheet: 'S', range: 'I1:I12', add: 'top10', n: 3 },
  { sheet: 'S', range: 'J1:J12', add: 'aboveAverage', n: 0 },
  { sheet: 'S', range: 'K1:K12', add: 'unique' },
  { sheet: 'S', range: 'L1:L12', add: 'colorScale', n: 3 },
  { sheet: 'S', range: 'M1:M12', add: 'colorScale', n: 2 },
  { sheet: 'S', range: 'N1:N12', add: 'dataBar' },
  { sheet: 'S', range: 'O1:O12', add: 'iconSet', n: 1 },
  { sheet: 'S', range: 'P1:P5,R1:R5', add: 'expression', f1: '=P1>2' },
]

const ranges = (m: CondMake) => m.range.split(',')

/** The rule Excel is expected to make for each of `MAKE`, with its defaults. */
export function madeRules(): ConditionalFormat[] {
  return MAKE.map((m): ConditionalFormat => {
    const base = { ranges: ranges(m) }
    const strip = (x: string) => x.replace(/^=/, '')
    switch (m.add) {
      case 'cellIs':
        return { ...base, type: 'cellIs', operator: m.operator === 1 ? 'between' : 'greaterThan', formulas: [m.f1!, ...(m.f2 ? [m.f2] : [])].map(strip), style: MADE_STYLE }
      case 'expression':
        return { ...base, type: 'expression', formula: strip(m.f1!), style: MADE_STYLE }
      case 'text':
        return { ...base, type: 'containsText', text: m.text!, style: MADE_STYLE }
      case 'blanks':
        return { ...base, type: 'containsBlanks', style: MADE_STYLE }
      case 'errors':
        return { ...base, type: 'containsErrors', style: MADE_STYLE }
      case 'period':
        return { ...base, type: 'timePeriod', period: 'yesterday', style: MADE_STYLE }
      case 'top10':
        return { ...base, type: 'top10', rank: m.n!, style: MADE_STYLE }
      case 'aboveAverage':
        return { ...base, type: 'aboveAverage', style: MADE_STYLE }
      case 'unique':
        return { ...base, type: 'uniqueValues', style: MADE_STYLE }
      case 'colorScale':
        return m.n === 3
          ? { ...base, type: 'colorScale', points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'percentile', value: '50' }, color: '#FFEB84' }, { value: { type: 'max' }, color: '#63BE7B' }] }
          : { ...base, type: 'colorScale', points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] }
      case 'dataBar':
        return { ...base, type: 'dataBar', min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', border: '#638EC6', negativeColor: '#FF0000', negativeBorder: '#FF0000' }
      case 'iconSet':
        return { ...base, type: 'iconSet', icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }] }
      default:
        throw new Error(`made: ${m.add}`)
    }
  })
}

export function buildCondCases(): CondCases {
  const cases: CondCase[] = [{ id: 'made', book: 'plain', make: MAKE, read: [['S', 'B1:R12']] }]
  return { books: { plain: plainBook() }, cases }
}

/** The plain book with the rules Excel is expected to make. */
export function expectedBook(): BookInput {
  const book = plainBook()
  return { ...book, sheets: [{ ...book.sheets[0], conditionalFormats: madeRules() }] }
}
