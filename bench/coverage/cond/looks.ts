/**
 * What conditional formats show, en-US: books written by `writeXlsx` with rules of each kind over cells of each kind
 * (numbers, text, booleans, blanks, empty text, errors, dates), Excel reading every cell's display format. Some cases
 * decide between readings of Excel's rules: the cell a formula of several ranges is written for, a formula answering
 * a range, the text of a text rule against its formula, `between` with its ends reversed, an error among a colour
 * scale's or a top rule's cells, the standard deviation of an average rule, a colour scale over equal numbers.
 * Predicted before Excel ran it (`looks.predicted.json`).
 */
import type { BookInput, CellInput, ConditionalFormat, ConditionalStyle, ConditionalValue } from '../../../src/contract.js'
import type { CondCase, CondCases } from './cases.js'
import type { StoredRule, StoredStyle } from '../stored.js'
import { formulaCell as f, valueCell as v } from '../stored.js'

const RED: ConditionalStyle = { fill: { color: '#FFC7CE' }, font: { color: '#9C0006' } }

/** The same cells in each of `cols`, from row 1 down; an undefined entry stays blank. */
function columns(values: (CellInput | undefined)[], cols: string[]): Record<string, CellInput> {
  const out: Record<string, CellInput> = {}
  for (const c of cols) values.forEach((v, i) => v && (out[`${c}${i + 1}`] = v))
  return out
}

/** A mix of kinds: numbers, text, empty text, a boolean, an error, numeric text, a blank. */
const MIX = [v(1), v(5), v(10), v('abc'), v(''), v(true), f('NA()'), v(0), v(-3), v('10'), undefined, v(7)]

function cellIsBook(): BookInput {
  const cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
  const rule = (col: string, r: Omit<Extract<ConditionalFormat, { type: 'cellIs' }>, 'type' | 'ranges' | 'style'>): ConditionalFormat => ({ type: 'cellIs', ranges: [`${col}1:${col}12`], ...r, style: RED })
  return {
    sheets: [
      {
        name: 'S',
        cells: { ...columns(MIX, cols), Z2: v(3) },
        conditionalFormats: [
          rule('B', { operator: 'greaterThan', formulas: ['5'] }),
          rule('C', { operator: 'lessThan', formulas: ['5'] }),
          rule('D', { operator: 'equal', formulas: ['"ABC"'] }),
          // its ends reversed
          rule('E', { operator: 'between', formulas: ['10', '1'] }),
          rule('F', { operator: 'notBetween', formulas: ['1', '5'] }),
          rule('G', { operator: 'notEqual', formulas: ['5'] }),
          rule('H', { operator: 'greaterThanOrEqual', formulas: ['"abc"'] }),
          // a reference to a blank cell, and a relative one that moves down
          rule('I', { operator: 'lessThanOrEqual', formulas: ['$Z1'] }),
        ],
      },
    ],
  }
}

function expressionBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 6; r++) cells[`A${r}`] = v(r - 1)
  for (const c of ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']) for (let r = 1; r <= 6; r++) cells[`${c}${r}`] = v(`${c}${r}`)
  const rule = (ranges: string[], formula: string): ConditionalFormat => ({ type: 'expression', ranges, formula, style: RED })
  return {
    sheets: [
      {
        name: 'S',
        cells,
        conditionalFormats: [
          rule(['B1:B6'], '$A1>3'),
          // a formula answering a range: Excel's implicit intersection (row by row) or its first value
          rule(['C1:C6'], '$A$1:$A$6>3'),
          // text "TRUE", a number, an error
          rule(['D1:D6'], 'IF($A1>3,"TRUE","")'),
          rule(['E1:E6'], '$A1'),
          rule(['F1:F6'], '1/($A1-3)'),
          rule(['G1:G6'], 'MOD(ROW(),2)=0'),
          // two ranges, the first not the top one: its formula is written for H3, or for the top-left H1
          rule(['H3:H6', 'J1:J3'], '$A3>2'),
          rule(['K1:K6'], '$A1>lim'),
          // another sheet's cell: Excel 2010's rule
          rule(['L1:L6'], "$A1>'T'!$A$1"),
        ],
      },
      { name: 'T', cells: { A1: v(2) } },
    ],
    names: [{ name: 'lim', ref: '3' }],
  }
}

function textBook(): BookInput {
  const words = [v('Apple'), v('apple pie'), v('PINEAPPLE'), v('banana'), v(123), v('a*b'), v('axb'), v(''), undefined, v('ap')]
  const blanks = [v(''), v(' '), undefined, v('x'), v(0), f('""'), f('NA()'), f('1/0'), v(true), v('#N/A')]
  const t = (type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith', col: string, text: string, formula?: string): ConditionalFormat => ({ type, ranges: [`${col}1:${col}10`], text, ...(formula ? { formula } : {}), style: RED })
  const b = (type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors', col: string): ConditionalFormat => ({ type, ranges: [`${col}1:${col}10`], style: RED })
  return {
    sheets: [
      {
        name: 'S',
        cells: { ...columns(words, ['B', 'C', 'D', 'E', 'F', 'G', 'H']), ...columns(blanks, ['I', 'J', 'K', 'L']) },
        conditionalFormats: [
          t('containsText', 'B', 'apple'),
          t('notContainsText', 'C', 'apple'),
          t('beginsWith', 'D', 'ap'),
          t('endsWith', 'E', 'LE'),
          // a wildcard in the text
          t('containsText', 'F', 'a*b'),
          // a number's digits
          t('containsText', 'G', '2'),
          // the text says one thing and the formula another: which does Excel compute?
          t('containsText', 'H', 'zzz', 'NOT(ISERROR(SEARCH("apple",H1)))'),
          b('containsBlanks', 'I'),
          b('notContainsBlanks', 'J'),
          b('containsErrors', 'K'),
          b('notContainsErrors', 'L'),
        ],
      },
    ],
  }
}

function periodBook(): BookInput {
  const days = [f('TODAY()'), f('TODAY()-1'), f('TODAY()+1'), f('TODAY()-6'), f('TODAY()-7'), f('TODAY()-WEEKDAY(TODAY())'), f('EDATE(TODAY(),-1)'), f('EDATE(TODAY(),1)'), f('TODAY()+0.5'), v('x')]
  const periods = ['today', 'yesterday', 'tomorrow', 'last7Days', 'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth'] as const
  const cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
  return {
    sheets: [
      {
        name: 'S',
        cells: columns(days, cols),
        conditionalFormats: periods.map((period, i): ConditionalFormat => ({ type: 'timePeriod', ranges: [`${cols[i]}1:${cols[i]}10`], period, style: RED })),
      },
    ],
  }
}

function topBook(): BookInput {
  const nums = [10, 20, 20, 30, 40, 50, 60, 70, 80, 90, 100, 30].map(v)
  const withText = [...nums.slice(0, 10), v('x'), undefined]
  const withError = [...nums.slice(0, 10), f('NA()'), v(5)]
  // 1..9, 30 and 16.4 (or -1.3): the one lies above the mean by more than one population deviation and less than one sample deviation (the other below)
  const spread = [1, 2, 3, 4, 5, 6, 7, 8, 9, 30].map(v)
  const dups = [v(1), v('1'), v('a'), v('A'), v(true), v(true), v(2), v(2), f('NA()'), f('NA()'), undefined, v('')]
  const rule = (col: string, r: Partial<ConditionalFormat> & { type: ConditionalFormat['type'] }): ConditionalFormat => ({ ranges: [`${col}1:${col}12`], style: RED, ...r }) as ConditionalFormat
  return {
    sheets: [
      {
        name: 'S',
        cells: { ...columns(withText, ['B', 'C']), ...columns(nums, ['D', 'E', 'G', 'H', 'I']), ...columns(withError, ['F']), ...columns(spread, ['J', 'K']), ...columns(dups, ['L', 'M']), J11: v(16.4), K11: v(-1.3) },
        conditionalFormats: [
          rule('B', { type: 'top10', rank: 3 }),
          rule('C', { type: 'top10', rank: 2, bottom: true }),
          // 12 numbers: 10 percent is 1.2 of them, 25 percent 3
          rule('D', { type: 'top10', rank: 10, percent: true }),
          rule('E', { type: 'top10', rank: 25, percent: true }),
          rule('F', { type: 'top10', rank: 3 }),
          rule('G', { type: 'aboveAverage' }),
          rule('H', { type: 'aboveAverage', below: true }),
          rule('I', { type: 'aboveAverage', equal: true }),
          rule('J', { type: 'aboveAverage', stdDev: 1 }),
          rule('K', { type: 'aboveAverage', below: true, stdDev: 1 }),
          rule('L', { type: 'duplicateValues' }),
          rule('M', { type: 'uniqueValues' }),
        ],
      },
    ],
  }
}

function scaleBook(): BookInput {
  const ten = Array.from({ length: 10 }, (_, i) => v(i + 1))
  const skew = [1, 2, 3, 4, 5, 6, 7, 8, 100, 200].map(v)
  const wide = [-5, 0, 10, 50, 90, 100, 150, 25, 75, 60].map(v)
  const same = [v(5), v(5), v(5)]
  const mixed = [v(1), v('x'), undefined, v(true), v(5), v(10), v(''), v(3), v('7'), v(9)]
  const errs = [v(1), v(2), f('NA()'), v(4), v(5), v(6), v(7), v(8), v(9), v(10)]
  const dates = [45658, 45689, 45717, 45748, 45778, 45809, 45839, 45870, 45901, 45931].map(v)
  const two = (col: string, min: ConditionalValue, max: ConditionalValue): ConditionalFormat => ({ type: 'colorScale', ranges: [`${col}1:${col}10`], points: [{ value: min, color: '#F8696B' }, { value: max, color: '#63BE7B' }] })
  const three = (col: string, mid: ConditionalValue): ConditionalFormat => ({
    type: 'colorScale',
    ranges: [`${col}1:${col}10`],
    points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: mid, color: '#FFEB84' }, { value: { type: 'max' }, color: '#63BE7B' }],
  })
  return {
    sheets: [
      {
        name: 'S',
        cells: { ...columns(ten, ['B']), ...columns(skew, ['C']), ...columns(wide, ['D', 'E', 'F']), ...columns(same, ['G']), ...columns(mixed, ['H']), ...columns(errs, ['I']), ...columns(dates, ['J']), ...columns(ten, ['K']), Z1: v(20), Z2: v(80) },
        conditionalFormats: [
          two('B', { type: 'min' }, { type: 'max' }),
          three('C', { type: 'percentile', value: '50' }),
          two('D', { type: 'num', value: '0' }, { type: 'num', value: '100' }),
          two('E', { type: 'percent', value: '20' }, { type: 'percent', value: '80' }),
          two('F', { type: 'formula', value: '$Z$1' }, { type: 'formula', value: '$Z$2' }),
          two('G', { type: 'min' }, { type: 'max' }),
          two('H', { type: 'min' }, { type: 'max' }),
          two('I', { type: 'min' }, { type: 'max' }),
          two('J', { type: 'min' }, { type: 'max' }),
          three('K', { type: 'num', value: '3' }),
        ],
      },
    ],
  }
}

function priorityBook(): BookInput {
  const ten = Array.from({ length: 10 }, (_, i) => v(i + 1))
  const cells = columns(ten, ['A', 'B', 'C', 'D', 'E', 'F'])
  for (let r = 1; r <= 10; r++) cells[`E${r}`] = { value: r, style: 0 }
  for (let r = 1; r <= 10; r++) cells[`F${r}`] = { value: r * 1.5 }
  const gt = (col: string, x: number, style: StoredStyle, stopIfTrue = false): StoredRule => ({ type: 'cellIs', ranges: [`${col}1:${col}10`], operator: 'greaterThan', formulas: [String(x)], style, ...(stopIfTrue ? { stopIfTrue } : {}) })
  const scale = (col: string): ConditionalFormat => ({ type: 'colorScale', ranges: [`${col}1:${col}10`], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] })
  return {
    styles: [{ fill: { color: '#FFFF00' }, font: { italic: true } }],
    sheets: [
      {
        name: 'S',
        cells,
        conditionalFormats: [
          // A: two rules hold, their fields merge, the first one's fill wins
          gt('A', 5, { fill: { color: '#FF0000' }, font: { bold: true } }),
          gt('A', 3, { fill: { color: '#00FF00' }, font: { color: '#0000FF', italic: true } }),
          // B: the first stops the second where it holds
          gt('B', 5, { fill: { color: '#FF0000' } }, true),
          gt('B', 3, { font: { italic: true } }),
          // C: a fill before a colour scale; D: a colour scale before a fill
          gt('C', 5, { fill: { color: '#FF0000' } }),
          scale('C'),
          scale('D'),
          gt('D', 5, { fill: { color: '#FF0000' } }),
          // E: cells with a yellow italic style of their own, a rule's red font over it
          gt('E', 5, { font: { color: '#FF0000', italic: false } }),
          // F: a rule's number format
          gt('F', 5, { format: '0.00%', font: { bold: true } }),
        ],
      },
    ],
  }
}

export function buildCondCases(): CondCases {
  const books = { cellIs: cellIsBook(), expression: expressionBook(), text: textBook(), period: periodBook(), top: topBook(), scale: scaleBook(), priority: priorityBook() }
  const cases: CondCase[] = [
    { id: 'cell-is', book: 'cellIs', read: [['S', 'B1:I12']] },
    { id: 'expression', book: 'expression', read: [['S', 'B1:L6']] },
    { id: 'text', book: 'text', read: [['S', 'B1:L10']] },
    { id: 'period', book: 'period', read: [['S', 'B1:K10']] },
    { id: 'top', book: 'top', read: [['S', 'B1:M12']] },
    { id: 'scale', book: 'scale', read: [['S', 'B1:K10']] },
    { id: 'priority', book: 'priority', read: [['S', 'A1:F10']] },
  ]
  return { books, cases }
}
