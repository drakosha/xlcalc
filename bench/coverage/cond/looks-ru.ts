/**
 * What conditional formats show in a book of ru-RU, Excel in ru-RU: text rules over Cyrillic in any case and over
 * numbers that read with a decimal comma, a cell value rule over text in the locale's order, dates between two, a rule's
 * number format shown by the locale. Predicted before Excel ran it (`looks-ru.predicted.json`).
 */
import type { CellInput, ConditionalFormat, ConditionalStyle } from '../../../src/contract.js'
import type { CondCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'

const RED: ConditionalStyle = { fill: { color: '#FFC7CE' }, font: { color: '#9C0006' } }
const v = (value: CellInput['value']): CellInput => ({ value })

function ruBook(): StoredBook {
  const words = [v('Яблоко'), v('ЯБЛОКО'), v('яблоня'), v('груша'), v(1.5), v('1,5'), v(1234.5), v('ёлка'), v('Ёж'), v('b')]
  const cells: Record<string, StoredCell> = {}
  for (const c of ['A', 'B', 'C', 'D', 'E', 'F', 'G']) words.forEach((w, i) => (cells[`${c}${i + 1}`] = w))
  const dates = [45658, 45689, 45717, 45748, 45778, 45809, 45839, 45870, 45901, 45931]
  dates.forEach((d, i) => (cells[`H${i + 1}`] = { value: d, format: 14 }))
  const t = (type: ConditionalFormat['type'], col: string, rest: object): ConditionalFormat => ({ type, ranges: [`${col}1:${col}10`], ...rest, style: RED }) as ConditionalFormat
  return {
    locale: 'ru-RU',
    sheets: [
      {
        name: 'Лист',
        cells,
        conditionalFormats: [
          t('containsText', 'A', { text: 'яблок' }),
          // a number reads "1,5" in ru-RU
          t('containsText', 'B', { text: ',' }),
          t('beginsWith', 'C', { text: 'ё' }),
          // the order of text: Latin before Cyrillic (Windows) or after (Mac)
          t('cellIs', 'D', { operator: 'greaterThan', formulas: ['"в"'] }),
          t('cellIs', 'E', { operator: 'equal', formulas: ['1.5'] }),
          t('cellIs', 'F', { operator: 'lessThan', formulas: ['"1,6"'] }),
          { type: 'cellIs', ranges: ['G1:G10'], operator: 'greaterThan', formulas: ['1'], style: { format: '#,##0.00', font: { bold: true } } },
          { type: 'cellIs', ranges: ['H1:H10'], operator: 'between', formulas: ['DATE(2025,3,1)', 'DATE(2025,6,30)'], style: RED },
        ],
      },
    ],
  }
}

export function buildCondCases(): CondCases {
  return { locale: 'ru-RU', books: { ru: ruBook() }, cases: [{ id: 'ru', book: 'ru', read: [['Лист', 'A1:H10']] }] }
}
