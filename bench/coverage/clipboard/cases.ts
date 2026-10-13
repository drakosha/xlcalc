/** Clipboard copy cases, predicted before either Excel reads them. */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { Step } from '../../windows/remote.js'

export interface ClipboardCase {
  id: string
  book: BookInput
  range: string
  rows: string[][]
  steps?: Step[]
}

const bookOf = (cells: Record<string, CellInput>): BookInput => ({ sheets: [{ name: 'S', cells }] })
const grid = (): BookInput => bookOf({ A1: { value: 'a' }, B1: { value: 'b' }, A2: { value: 'c' }, B2: { value: 'd' }, A3: { value: 'e' }, B3: { value: 'f' } })

export function clipboardCases(tag: string): ClipboardCase[] {
  const chars = ['line\nnext', 'tab\tstop', 'a"b', '"first', 'last"', 'line\n"quote"', 'carriage\rreturn', 'pair\r\nnext', '', ' tail ']
  return [
    { id: 'grid', book: grid(), range: 'A1:B3', rows: [['a', 'b'], ['c', 'd'], ['e', 'f']] },
    { id: 'special', book: bookOf(Object.fromEntries(chars.map((value, i) => [`A${i + 1}`, { value }]))), range: 'A1:A10', rows: chars.map((x) => [x]) },
    { id: 'single', book: bookOf({ A1: { value: 'a' } }), range: 'A1', rows: [['a']] },
    { id: 'empty', book: bookOf({}), range: 'A1:B2', rows: [['', ''], ['', '']] },
    { id: 'narrow', book: { sheets: [{ name: 'S', cells: { A1: { value: 46096, numberFormat: 'yyyy-mm-dd' }, A2: { value: 123456789012 }, A3: { value: 1.23456789 } }, columns: { A: { width: 3 } } }] }, range: 'A1:A3', rows: [['2026-03-15'], [tag === 'ru-RU' ? '123456789012' : '123456789012'], [tag === 'ru-RU' ? '1,23456789' : '1.23456789']] },
    { id: 'hidden-row', book: grid(), range: 'A1:B3', rows: [['a', 'b'], ['c', 'd'], ['e', 'f']], steps: [{ do: 'hide', sheet: 'S', range: '2:2', hidden: true }] },
    { id: 'hidden-column', book: grid(), range: 'A1:B3', rows: [['a', 'b'], ['c', 'd'], ['e', 'f']], steps: [{ do: 'hideColumns', sheet: 'S', range: 'B:B', hidden: true }] },
    { id: 'filtered', book: grid(), range: 'A1:B3', rows: [['a', 'b'], ['c', 'd']], steps: [{ do: 'filter', sheet: 'S', range: 'A1:B3', field: 1, criterion: 'c' }] },
    { id: 'merged', book: { sheets: [{ name: 'S', cells: { A1: { value: 'merged' }, C1: { value: 'right' }, C2: { value: 'below' } }, merges: ['A1:B2'] }] }, range: 'A1:C2', rows: [['merged', '', 'right'], ['', '', 'below']] },
    { id: 'spill', book: bookOf({ A1: { formula: 'SEQUENCE(2,2)' } }), range: 'A1:B2', rows: [['1', '2'], ['3', '4']] },
  ]
}

/** Draft engine rule to test against the oracle, with a terminal CRLF and quoted control characters or quotes. */
export const predictedText = (rows: string[][]): string => rows.map((row) => row.map((s) => /[\t\r\n"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s).join('\t')).join('\r\n') + '\r\n'

/** Quote-only text and trailing blank cells test whether the plain-text copy is reversible. */
export function quoteCases(): ClipboardCase[] {
  return [
    { id: 'quoted-controls', book: bookOf({ A1: { value: 'a\tb' }, B1: { value: 'a\nb' }, A2: { value: 'a\n"b"' }, B2: { value: 'end' } }), range: 'A1:B2', rows: [['a\tb', 'a\nb'], ['a\n"b"', 'end']] },
    { id: 'quote-only', book: bookOf({ A1: { value: '"abc"' }, A2: { value: '"first' }, A3: { value: 'last"' } }), range: 'A1:A3', rows: [['"abc"'], ['"first'], ['last"']] },
    { id: 'trailing-blanks', book: bookOf({ A1: { value: 'a' } }), range: 'A1:B3', rows: [['a', ''], ['', ''], ['', '']] },
  ]
}
