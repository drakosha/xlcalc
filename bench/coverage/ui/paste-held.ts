/**
 * Held out, predicted by `Session.setCellTexts` and `previewCellTexts` before Excel ran it: text pasted from the clipboard
 * (tabs between columns, line breaks between rows) at B1, each field a cell's text. How each value is read (numbers,
 * fractions, dates and times of both locales, percents, money, booleans and errors, text after an apostrophe, formulas
 * in both locales and one that does not parse), whether a column widens and to what width (long numbers, dates with a
 * time, long text, a column whose width was set by hand), whether a rule of data validation stops a value, cells
 * formatted before as text, percent and date, empty fields and ragged rows over cells that held values, a line break
 * after the last row, and the other line breaks. Both locales' forms are in the set, which runs in each.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, DATA, type UiAction, type UiCase } from './cases.js'

/** One text per row. */
const column = (...texts: string[]) => texts.map((t) => [t])

/** Cells B1 down with a format. */
const formatted = (n: number, numberFormat: string | number): Record<string, CellInput> => Object.fromEntries(Array.from({ length: n }, (_, i) => [`B${i + 1}`, { numberFormat }]))

/** Cells holding text before the paste. */
const old = (...cells: string[]): Record<string, CellInput> => Object.fromEntries(cells.map((a) => [a, { value: 'old' }]))

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const paste = (group: string, rows: string[][], more: Partial<Omit<UiCase, 'id' | 'act'>> = {}, act: Partial<Extract<UiAction, { do: 'paste' }>> = {}) =>
    add(group, { ...more, cells: { ...DATA, ...more.cells }, act: { do: 'paste', at: 'B1', rows, ...act } })

  paste('numbers', column('1234.5', '1,234.5', '1234,5', '1 234,5', '-5', '(5)', '1E+3', '007', ' 5', '5 ', '5 abc'))
  paste('wide', [['1234567890123', '12/31/2025 18:00', '31.12.2025 18:00', 'abcdefghijklmnopqrstuvwxyz']])
  paste('wide', [['1234567890123456789', '3.14159265358979', '3,14159265358979', '12345678.9']])
  paste('fractions', column('1 1/2', '0 1/3', '3/4', '1 1/12'))
  paste('dates', column('1/31/2025', '31.01.2025', '2025-01-31', '31-Jan-2025', 'Jan 2025', '15.3', '3/15', '31 янв 2025'))
  paste('times', column('18:30', '6:30 PM', '18:30:15', '25:30', '18:30:15,5'))
  paste('percents', column('12%', '12.5%', '12,5%', '-5%', '5 1/2%', '%5'))
  paste('money', column('$1,234.50', '$5', '($5)', '1 234,50 ₽', '5 ₽', '€5'))
  paste('bools', column('TRUE', 'false', 'ИСТИНА', 'ложь', '#N/A', '#Н/Д', '#DIV/0!'))
  paste('apostrophe', column("'007", "'1.5", "'=1+2", "'abc", "'TRUE"))
  paste('formulas', column('=A1*2', '=A1+A3', '-A1', '+A1', '=A1&"x"', '=1/0'))
  paste('formula-en', column('=SUM(A1:A3)', '=SUM(A1,1.5)', '5'))
  paste('formula-ru', column('=СУММ(A1:A3)', '=СУММ(A1;1,5)', '5'))
  paste('bad', column('=1+', '5'))
  paste('grid', [['1', 'abc', '1/31/2025'], ['2.5', '=A1*2', 'TRUE']])
  // empty fields and a short row over cells that held text: which are cleared
  paste('empty', [['1', '', '3'], ['4']], { cells: old('B1', 'C1', 'D1', 'B2', 'C2', 'D2'), read: ['C2', 'D2'] })
  // a line break after the last row: the cell below keeps its text
  paste('trailing', [['5']], { cells: old('B2'), read: ['B2'] }, { trailing: true })
  paste('newline', [['1', '2'], ['3', '4']], {}, { newline: '\r\n' })
  paste('newline', [['1', '2'], ['3', '4']], {}, { newline: '\r' })
  paste('text-cell', column('007', '1.5', '=A1*2', '1/31/2025', '12%', 'TRUE'), { cells: formatted(6, '@') })
  paste('percent-cell', column('5', '0.5', '5%', 'abc'), { cells: formatted(4, '0%') })
  paste('date-cell', column('45688', '18:30', '5%'), { cells: formatted(3, 14) })
  paste('hand-width', [['1234567890123']], { columns: { B: { width: 5 } } })
  paste('valid', column('-2', '5', 'abc'), { validations: [{ range: 'B1:B3', type: 'whole', operator: 'greaterThan', formula1: '0' }] })
  paste('valid', column('c', 'a'), { validations: [{ range: 'B1:B2', type: 'list', formula1: '"a,b"' }] })
  return out
}
