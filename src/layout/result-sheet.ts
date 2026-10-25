/** A sheet of a computed book read for AutoFit (`fitColumns`, `fitRows`): its cells with their values, and its lines checked. */
import type { BookInput, CellValue, ComputeResult } from '../contract.js'
import { checkBook, checkSheetOf } from '../input.js'
import { InputError } from '../input-error.js'
import { parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { parseSpan } from './spans.js'

/** A cell of a sheet and its value: zero-based row and column. */
export type Placed = readonly [r: number, c: number, v: CellValue]

/**
 * A sheet of a computed book and its result's values, its lines `C`, `B:D` (`5`, `2:9` along rows) checked: the book as
 * checked, the sheet's index, the cells and the span; for `fitColumns` and `fitRows`.
 */
export function resultSheet(book: BookInput, result: ComputeResult, sheet: string, span: string, axis: 'c' | 'r'): { book: BookInput; i: number; cells: Placed[]; lo: number; hi: number } {
  book = checkBook(book)
  const i = book.sheets.indexOf(checkSheetOf(book, sheet, 'sheet'))
  const lines = typeof span === 'string' ? parseSpan(span, axis) : null
  const [path, example] = axis === 'c' ? ['columns', '"C" or "B:D"'] : ['rows', '"5" or "2:9"']
  if (!lines) throw new InputError(path, `expected ${path} such as ${example}, got ${JSON.stringify(span)}`)
  const values = result.sheets.find((x) => sheetKey(x.name) === sheetKey(book.sheets[i].name))
  if (!values) throw new InputError('result', `the result has no sheet ${JSON.stringify(book.sheets[i].name)}`)
  const cells = Object.entries(values.cells).map(([a, x]): Placed => {
    const { row, col } = parseAddress(a)
    return [row, col, x.value]
  })
  return { book, i, cells, lo: lines[0], hi: lines[1] }
}
