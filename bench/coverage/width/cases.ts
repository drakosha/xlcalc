/** Fixed xlsx widths and ColumnWidth settings, predicted from the current model before Excel runs. */
import type { BookInput } from '../../../src/contract.js'
export const FILE_WIDTHS = [8.83203125, 9.140625, 3, 5, 10, 15.5, 0.5, 0, 50, 255]
export const SET_WIDTHS = [0.5, 1, 3, 8, 8.43, 10, 15.5]
export const COLUMNS = 'ABCDEFGHIJKLMNOPQR'.split('')
export const widthBook = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 'width' } }, columns: Object.fromEntries(FILE_WIDTHS.map((width, i) => [COLUMNS[i + 1], { width }])) }] })
