import { CELL_FONTS } from '../fit/cell-fonts-held.js'
import type { RowCase } from './cases.js'

export function buildRowCases(): RowCase[] {
  return CELL_FONTS.flatMap(({ id, font }) => [
    { id: `${id}-one-line`, row: { look: { font } }, cells: { 0: { value: 'Один' } } },
    { id: `${id}-two-lines`, row: { look: { font } }, cells: { 0: { value: 'Один\nДва', look: { font, alignment: { wrap: true } } } } },
  ])
}
