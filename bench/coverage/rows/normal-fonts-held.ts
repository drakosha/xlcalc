import { NORMAL_FONTS } from '../normal-fonts.js'
import type { RowCase } from './cases.js'

export function buildRowCases(): RowCase[] {
  return NORMAL_FONTS.flatMap(({ id, bookFont }) => [
    { id: `${id}-one-line`, bookFont, cells: { 0: { value: 'Один' } } },
    { id: `${id}-two-lines`, bookFont, cells: { 0: { value: 'One\nДва', look: { alignment: { wrap: true } } } } },
  ])
}
