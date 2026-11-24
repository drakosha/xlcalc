import type { FitCase } from './cases.js'
import { NORMAL_FONTS } from '../normal-fonts.js'

export function buildFitCases(): FitCase[] {
  return NORMAL_FONTS.flatMap(({ id, bookFont }) => [
    { id: `${id}-standard`, bookFont, readWidth: true, cells: { 1: { value: 12345678 } } },
    { id: `${id}-file-width`, bookFont, readWidth: true, column: { width: 80 }, cells: { 1: { value: 18 } } },
    { id: `${id}-fraction`, bookFont, readWidth: true, column: { width: 0.75 }, cells: { 1: { value: 8 } } },
    { id: `${id}-number`, bookFont, readWidth: true, column: { width: 5 }, cells: { 1: { value: 8.25, format: '0.00' } } },
  ])
}
