/** The calculation option: written by `writeXlsx` as each of its values, and by `writeKept` set and removed over a file. */
import type { FieldsCase } from './cases.js'

const sheets = [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1*2' } } }]

export function buildFieldsCases(): FieldsCase[] {
  return [
    { id: 'manual', book: { sheets, calculation: 'manual' } },
    { id: 'auto-no-table', book: { sheets, calculation: 'autoNoTable' } },
    { id: 'automatic', book: { sheets } },
    { id: 'kept-manual', book: { sheets }, kept: { calculation: 'manual' } },
    { id: 'kept-automatic', book: { sheets, calculation: 'manual', iteration: { maxIterations: 50 } }, kept: { calculation: null } },
  ]
}
