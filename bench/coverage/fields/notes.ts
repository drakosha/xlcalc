/** Notes shown all the time: written by `writeXlsx` beside notes shown on hover, and by `writeKept` shown and hidden over a file. */
import type { SheetInput } from '../../../src/contract.js'
import type { FieldsCase } from './cases.js'

const sheets: SheetInput[] = [
  { name: 'S', cells: { A1: { value: 1, note: { author: 'Ann', text: 'shown', shown: true } }, B2: { value: 2, note: { text: 'on hover' } }, C3: { note: { text: 'shown too', shown: true } } } },
  { name: 'T', cells: { D4: { value: 'x', note: { text: 'on hover' } } } },
]

export function buildFieldsCases(): FieldsCase[] {
  return [
    { id: 'written', book: { sheets } },
    { id: 'kept-shown', book: { sheets }, kept: { cells: [{ sheet: 'T', cell: 'D4', input: { note: { text: 'on hover', shown: true } } }, { sheet: 'S', cell: 'A1', input: { note: { author: 'Ann', text: 'shown' } } }] } },
  ]
}
