/** Sheets the tests of Find and Replace share. */
import type { SheetInput } from '../src/contract.js'

/** F: x1 to x8 in A1:C3, row 2 hidden by a filter (filter mode on) and column B hidden. */
export const filteredSheet = (): SheetInput => ({
  name: 'F',
  cells: { A1: { value: 'x1' }, B1: { value: 'x2' }, C1: { value: 'x3' }, A2: { value: 'x4' }, B2: { value: 'x5' }, C2: { value: 'x6' }, A3: { value: 'x7' }, C3: { value: 'x8' } },
  rows: { 2: { hidden: true } },
  columns: { B: { hidden: true } },
  filtered: true,
})
