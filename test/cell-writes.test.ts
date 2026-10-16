/**
 * The log of a sheet's writes (`src/cell-writes.ts`), which `putInput` notes and the index of a sheet's cells reads: the
 * places written since a mark, in order; a log grown past its bound starts again, and a reader of the old one is told.
 */
import { expect, it } from 'vitest'
import { writesMark, writesSince, wrote } from '../src/cell-writes.js'
import { putInput } from '../src/looks/sheet.js'

it('gives the places written since a mark, and null once the log started again', () => {
  const sheet = { name: 'S', cells: {} }
  const mark = writesMark(sheet)
  expect(writesSince(sheet, mark)).toEqual([])
  putInput(sheet, 'B3', 2, 1, { value: 1 })
  putInput(sheet, 'A1', 0, 0, null)
  expect(writesSince(sheet, mark)).toEqual([2 * 16_384 + 1, 0])
  const later = writesMark(sheet)
  expect(writesSince(sheet, later)).toEqual([])
  // past 65,536 places the log starts again
  for (let i = 0; i < 65_536; i++) wrote(sheet, i)
  expect(writesSince(sheet, later)).toBeNull()
  expect(writesSince({}, mark)).toBeNull()
})
