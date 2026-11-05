/**
 * A call with fewer arguments than its function takes, in a file or in a workbook. Excel refuses such a file, other
 * writers make them, and one such cell must not cost the rest of the workbook: the conversion of formulas of the old
 * kind (`legacy/convert.ts`) reads each function's arguments by place.
 */
import { describe, expect, it } from 'vitest'
import { compute, FUNCTION_NAMES } from '../src/index.js'
import { fromLegacy, readXlsx, writeXlsx } from '../src/xlsx.js'
import { sheet } from './xlsx-fixtures.js'

const CLOCK = { now: '2026-10-04T12:00:00Z', seed: 7 }
const FEW = ['VLOOKUP(1,A2:B3)', 'VLOOKUP()', 'HLOOKUP(1,A2:B3)', 'HLOOKUP(1)', 'INDEX()', 'XLOOKUP(1)', 'XLOOKUP(1,A2:A3)', 'XLOOKUP()', 'OFFSET()']

describe('a call with too few arguments', () => {
  it.each(FEW)('%s: the file reads, the other cells are whole, and the formula computes as it does in a workbook', async (formula) => {
    const r = await sheet(`<row r="1"><c r="A1"><f>${formula}</f><v>0</v></c><c r="B1"><v>5</v></c><c r="C1"><f>B1+1</f><v>6</v></c></row>`)
    expect(r.book.sheets[0].cells).toEqual({ A1: { formula }, B1: { value: 5 }, C1: { formula: 'B1+1' } })
    const book = { sheets: [{ name: 'S', cells: { A1: { formula }, B1: { value: 5 }, C1: { formula: 'B1+1' } } }] }
    const result = compute(book, { clock: CLOCK })
    expect(compute(r.book, { clock: CLOCK }).sheets[0].cells.A1.value).toEqual(result.sheets[0].cells.A1.value)
    expect(result.sheets[0].cells.C1.value).toBe(6)
  })

  it.each(FEW)('%s: the workbook writes and reads back', async (formula) => {
    const book = { sheets: [{ name: 'S', cells: { A1: { formula }, B1: { value: 5 } } }] }
    const { data } = await writeXlsx({ book })
    expect((await readXlsx(data)).book.sheets[0].cells.B1).toEqual({ value: 5 })
  })

  it('no function takes the conversion down, whatever the number of arguments', () => {
    const thrown: string[] = []
    for (const fn of FUNCTION_NAMES)
      for (const arg of ['1', 'A2:B3'])
        for (let n = 0; n <= 6; n++) {
          const formula = `${fn}(${Array(n).fill(arg).join(',')})`
          try {
            fromLegacy(formula)
          } catch (e) {
            thrown.push(`${formula}: ${(e as Error).message}`)
          }
        }
    expect(thrown).toEqual([])
  })
})
