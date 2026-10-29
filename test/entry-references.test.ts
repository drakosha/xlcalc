/**
 * A typed formula in General takes the format of the cells it refers to (the entry sets `refs`, `refs-held`,
 * `refs-held-2`, the ui set `fill-refs-held`): `parseEntry` by `numberFormatAt`, and every kind of session by its cells,
 * their rows and columns, one step of undo with the entry; a fill gives every cell its active cell's format, AutoSum
 * the format of its range's first cell.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { InputError, parseEntry, type BookInput, type NumberFormat } from '../src/index.js'
import { parseEntry as entryParse } from '../src/typed-entry.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

describe('parseEntry numberFormatAt', () => {
  const formats: Record<string, NumberFormat> = { C1: 9, C2: 14, C3: 7, 'T!A1': '0.0%' }
  const calls: [string | null, string][] = []
  const numberFormatAt = (sheet: string | null, cell: string) => {
    calls.push([sheet, cell])
    return formats[sheet === null ? cell : `${sheet}!${cell}`]
  }
  const typed = (text: string, cell?: NumberFormat) => parseEntry(text, { numberFormatAt, ...(cell === undefined ? {} : { cell: { numberFormat: cell } }) })

  it('gives a formula the format of its first reference with one, through the operators and functions Excel passes it by', () => {
    expect(typed('=C1+C9')).toEqual({ ok: true, input: { formula: 'C1+C9', numberFormat: 9 } })
    expect(typed('=C9-C3')).toEqual({ ok: true, input: { formula: 'C9-C3', numberFormat: 7 } })
    expect(typed('=C2-C2')).toEqual({ ok: true, input: { formula: 'C2-C2' } })
    expect(typed('=C1*2')).toEqual({ ok: true, input: { formula: 'C1*2' } })
    expect(typed('=MOD(C3,7)')).toEqual({ ok: true, input: { formula: 'MOD(C3,7)', numberFormat: 7 } })
    expect(typed('=ABS(C3)')).toEqual({ ok: true, input: { formula: 'ABS(C3)' } })
    expect(typed('=T!A1')).toEqual({ ok: true, input: { formula: "'T'!A1", numberFormat: '0.0%' } })
    expect(entryParse('=C2+1', { numberFormatAt })).toEqual({ ok: true, input: { formula: 'C2+1', numberFormat: 14 } })
  })

  it('asks at most once for each reference, by the first cell of a range, the sheet as the formula names it, and stops at the first format', () => {
    calls.length = 0
    expect(typed('=SUM(C:C)+SUM(C9:C1)')).toEqual({ ok: true, input: { formula: 'SUM(C:C)+SUM(C1:C9)', numberFormat: 9 } })
    expect(typed('=SUM(t!A1:A9)')).toEqual({ ok: true, input: { formula: "SUM('t'!A1:A9)" } })
    expect(calls).toEqual([
      [null, 'C1'],
      ['t', 'A1'],
    ])
  })

  it("keeps a cell's format as typing a constant of the format's kind keeps it", () => {
    expect(typed('=C1', 2)).toEqual({ ok: true, input: { formula: 'C1', numberFormat: 2 } })
    expect(typed('=C2', 9)).toEqual({ ok: true, input: { formula: 'C2', numberFormat: 14 } })
    expect(typed('=C9', 9)).toEqual({ ok: true, input: { formula: 'C9', numberFormat: 9 } })
  })

  it('gives a name and no function nothing', () => {
    calls.length = 0
    expect(typed('=Rate+1')).toEqual({ ok: true, input: { formula: 'Rate+1' } })
    expect(parseEntry('=C1')).toEqual({ ok: true, input: { formula: 'C1' } })
    expect(calls).toEqual([])
  })

  it('refuses what is not a function, and a format the function gives that is none', () => {
    expect(() => parseEntry('=C1', { numberFormatAt: 9 as never })).toThrow(InputError)
    expect(() => parseEntry('=C1', { numberFormatAt: () => -1 })).toThrow(/numberFormatAt/)
  })
})

const BOOK: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 0.1, numberFormat: 9 }, A2: { value: 0.2, numberFormat: 9 }, A3: { value: 2 }, A4: { value: 46096, numberFormat: 14 }, A5: { value: 100, numberFormat: 7 }, E1: { value: 1 } },
      columns: { E: { numberFormat: '0.0%' } },
    },
  ],
}

describe.each(['closures', 'eval'] as const)('formats by reference, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)

    it('typed, edited in place and pasted, one step of undo with the entry', async () => {
      const s = await open(BOOK, { clock: CLOCK, execution })
      await s.setCellText('S', 'B1', '=A3+A1')
      await s.setCellText('S', 'B2', '=E1', { inPlace: true })
      await s.pasteText('S', 'B3', '=A5-A3')
      const cells = (await s.getBook()).sheets[0].cells
      expect([cells.B1, cells.B2, cells.B3]).toEqual([
        { formula: 'A3+A1', numberFormat: 9 },
        { formula: 'E1', numberFormat: '0.0%' },
        { formula: 'A5-A3', numberFormat: 7 },
      ])
      await s.undo()
      expect((await s.getBook()).sheets[0].cells.B3).toBeUndefined()
      await s.close()
    })

    it('a formula among several typed at once refers to the cells typed before it as they are then', async () => {
      const s = await open(BOOK, { clock: CLOCK, execution })
      await s.setCellTexts([
        { sheet: 'S', cell: 'C1', text: '5%' },
        { sheet: 'S', cell: 'C2', text: '=C1+1' },
        { sheet: 'S', cell: 'A1', text: '' },
        { sheet: 'S', cell: 'C3', text: '=A1' },
      ])
      const cells = (await s.getBook()).sheets[0].cells
      expect([cells.C2, cells.C3]).toEqual([{ formula: 'C1+1', numberFormat: 9 }, { formula: 'A1', numberFormat: 9 }])
      await s.close()
    })

    it("a fill takes its active cell's format in every cell, AutoSum its range's first cell's", async () => {
      const s = await open(BOOK, { clock: CLOCK, execution })
      await s.fillCellText('S', 'B1:B3', '=A1+A3')
      const sum = await s.getAutoSum('S', 'A6')
      await s.setCellText('S', 'A6', `=${sum.formulaLocal}`)
      const cells = (await s.getBook()).sheets[0].cells
      expect([cells.B1, cells.B2, cells.B3].map((c) => c?.numberFormat)).toEqual([9, 9, 9])
      expect(cells.A6).toEqual({ formula: sum.formula, numberFormat: 7 })
      await s.close()
    })
  })
})
