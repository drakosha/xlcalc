/** A manually assigned height leaves AutoFit, including after undo, batched layout edits and xlsx roundtrip. */
import { afterAll, describe, expect, it } from 'vitest'
import type { BookInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({ styles: [{ alignment: { wrap: true } }], sheets: [{ name: 'S', cells: { A1: { value: 'x\ny' }, A2: { value: 'x\ny' }, A3: { value: 'x\ny' } }, rows: { '1:3': { height: 32, bestFit: true, style: 0, numberFormat: '@', outlineLevel: 1 } } }] })

describe.each(['closures', 'eval'] as const)('manual row height, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)

    for (const entry of ['setCellText', 'setCellTexts', 'fillCellText'] as const) it(`keeps a manual height on ${entry}, preserving the other row fields and neighboring rows`, async () => {
      const s = await open(book(), { clock: CLOCK, execution })
      await s.setLayout([{ sheet: 'S', rows: '2', set: { height: 40 } }])
      if (entry === 'setCellText') await s.setCellText('S', 'A2', 'x\ny\nz')
      else if (entry === 'setCellTexts') await s.setCellTexts([{ sheet: 'S', cell: 'A2', text: 'x\ny\nz' }])
      else await s.fillCellText('S', 'A2:B2', 'x\ny\nz')
      expect((await s.getViewport('S', 'A1:A3')).rows.map((r) => r.height)).toEqual([32, 40, 32])
      expect((await s.getLayout()).sheets[0].rows).toEqual({
        1: { height: 32, bestFit: true, style: 0, numberFormat: '@', outlineLevel: 1 },
        2: { height: 40, style: 0, numberFormat: '@', outlineLevel: 1 },
        3: { height: 32, bestFit: true, style: 0, numberFormat: '@', outlineLevel: 1 },
      })
      await s.close()
    })

    it('undo and redo restore whether a height was fitted, and AutoFit enables fitting on entry again', async () => {
      const s = await open(book(), { clock: CLOCK, execution })
      await s.setLayout([{ sheet: 'S', rows: '2', set: { height: 40 } }])
      await s.undo()
      expect((await s.getLayout()).sheets[0].rows).toEqual(book().sheets[0].rows)
      await s.redo()
      await s.setCellText('S', 'A2', 'x\ny\nz')
      expect((await s.getLayout()).sheets[0].rows?.[2].height).toBe(40)
      await s.setLayout([{ sheet: 'S', rows: '2', fit: true }])
      expect((await s.getLayout()).sheets[0].rows?.[2]).toMatchObject({ height: 48, bestFit: true })
      await s.setCellText('S', 'A2', 'x\ny\nz\nw')
      expect((await s.getLayout()).sheets[0].rows?.[2].height).toBe(64)
      await s.close()
    })

    for (const height of [32, 0, null]) it(`clears the fitted mark when height is assigned ${height}`, async () => {
      const s = await open(book(), { clock: CLOCK, execution })
      await s.setLayout([{ sheet: 'S', rows: '2', set: { height } }])
      expect((await s.getLayout()).sheets[0].rows?.[2].bestFit).toBeUndefined()
      await s.setCellText('S', 'A2', 'x\ny\nz')
      const row = (await s.getViewport('S', 'A2')).rows[0]
      expect(row.height).toBe(height === null ? 48 : height)
      expect(row.hidden ?? false).toBe(height === 0)
      await s.close()
    })

    it('preserves an explicit bestFit and leaves it alone on patches without height', async () => {
      const s = await open(book(), { clock: CLOCK, execution })
      await s.setLayout([{ sheet: 'S', rows: '2', set: { outlineLevel: 2 } }])
      expect((await s.getLayout()).sheets[0].rows?.[2]).toMatchObject({ height: 32, bestFit: true, outlineLevel: 2 })
      await s.setLayout([{ sheet: 'S', rows: '2', set: { height: 40, bestFit: true } }])
      await s.setCellText('S', 'A2', 'x\ny\nz')
      expect((await s.getLayout()).sheets[0].rows?.[2]).toMatchObject({ height: 48, bestFit: true })
      await s.setLayout([{ sheet: 'S', rows: '2', set: { height: 40, bestFit: false } }])
      await s.setCellText('S', 'A2', 'x\ny\nz\nw')
      expect((await s.getLayout()).sheets[0].rows?.[2]).toMatchObject({ height: 40, bestFit: false })
      await s.close()
    })

    it('applies the same rule through applyChanges and previewChanges, and retains manual heights in xlsx', async () => {
      const s = await open(book(), { clock: CLOCK, execution })
      const ops = [{ op: 'setLayout' as const, edits: [{ sheet: 'S', rows: '2', set: { height: 40 } }] }]
      await s.previewChanges(ops)
      expect((await s.getLayout()).sheets[0].rows).toEqual(book().sheets[0].rows)
      await s.applyChanges(ops)
      await s.setCellText('S', 'A2', 'x\ny\nz')
      expect((await s.getLayout()).sheets[0].rows?.[2].height).toBe(40)
      const { data } = await writeXlsx(await s.getSnapshot())
      const read = await readXlsx(data)
      expect(read.book.sheets[0].rows?.[2]).toMatchObject({ height: 40 })
      expect(read.book.sheets[0].rows?.[2].bestFit).toBeUndefined()
      await s.close()
    })

    it('honors the order of a manual height and AutoFit within a layout call', async () => {
      const s = await open(book(), { clock: CLOCK, execution })
      await s.setLayout([{ sheet: 'S', rows: '2', fit: true }, { sheet: 'S', rows: '2', set: { height: 40 } }])
      await s.setCellText('S', 'A2', 'x\ny\nz')
      expect((await s.getLayout()).sheets[0].rows?.[2].height).toBe(40)
      await s.setLayout([{ sheet: 'S', rows: '2', set: { height: 40 } }, { sheet: 'S', rows: '2', fit: true }])
      expect((await s.getLayout()).sheets[0].rows?.[2]).toMatchObject({ height: 48, bestFit: true })
      await s.close()
    })
  })
})
