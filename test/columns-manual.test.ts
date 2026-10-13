/** A width set by hand leaves AutoFit: a number typed later does not widen the column, as in Excel. */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import type { BookInput } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 } }, columns: { A: { width: 3.83203125, bestFit: true, numberFormat: '0' } } }] })

describe.each(['closures', 'eval'] as const)('manual column width, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    let open: Awaited<ReturnType<typeof opener>>
    beforeAll(async () => { open = await opener(kind) })

    it('a width set after AutoFit clears bestFit, keeps the other fields, and a typed number leaves it', async () => {
      const s = await open(book(), { clock: CLOCK, execution })
      await s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 3 } }])
      expect((await s.getLayout()).sheets[0].columns).toEqual({ A: { width: 3, numberFormat: '0' } })
      await s.setCellText('S', 'A2', '123456789012')
      expect((await s.getLayout()).sheets[0].columns).toEqual({ A: { width: 3, numberFormat: '0' } })
      await s.undo()
      await s.undo()
      expect((await s.getLayout()).sheets[0].columns).toEqual(book().sheets[0].columns)
      await s.close()
    })

    it('an explicit bestFit is kept, and a patch without width keeps the mark', async () => {
      const s = await open(book(), { clock: CLOCK, execution })
      await s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 5, bestFit: true } }])
      expect((await s.getLayout()).sheets[0].columns).toEqual({ A: { width: 5, bestFit: true, numberFormat: '0' } })
      await s.setLayout([{ sheet: 'S', columns: 'A', set: { numberFormat: '0.0' } }])
      expect((await s.getLayout()).sheets[0].columns).toEqual({ A: { width: 5, bestFit: true, numberFormat: '0.0' } })
      await s.close()
    })
  })
})
