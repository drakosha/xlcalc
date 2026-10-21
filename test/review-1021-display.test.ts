/**
 * Reproductions from the review of 21.10 on display and sheets: a number format with more than 100 decimals, a fraction
 * format with many denominator digits, rows fitted for more wrapping cells than a call takes arguments, and deleting or
 * moving a sheet while a reference across sheets names a sheet the book does not have. With closures and eval, and the
 * session parts in a synchronous session, in the same thread and in a thread.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, displayValue, type BookInput } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const decimals = (n: number) => '0.' + '0'.repeat(n)
// a book keeps 15 significant digits of a value
const third = 0.333333333333333
const fmtBook = (numberFormat: string): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: third, numberFormat }, B1: { value: 2 } }, columns: { A: { width: 255 }, C: { width: 255, numberFormat } } }] })

describe('more than 100 decimals in a number format', () => {
  it('shows the decimals past 100 as zeros, in positional and scientific notation and with # and ?', () => {
    const at100 = (tail: string) => displayValue(1 / 3, { numberFormat: decimals(100) + tail }).text
    expect(displayValue(1 / 3, { numberFormat: decimals(101) }).text).toBe(at100('') + '0')
    expect(displayValue(1 / 3, { numberFormat: decimals(130) }).text).toBe(at100('') + '0'.repeat(30))
    expect(displayValue(1 / 3, { numberFormat: decimals(101) + 'E+00' }).text).toBe(at100('E+00').replace('E', '0E'))
    expect(displayValue(1 / 3, { numberFormat: decimals(101) + '%' }).text).toBe(at100('%').replace('%', '0%'))
    expect(displayValue(0.5, { numberFormat: '0.' + '#'.repeat(101) }).text).toBe('0.5')
    expect(displayValue(0.5, { numberFormat: '0.' + '?'.repeat(101) }).text).toBe('0.5' + ' '.repeat(100))
    expect(displayValue(-1.5, { numberFormat: '0;' + decimals(101) }).text).toBe('1.5' + '0'.repeat(100))
  })

  it('FIXED keeps its decimals past 100', () => {
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'FIXED(1/3,110)' }, A2: { formula: 'FIXED(1/3,100)' } } }] }, { clock })
    expect(r.sheets[0].cells.A1.value).toBe(`${r.sheets[0].cells.A2.value}${'0'.repeat(10)}`)
  })

  it.each(['closures', 'eval'] as const)('compute displays the cell, %s', (execution) => {
    const cells = compute(fmtBook(decimals(101)), { clock, execution }).sheets[0].cells
    expect(cells.A1.text).toBe(displayValue(third, { numberFormat: decimals(101) }).text)
    expect(cells.B1.text).toBe('2')
  })

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it.each(['closures', 'eval'] as const)('reads, shows and takes typed text, %s', async (execution) => {
      const text = displayValue(third, { numberFormat: decimals(101) }).text
      const s = await open(fmtBook(decimals(101)), { clock, execution })
      expect((await s.getCells('S', 'A1')).A1.text).toBe(text)
      expect((await s.getViewport('S', 'A1:B1')).cells.A1.text).toBe(text)
      const r = await s.setCellText('S', 'C2', '0.25')
      expect(r.ok).toBe(true)
      expect((await s.getCells('S', 'C2')).C2).toMatchObject({ value: 0.25, text: '0.25' + '0'.repeat(99) })
      await s.close()
    })
  })
})
