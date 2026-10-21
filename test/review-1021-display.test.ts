/**
 * Reproductions from the review of 21.10 on display and sheets: a number format with more than 100 decimals, a fraction
 * format with many denominator digits, rows fitted for more wrapping cells than a call takes arguments, and deleting or
 * moving a sheet while a reference across sheets names a sheet the book does not have. With closures and eval, and the
 * session parts in a synchronous session, in the same thread and in a thread.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, displayValue, type BookInput } from '../src/index.js'
import { createState } from '../src/engine.js'
import { closestFraction } from '../src/numfmt/number.js'
import { typedHeights } from '../src/session/wrap.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const decimals = (n: number) => '0.' + '0'.repeat(n)
// a book keeps 15 significant digits of a value
const third = 0.333333333333333
const fmtBook = (numberFormat: string): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: third, numberFormat }, B1: { value: 2 } }, columns: { A: { width: 255 }, C: { width: 255, numberFormat } } }] })

/** The old search for the closest fraction: every denominator in turn. */
function scan(x: number, max: number): [number, number] {
  let [best, num, den] = [Infinity, 0, 0]
  for (let d = 1; d <= max; d++) {
    const n = Math.round(x * d)
    const err = Math.abs(x - n / d)
    if (err < best - 1e-12) [best, num, den] = [err, n, d]
  }
  return [num, den]
}

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

describe('a fraction format with many denominator digits', () => {
  it('takes the fraction the old scan over every denominator took, for 1 to 5 digits', () => {
    let seed = 1
    const rnd = () => ((seed = (seed * 48271) % 2147483647), seed / 2147483647)
    const values = [0, 0.5, 1 / 3, Math.PI - 3, 0.999999, 1e-7, 123.456, 2 / 7, 0.1 + 0.2, 1e15]
    for (let i = 0; i < 600; i++) values.push(rnd(), rnd() * 100, Math.floor(rnd() * 100) / (1 + Math.floor(rnd() * 99)), rnd() ** 8)
    for (const v of values) for (let k = 1; k <= 5; k++) expect(closestFraction(v, 10 ** k - 1), `${v} in ${k} digits`).toEqual(scan(v, 10 ** k - 1))
  })

  it.each(['closures', 'eval'] as const)('TEXT and a cell answer at once with 15 and 200 places, %s', (execution) => {
    // the old scan took 0.7 s at 9 places and grew tenfold with each one
    const started = performance.now()
    const r = compute(
      {
        sheets: [
          {
            name: 'S',
            cells: {
              A1: { formula: `TEXT(PI(),"# ?/${'?'.repeat(15)}")` },
              A2: { formula: `TEXT(1E+15,"?/${'?'.repeat(15)}")` },
              A3: { formula: `TEXT(1E-9,"# ?/${'?'.repeat(15)}")` },
              A4: { value: 1e-7, numberFormat: `?/${'?'.repeat(200)}` },
              A5: { value: Math.E, numberFormat: `# ?/${'?'.repeat(15)}` },
            },
            columns: { A: { width: 255 } },
          },
        ],
      },
      { clock, execution },
    )
    expect(performance.now() - started).toBeLessThan(2000)
    const c = r.sheets[0].cells
    expect(c.A1.value).toBe('3 192583/1360120' + ' '.repeat(8))
    expect(c.A2.value).toEqual({ error: 'VALUE' })
    // the old scan's answers (9 s and 0.9 s of it): from an error of 1e-12 no denominator is taken
    expect(c.A3.value).toBe(' 1/999002167' + ' '.repeat(6))
    expect(c.A4.text).toBe('1/9999994' + ' '.repeat(193))
    expect(c.A5.text?.trimEnd()).toMatch(/^2 \d+\/\d+$/)
  })
})

describe.each(['closures', 'eval'] as const)('rows fitted for many wrapping cells, %s', (execution) => {
  it('fits 200,000 rows of one call', () => {
    const rows = 200_000
    const cells: BookInput['sheets'][number]['cells'] = {}
    for (let r = 1; r <= rows; r++) cells[`A${r}`] = { value: 'a\nb' }
    const state = createState({ styles: [{ alignment: { wrap: true } }], sheets: [{ name: 'S', cells, columns: { A: { style: 0 } } }] }, { now: 0, seed: 1, execution })
    const typed = Array.from({ length: rows }, (_, r) => ({ i: 0, r, c: 0, edit: { sheet: 'S', cell: `A${r + 1}`, input: { value: 'a\nb' } } }))
    const edits = typedHeights(state, typed as Parameters<typeof typedHeights>[1])
    expect(edits.length).toBe(rows)
    const heights = new Set(edits.map((e) => ('rows' in e ? e.set.height : undefined)))
    expect(heights.size).toBe(1)
    expect([...heights][0]).toBeGreaterThan(15)
  })
})
