/**
 * `internStyle`: a style's number in a style table without a session, by the equality `patchCells` finds styles by (the
 * fields in any order, fields set to undefined left out): an equal style's number and the same array, or a new array with
 * a copy of the style at the end; the number `patchCells` gives the same style. A table is read once per array: a chain of
 * calls, tables grown apart from one table and tables changed in place answer as a table read anew does, and a chain of
 * 10,000 calls over 100 styles takes under 50 ms.
 */
import { performance } from 'node:perf_hooks'
import { describe, expect, it } from 'vitest'
import { createSession, InputError, internStyle, type CellStyle } from '../src/index.js'

/** The answer of a table read anew: the first style with the same fields, else the style added at the end. */
function reference(styles: readonly CellStyle[], style: CellStyle): { index: number; styles: CellStyle[] } {
  const key = (v: unknown): string => (v === null || typeof v !== 'object' ? JSON.stringify(v) : `{${Object.keys(v).filter((k) => (v as Record<string, unknown>)[k] !== undefined).sort().map((k) => `${k}:${key((v as Record<string, unknown>)[k])}`).join(',')}}`)
  const i = styles.findIndex((s) => key(s) === key(style))
  return i >= 0 ? { index: i, styles: styles as CellStyle[] } : { index: styles.length, styles: [...styles, style] }
}

/** `n` styles that differ, their fields in the order of `seed`. */
const many = (n: number, seed = 0): CellStyle[] =>
  Array.from({ length: n }, (_, i) => {
    const fields: [keyof CellStyle, unknown][] = [['font', { bold: i % 2 === 0, size: 8 + (i % 30) }], ['fill', { color: `#${String(i).padStart(6, '0')}` }], ['alignment', { horizontal: 'center' }]]
    return Object.fromEntries((i + seed) % 2 ? fields : fields.reverse()) as CellStyle
  })

const table = (): CellStyle[] => [{ font: { bold: true } }, { fill: { color: '#C6EFCE' } }, { font: { italic: true, color: '#C00000' }, alignment: { horizontal: 'center' } }]

describe('internStyle', () => {
  it('an equal style: its number and the same array', () => {
    const styles = table()
    const r = internStyle(styles, { fill: { color: '#C6EFCE' } })
    expect(r.index).toBe(1)
    expect(r.styles).toBe(styles)
  })

  it('the order of fields does not count, nor fields set to undefined; the first of equal styles', () => {
    const styles = [...table(), { font: { bold: true } }]
    expect(internStyle(styles, { alignment: { horizontal: 'center' }, font: { color: '#C00000', italic: true } }).index).toBe(2)
    expect(internStyle(styles, { font: { bold: true, italic: undefined }, fill: undefined }).index).toBe(0)
  })

  it('a style the table lacks: a new array with a copy at the end, the table given unchanged', () => {
    const styles = table()
    const style: CellStyle = { font: { bold: true, italic: true } }
    const r = internStyle(styles, style)
    expect(r.index).toBe(3)
    expect(r.styles).toEqual([...table(), style])
    expect(styles).toEqual(table())
    style.font!.bold = false
    expect(r.styles[3]).toEqual({ font: { bold: true, italic: true } })
    // the plain style is a style too, and an empty font is not the plain style, as patchCells compares them
    expect(internStyle([], {})).toEqual({ index: 0, styles: [{}] })
    expect(internStyle([{ font: {} }], {}).index).toBe(1)
  })

  it('the number patchCells gives the same style', () => {
    const styles = [{ fill: { color: '#C6EFCE' }, font: { bold: true } }, { font: { bold: true } }]
    const session = createSession({ styles, sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] })
    session.patchCells([{ sheet: 'S', cell: 'A1', set: { style: { font: { bold: true }, fill: { color: '#C6EFCE' } } } }])
    session.patchCells([{ sheet: 'S', cell: 'B1', set: { style: { font: { italic: true } } } }])
    const book = session.getBook()
    expect(internStyle(styles, { font: { bold: true }, fill: { color: '#C6EFCE' } }).index).toBe(book.sheets[0].cells.A1.style)
    const added = internStyle(styles, { font: { italic: true } })
    expect([added.index, added.styles]).toEqual([book.sheets[0].cells.B1.style, book.styles])
  })

  it('a chain of calls, each given the last answer, answers as tables read anew', () => {
    const pool = [...many(30), ...many(30, 1)]
    let styles: CellStyle[] = table()
    for (let n = 0; n < 2000; n++) {
      const style = pool[(n * 7919) % pool.length]
      const want = reference(styles, style)
      const got = internStyle(styles, style)
      expect([got.index, got.styles === styles, got.styles]).toEqual([want.index, want.styles === styles, want.styles])
      styles = got.styles
    }
    expect(styles).toHaveLength(33)
  })

  it('tables grown apart from one table, and the table they grew from, answer as read anew', () => {
    const base = table()
    const [x, y, z] = many(3)
    const a = internStyle(base, x).styles
    const b = internStyle(base, y).styles
    const c = internStyle(internStyle(b, z).styles, x).styles
    for (const [styles, style] of [[a, y], [b, x], [a, x], [b, y], [base, x], [c, x], [c, z], [base, z], [a, z]] as const) {
      const got = internStyle(styles, style)
      expect([got.index, got.styles], JSON.stringify(style)).toEqual([reference(styles, style).index, reference(styles, style).styles])
    }
  })

  it('a table changed in place is read again: a style pushed, replaced or changed', () => {
    const [x, y] = many(2)
    const styles = internStyle(internStyle(table(), x).styles, y).styles
    const z: CellStyle = { font: { strike: true } }
    styles.push(z)
    expect(internStyle(styles, { font: { strike: true } })).toEqual({ index: 5, styles })
    // an equal style put before the one found: the first is its number
    styles[0] = { ...y }
    expect(internStyle(styles, y)).toEqual({ index: 0, styles })
    // the style found gone, and one changed in place
    styles[3] = { fill: { color: '#000000' } }
    expect(internStyle(styles, x)).toEqual(reference(styles, x))
    styles[1].fill!.color = '#111111'
    expect(internStyle(styles, { fill: { color: '#C6EFCE' } })).toEqual(reference(styles, { fill: { color: '#C6EFCE' } }))
    expect(internStyle(styles, { fill: { color: '#111111' } })).toEqual({ index: 1, styles })
  })

  it('a chain of 10,000 calls over 100 styles in under 50 ms', () => {
    const pool = many(100)
    const order = Array.from({ length: 10_000 }, (_, n) => pool[(n * 7919 + ((n * n) % 13)) % 100])
    const times = [0, 1, 2].map(() => {
      let styles: CellStyle[] = []
      const t = performance.now()
      for (const s of order) styles = internStyle(styles, s).styles
      expect(styles).toHaveLength(100)
      return performance.now() - t
    })
    expect(Math.min(...times)).toBeLessThan(50)
  })

  it('wrong input is an InputError naming its place', () => {
    expect(() => internStyle(table(), { font: { size: 1000 } })).toThrow(new InputError('style.font.size', 'expected a number from 1 to 409, got number 1000'))
    expect(() => internStyle([{ colour: 'red' } as CellStyle], {})).toThrow(InputError)
    expect(() => internStyle(undefined as never, {})).toThrow(/^styles: expected an array of styles/)
    expect(() => internStyle(table(), null as never)).toThrow(/^style: /)
  })
})
