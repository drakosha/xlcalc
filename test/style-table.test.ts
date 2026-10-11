/**
 * `internStyle`: a style's number in a style table without a session, by the equality `patchCells` finds styles by (the
 * fields in any order, fields set to undefined left out): an equal style's number and the same array, or a new array with
 * a copy of the style at the end; the number `patchCells` gives the same style.
 */
import { describe, expect, it } from 'vitest'
import { createSession, InputError, internStyle, type CellStyle } from '../src/index.js'

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

  it('wrong input is an InputError naming its place', () => {
    expect(() => internStyle(table(), { font: { size: 1000 } })).toThrow(new InputError('style.font.size', 'expected a number from 1 to 409, got number 1000'))
    expect(() => internStyle([{ colour: 'red' } as CellStyle], {})).toThrow(InputError)
    expect(() => internStyle(undefined as never, {})).toThrow(/^styles: expected an array of styles/)
    expect(() => internStyle(table(), null as never)).toThrow(/^style: /)
  })
})
