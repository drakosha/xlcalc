/**
 * A style's field or a border's side given as undefined is no field, in a cell's style as in a conditional format's, as
 * every field of input is: `{ border: { top: undefined } }` was refused in a conditional format alone. In a book, through
 * `setStyles`, `patchCells` with a style given as its fields and `setLayout` with conditional formats, in a synchronous
 * session, in the same thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import type { BookInput, CellStyle, ConditionalFormat } from '../src/contract.js'
import { compute } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const THIN = { style: 'thin' } as const
/** A style with a side undefined, and the same without it. */
const loose: CellStyle = { border: { top: undefined, bottom: THIN }, font: { bold: true, italic: undefined } }
const plain: CellStyle = { border: { bottom: THIN }, font: { bold: true } }
const rule = (style: CellStyle): ConditionalFormat => ({ type: 'expression', ranges: ['A1:A2'], formula: 'TRUE', style })
const book = (style: CellStyle): BookInput => ({ styles: [style], sheets: [{ name: 'S', cells: { A1: { value: 1, style: 0 }, A2: { value: 2 } }, conditionalFormats: [rule(style)] }] })

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  it('a book with a side undefined is the book without it', () => {
    expect(compute(book(loose), { execution }).inputHash).toBe(compute(book(plain), { execution }).inputHash)
  })

  it.each(KINDS)('in %s, setStyles, patchCells and setLayout take a side undefined as no side, in a cell style and a conditional one', async (kind) => {
    const open = await opener(kind)
    const s = await open({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 } } }] }, { execution })
    await s.setStyles([loose])
    await s.patchCells([{ sheet: 'S', cell: 'A2', set: { style: loose } }])
    await s.setLayout([{ sheet: 'S', conditionalFormats: [rule(loose)] }])
    const cells = await s.getCells('S', 'A1:A2')
    expect(cells.A2?.appearance?.style).toEqual({ border: { bottom: THIN }, font: { bold: true } })
    const styles = (await s.getBook()).styles!
    expect(styles[cells.A2!.style!]).toEqual(plain)
    await s.close()
  })
})
