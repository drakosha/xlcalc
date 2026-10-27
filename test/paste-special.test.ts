/**
 * Paste Special through `copyRange`'s `paste` and `transpose` (`test/paste-excel.test.ts` checks it against Excel): its
 * input checked; in every kind of session a paste of values takes the session's values as they are after its edits,
 * a paste of formats leaves the cells' contents and their values, each paste is one step of undo, and the same
 * paste gives the same book and hash in each kind; a copy that brings its looks fits the rows of its cells to their fonts
 * in the same step and tells them (`fitted`), and a paste of values fits none; a transposed copy of a reference into the
 * source follows the cell.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { editStructure, InputError, type BookInput, type StructureOp } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** S: a value, a formula over it, a bold cell, a target with a value and a percent format. */
const book = (): BookInput => ({
  styles: [{}, { font: { bold: true } }],
  sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' }, B1: { value: 'b', style: 1 }, D2: { value: 5, numberFormat: '0%' } } }],
})

const copy = (from: string, to: string, more: Partial<Extract<StructureOp, { op: 'copyRange' }>> = {}): StructureOp => ({ op: 'copyRange', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, ...more })

describe('copyRange with paste and transpose', () => {
  it('refuses a kind it does not have and a transpose that is no boolean with an InputError at the field', () => {
    for (const [op, path] of [
      [copy('A1', 'D1', { paste: 'comments' as never }), 'op.paste'],
      [copy('A1', 'D1', { transpose: 1 as never }), 'op.transpose'],
    ] as const) {
      const e = (() => {
        try {
          editStructure(book(), op)
        } catch (x) {
          return x
        }
      })()
      expect(e).toBeInstanceOf(InputError)
      expect((e as InputError).path).toBe(path)
    }
  })

  it('a transposed copy turns a reference into the source with its cell, and its $ with it', () => {
    const r = editStructure({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: '$A$1+A1' } } }] }, copy('A1:B1', 'D1', { transpose: true }))
    expect(r.ok && r.book.sheets[0].cells).toMatchObject({ D1: { value: 1 }, D2: { formula: '$D$1+D1' } })
  })
})

describe.each(KINDS)('Paste Special in %s', (kind) => {
  it('values as the session computes them now, formats without the contents; one step of undo each; the same in each kind', async () => {
    const open = await opener(kind)
    const s = await open(book(), { clock })
    await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
    const values = await s.editStructure(copy('A1:A2', 'D1', { paste: 'values' }))
    expect(values.ok).toBe(true)
    expect(await s.getCells('S', 'D1:D2')).toMatchObject({ D1: { value: 3 }, D2: { value: 30, text: '3000%' } })
    const input = (await s.getBook()).sheets[0].cells
    expect([input.D1, input.D2]).toEqual([{ value: 3 }, { value: 30, numberFormat: '0%' }])
    const formats = await s.editStructure(copy('B1', 'D1:D2', { paste: 'formats' }))
    // the cells it changed show another look and text, with the values they had
    expect(formats.ok && formats.changes.map((c) => [c.cell, c.value, c.text])).toEqual([
      ['D1', 3, '3'],
      ['D2', 30, '30'],
    ])
    expect((await s.getBook()).sheets[0].cells.D2).toEqual({ value: 30, style: 1 })
    expect(s.history).toEqual({ undo: 3, redo: 0 })
    await s.undo()
    expect((await s.getBook()).sheets[0].cells.D2).toEqual({ value: 30, numberFormat: '0%' })
    await s.undo()
    expect((await s.getBook()).sheets[0].cells.D2).toEqual({ value: 5, numberFormat: '0%' })
    await s.redo()
    const sync = await (await opener('a synchronous session'))(book(), { clock })
    await sync.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
    await sync.editStructure(copy('A1:A2', 'D1', { paste: 'values' }))
    // undo and redo keep the cells, not the order of their keys
    expect([await s.getBook(), await s.getHash()]).toEqual([await sync.getBook(), await sync.getHash()])
    await Promise.all([s.close(), sync.close()])
  })
})

describe.each(KINDS)('rows fitted after a copy in %s', (kind) => {
  it('formats, all and a plain copy of a larger font fit the row in the same step; values fit none', async () => {
    const open = await opener(kind)
    const big: BookInput = { styles: [{}, { font: { size: 20, bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 'big', style: 1 }, C5: { value: 'x' } } }] }
    for (const more of [{ paste: 'formats' as const }, { paste: 'all' as const }, {}]) {
      const s = await open(big, { clock })
      const r = await s.editStructure(copy('A1', 'C5', more))
      expect(r.ok && r.fitted).toEqual([{ sheet: 'S', rows: '5', height: 26 }])
      expect((await s.getBook()).sheets[0].rows).toEqual({ 5: { height: 26, bestFit: true } })
      await s.undo()
      expect((await s.getBook()).sheets[0]).toMatchObject({ cells: { C5: { value: 'x' } } })
      expect((await s.getBook()).sheets[0].rows).toBeUndefined()
      await s.close()
    }
    const s = await open(big, { clock })
    const r = await s.editStructure(copy('A1', 'C5', { paste: 'values' }))
    expect(r.ok && r.fitted).toBeUndefined()
    await s.close()
  })
})
