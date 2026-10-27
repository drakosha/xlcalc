/**
 * Protection in the model: a cell's in its style (`CellStyle.protection`), a sheet's in `SheetInput.protection`. Both
 * hold only what differs from Excel's defaults, so a style `{ protection: { locked: true } }` is the style `{}` and a
 * sheet's `{ formatCells: false }` is Excel's default protection; neither changes a value or the hash, nor does the
 * sheet's enter `inputHash`. A sheet's protection is set and removed by `setLayout`, undone and redone; a cell's goes
 * with its style (`patchCells`, `clearCells` of formats). In a synchronous session, in the same thread and in a thread,
 * with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createSession, InputError, internStyle, type BookInput } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const book = (): BookInput => ({
  styles: [{ protection: { locked: false } }, { font: { bold: true }, protection: { locked: true, hidden: false } }],
  sheets: [{ name: 'S', cells: { A1: { value: 2, style: 0 }, A2: { formula: 'A1*3', style: 1 }, A3: { value: 'x' } } }],
})

describe('protection as input', () => {
  it('a style and a sheet take only booleans of their own fields', () => {
    const fails = (b: BookInput, message: string) => expect(() => compute(b)).toThrow(message)
    fails({ styles: [{ protection: { locked: 0 } as never }], sheets: [{ name: 'S', cells: {} }] }, 'book.styles[0].protection.locked: expected a boolean, got number 0')
    fails({ styles: [{ protection: { lock: false } as never }], sheets: [{ name: 'S', cells: {} }] }, 'book.styles[0].protection: unknown field "lock" (expected locked, hidden)')
    fails({ sheets: [{ name: 'S', cells: {}, protection: { formatCell: true } as never }] }, 'book.sheets[0].protection: unknown field "formatCell" (expected selectLockedCells, selectUnlockedCells, formatCells')
    fails({ sheets: [{ name: 'S', cells: {}, protection: true as never }] }, 'book.sheets[0].protection: expected a protection such as { formatCells: true } ({} protects the sheet as Excel does by default), got boolean true')
    const s = createSession({ sheets: [{ name: 'S', cells: {} }] })
    expect(() => s.setLayout([{ sheet: 'S', protection: { sort: 'yes' } as never }])).toThrow(InputError)
    expect(() => s.setLayout([{ sheet: 'S', protection: {}, frozen: null } as never])).toThrow('edits[0]: unknown field "frozen" (expected sheet, protection)')
    expect(() => s.patchCells([{ sheet: 'S', cell: 'A1', set: { style: { protection: { hidden: 1 } } as never } }])).toThrow('patches[0].set.style.protection.hidden: expected a boolean, got number 1')
    s.close()
  })

  it("Excel's defaults are no field: the same styles, the same protection and the same hash as without them", () => {
    const s = createSession(book())
    expect(s.getBook().styles).toEqual([{ protection: { locked: false } }, { font: { bold: true } }])
    const plain = { ...book(), styles: [{ protection: { locked: false } }, { font: { bold: true } }] }
    expect(compute(book()).inputHash).toBe(compute(plain).inputHash)
    // a sheet's protection, defaults or not, is no part of inputHash
    const guarded = { ...book(), sheets: [{ ...book().sheets[0], protection: { formatCells: false, sort: true } }] }
    expect(createSession(guarded).getBook().sheets[0].protection).toEqual({ sort: true })
    expect(compute(guarded).inputHash).toBe(compute(book()).inputHash)
    // internStyle finds a style by its fields as the workbook holds them
    expect(internStyle([{}, { protection: { locked: false } }], { protection: { locked: true } })).toEqual({ index: 0, styles: [{}, { protection: { locked: false } }] })
    expect(internStyle([{ protection: { hidden: false } }], {}).index).toBe(0)
    const grown = internStyle([{}], { protection: { locked: false, hidden: false } })
    expect(grown.styles[1]).toEqual({ protection: { locked: false } })
    s.close()
  })
})

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  it.each(KINDS)('in %s, setLayout protects and unprotects a sheet, undone and redone, the values and the hash as they were', async (kind) => {
    const open = await opener(kind)
    const s = await open(book(), { execution })
    const hash = await s.getHash()
    const r = await s.setLayout([{ sheet: 'S', protection: { formatCells: true, selectLockedCells: false, objects: false } }])
    expect([r.changes, r.layout]).toEqual([[], ['S']])
    expect((await s.getLayout()).sheets[0].protection).toEqual({ formatCells: true, selectLockedCells: false })
    expect(await s.getHash()).toBe(hash)
    await s.setLayout([{ sheet: 'S', protection: null }])
    expect('protection' in (await s.getBook()).sheets[0]).toBe(false)
    await s.undo()
    expect((await s.getBook()).sheets[0].protection).toEqual({ formatCells: true, selectLockedCells: false })
    await s.undo()
    expect((await s.getBook()).sheets[0].protection).toBeUndefined()
    await s.redo()
    expect((await s.getBook()).sheets[0].protection).toEqual({ formatCells: true, selectLockedCells: false })
    expect(await s.getValues('S', 'A1:A2')).toEqual([[2], [6]])
    await s.close()
  })

  it.each(KINDS)("in %s, a cell's protection goes with its style: patchCells sets and removes it, clearing formats takes it", async (kind) => {
    const open = await opener(kind)
    const s = await open(book(), { execution })
    const styleOf = async (cell: string) => {
      const [v, b] = [(await s.getCells('S', cell))[cell], await s.getBook()]
      return v?.style === undefined ? {} : b.styles![v.style]
    }
    await s.patchCells([{ sheet: 'S', range: 'A2:A3', set: { style: { protection: { locked: false, hidden: true } } } }])
    expect(await styleOf('A2')).toEqual({ font: { bold: true }, protection: { locked: false, hidden: true } })
    expect(await styleOf('A3')).toEqual({ protection: { locked: false, hidden: true } })
    // locked again is the default: the style without protection, the first of its fields in the table
    await s.patchCells([{ sheet: 'S', cell: 'A3', set: { style: { protection: { locked: true, hidden: null } } } }])
    expect(await styleOf('A3')).toEqual({})
    await s.patchCells([{ sheet: 'S', cell: 'A2', set: { style: { protection: null } } }])
    expect((await s.getCells('S', 'A2')).A2.style).toBe(1)
    await s.clearCells([{ sheet: 'S', range: 'A1' }], 'formats')
    expect(await styleOf('A1')).toEqual({})
    await s.undo()
    expect(await styleOf('A1')).toEqual({ protection: { locked: false } })
    await s.close()
  })
})
