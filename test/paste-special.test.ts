/**
 * Paste Special through `copyRange`'s `paste`, `transpose`, `operation` and `skipBlanks` (`test/paste-excel.test.ts`
 * checks it against Excel): its input checked; in every kind of session a paste of values takes the session's values as they are after its edits,
 * a paste of formats leaves the cells' contents and their values, each paste is one step of undo, and the same
 * paste gives the same book and hash in each kind; a copy that brings its looks fits the rows of its cells to their fonts
 * in the same step and tells them (`fitted`), and a paste of values fits none; a transposed copy of a reference into the
 * source follows the cell; an operation combines the session's values as they are and formulas as text, Skip Blanks
 * keeps the target's cells, the other kinds take their part, each in one step of undo, in both kinds of execution.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { editStructure, InputError, type BookInput, type Execution, type StructureOp } from '../src/index.js'
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
      [copy('A1', 'D1', { operation: 'power' as never }), 'op.operation'],
      [copy('A1', 'D1', { skipBlanks: 'yes' as never }), 'op.skipBlanks'],
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

  it('an operation or Skip Blanks without paste is Paste Special with all', () => {
    // 5 plus 2 with the source's look; the text b leaves E2 blank, with its look
    const r = editStructure(book(), copy('A1:B1', 'D2', { operation: 'add' }))
    expect(r.ok && [r.book.sheets[0].cells.D2, r.book.sheets[0].cells.E2, r.book.sheets[0].looks]).toEqual([{ value: 7 }, undefined, [{ range: 'E2', style: 1 }]])
    const skipped = editStructure(book(), copy('C1', 'D2', { skipBlanks: true }))
    expect(skipped.ok && skipped.book.sheets[0].cells.D2).toEqual({ value: 5, numberFormat: '0%' })
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

describe.each(['closures', 'eval'] as const)('operations, Skip Blanks and the other kinds, execution %s', (execution: Execution) => {
  /** S: numbers and a formula A1:A3 with a blank A4, targets D1:D4 with a formula, a note and a format; column A wide. */
  const opsBook = (): BookInput => ({
    sheets: [
      {
        name: 'S',
        cells: { A1: { value: 2, numberFormat: '0.0' }, A2: { formula: 'A1*10' }, A3: { value: 'x', note: { text: 'n' } }, D1: { value: 10 }, D2: { formula: 'D1+1' }, D3: { value: 5, numberFormat: '0%' }, D4: { value: 7 } },
        columns: { A: { width: 20 } },
      },
    ],
  })

  describe.each(KINDS)('%s', (kind) => {
    it('adds the values the session computes now, combines formulas as text, skips blanks; one step of undo each; the same book and hash in each kind', async () => {
      const open = await opener(kind)
      const s = await open(opsBook(), { clock, execution })
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
      const added = await s.editStructure(copy('A1:A4', 'D1', { paste: 'values', operation: 'add', skipBlanks: true }))
      expect(added.ok).toBe(true)
      // 10+3, the formula D1+1 plus 30, text left, a blank skipped
      expect((await s.getBook()).sheets[0].cells).toMatchObject({ D1: { value: 13 }, D2: { formula: '(D1+1)+30' }, D3: { value: 5, numberFormat: '0%' }, D4: { value: 7 } })
      expect(await s.getCells('S', 'D2')).toMatchObject({ D2: { value: 44 } })
      const formulas = await s.editStructure(copy('A2', 'D4', { paste: 'formulas', operation: 'multiply' }))
      expect(formulas.ok && (await s.getBook()).sheets[0].cells.D4).toEqual({ formula: '7*(D3*10)' })
      await s.editStructure(copy('A1:A4', 'F1', { paste: 'valuesAndNumberFormats' }))
      await s.editStructure(copy('A1:A4', 'G1', { paste: 'notes' }))
      await s.editStructure(copy('A1', 'H1', { paste: 'columnWidths' }))
      const sheet = (await s.getBook()).sheets[0]
      expect([sheet.cells.F1, sheet.cells.G3, sheet.columns]).toEqual([{ value: 3, numberFormat: '0.0' }, { note: { text: 'n' } }, { A: { width: 20 }, H: { width: 20 } }])
      expect(s.history).toEqual({ undo: 6, redo: 0 })
      await s.undo()
      expect((await s.getBook()).sheets[0].columns).toEqual({ A: { width: 20 } })
      await s.undo()
      await s.undo()
      await s.undo()
      expect((await s.getBook()).sheets[0].cells.D4).toEqual({ value: 7 })
      await s.redo()
      const sync = await (await opener('a synchronous session'))(opsBook(), { clock, execution })
      await sync.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
      await sync.editStructure(copy('A1:A4', 'D1', { paste: 'values', operation: 'add', skipBlanks: true }))
      await sync.editStructure(copy('A2', 'D4', { paste: 'formulas', operation: 'multiply' }))
      expect([await s.getBook(), await s.getHash()]).toEqual([await sync.getBook(), await sync.getHash()])
      await Promise.all([s.close(), sync.close()])
    })
  })
})

describe.each(KINDS)('All Merging Conditional Formats in %s', (kind) => {
  /** S: A1:A2 under a rule, D1:D3 under one of its own. */
  const rulesBook = (): BookInput => ({
    sheets: [
      {
        name: 'S',
        cells: { A1: { value: 2 }, A2: { value: 3 }, D1: { value: 1 } },
        conditionalFormats: [
          { type: 'cellIs', ranges: ['A1:A2'], operator: 'greaterThan', formulas: ['1'], style: { font: { bold: true } } },
          { type: 'cellIs', ranges: ['D1:D3'], operator: 'lessThan', formulas: ['5'], style: { font: { italic: true } } },
        ],
      },
    ],
  })
  const ranges = async (s: { getBook(): Promise<BookInput> | BookInput }) => (await s.getBook()).sheets[0].conditionalFormats?.map((r) => r.ranges)

  it('keeps the rules of the cells it lands on and grows a rule it pastes beside its own cells; one step of undo; the same book and hash as a synchronous session', async () => {
    const s = await (await opener(kind))(rulesBook(), { clock })
    expect((await s.editStructure(copy('A1:A2', 'D1', { paste: 'allMergingConditionalFormats' }))).ok).toBe(true)
    expect(await ranges(s)).toEqual([['D1:D2'], ['A1:A2'], ['D1:D3']])
    await s.editStructure(copy('A1:A2', 'A3', { paste: 'allMergingConditionalFormats' }))
    expect(await ranges(s)).toEqual([['D1:D2'], ['A1:A4'], ['D1:D3']])
    // with Skip Blanks Excel pastes as all
    await s.editStructure(copy('A1:A2', 'D2', { paste: 'allMergingConditionalFormats', skipBlanks: true }))
    expect(await ranges(s)).toEqual([['D2:D3'], ['D1'], ['A1:A4'], ['D1']])
    expect(s.history).toEqual({ undo: 3, redo: 0 })
    await s.undo()
    await s.undo()
    expect(await ranges(s)).toEqual([['D1:D2'], ['A1:A2'], ['D1:D3']])
    const sync = await (await opener('a synchronous session'))(rulesBook(), { clock })
    await sync.editStructure(copy('A1:A2', 'D1', { paste: 'allMergingConditionalFormats' }))
    expect([await s.getBook(), await s.getHash()]).toEqual([await sync.getBook(), await sync.getHash()])
    await Promise.all([s.close(), sync.close()])
  })
})

describe.each(KINDS)('an operation onto a spilled array in %s', (kind) => {
  it('is refused with array, the book as it was; a paste of values without one is not', async () => {
    const s = await (await opener(kind))({ sheets: [{ name: 'S', cells: { A1: { value: 5 }, D1: { formula: 'SEQUENCE(3)' } } }] }, { clock })
    const r = await s.editStructure(copy('A1', 'D2', { paste: 'values', operation: 'add' }))
    expect(r.ok ? null : r.code).toBe('array')
    expect([(await s.getBook()).sheets[0].cells.D1, s.history]).toEqual([{ formula: 'SEQUENCE(3)' }, { undo: 0, redo: 0 }])
    expect((await s.editStructure(copy('A1', 'D2', { paste: 'values' }))).ok).toBe(true)
    await s.close()
  })
})
