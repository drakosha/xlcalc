/**
 * Formatting under structure operations, from the review of 21.10: looks reaching the sheet's edge no longer refuse an
 * insert, inserted lines take the look of the line before them as rectangles, copies and moves carry looks as
 * rectangles. Each result's book computes, a session's hash is a full computation's, and undo and redo give the books
 * back.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, editStructure, type AsyncSession, type BookInput, type Execution, type LookRange, type SessionOptions, type StructureOp } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']
const styles = [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }]

afterAll(closeCalcs)

/** The operation's book, which computes. */
function edited(book: BookInput, op: StructureOp, execution: Execution): BookInput {
  const r = editStructure(book, op)
  if (!r.ok) throw new Error(`refused: ${r.message}`)
  compute(r.book, { execution, clock })
  return r.book
}

/** What a book holds of a sheet: its cells and looks. */
const held = (book: BookInput) => book.sheets.map((s) => ({ cells: s.cells, looks: s.looks }))

/**
 * The operation in a session: its book computes and hashes as the session, undo gives the book before and redo the
 * book after. The book after.
 */
async function inSession(open: (book: BookInput, opts: SessionOptions) => Promise<AsyncSession>, book: BookInput, op: StructureOp, execution: Execution): Promise<BookInput> {
  const s = await open(book, { execution, clock })
  const before = held(await s.getBook())
  const r = await s.editStructure(op)
  expect(r.ok).toBe(true)
  const after = await s.getBook()
  expect(compute(after, { execution, clock }).hash).toBe(await s.getHash())
  await s.undo()
  expect(held(await s.getBook())).toEqual(before)
  await s.redo()
  expect(held(await s.getBook())).toEqual(held(after))
  expect(compute(await s.getBook(), { execution, clock }).hash).toBe(await s.getHash())
  await s.close()
  return after
}

const sheet = (cells: BookInput['sheets'][number]['cells'], looks?: LookRange[]): BookInput => ({ styles, sheets: [{ name: 'S', cells, ...(looks ? { looks } : {}) }] })

/** The wide shape of docs/grid.md: rows formatted out to XFD. */
const wide = (): BookInput => ({
  styles,
  sheets: [{ name: 'Plan', cells: { A1: { value: 'Region', style: 0 }, B2: { value: 0.25 } }, looks: [{ range: 'A2:XFD4', numberFormat: '#,##0' }, { range: 'B2:B4', style: 1, numberFormat: '0.0%' }, { range: 'XFD4' }] }],
})

describe.each(EXECUTIONS)('looks at the sheet\'s edge (x-fix-structure-looks-4), %s', (execution) => {
  it('an entire column\'s look stays over the column under inserted rows, an entire row\'s under inserted columns', () => {
    expect(edited(sheet({ A1: { value: 1 } }, [{ range: 'D:D', style: 0 }]), { op: 'insertRows', sheet: 'S', row: 2 }, execution).sheets[0].looks).toEqual([{ range: 'D:D', style: 0 }])
    expect(edited(sheet({ A1: { value: 1 } }, [{ range: '5:5', style: 0 }]), { op: 'insertColumns', sheet: 'S', column: 'B' }, execution).sheets[0].looks).toEqual([{ range: '5:5', style: 0 }])
  })

  it('the wide shape of docs/grid.md takes an inserted column: the looks pushed past XFD are dropped', () => {
    const out = edited(wide(), { op: 'insertColumns', sheet: 'Plan', column: 'B' }, execution)
    expect(out.sheets[0].looks).toEqual([
      { range: 'B1', style: 0 },
      { range: 'A2:B4', numberFormat: '#,##0' },
      { range: 'D2:XFD4', numberFormat: '#,##0' },
      { range: 'C3:C4', style: 1, numberFormat: '0.0%' },
    ])
    expect(out.sheets[0].cells).toEqual({ A1: { value: 'Region', style: 0 }, C2: { value: 0.25 } })
  })

  it('a cell with content, meta, a note or a link at the edge still refuses', () => {
    for (const cell of [{ value: 1 }, { meta: 1 }, { note: { text: 'n' } }, { link: { url: 'https://example.com' } }])
      expect(editStructure(sheet({ XFD3: cell }, [{ range: 'A1:XFD9', style: 0 }]), { op: 'insertColumns', sheet: 'S', column: 'B' })).toMatchObject({ ok: false, code: 'off_sheet' })
  })

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('inserts columns into the wide shape and rows under an entire column\'s look, with undo and redo', async () => {
      await inSession(open, wide(), { op: 'insertColumns', sheet: 'Plan', column: 'B', count: 3 }, execution)
      const after = await inSession(open, sheet({ A1: { value: 1 } }, [{ range: 'D:D', style: 0 }]), { op: 'insertRows', sheet: 'S', row: 2, count: 4 }, execution)
      expect(after.sheets[0].looks).toEqual([{ range: 'D:D', style: 0 }])
    })
  })
})

/** Thirty cells of a header row, A1:AD1, each with a value and a style. */
function header(): BookInput {
  const cells: BookInput['sheets'][number]['cells'] = {}
  for (let c = 0; c < 30; c++) cells[`${c < 26 ? '' : 'A'}${String.fromCharCode(65 + (c % 26))}1`] = { value: c, style: 0 }
  return sheet(cells)
}

describe.each(EXECUTIONS)('inserted lines take the look of the line before them as rectangles (fix-u298), %s', (execution) => {
  it('a million rows under a styled header: one look, in a fraction of a second', () => {
    const t0 = performance.now()
    const out = edited(header(), { op: 'insertRows', sheet: 'S', row: 2, count: 1_000_000 }, execution)
    expect(performance.now() - t0).toBeLessThan(2000)
    expect(out.sheets[0].looks).toEqual([{ range: 'A2:AD1000001', style: 0 }])
    expect(Object.keys(out.sheets[0].cells)).toHaveLength(30)
  })

  it('100 columns beside a formatted column of 34,000 cells: one look', () => {
    const cells: BookInput['sheets'][number]['cells'] = {}
    for (let r = 1; r <= 34_000; r++) cells[`A${r}`] = { value: r, numberFormat: '#,##0.00' }
    const t0 = performance.now()
    const out = edited(sheet(cells), { op: 'insertColumns', sheet: 'S', column: 'B', count: 100 }, execution)
    expect(performance.now() - t0).toBeLessThan(2000)
    expect(out.sheets[0].looks).toEqual([{ range: 'B1:CW34000', numberFormat: '#,##0.00' }])
  })

  it('a table\'s calculated column in the new rows keeps the look of the row before', () => {
    const book: BookInput = {
      styles,
      sheets: [{
        name: 'S',
        cells: { A1: { value: 'k' }, B1: { value: 'v' }, A2: { value: 1, style: 1 }, B2: { formula: 'A2*2', style: 0, numberFormat: '0.0' }, A3: { value: 2 }, B3: { formula: 'A3*2' } },
        tables: [{ name: 'T', range: 'A1:B3', columns: [{ name: 'k' }, { name: 'v', formula: '[@k]*2' }] }],
      }],
    }
    const out = edited(book, { op: 'insertRows', sheet: 'S', row: 3, count: 2 }, execution)
    expect(out.sheets[0].cells.B3).toEqual({ formula: '[@k]*2', style: 0, numberFormat: '0.0' })
    expect(out.sheets[0].cells.B4).toEqual({ formula: '[@k]*2', style: 0, numberFormat: '0.0' })
    expect(out.sheets[0].looks).toEqual([{ range: 'A3:A4', style: 1 }])
  })

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('inserts a million rows under a styled header and 100 columns, with undo and redo', async () => {
      const after = await inSession(open, header(), { op: 'insertRows', sheet: 'S', row: 2, count: 1_000_000 }, execution)
      expect(after.sheets[0].looks).toEqual([{ range: 'A2:AD1000001', style: 0 }])
      await inSession(open, header(), { op: 'insertColumns', sheet: 'S', column: 'B', count: 100 }, execution)
    })
  })
})

describe.each(EXECUTIONS)('a copy, a fill and a move lay looks as rectangles (fix-u299, copy-looks), %s', (execution) => {
  const formatted = () => sheet({}, [{ range: 'A1:AD1', style: 0 }])
  const columns = () => sheet({ A1: { value: 1 } }, [{ range: 'A:H', style: 0 }])

  it('a row of looks copied onto A2:AD1048576 and filled down a million rows: one look, in a fraction of a second', () => {
    const t0 = performance.now()
    const copy = edited(formatted(), { op: 'copyRange', from: { sheet: 'S', range: 'A1:AD1' }, to: { sheet: 'S', range: 'A2:AD1048576' } }, execution)
    expect(copy.sheets[0].looks).toEqual([{ range: 'A:AD', style: 0 }])
    const fill = edited(formatted(), { op: 'fillRange', sheet: 'S', range: 'A1:AD1', to: 'A1:AD1000000' }, execution)
    expect(fill.sheets[0].looks).toEqual([{ range: 'A1:AD1000000', style: 0 }])
    expect(performance.now() - t0).toBeLessThan(3000)
  })

  it('entire formatted columns copied and moved', () => {
    expect(edited(columns(), { op: 'copyRange', from: { sheet: 'S', range: 'A:B' }, to: { sheet: 'S', range: 'J:K' } }, execution).sheets[0]).toEqual({
      name: 'S', cells: { A1: { value: 1 }, J1: { value: 1 } }, looks: [{ range: 'B1:H1', style: 0 }, { range: 'K1', style: 0 }, { range: 'A2:H1048576', style: 0 }, { range: 'J2:K1048576', style: 0 }],
    })
    expect(edited(columns(), { op: 'moveRange', from: { sheet: 'S', range: 'A:B' }, to: { sheet: 'S', cell: 'J1' } }, execution).sheets[0]).toEqual({
      name: 'S', cells: { J1: { value: 1 } }, looks: [{ range: 'C1:H1', style: 0 }, { range: 'K1', style: 0 }, { range: 'C2:H1048576', style: 0 }, { range: 'J2:K1048576', style: 0 }],
    })
  })

  it('a source over rows a filter hid copies the looks of its visible rows, a spilled value over a look taking it', () => {
    const book: BookInput = {
      styles,
      sheets: [{
        name: 'S',
        cells: { A1: { value: 'k' }, A2: { value: 1 }, A3: { value: 2 }, A4: { value: 3 }, C1: { formula: 'SEQUENCE(4)' } },
        rows: { 3: { hidden: true } },
        looks: [{ range: 'A2:A4', style: 1 }, { range: 'B2', numberFormat: '0.0' }, { range: 'B4', style: 0 }, { range: 'C2:C4', numberFormat: '0.00' }],
        filtered: true,
      }],
    }
    const out = edited(book, { op: 'copyRange', from: { sheet: 'S', range: 'A2:C4' }, to: { sheet: 'S', range: 'E2' } }, execution)
    expect(out.sheets[0].cells).toMatchObject({ E2: { value: 1 }, E3: { value: 3 }, G2: { value: 2, numberFormat: '0.00' }, G3: { value: 4, numberFormat: '0.00' } })
    expect(out.sheets[0].looks).toContainEqual({ range: 'F2', numberFormat: '0.0' })
    expect(out.sheets[0].looks).toContainEqual({ range: 'F3', style: 0 })
  })

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('copies, fills and moves under limits.writes, with undo and redo', async () => {
      const limited = (book: BookInput, opts: SessionOptions) => open(book, { ...opts, limits: { writes: 100_000 } })
      const copied = await inSession(limited, formatted(), { op: 'copyRange', from: { sheet: 'S', range: 'A1:AD1' }, to: { sheet: 'S', range: 'A2:AD1048576' } }, execution)
      expect(copied.sheets[0].looks).toEqual([{ range: 'A:AD', style: 0 }])
      await inSession(limited, formatted(), { op: 'fillRange', sheet: 'S', range: 'A1:AD1', to: 'A1:AD100000' }, execution)
      await inSession(limited, columns(), { op: 'copyRange', from: { sheet: 'S', range: 'A:B' }, to: { sheet: 'S', range: 'J:K' } }, execution)
      await inSession(limited, columns(), { op: 'moveRange', from: { sheet: 'S', range: 'A:B' }, to: { sheet: 'S', cell: 'J1' } }, execution)
    })
  })
})

describe.each(EXECUTIONS)('formatting whole columns over large looks (patch-looks), %s', (execution) => {
  /** A merged area of entire columns: its looks cover A2:H1048576, two cells of them with a format of their own. */
  const merged = (): BookInput => ({ styles, sheets: [{ name: 'S', cells: { A1: { value: 'x y', style: 1 } }, looks: [{ range: 'B1:H1', style: 1 }, { range: 'A2:H1048576', style: 1 }, { range: 'B3:B4', style: 1, numberFormat: '0.0' }], merges: ['A1:H1048576'] }] })

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('a format over the columns changes the cells of the looks it changes, in a fraction of a second, with undo and redo', async () => {
      const s = await open(merged(), { execution, clock })
      const before = held(await s.getBook())
      const t0 = performance.now()
      const r = await s.patchCells([{ sheet: 'S', range: 'A:H', set: { numberFormat: '@' } }])
      expect(performance.now() - t0).toBeLessThan(3000)
      expect(r.changes.map((c) => [c.cell, c.numberFormat])).toEqual([['B3', '@'], ['B4', '@']])
      const after = await s.getBook()
      expect(after.sheets[0].columns).toEqual({ 'A:H': { numberFormat: '@' } })
      expect(compute(after, { execution, clock }).hash).toBe(await s.getHash())
      await s.undo()
      expect(held(await s.getBook())).toEqual(before)
      await s.redo()
      expect(held(await s.getBook())).toEqual(held(after))
      await s.close()
    })
  })
})
