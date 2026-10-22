/**
 * Formatting under structure operations, from the review of 21.10: looks reaching the sheet's edge no longer refuse an
 * insert. Each result's book computes, a session's hash is a full computation's, and undo and redo give the books
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
