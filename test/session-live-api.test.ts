import { afterAll, describe, expect, it } from 'vitest'
import { compute, type AsyncSession, type BookInput, type SessionOperation } from '../src/index.js'
import { pool } from './fixtures/pool.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'A1*2' }, C1: { formula: 'B1+3' }, D1: { value: 'hello' } } }] })
const check = async (s: AsyncSession) => {
  const got = await s.getResult(),
    full = compute(await s.getBook(), { clock: s.clock, display: false })
  expect(got.hash).toBe(full.hash)
}
afterAll(closeCalcs)
describe.each(['closures', 'eval'] as const)('live workbook additions, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    const session = () => open(book(), { clock, execution, display: false })
    it('applies cells, names and formatting as one revision and history step', async () => {
      const s = await session(),
        before = await s.getHash()
      try {
        const operations: SessionOperation[] = [
          { op: 'setCells', edits: [{ sheet: 'S', cell: 'B1', input: { formula: 'A1*Rate' } }] },
          { op: 'setNames', names: [{ name: 'Rate', ref: '3' }] },
          { op: 'patchCells', patches: [{ sheet: 'S', range: 'B1', set: { numberFormat: '0.00', style: { font: { bold: true } } } }] },
        ]
        const got = await s.applyChanges(operations, { baseRevision: 0 })
        expect(got.revision).toBe(1)
        expect(s.history.undo).toBe(1)
        expect((await s.getCell('S', 'C1'))!.value).toBe(9)
        expect((await s.getCell('S', 'B1'))!.numberFormat).toBe('0.00')
        await check(s)
        await s.undo()
        expect(await s.getHash()).toBe(before)
        await s.redo()
        await check(s)
        await expect(s.applyChanges([], { baseRevision: 0 })).rejects.toMatchObject({ code: 'stale_revision' })
      } finally {
        await s.close()
      }
    })
    it('previews dependent changes and errors, with bounded output and no session change', async () => {
      const s = await session(),
        before = await s.getSnapshot()
      try {
        const ops: SessionOperation[] = [{ op: 'setCells', edits: [{ sheet: 'S', cell: 'B1', input: { formula: 'A1*3' } }] }]
        const p = await s.previewChanges(ops, { clock, maxChanges: 1 })
        expect(p.revision).toBe(0)
        expect(p.totalChanges).toBe(2)
        expect(p.truncated).toBe(true)
        expect(s.history.undo).toBe(0)
        expect(await s.getBook()).toEqual(before.book)
        expect(await s.getHash()).toBe(before.result.hash)
        const all = await s.previewChanges(ops, { clock })
        expect(all.impacts.find((c) => c.cell === 'C1')).toMatchObject({ before: 7, after: 9 })
        const a = await s.applyChanges(ops, { baseRevision: all.revision, clock: all.clock })
        expect(a.changes).toEqual(all.changes)
        const errors = await s.previewChanges([{ op: 'setCells', edits: [{ sheet: 'S', cell: 'A1', input: { formula: '1/0' } }] }])
        expect(errors.impacts.find((c) => c.cell === 'C1')!.issue).toBeDefined()
        await check(s)
      } finally {
        await s.close()
      }
    })
    it('uses an explicit live preview clock for the application and preserves the old clock afterward', async () => {
      const s = await open({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'NOW()+A1' }, C1: { formula: 'RAND()' } } }] }, { live: true, execution, display: false })
      try {
        const oldClock = s.clock,
          oldHash = await s.getHash()
        const ops: SessionOperation[] = [{ op: 'setCells', edits: [{ sheet: 'S', cell: 'A1', input: { value: 2 } }] }]
        const p = await s.previewChanges(ops, { clock })
        expect(s.clock).toEqual(oldClock)
        expect(await s.getHash()).toBe(oldHash)
        const a = await s.applyChanges(ops, { clock: p.clock, baseRevision: p.revision })
        expect(a.changes).toEqual(p.changes)
        expect(s.clock).toEqual(clock)
        await check(s)
      } finally {
        await s.close()
      }
    })
    it('undoes a fast mixed cell and look batch, including leading equals', async () => {
      const s = await session(),
        before = await s.getSnapshot()
      try {
        await s.applyChanges([
          { op: 'setCells', edits: [{ sheet: 'S', cell: 'B1', input: { formula: '=A1*3' } }] },
          { op: 'patchCells', patches: [{ sheet: 'S', range: 'B1', set: { style: { font: { bold: true } } } }] },
          { op: 'setLayout', edits: [{ sheet: 'S', columns: 'A', set: { width: 12 } }] },
        ])
        expect((await s.getBook()).sheets[0].cells.B1.formula).toBe('A1*3')
        expect(s.history.undo).toBe(1)
        await s.undo()
        expect(await s.getBook()).toEqual(before.book)
        expect(await s.getHash()).toBe(before.result.hash)
        await s.redo()
        await check(s)
      } finally {
        await s.close()
      }
    })
    it('summarizes overflow explicitly instead of serializing it as null', async () => {
      const s = await open({ sheets: [{ name: 'S', cells: { A1: { value: 1e308 }, A2: { value: 1e308 } } }] }, { clock, execution })
      try {
        expect((await s.summarizeRanges([{ sheet: 'S', range: 'A1:A2' }])).sum).toEqual({ error: 'NUM' })
      } finally {
        await s.close()
      }
    })
    it('keeps previews and reads bounded over large spills and preserves range dependency references', async () => {
      const s = await open({ sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(100000,1,5,0)' }, B1: { formula: 'SUM(A:A)' } } }] }, { clock, execution, display: false })
      try {
        const p = await s.readCells([{ sheet: 'S', range: 'A:A' }], { maxCells: 3, fields: ['value'] })
        expect(p.cells).toHaveLength(3)
        expect(p.truncated).toBe(true)
        expect(p.cells.every((c) => c.value === 5)).toBe(true)
        const deps = await s.getDependents([{ sheet: 'S', range: 'A99999' }])
        expect(deps.dependents[0]!.references[0]!.range).toBe('A1:A1048576')
      } finally {
        await s.close()
      }
    })
    it('leaves the book, clock and history intact when an asynchronous batch is aborted', async () => {
      if (kind === 'a synchronous session') return
      const cells: Record<string, { value?: number; formula?: string }> = { A1: { value: 1 } }
      for (let r = 2; r <= 3000; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
      const calc = kind === 'a thread' ? await pool({ threads: 1, sliceMs: 0 }) : null
      const s = await (calc ? calc.createSession({ sheets: [{ name: 'S', cells }] }, { clock, execution, display: false }) : open({ sheets: [{ name: 'S', cells }] }, { clock, execution, display: false }))
      try {
        const hash = await s.getHash(),
          controller = new AbortController()
        const task = s.previewChanges([{ op: 'setCells', edits: [{ sheet: 'S', cell: 'A1', input: { value: 9 } }] }], { clock: { ...clock, seed: 9 }, signal: controller.signal })
        controller.abort()
        const result = await task.then(
          () => ({ code: 'resolved' }),
          (e: unknown) => e as { code: string },
        )
        expect(result.code).toBe('aborted')
        expect(s.clock).toEqual(clock)
        expect(s.revision).toBe(0)
        expect(s.history.undo).toBe(0)
        expect(await s.getHash()).toBe(hash)
        await s.applyChanges([{ op: 'setCells', edits: [{ sheet: 'S', cell: 'A1', input: { value: 2 } }] }])
        await check(s)
      } finally {
        await s.close()
        await calc?.close()
      }
    })
    it('previews linked formulas without leaking refresh markers or linked values', async () => {
      const s = await open({ sheets: [{ name: 'S', cells: { A1: { formula: "'[1]Sheet1'!A9" } } }], externalBooks: [{ target: 'Other.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1 } }] }] }, { clock, execution, display: false })
      try {
        const before = await s.getSnapshot()
        const ops: SessionOperation[] = [{ op: 'setCells', edits: [{ sheet: 'S', cell: 'B1', input: { formula: "'[1]Sheet1'!A1" } }] }]
        const p = await s.previewChanges(ops, { clock })
        expect(await s.getBook()).toEqual(before.book)
        expect(await s.getHash()).toBe(before.result.hash)
        expect(p.impacts.find((c) => c.cell === 'A1')!.after).toEqual({ error: 'REF' })
        const a = await s.applyChanges(ops, { clock: p.clock, baseRevision: p.revision })
        expect(a.changes).toEqual(p.changes)
        await check(s)
      } finally {
        await s.close()
      }
    })
    it('refuses an invalid later operation without applying the earlier one', async () => {
      const s = await session(),
        before = await s.getHash()
      try {
        await expect(
          s.applyChanges([
            { op: 'setCells', edits: [{ sheet: 'S', cell: 'A1', input: { value: 9 } }] },
            { op: 'setLayout', edits: [{ sheet: 'missing', columns: 'A', set: { width: 12 } }] },
          ]),
        ).rejects.toMatchObject({ name: 'InputError', path: 'operations[1].edits[0].sheet', message: expect.stringContaining('unknown sheet "missing"') })
        expect(s.revision).toBe(0)
        expect(s.history.undo).toBe(0)
        expect(await s.getHash()).toBe(before)
      } finally {
        await s.close()
      }
    })
    it('reads several ranges at one revision with projection, paging and overlap deduplication', async () => {
      const s = await session()
      try {
        const refs = [
          { sheet: 'S', range: 'A1:C1' },
          { sheet: 'S', range: 'B1:D1' },
        ]
        const first = await s.readCells(refs, { maxCells: 2, fields: ['value'] })
        expect(first.cells).toHaveLength(2)
        expect(first.truncated).toBe(true)
        expect(Object.keys(first.cells[0]).sort()).toEqual(['cell', 'sheet', 'value'])
        const second = await s.readCells(refs, { maxCells: 2, fields: ['value'], cursor: first.cursor })
        expect([...first.cells, ...second.cells].map((c) => c.cell)).toEqual(['A1', 'B1', 'C1', 'D1'])
        expect(second.truncated).toBe(false)
        await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
        await expect(s.readCells(refs, { maxCells: 2, fields: ['value'], cursor: first.cursor })).rejects.toMatchObject({ code: 'stale_revision' })
      } finally {
        await s.close()
      }
    })
    it('returns bounded immediate dependents with range references and performs find and summary in the session', async () => {
      const s = await session()
      try {
        const got = await s.getDependents([{ sheet: 'S', range: 'A1' }])
        expect(got.dependents).toEqual([{ sheet: 'S', cell: 'B1', references: [{ sheet: 'S', range: 'A1' }] }])
        const all = [{ sheet: 'S', range: 'A1:D1' }]
        const found = await s.findCells(all, 'HEL', { fields: ['value'], count: true })
        expect(found.total).toBe(1)
        expect(found.cells[0]).toMatchObject({ cell: 'D1', value: 'hello' })
        const stats = await s.summarizeRanges(all)
        expect(stats).toMatchObject({ revision: 0, cells: 4, numbers: 3, errors: 0, sum: 13, min: 2, max: 7 })
        // the rows a filter shows alone, as readCells takes them
        await s.setCells([{ sheet: 'S', cell: 'A3', input: { value: 100 } }])
        expect(await s.summarizeRanges([{ sheet: 'S', range: 'A1:A3' }], { rows: [3] })).toMatchObject({ cells: 1, numbers: 1, sum: 100 })
        expect(await s.summarizeRanges([{ sheet: 'S', range: 'A1:A3' }], { rows: [2] })).toMatchObject({ cells: 0, sum: 0, min: null })
        await expect(s.summarizeRanges(all, { rows: [2, 1] })).rejects.toThrow('opts.rows')
      } finally {
        await s.close()
      }
    })
  })
})
