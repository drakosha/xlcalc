/**
 * The thread pool: the same values and hashes as the synchronous engine on the case books in both execution modes;
 * sessions whose book lives in a thread (edits go there, changes come back), two of them on one thread with their
 * calls mixed; `mapCell` of a structure operation as the synchronous session's; soft deadlines and aborts, waiting
 * and running.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellInput, type StructureOp } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { afterStopped, BOOKS, chainBook, CLOCK, fails, heavyBook, sameOverEdits, setA1, slowBook } from './async-books.js'
import { closedAfter } from './fixtures/pool.js'
import { randomBook } from './seeded-books.js'
import { edits } from './session-edits.js'
import './locales.js'
import { cellIssues } from './diagnostics.js'

const one = (formula: string): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { formula } } }] })

const { made } = closedAfter()

describe.each(['closures', 'eval'] as const)('in threads, execution %s', (execution) => {
  it('the case books: values and hashes as the synchronous engine', async () => {
    const p = await made({ threads: 2 })
    const results = await Promise.all(BOOKS.map(([, book]) => p.compute(book, { clock: CLOCK, execution })))
    for (const [i, [name, book]] of BOOKS.entries()) {
      const sync = compute(book, { clock: CLOCK, execution })
      expect(results[i].hash, name).toBe(sync.hash)
      expect(results[i].sheets, name).toEqual(sync.sheets)
      expect(cellIssues(results[i]), name).toEqual(cellIssues(sync))
    }
  }, 120_000)

  it('a session over 40 edits: the same changes and hash after each as a synchronous session', async () => {
    const synth = synthBook({ seed: 1, rows: 340 })
    const p = await made({ threads: 1 })
    const s = await p.createSession(synth.book, { clock: CLOCK, execution })
    const sync = await sameOverEdits(s, synth, execution)
    const report = synth.book.sheets.at(-1)!
    const cell = Object.keys(report.cells).find((k) => report.cells[k].formula)!
    expect(await s.explain(report.name, cell)).toEqual(sync.explain(report.name, cell))
    expect(await s.getBook()).toEqual(sync.getBook())
    await s.close()
  }, 120_000)
})

describe('in threads', () => {
  it('two sessions on one thread, their calls mixed without waiting: each has its own changes and hash', async () => {
    const p = await made({ threads: 1 })
    const [a, b] = [chainBook(400), synthBook({ seed: 2, rows: 200 }).book]
    const [sa, sb] = await Promise.all([p.createSession(a, { clock: CLOCK }), p.createSession(b, { clock: CLOCK })])
    const [ya, yb] = [createSession(a, { clock: CLOCK }), createSession(b, { clock: CLOCK })]
    const eb = edits(9, 8, synthBook({ seed: 2, rows: 200 }))
    const pending: Promise<unknown>[] = []
    const expected: unknown[] = []
    for (let i = 0; i < 8; i++) {
      pending.push(sa.setCells(setA1(i + 2)), sb.setCells(eb[i]), sa.getResult().then((r) => r.hash))
      expected.push(ya.setCells(setA1(i + 2)), yb.setCells(eb[i]), ya.getResult().hash)
    }
    expect(await Promise.all(pending)).toEqual(expected)
    expect([(await sa.getResult()).hash, (await sb.getResult()).hash]).toEqual([ya.getResult().hash, yb.getResult().hash])
  }, 60_000)

  it('names, the clock, layout, styles and a live session as in the same thread', async () => {
    const p = await made({ threads: 1 })
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'TODAY()' }, A2: { formula: 'A1+1' }, A3: { value: 5 }, B1: { formula: 'x*2' }, B2: { formula: 'SUBTOTAL(109,A3:A4)' }, A4: { value: 7 } } }], names: [{ name: 'x', ref: "'S'!$A$3" }] }
    const s = await p.createSession(book, { clock: CLOCK })
    const y = createSession(book, { clock: CLOCK })
    const next = { now: '2026-10-05T00:00:00Z', seed: 1 }
    expect(await s.setClock(next)).toEqual(y.setClock(next))
    expect(s.clock).toEqual(next)
    expect(await s.setNames([{ name: 'x', ref: "'S'!$A$2" }])).toEqual(y.setNames([{ name: 'x', ref: "'S'!$A$2" }]))
    expect(await s.setLayout([{ sheet: 'S', rows: '4', set: { hidden: true } }])).toEqual(y.setLayout([{ sheet: 'S', rows: '4', set: { hidden: true } }]))
    await s.setStyles([{ font: { bold: true } }])
    y.setStyles([{ font: { bold: true } }])
    expect((await s.getResult()).hash).toBe(y.getResult().hash)
    expect(await s.getBook()).toEqual(y.getBook())
    const live = await p.createSession(chainBook(5), { live: true })
    await live.setCells(setA1(3))
    expect((await live.getResult()).hash).toBe(compute(await live.getBook(), { clock: live.clock }).hash)
  })

  it('editStructure: the changes and mapCell of the synchronous session, the book from book()', async () => {
    const p = await made({ threads: 1 })
    const book = randomBook(5)
    book.sheets[1] = { ...book.sheets[1], rows: { '3': { hidden: true }, '7:9': { hidden: true, height: 20 } }, filtered: true }
    const ops: StructureOp[] = [
      { op: 'insertRows', sheet: 'S1', row: 4, count: 2 },
      { op: 'deleteRows', sheet: 'S2', row: 2, count: 8 },
      { op: 'insertRows', sheet: 'S2', row: 6, count: 3 },
      { op: 'deleteColumns', sheet: 'S 3', column: 'B', count: 2 },
      { op: 'moveRange', from: { sheet: 'S1', range: 'A1:C5' }, to: { sheet: 'S 3', cell: 'D10' } },
      { op: 'copyRange', from: { sheet: 'S2', range: 'A1:B10' }, to: { sheet: 'S1', range: 'F1:G10' } },
      { op: 'renameSheet', sheet: 's1', name: 'One' },
      { op: 'insertSheet', name: 'New', before: 'S2' },
      { op: 'moveSheet', sheet: 'S2', before: 'One' },
      { op: 'deleteSheet', sheet: 'New' },
      { op: 'insertRows', sheet: 'S 3', row: 1048576, count: 1 },
    ]
    const s = await p.createSession(book, { clock: CLOCK })
    const y = createSession(book, { clock: CLOCK })
    for (const op of ops) {
      const names = y.getBook().sheets.map((x) => x.name)
      const [a, b] = [await s.editStructure(op), y.editStructure(op)]
      expect(a.ok, JSON.stringify(op)).toBe(b.ok)
      if (!a.ok || !b.ok) continue
      expect(a.changes).toEqual(b.changes)
      for (const sheet of names) for (const cell of ['A1', 'B2', 'C5', 'D9', 'A12', 'F3', 'G10', 'H20']) expect(a.mapCell({ sheet, cell }), `${JSON.stringify(op)} ${sheet}!${cell}`).toEqual(b.mapCell({ sheet, cell }))
    }
    expect(await s.getBook()).toEqual(y.getBook())
  }, 60_000)

  it('editStructure over tables: mapCell and changes of the synchronous session, refusals as its', async () => {
    const p = await made({ threads: 1 })
    const head = { A1: { value: 'Region' }, B1: { value: 'Qty' }, C1: { value: 'Price' } }
    const rows = Object.fromEntries([2, 3, 4, 5].flatMap((r) => [[`A${r}`, { value: `r${r}` }], [`B${r}`, { value: r }], [`C${r}`, { formula: 'Sales[[#This Row],[Qty]]*2' }]]))
    const book: BookInput = {
      sheets: [
        { name: 'S1', cells: { ...head, ...rows, F3: { value: 'a' }, F4: { value: 1 }, J1: { formula: 'SUM(Sales[Qty])' }, J2: { formula: 'Sales[@Price]' } }, tables: [{ name: 'Sales', range: 'A1:C5', columns: [{ name: 'Region' }, { name: 'Qty' }, { name: 'Price', formula: 'Sales[[#This Row],[Qty]]*2' }] }, { name: 'Side', range: 'F3:F4', columns: [{ name: 'a' }] }] },
        { name: 'S2', cells: { A1: { formula: 'SUM(Sales[Price])' }, B1: { value: 'x' }, B2: { value: 'y' } } },
      ],
    }
    const ops: StructureOp[] = [
      { op: 'insertRows', sheet: 'S1', row: 2 },
      { op: 'insertRows', sheet: 'S1', row: 1 },
      { op: 'insertRows', sheet: 'S1', row: 6 },
      { op: 'deleteRows', sheet: 'S1', row: 3, count: 5 },
      { op: 'insertColumns', sheet: 'S1', column: 'B' },
      { op: 'addTable', sheet: 'S2', table: { name: 'Two', range: 'B1:B2', columns: [{ name: 'x' }] } },
      { op: 'renameTable', table: 'two', name: 'Deux' },
      { op: 'copyRange', from: { sheet: 'S2', range: 'A1' }, to: { sheet: 'S2', range: 'B1' } },
      { op: 'moveRange', from: { sheet: 'S1', range: 'J1:J3' }, to: { sheet: 'S2', cell: 'D4' } },
      { op: 'convertTable', table: 'Sales' },
      { op: 'deleteTable', table: 'Deux' },
    ]
    const s = await p.createSession(book, { clock: CLOCK })
    const y = createSession(book, { clock: CLOCK })
    for (const op of ops) {
      const names = y.getBook().sheets.map((x) => x.name)
      const [a, b] = [await s.editStructure(op), y.editStructure(op)]
      expect(a.ok, JSON.stringify(op)).toBe(b.ok)
      if (!a.ok || !b.ok) continue
      expect(a.changes).toEqual(b.changes)
      for (const sheet of names) for (const cell of ['A1', 'A2', 'B1', 'B2', 'C3', 'D4', 'F2', 'J1', 'J2']) expect(a.mapCell({ sheet, cell }), `${JSON.stringify(op)} ${sheet}!${cell}`).toEqual(b.mapCell({ sheet, cell }))
    }
    expect(await s.getBook()).toEqual(y.getBook())
  }, 60_000)

  it('soft deadlines and aborts, waiting and running; input errors come back as InputError', async () => {
    // slices of 5 ms: the thread reads a cancel within one, however fast it computes
    const p = await made({ threads: 1, sliceMs: 5 })
    const big = synthBook({ seed: 1, rows: 1700 }).book
    expect((await fails(p.compute(big, { clock: CLOCK, timeoutMs: 5 }))).code).toBe('soft_timeout')
    // one running, one waiting: both aborted
    const [run, wait] = [new AbortController(), new AbortController()]
    const running = p.compute(big, { signal: run.signal })
    const waiting = p.compute(big, { signal: wait.signal })
    wait.abort('not needed')
    const w = await fails(waiting)
    expect([w.code, w.cause]).toEqual(['aborted', 'not needed'])
    setImmediate(() => run.abort('stop'))
    const r = await fails(running)
    expect([r.code, r.cause]).toEqual(['aborted', 'stop'])
    // the thread goes on
    expect((await p.compute(chainBook(3), { clock: CLOCK })).hash).toBe(compute(chainBook(3), { clock: CLOCK }).hash)
    const s = await p.createSession(chainBook(3))
    await expect(s.setCells([{ sheet: 'Nope', cell: 'A1', input: { value: 1 } }])).rejects.toThrow(/^edits\[0\]\.sheet: unknown sheet/)
    await expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 1 } }], { deadline: 1 } as never)).rejects.toThrow('opts: unknown field "deadline"')
    await expect(p.compute(chainBook(3), { clok: 1 } as never)).rejects.toThrow('opts: unknown field "clok"')
  }, 60_000)

  it('a session call past its soft deadline in a thread changes nothing', async () => {
    const p = await made({ threads: 1 })
    const s = await p.createSession(slowBook(), { clock: CLOCK })
    const before = (await s.getResult()).hash
    expect((await fails(s.setCells(setA1(2), { timeoutMs: 50 }))).code).toBe('soft_timeout')
    expect((await s.getBook()).sheets[0].cells.A1).toEqual({ value: 1 })
    expect((await s.getResult()).hash).toBe(before)
    const y = createSession(slowBook(), { clock: CLOCK })
    expect(await s.setCells(setA1(0.5))).toEqual(y.setCells(setA1(0.5)))
  }, 60_000)

  it('the call after a stopped one: computing the book again in the thread does not count in its soft deadline', async () => {
    const p = await made({ threads: 1 })
    const s = await p.createSession(heavyBook(), { clock: CLOCK })
    await afterStopped(s)
  }, 60_000)

  it('maxHeapMb without limits: arrays of more values than the heap has 64 bytes, spills of more than it has KB give CalcError', async () => {
    const p = await made({ threads: 1, maxHeapMb: 64 })
    const array = await fails(p.compute(one('SUM(SEQUENCE(1048577))')))
    expect([array.limit, array.max]).toEqual(['cells', 1_048_576])
    const spill = await fails(p.compute(one('SEQUENCE(40000,2)')))
    expect([spill.code, spill.limit, spill.max, spill.size, spill.cell]).toEqual(['limit', 'spill', 65_536, 80_000, { sheet: 'S', cell: 'A1' }])
    for (const fits of [one('SUM(SEQUENCE(1048576))'), one('SEQUENCE(32768,2)')]) expect((await p.compute(fits, { clock: CLOCK })).hash).toBe(compute(fits, { clock: CLOCK }).hash)
  })

  it("a call's limits keep the bounds the heap gives where they do not give their own", async () => {
    const p = await made({ threads: 1, maxHeapMb: 64 })
    const array = await fails(p.compute(one('SUM(SEQUENCE(2000000))'), { limits: { formulas: 1_000 } }))
    expect([array.code, array.limit, array.max]).toEqual(['limit', 'cells', 1_048_576])
    const s = await p.createSession(one('1'), { clock: CLOCK, limits: { formulas: 1_000 } })
    const spill = await fails(s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'SEQUENCE(40000,2)' } }]))
    expect([spill.code, spill.limit, spill.max]).toEqual(['limit', 'spill', 65_536])
  }, 30_000)

  it('a pool of 384 MB computes SUM(SEQUENCE(500000)) and TOCOL of a whole column; its spills stay bounded by the heap', async () => {
    const p = await made({ threads: 1, maxHeapMb: 384 })
    const data = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 } }
    for (const formula of ['SUM(SEQUENCE(500000))', 'TOCOL(A:A)', 'ROWS(TOCOL(A:C))']) {
      const b: BookInput = { sheets: [{ name: 'S', cells: { ...data, B1: { formula } } }] }
      expect((await p.compute(b, { clock: CLOCK })).hash, formula).toBe(compute(b, { clock: CLOCK }).hash)
    }
    expect((await fails(p.compute(one('SEQUENCE(200000,2)')))).max).toBe(393_216)
    // a host that raises cells keeps the spill's bound, one that gives spill sets it
    const raised = await made({ threads: 1, maxHeapMb: 64, limits: { cells: 10_000_000 } })
    expect((await raised.compute(one('SUM(SEQUENCE(1100000))'), { clock: CLOCK })).sheets[0].cells.A1.value).toBe(605_000_550_000)
    const kept = await fails(raised.compute(one('SEQUENCE(40000,2)')))
    expect([kept.limit, kept.max]).toEqual(['spill', 65_536])
    const given = await made({ threads: 1, maxHeapMb: 64, limits: { spill: 100_000 } })
    expect((await given.compute(one('SEQUENCE(40000,2)'), { clock: CLOCK })).hash).toBe(compute(one('SEQUENCE(40000,2)'), { clock: CLOCK }).hash)
    expect((await fails(given.compute(one('SUM(SEQUENCE(1048577))')))).max).toBe(1_048_576)
  }, 30_000)

  it('a pool of 384 MB opens a running total over 10,000 rows: a sum of one range is not made whole', async () => {
    const p = await made({ threads: 1, maxHeapMb: 384 })
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 10_000; r++) { cells[`A${r}`] = { value: 1 }; cells[`B${r}`] = { formula: `SUM($A$1:A${r})` } }
    const b: BookInput = { sheets: [{ name: 'S', cells }] }
    const s = await p.createSession(b, { clock: CLOCK })
    expect(await s.getHash()).toBe(compute(b, { clock: CLOCK }).hash)
  }, 30_000)

  it("a pool of 384 MB opens and edits running AVERAGE, COUNTIF and AVERAGE with COUNTA over 10,000 rows: the ranges' caches stay within a quarter of the heap", async () => {
    const p = await made({ threads: 1, maxHeapMb: 384 })
    for (const execution of ['closures', 'eval'] as const)
      for (const formula of ['AVERAGE($A$1:A<r>)', 'COUNTIF($A$1:A<r>,">0")', 'AVERAGE($A$1:A<r>)+COUNTA($A$1:A<r>)']) {
        const cells: Record<string, CellInput> = {}
        for (let r = 1; r <= 10_000; r++) { cells[`A${r}`] = { value: 1 }; cells[`B${r}`] = { formula: formula.replaceAll('<r>', String(r)) } }
        const s = await p.createSession({ sheets: [{ name: 'S', cells }] }, { clock: CLOCK, execution })
        await s.setCells([{ sheet: 'S', cell: 'A5000', input: { value: 2 } }])
        expect(await s.getHash(), `${formula}, ${execution}`).toBe(compute(await s.getBook(), { clock: CLOCK, execution }).hash)
        await s.close()
      }
  }, 120_000)

  it('limits in a thread: CalcError with the cell', async () => {
    const p = await made({ threads: 1, limits: { cells: 1000 } })
    const e = await fails(p.compute({ sheets: [{ name: 'S', cells: { C3: { formula: 'SUM(SEQUENCE(100,100))' } } }] }))
    expect([e.code, e.limit, e.cell, e.size]).toEqual(['limit', 'cells', { sheet: 'S', cell: 'C3' }, 10_000])
  })
})
