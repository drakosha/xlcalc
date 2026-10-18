/**
 * The heap after sessions of a chain of 10,000 formulas are closed, in a fresh process (for `session-close-memory.test.ts`;
 * `--expose-gc`): ten sessions opened and closed one after another (`series`), one opened, read and closed (`read`), or
 * sessions of a calculator that stays open, ten in the same thread (`inline`) or twenty in a pool's thread of 64 MB
 * (`pool`), each with a listener and a call still waiting when it is closed. A pool's session first sets 3,000 texts of
 * its own and undoes them, so that its history holds about 3 MB: twenty that its thread kept, even stopped, take the
 * thread past 64 MB.
 * After each close the heap is read at once and then every 10 ms (two turns of the event loop and two collections before
 * each read) until it is within 3 MB of the heap before the first session and the closed session's object is collected,
 * for at most 500 ms. Prints, for each close, how long it waited, the heap above the start then, in MB, whether the object
 * was still held, and the records of sessions the pool still keeps for its threads (`PoolInside`), as JSON.
 */
const mode = process.argv[2] ?? 'series'
const { libUrl } = await import('./target.js')
const { createInlineCalc, createSession } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')
const { heap } = await import('./heap.js')

const cells: Record<string, { value?: number; formula?: string }> = { A1: { value: 1 } }
for (let r = 2; r <= 10_000; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
const book = { sheets: [{ name: 'S', cells }] }
const opts = { clock: { now: '2026-10-04T12:00:00Z', seed: 7 }, display: false }

const pool = mode === 'pool' ? await (await import('./pool.js')).inside({ threads: 1, maxHeapMb: 64 }) : undefined
const calc = mode === 'inline' ? createInlineCalc() : pool?.calc
/** The records of sessions the pool keeps: its sessions, and those placed on its threads or held there. */
const records = () => (pool ? pool.inside.recs.size + pool.inside.slots.reduce((n, x) => n + x.sessions.size + x.resident.size, 0) : 0)
/** Texts of session `i` in column C, which its history keeps once they are undone. */
const texts = (i: number) => Array.from({ length: 3000 }, (_, k) => ({ sheet: 'S', cell: `C${k + 1}`, input: { value: `${i}:${k}:${'x'.repeat(1000)}` } }))
/** A session opened and closed; a reference to it that does not keep it. */
const openClose = async (): Promise<WeakRef<object>> => {
  if (calc) {
    const s = await calc.createSession(book, opts)
    s.subscribe(() => {})
    if (pool) {
      await s.setCells(texts(closes.length))
      await s.undo()
    }
    const waiting = s.getHash()
    await s.close()
    await waiting
    return new WeakRef(s)
  }
  const s = createSession(book, opts)
  if (mode === 'read') s.getHash()
  s.close()
  return new WeakRef(s)
}

const start = await heap()
const closes: { ms: number; above: number; held: boolean; records: number }[] = []
for (let i = 0; i < (mode === 'read' ? 1 : pool ? 20 : 10); i++) {
  const ref = await openClose()
  // the engine's background compilation may hold a closed state for some milliseconds: wait for the heap, not a time
  const t = performance.now()
  let above = (await heap()) - start
  while ((above > 3 || ref.deref()) && performance.now() - t < 500) {
    await new Promise((done) => setTimeout(done, 10))
    above = (await heap()) - start
  }
  closes.push({ ms: Math.round(performance.now() - t), above: +above.toFixed(2), held: ref.deref() !== undefined, records: records() })
}
await calc?.close()
console.log(JSON.stringify(closes))
