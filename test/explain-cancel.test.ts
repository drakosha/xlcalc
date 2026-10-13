/** Explain calls use the same cancellation and deadlines as other async calls, with both engines. */
import { describe, expect, it } from 'vitest'
import { CalcError, createInlineCalc, createSession, type BookInput } from '../src/index.js'
import { pool } from './fixtures/pool.js'

const count = 12_000
const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { ...Object.fromEntries(Array.from({ length: count }, (_, i) => [`A${i + 1}`, { value: i + 1 }])), B1: { formula: `SUM(A1:A${count})` } } }] })
const expanded = { expand: true, depth: 3, maxNodes: count + 2 }

describe.each(['closures', 'eval'] as const)('explain, %s', (execution) => {
  for (const kind of ['inline', 'pool']) {
    const open = async () => {
      const calc = kind === 'inline' ? createInlineCalc({ sliceMs: 0 }) : await pool({ threads: 1, sliceMs: 0 })
      const s = await calc.createSession(book(), { execution })
      return { calc, s }
    }
    it(`${kind}: an already aborted signal refuses without changing the session`, async () => {
      const { calc, s } = await open()
      try {
        await expect(s.explain('S', 'B1', { signal: AbortSignal.abort('early') } as never)).rejects.toMatchObject({ name: 'CalcError', code: 'aborted', cause: 'early' })
        expect([s.revision, s.history]).toEqual([0, { undo: 0, redo: 0 }])
      } finally { await calc.close() }
    })
    it(`${kind}: a queued explanation aborts while another explanation runs`, async () => {
      const { calc, s } = await open()
      try {
        const first = s.explain('S', 'B1', expanded)
        const ac = new AbortController()
        const waiting = s.explain('S', 'A1', { signal: ac.signal } as never)
        const rejected = expect(waiting).rejects.toMatchObject({ code: 'aborted', cause: 'waiting' })
        ac.abort('waiting')
        await rejected
        expect((await first).nodes).toBe(count + 2)
        expect(s.revision).toBe(0)
      } finally { await calc.close() }
    })
    it(`${kind}: an expanded running tree yields so cancellation reaches it`, async () => {
      const { calc, s } = await open()
      try {
        const hash = await s.getHash()
        const ac = new AbortController()
        const p = s.explain('S', 'B1', { ...expanded, signal: ac.signal } as never)
        const rejected = expect(p).rejects.toBeInstanceOf(CalcError)
        setImmediate(() => ac.abort('running'))
        await rejected
        await expect(p).rejects.toMatchObject({ code: 'aborted', cause: 'running' })
        expect(await s.getHash()).toBe(hash)
        expect([s.revision, s.history]).toEqual([0, { undo: 0, redo: 0 }])
        expect((await s.explain('S', 'B1')).root.formula).toBe(`SUM(A1:A${count})`)
      } finally { await calc.close() }
    })
    it(`${kind}: timeoutMs stops the running traversal and invalid call options are refused`, async () => {
      const { calc, s } = await open()
      try {
        await expect(s.explain('S', 'B1', { ...expanded, timeoutMs: 0 } as never)).rejects.toMatchObject({ code: 'soft_timeout' })
        await expect(s.explain('S', 'B1', { signal: {} } as never)).rejects.toMatchObject({ name: 'InputError', path: 'opts.signal' })
        await expect(s.explain('S', 'B1', { timeoutMs: -1 } as never)).rejects.toMatchObject({ name: 'InputError', path: 'opts.timeoutMs' })
      } finally { await calc.close() }
    })
  }
  it('structured nodes identify their local sheet and range without parsing rendered text', () => {
    const s = createSession(book(), { execution })
    try {
      const root = s.explain('S', 'B1').root
      expect(root.location).toEqual({ sheet: 'S', range: 'B1' })
      expect(root.children[0].location).toEqual({ sheet: 'S', range: `A1:A${count}` })
    } finally { s.close() }
  })
})
