/**
 * Session: `setClock` recomputes volatile cells and `setNames` rebuilds, a live session takes a clock on every
 * recomputation, `explain` walks the precedents, and `debugGeneratedCode` dumps the code in evaluation order.
 * The hash of a session after each of 200 seeded edits, and its reported changes, are checked against full
 * computations by `execution-session-*.test.ts` (`execution-compare.ts`), in both execution modes.
 */
import { describe, expect, it } from 'vitest'
import { InputError, compute, debugGeneratedCode, createSession, type BookInput } from '../src/index.js'
import { apply } from './session-edits.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

describe('session', () => {
  it('setClock recomputes volatile cells; setNames rebuilds', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'TODAY()' }, A2: { formula: 'A1+1' }, A3: { value: 5 }, B1: { formula: 'x*2' } } }], names: [{ name: 'x', ref: "'S'!$A$3" }] }
    const s = createSession(book, { clock, display: false })
    expect(s.setClock({ now: '2026-10-05T00:00:00Z', seed: 1 }).changes.map((c) => c.cell).sort()).toEqual(['A1', 'A2'])
    expect(s.getResult().hash).toBe(compute(book, { clock: { now: '2026-10-05T00:00:00Z', seed: 1 } }).hash)
    expect(s.setNames([{ name: 'x', ref: "'S'!$A$2" }]).changes).toEqual([{ sheet: 'S', cell: 'B1', value: 2 * (46300 + 1) }])
  })

  it('a live session takes a clock on every recomputation; a full computation with its clock gives the same hash', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'NOW()' }, A2: { formula: 'RAND()' }, A3: { value: 1 }, A4: { formula: 'A3*2' } } }] }
    const s = createSession(book, { live: true })
    const c1 = s.clock
    expect(Math.abs(Date.parse(c1.now) - Date.now())).toBeLessThan(60_000)
    const { changes: ch } = s.setCells([{ sheet: 'S', cell: 'A3', input: { value: 2 } }])
    const c2 = s.clock
    expect(c2.seed).not.toBe(c1.seed)
    expect(ch.map((c) => c.cell).sort()).toEqual(['A1', 'A2', 'A3', 'A4'].filter((c) => c !== 'A1' || c1.now !== c2.now))
    const saved = s.getResult()
    expect(saved.clock).toEqual(c2)
    expect(compute(apply(book, [{ sheet: 'S', cell: 'A3', input: { value: 2 } }]), { clock: c2 }).hash).toBe(saved.hash)
    expect(() => s.setClock(c1)).toThrow(InputError)
  })

  it('explain: the precedents tree with depth, node limit and text', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { formula: 'SUM(A1:A2)' }, A4: { formula: 'A3*2+A1' } } }] }
    const s = createSession(book, { clock })
    const e = s.explain('S', 'A4')
    expect(e.root).toMatchObject({ ref: "'S'!A4", value: 7, formula: "A3*2+A1" })
    expect(e.root.children.map((c) => c.ref)).toEqual(["'S'!A3", "'S'!A1"])
    expect(e.root.children[0].children[0]).toMatchObject({ ref: "'S'!A1:A2", cells: 2, children: [] })
    expect(e.text).toBe(["S!A4 = 7  ← A3*2+A1", "  S!A3 = 3  ← SUM(A1:A2)", "    S!A1:A2 (2 cells)", "  S!A1 = 1"].join('\n'))
    const x = s.explain('S', 'A4', { expand: true })
    expect(x.root.children[0].children[0].children.map((c) => c.ref)).toEqual(["'S'!A1", "'S'!A2"])
    const t = s.explain('S', 'A4', { depth: 1, maxNodes: 2 })
    expect(t.truncated).toBe(true)
    expect(t.nodes).toBeLessThanOrEqual(3)
  })

  it('debug dump of the code in evaluation order', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A2: { formula: 'A1*2' }, A1: { formula: '1+1' } } }] }
    const code = debugGeneratedCode(book, { clock })
    expect(code.indexOf("// 'S'!A1 = 1+1")).toBeLessThan(code.indexOf("// 'S'!A2 = A1*2"))
    expect(code).toContain('return ()=>addLast(1,1)')
  })
})
