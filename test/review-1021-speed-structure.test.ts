/**
 * Structure operations, fills and pastes at size, from the review of 21.10 (speed): the work of a fill or a paste grows
 * with what it writes, not with the square of it or with the sheet times the tiles. Each test takes a size at which the
 * old code took seconds and the new takes a small part of the bound.
 */
import { describe, expect, it } from 'vitest'
import { createSession, type BookInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** Milliseconds `f` takes. */
function timed(f: () => void): number {
  const t = performance.now()
  f()
  return performance.now() - t
}

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  it('a fill of a validated cell joins the pasted rules in linear time (speed-u123)', () => {
    const n = 20_000
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } }, validations: [{ range: 'A1', type: 'whole', operator: 'greaterThan', formula1: 'B1' }] }] }
    const s = createSession(book, { execution, clock, display: false })
    // the old join re-parsed the growing range per pasted rule: 16-18 s at 20,000 rows
    const ms = timed(() => expect(s.editStructure({ op: 'fillRange', sheet: 'S', range: 'A1', to: `A1:A${n}`, type: 'copy' }).ok).toBe(true))
    expect(ms).toBeLessThan(3000)
    expect(s.getBook().sheets[0].validations).toEqual([{ range: `A1:A${n}`, type: 'whole', operator: 'greaterThan', formula1: 'B1' }])
  })

  it('a fill over rows with their own entries looks up only the lines of each tile (speed-u91)', () => {
    const n = 12_000
    const [cells, rows]: [BookInput['sheets'][0]['cells'], Record<string, { height: number }>] = [{ B1: { formula: 'A1*2' } }, {}]
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: r }
      rows[String(r)] = { height: 15 + (r % 3) }
    }
    const s = createSession({ sheets: [{ name: 'S', cells, rows }] }, { execution, clock, display: false })
    // the old fill walked every row entry per tile: 9-10 s at 12,000 rows
    const ms = timed(() => expect(s.editStructure({ op: 'fillRange', sheet: 'S', range: 'B1', to: `B1:B${n}`, type: 'copy' }).ok).toBe(true))
    expect(ms).toBeLessThan(3000)
    expect(s.getCell('S', `B${n}`)).toMatchObject({ formula: `A${n}*2`, value: n * 2 })
  })
})
