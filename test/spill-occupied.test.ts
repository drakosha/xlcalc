/**
 * The cells an array cannot spill over, kept sorted by column: many cells written into one column in one call stay
 * linear (each cell sorted the column again, so 40,000 cells took 8 s), and a spill sees every one of them, written
 * top to bottom, bottom to top or between.
 */
import { describe, expect, it } from 'vitest'
import { createSession } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const column = (rows: number[]) => rows.map((r) => ({ sheet: 'S', cell: `A${r}`, input: { value: r } }))

describe('cells a spill cannot pass, written into one column', () => {
  it('40,000 cells into one column in one call within 2 s, top to bottom and bottom to top', () => {
    for (const reverse of [false, true]) {
      const s = createSession({ sheets: [{ name: 'S', cells: {} }] }, { clock, display: false })
      const rows = Array.from({ length: 40_000 }, (_, i) => i + 1)
      const t = performance.now()
      s.setCells(column(reverse ? rows.reverse() : rows))
      expect(performance.now() - t).toBeLessThan(2000)
    }
  })

  it('a spill stops at a cell written after it, in any order, and goes on when it is cleared', () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(100)' } } }] }, { clock })
    s.setCells(column([90, 50, 70, 60]))
    expect(s.getCell('S', 'A1')?.value).toEqual({ error: 'SPILL' })
    s.setCells([50, 60, 70].map((r) => ({ sheet: 'S', cell: `A${r}`, input: null })))
    expect(s.getCell('S', 'A1')?.value).toEqual({ error: 'SPILL' })
    s.setCells([{ sheet: 'S', cell: 'A90', input: null }])
    expect(s.getCell('S', 'A1')?.value).toBe(1)
    expect(s.getCell('S', 'A90')?.value).toBe(90)
  })
})
