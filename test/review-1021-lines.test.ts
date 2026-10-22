/**
 * A sheet's rows and columns, from the review of 21.10 (speed): a map of many entries is read once per call, not once
 * per cell, edit or line. Each case took seconds before, on a size where it now takes a few milliseconds or tens of them.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { editStructure, type BookInput, type Execution, type RowInput } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']

afterAll(closeCalcs)

/** Milliseconds a call takes. */
async function took(f: () => unknown): Promise<number> {
  const t0 = performance.now()
  await f()
  return performance.now() - t0
}

/** Milliseconds a synchronous call takes. */
function tookSync(f: () => unknown): number {
  const t0 = performance.now()
  f()
  return performance.now() - t0
}

/** Rows 1 to `n`, each its own entry: heights that differ from the next row's, as an xlsx with fitted rows keeps them. */
const manyRows = (n: number, entry: (r: number) => RowInput = () => ({})): Record<string, RowInput> =>
  Object.fromEntries(Array.from({ length: n }, (_, i) => [String(i + 1), { height: 15 + (i % 3), ...entry(i + 1) }]))

describe.each(EXECUTIONS)('review 21.10 rows and columns, %s', (execution) => {
  it.each([
    ['8,000 rows of their own entries in one group', manyRows(8000, () => ({ outlineLevel: 1 }))],
    ['3,000 groups of three rows', Object.fromEntries(Array.from({ length: 3000 }, (_, i) => [`${4 * i + 1}:${4 * i + 3}`, { outlineLevel: 1 }]))],
  ])('outline commands over %s patch the rows once (speed-u92)', (_, rows) => {
    const book: BookInput = { sheets: [{ name: 'S', rows, cells: {} }] }
    let hidden: BookInput | undefined
    // a patch of the whole map per entry and per mark: 7 to 12 s for each
    expect(
      tookSync(() => {
        const r = editStructure(book, { op: 'showLevels', sheet: 'S', rows: 1 })
        if (!r.ok) throw new Error(r.code)
        hidden = r.book
        const shown = editStructure(hidden, { op: 'showLevels', sheet: 'S', rows: 2 })
        if (!shown.ok) throw new Error(shown.code)
      }),
    ).toBeLessThan(1000)
    expect(Object.values(hidden!.sheets[0].rows!).every((e) => e.outlineLevel !== 1 || e.hidden === true)).toBe(true)
  })
})
