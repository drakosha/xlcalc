/**
 * Seeded random sessions against full computations: after each edit or structure operation the session's hash is the
 * hash of `compute(session.book())` with the same clock. The books and edits (`session-fuzz.ts`) aim at the range
 * caches: a data sheet of one column read through rectangles wider than it and whole columns, written past its edge,
 * then read again through its cached ranges; spills that appear and are blocked; rows inserted and deleted. Once under
 * the default budget, where cached ranges live through the edits, and once under one that drops them at each step.
 */
import { describe, expect, it } from 'vitest'
import { compute, openSession, type Execution } from '../src/index.js'
import * as rangeCache from '../src/range-cache.js'
import { differences, fuzz, patches } from './session-fuzz.js'
import { extentOf } from './sheet-extent.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const STEPS = 60
// absent before the cache budget (0.7.0): the calibration on older code runs the rest of the file
const heapCaches = (rangeCache as Partial<typeof rangeCache>).heapCaches
/** A few ranges of 300 values: each step that reads ranges drops some (`docs/cache.md`). */
const SMALL = 8 * 1024

/** Runs `f` with this heap's budget cut to `SMALL`. */
function tight<T>(f: () => T): T {
  const cap = heapCaches!.cap
  heapCaches!.cap = SMALL
  try {
    return f()
  } finally {
    heapCaches!.cap = cap
  }
}

/**
 * The steps of one seed, each followed by patches of cells' looks, then the session's hash against a full computation's
 * and each sheet's extent against the full computation's cells; `around` wraps the session's calls.
 */
function run(seed: number, execution: Execution, around: <T>(f: () => T) => T = (f) => f()) {
  const opts = { clock, execution, display: false }
  const { book, step } = fuzz(seed)
  const patch = patches(seed)
  const s = around(() => openSession(book, opts))
  // absent before patches and extents (0.12.0): the calibration on older code runs the steps alone
  const looks = 'patchCells' in s
  try {
    for (let k = 0; k <= STEPS; k++) {
      const st = k ? step() : undefined
      if (st && 'op' in st) {
        if (!around(() => s.restructure(st.op)).ok) continue
      } else if (st) around(() => s.setCells(st.edits))
      const p = looks && k ? patch() : []
      if (p.length) expect(around(() => s.patchCells(p))).toEqual([])
      const hash = around(() => s.hash())
      const full = compute(s.book(), opts)
      const at = `seed ${seed}, step ${k}: ${JSON.stringify(st ?? 'open')}, patches ${JSON.stringify(p)}`
      if (looks) for (const sheet of full.sheets) expect(around(() => s.extent(sheet.name)), `${at}: ${sheet.name}`).toEqual(extentOf(full, s.book(), sheet.name))
      if (hash === full.hash) continue
      expect(around(() => differences(s.result(), full)), at).toEqual([])
      expect(hash, at).toBe(full.hash)
    }
  } finally {
    s.close()
  }
}

describe.each(['closures', 'eval'] as const)('seeded sessions against full computations, %s', (execution) => {
  it.each([1, 2, 3, 4, 5, 6])('seed %i, default budget', (seed) => run(seed, execution))

  describe.skipIf(!heapCaches)('a budget that drops ranges at each step', () => {
    it.each([7, 8, 9])('seed %i', (seed) => run(seed, execution, tight))
  })
})
