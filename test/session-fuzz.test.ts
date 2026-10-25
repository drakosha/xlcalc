/**
 * Seeded random sessions against full computations: after each edit or structure operation the session's hash is the
 * hash of `compute(session.getBook())` with the same clock. The books and edits (`session-fuzz.ts`) aim at the range
 * caches: a data sheet of one column read through rectangles wider than it and whole columns, written past its edge,
 * then read again through its cached ranges; spills that appear and are blocked; rows inserted and deleted; cells
 * replaced, cleared, formatted over ranges and typed into. Undoing every step gives the book opened. Once under the
 * default budget, where cached ranges live through the edits, and once under one that drops them at each step.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, createSession, type Execution, type Session } from '../src/index.js'
import { heapCaches } from '../src/range-cache.js'
import { differences, fuzz, lookEdits, patches, type LookEdit } from './session-fuzz.js'
import { extentOf } from './sheet-extent.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const STEPS = 60
/** A few ranges of 300 values: each step that reads ranges drops some (`docs/cache.md`). */
const SMALL = 8 * 1024

/** Runs `f` with this heap's budget cut to `SMALL`. */
function tight<T>(f: () => T): T {
  const cap = heapCaches.cap
  heapCaches.cap = SMALL
  try {
    return f()
  } finally {
    heapCaches.cap = cap
  }
}

/** A call of a look edit; an edit the session refuses (part of an array, a formula in a table's header) changes nothing. */
function lookEdit(s: Session, e: LookEdit) {
  try {
    if ('replace' in e) s.replaceCells(e.replace)
    else if ('clear' in e) s.clearCells(e.clear, e.what)
    else if ('patch' in e) s.patchCells(e.patch)
    else s.setCellText(...e.text)
  } catch (x) {
    if (!(x instanceof InputError)) throw x
  }
}

/**
 * The steps of one seed, each followed by patches of cells' looks and an edit of the calls that keep or set a look, then
 * the session's hash against a full computation's and each sheet's extent against the full computation's cells; at the
 * end every step undone, which gives the book opened. `around` wraps the session's calls.
 */
function run(seed: number, execution: Execution, around: <T>(f: () => T) => T = (f) => f()) {
  const opts = { clock, execution, display: false }
  const { book, step } = fuzz(seed)
  const patch = patches(seed)
  const look = lookEdits(seed)
  const s = around(() => createSession(book, { ...opts, undo: { depth: 10_000, maxMb: 1024 } }))
  try {
    for (let k = 0; k <= STEPS; k++) {
      const st = k ? step() : undefined
      if (st && 'op' in st) {
        if (!around(() => s.editStructure(st.op)).ok) continue
      } else if (st) around(() => s.setCells(st.edits))
      const p = k ? patch() : []
      // a patch lists the cells whose own number format, note or link it changed, and no value
      if (p.length) expect(around(() => s.patchCells(p).changes.filter((c) => !('numberFormat' in c || 'note' in c || 'link' in c)))).toEqual([])
      const e = k ? look() : undefined
      if (e) around(() => lookEdit(s, e))
      const hash = around(() => s.getHash())
      const full = compute(s.getBook(), opts)
      const at = `seed ${seed}, step ${k}: ${JSON.stringify(st ?? 'open')}, patches ${JSON.stringify(p)}, then ${JSON.stringify(e)}`
      for (const sheet of full.sheets) expect(around(() => s.getUsedRange(sheet.name)), `${at}: ${sheet.name}`).toEqual(extentOf(full, s.getBook(), sheet.name))
      if (hash === full.hash) continue
      expect(around(() => differences(s.getResult(), full)), at).toEqual([])
      expect(hash, at).toBe(full.hash)
    }
    // every step undone: the book opened, and its hash
    while (around(() => s.undo()));
    expect(s.getBook(), `seed ${seed}: every step undone`).toEqual(fuzz(seed).book)
    expect(around(() => s.getHash())).toBe(compute(fuzz(seed).book, opts).hash)
  } finally {
    s.close()
  }
}

describe.each(['closures', 'eval'] as const)('seeded sessions against full computations, %s', (execution) => {
  it.each([1, 2, 3, 4, 5, 6])('seed %i, default budget', (seed) => run(seed, execution))

  describe('a budget that drops ranges at each step', () => {
    it.each([7, 8, 9])('seed %i', (seed) => run(seed, execution, tight))
  })
})
