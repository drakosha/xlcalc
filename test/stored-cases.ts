/** Checks shared by the tests of stored sets of Excel's answers. */
import { expect, it } from 'vitest'
import type { StructCases } from '../bench/coverage/struct/cases.js'

/** A test that a structure set's stored cases, books and files are those its module builds now, each case answered. */
export function currentStructSet(set: StructCases & { answers: Record<string, unknown> }, built: StructCases): void {
  it('the stored answers are for the current cases and books', () => {
    expect(set.cases).toEqual(built.cases)
    expect(set.books).toEqual(built.books)
    expect(set.files).toEqual(built.files)
    expect(built.cases.filter((c) => !set.answers[c.id])).toEqual([])
  })
}

/** A test that the set's stored cases are those its module builds now (`built`), and that each has an answer. */
export function currentCases<C extends { id: string }>(set: { cases: C[]; answers: Record<string, unknown> }, built: C[]): void {
  it('the stored cases are the current ones, all answered', () => {
    expect(set.cases).toEqual(built)
    expect(built.filter((c) => !set.answers[c.id])).toEqual([])
  })
}

/**
 * The keys of the pinned differences of `known` that name no case of `sets` (`<set>:<id>`): stale entries to remove.
 * `byCase` takes a key's first two segments, for keys that name a case's cell (`<set>:<id>:<cell>`).
 */
export function pinnedCases<S extends { name: string; cases: { id: string }[] }>(sets: readonly S[], known: Readonly<Record<string, unknown>>, byCase = false): string[] {
  const ids = new Set(sets.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
  return Object.keys(known).filter((k) => !ids.has(byCase ? k.split(':').slice(0, 2).join(':') : k))
}
